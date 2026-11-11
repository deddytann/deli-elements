/* Deli Elements task pane: talks to PowerPoint through Office.js. Pure logic lives in core.js (DeliCore). */
/* global Office, PowerPoint, DeliCore */
"use strict";

const D = DeliCore;
const has = (v) => Office.context.requirements.isSetSupported("PowerPointApi", v);

function status(msg, kind) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status " + (kind || "");
}

async function run(label, fn) {
  try {
    await PowerPoint.run(fn);
    status(label + " ✓", "ok");
  } catch (e) {
    console.error(e);
    status(label + ": " + (e.message || e), "err");
  }
}

function currentSlide(context) {
  return has("1.5") ? context.presentation.getSelectedSlides().getItemAt(0) : context.presentation.slides.getItemAt(0);
}

// ---------- Render primitives as native shapes ----------
function styleText(shape, p) {
  const tf = shape.textFrame;
  tf.textRange.text = p.text;
  const f = p.font || {};
  if (f.name) tf.textRange.font.name = f.name;
  if (f.size) tf.textRange.font.size = f.size;
  if (f.color) tf.textRange.font.color = f.color;
  tf.textRange.font.bold = !!f.bold;
  tf.textRange.font.italic = !!f.italic;
  tf.textRange.paragraphFormat.horizontalAlignment = p.align || "Left";
  tf.verticalAlignment = p.valign || "Top";
  tf.wordWrap = true;
  const m = p.kind === "text" ? 0 : 8;
  tf.leftMargin = m; tf.rightMargin = m; tf.topMargin = p.kind === "text" ? 0 : 4; tf.bottomMargin = p.kind === "text" ? 0 : 4;
}

async function insertElement(key, opts) {
  const prims = D.buildElement(key, opts);
  await run("Inserted " + D.ELEMENTS[key].label, async (context) => {
    const shapes = currentSlide(context).shapes;
    const made = [];
    prims.forEach((p) => {
      let s;
      const box = { left: p.x, top: p.y, width: Math.max(p.w, 0.01), height: Math.max(p.h, 0.01) };
      if (p.kind === "line") {
        s = shapes.addLine("Straight", { left: p.x, top: p.y, width: p.w, height: p.h });
        s.lineFormat.color = p.line; s.lineFormat.weight = p.lineW || 1;
      } else if (p.kind === "text") {
        s = shapes.addTextBox(p.text || "", box);
        s.fill.clear(); s.lineFormat.visible = false;
        styleText(s, p);
      } else {
        s = shapes.addGeometricShape(p.geom, box);
        if (p.fill) s.fill.setSolidColor(p.fill); else s.fill.clear();
        if (p.line) { s.lineFormat.visible = true; s.lineFormat.color = p.line; s.lineFormat.weight = p.lineW || 1; } else s.lineFormat.visible = false;
        if (p.text) styleText(s, p);
        if (p.rotation && has("1.10")) s.rotation = p.rotation;
      }
      s.name = "Deli " + key + " " + p.name;
      made.push(s);
    });
    await context.sync();
    if (made.length > 1 && has("1.8")) {
      made.forEach((s) => s.load("id"));
      await context.sync();
      const g = shapes.addGroup(made.map((s) => s.id));
      g.name = "Deli " + D.ELEMENTS[key].label;
      await context.sync();
    }
  });
}

// ---------- Selection-based layout tools ----------
async function withSelection(label, min, fn) {
  if (!has("1.5")) return status("This PowerPoint version can't read the selection (needs PowerPointApi 1.5).", "err");
  await run(label, async (context) => {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items/id,items/name,items/left,items/top,items/width,items/height");
    await context.sync();
    if (sel.items.length < min) throw new Error(`select at least ${min} shape${min > 1 ? "s" : ""}`);
    fn(sel.items);
    await context.sync();
  });
}

const geom = (items) => items.map((s) => ({ left: s.left, top: s.top, width: s.width, height: s.height }));
const applyPos = (items, pos) => items.forEach((s, i) => { s.left = pos[i].left; s.top = pos[i].top; });

const LAYOUT = {
  "align-left": () => withSelection("Aligned left", 1, (it) => applyPos(it, D.align(geom(it), "left"))),
  "align-center": () => withSelection("Centered", 1, (it) => applyPos(it, D.align(geom(it), "center"))),
  "align-right": () => withSelection("Aligned right", 1, (it) => applyPos(it, D.align(geom(it), "right"))),
  "align-top": () => withSelection("Aligned top", 1, (it) => applyPos(it, D.align(geom(it), "top"))),
  "align-middle": () => withSelection("Aligned middle", 1, (it) => applyPos(it, D.align(geom(it), "middle"))),
  "align-bottom": () => withSelection("Aligned bottom", 1, (it) => applyPos(it, D.align(geom(it), "bottom"))),
  "dist-h": () => withSelection("Distributed horizontally", 3, (it) => applyPos(it, D.distribute(geom(it), "h"))),
  "dist-v": () => withSelection("Distributed vertically", 3, (it) => applyPos(it, D.distribute(geom(it), "v"))),
  "stack-h": () => withSelection("Stacked with 0.2\" gaps", 2, (it) => applyPos(it, D.stack(geom(it), "h", 0.2 * D.IN))),
  "stack-v": () => withSelection("Stacked with 0.2\" gaps", 2, (it) => applyPos(it, D.stack(geom(it), "v", 0.2 * D.IN))),
  "same-w": () => withSelection("Same width", 2, (it) => D.matchSize(geom(it), "width").forEach((z, i) => { it[i].width = z.width; })),
  "same-h": () => withSelection("Same height", 2, (it) => D.matchSize(geom(it), "height").forEach((z, i) => { it[i].height = z.height; })),
  "same-size": () => withSelection("Same size", 2, (it) => D.matchSize(geom(it), "size").forEach((z, i) => { it[i].width = z.width; it[i].height = z.height; })),
  "swap": () => withSelection("Swapped", 2, (it) => { const p = D.swap(geom(it)[0], geom(it)[1]); applyPos(it.slice(0, 2), p); }),
  "snap": () => withSelection("Snapped to 0.05\" grid", 1, (it) => it.forEach((s) => { s.left = D.snap(s.left, 3.6); s.top = D.snap(s.top, 3.6); s.width = D.snap(s.width, 3.6); s.height = D.snap(s.height, 3.6); })),
  "margins": () => withSelection("Fitted to content width", 1, (it) => { const b = D.bounds(geom(it)), k = (D.SLIDE.w - 2 * D.SLIDE.margin) / b.width; it.forEach((s) => { s.left = D.SLIDE.margin + (s.left - b.left) * k; s.width = s.width * k; }); }),
};

// ---------- Numbers in selected text ----------
async function formatSelected(style) {
  if (!has("1.5")) return status("Needs PowerPointApi 1.5.", "err");
  await run("Formatted numbers", async (context) => {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items/id");
    await context.sync();
    const ranges = sel.items.map((s) => { const r = s.textFrame.textRange; r.load("text"); return r; });
    await context.sync();
    ranges.forEach((r) => {
      const out = r.text.replace(/(−|-)?(Rp)?\d[\d.,]*(bn|m|k|%)?/gi, (m) => D.formatNumber(m, style));
      if (out !== r.text) r.text = out;
    });
  });
}

// ---------- Tables ----------
async function insertTable() {
  if (!has("1.8")) return status("Branded tables need PowerPoint with PowerPointApi 1.8 (Microsoft 365, current version).", "err");
  const rows = Math.max(2, Math.min(15, Number(document.getElementById("tbl-rows").value) || 5));
  const cols = Math.max(2, Math.min(8, Number(document.getElementById("tbl-cols").value) || 4));
  const spec = D.tableSpec(rows, cols);
  await run(`Inserted ${rows}×${cols} table`, async (context) => {
    const shapes = currentSlide(context).shapes;
    const t = shapes.addTable(rows, cols, Object.assign({ left: D.SLIDE.margin, top: 150, width: D.SLIDE.w - 2 * D.SLIDE.margin }, spec));
    t.name = "Deli table";
    await context.sync();
  });
}

// ---------- Library slides ----------
let libraryB64 = null;
async function libraryBase64() {
  if (libraryB64) return libraryB64;
  const buf = await (await fetch("assets/Deli_Elements_Library.pptx")).arrayBuffer();
  let bin = "";
  const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  libraryB64 = btoa(bin);
  return libraryB64;
}
async function insertLibrarySlide(slideId) {
  const b64 = await libraryBase64();
  await run("Inserted library slide", async (context) => {
    const opts = { formatting: "UseDestinationTheme", sourceSlideIds: [slideId + "#"] };
    if (has("1.5")) {
      const sel = context.presentation.getSelectedSlides();
      sel.load("items/id");
      await context.sync();
      if (sel.items.length) opts.targetSlideId = sel.items[0].id;
    }
    context.presentation.insertSlidesFromBase64(b64, opts);
    await context.sync();
  });
}

// ---------- Brand check ----------
let lastIssues = [];
async function brandCheck() {
  await run("Brand check finished", async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const all = [];
    const per = slides.items.map((sl) => { const sh = sl.shapes; sh.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height"); return sh; });
    await context.sync();
    const texty = [];
    per.forEach((sh, si) => sh.items.forEach((s) => {
      const rec = { slideIndex: si, slideId: slides.items[si].id, shapeId: s.id, name: s.name, left: s.left, top: s.top, width: s.width, height: s.height };
      all.push(rec);
      if (["GeometricShape", "TextBox", "Placeholder"].includes(s.type)) {
        const r = s.textFrame.textRange; r.load("text,font/name,font/size,font/color");
        texty.push([rec, r]);
      }
    }));
    await context.sync();
    texty.forEach(([rec, r]) => { rec.text = r.text; rec.font = r.font.name; rec.size = r.font.size; rec.color = r.font.color; });
    lastIssues = D.checkDeck(all).map((iss) => Object.assign(iss, { rec: all.find((a) => a.slideIndex === iss.slideIndex && a.name === iss.shapeName) }));
    renderIssues();
  });
}

function renderIssues() {
  const box = document.getElementById("issues");
  box.innerHTML = "";
  if (!lastIssues.length) { box.innerHTML = '<p class="good">No issues found. The deck is on brand.</p>'; return; }
  const fixable = lastIssues.filter((i) => i.fixable).length;
  const head = document.createElement("p");
  head.textContent = `${lastIssues.length} issue${lastIssues.length > 1 ? "s" : ""} found` + (fixable ? `, ${fixable} fixable automatically.` : ".");
  box.appendChild(head);
  const ul = document.createElement("ul");
  lastIssues.forEach((i) => {
    const li = document.createElement("li");
    li.textContent = `Slide ${i.slideIndex + 1} · ${i.shapeName || "shape"}: ${i.msg}`;
    ul.appendChild(li);
  });
  box.appendChild(ul);
  document.getElementById("fix").disabled = !fixable;
}

async function fixIssues() {
  await run("Fixed fonts and colors", async (context) => {
    for (const i of lastIssues.filter((x) => x.fixable && x.rec)) {
      const shape = context.presentation.slides.getItem(i.rec.slideId).shapes.getItem(i.rec.shapeId);
      const font = shape.textFrame.textRange.font;
      if (i.type === "font") font.name = (i.rec.size || 0) >= 24 ? "Cambria" : "Calibri";
      if (i.type === "color") font.color = D.nearestBrand(i.rec.color);
    }
    await context.sync();
  });
  await brandCheck();
}

// ---------- UI wiring ----------
function buildElementButtons() {
  const host = document.getElementById("elements");
  const groups = {};
  Object.entries(D.ELEMENTS).forEach(([k, e]) => { (groups[e.group] = groups[e.group] || []).push([k, e]); });
  Object.entries(groups).forEach(([g, list]) => {
    const h = document.createElement("h3"); h.textContent = g; host.appendChild(h);
    const wrap = document.createElement("div"); wrap.className = "grid";
    list.forEach(([k, e]) => {
      const b = document.createElement("button"); b.textContent = e.label; b.title = "Insert " + e.label;
      b.onclick = () => {
        const o = {};
        if (e.params && e.params.includes("text")) o.text = document.getElementById("opt-text").value || "DRAFT";
        if (e.params && e.params.includes("n")) o.n = Number(document.getElementById("opt-n").value) || 1;
        if (e.params && e.params.includes("steps")) o.steps = Number(document.getElementById("opt-steps").value) || 4;
        if (e.params && e.params.includes("current")) o.current = Number(document.getElementById("opt-current").value) || 1;
        if (e.params && e.params.includes("rows")) o.rows = Number(document.getElementById("opt-rows").value) || 4;
        insertElement(k, o);
      };
      wrap.appendChild(b);
    });
    host.appendChild(wrap);
  });
}

Office.onReady((info) => {
  buildElementButtons();
  if (info.host !== Office.HostType.PowerPoint) { status("Preview only: open the Deli tab inside PowerPoint to use these tools.", "err"); return; }
  document.querySelectorAll("[data-layout]").forEach((b) => (b.onclick = () => LAYOUT[b.dataset.layout]()));
  document.querySelectorAll("[data-num]").forEach((b) => (b.onclick = () => formatSelected(b.dataset.num)));
  document.querySelectorAll("[data-lib]").forEach((b) => (b.onclick = () => insertLibrarySlide(b.dataset.lib)));
  document.getElementById("tbl-insert").onclick = insertTable;
  document.getElementById("check").onclick = brandCheck;
  document.getElementById("fix").onclick = fixIssues;
  document.querySelectorAll(".tab").forEach((t) => (t.onclick = () => {
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("on", x === t));
    document.querySelectorAll(".panel").forEach((p) => p.classList.toggle("on", p.id === t.dataset.panel));
  }));
  const v = ["1.4", "1.5", "1.8", "1.10"].filter(has).pop() || "older";
  status(`Ready · PowerPoint API ${v}`, "ok");
});
