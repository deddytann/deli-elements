/* Deli Elements task pane. Pure logic: core.js (DeliCore) and elements.js (DeliElements). */
/* global Office, PowerPoint, DeliCore, DeliElements */
"use strict";

const D = DeliCore, EL = DeliElements;
const has = (v) => typeof Office !== "undefined" && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported("PowerPointApi", v);
const $ = (id) => document.getElementById(id);
const inPowerPoint = { ok: false };

function status(msg, kind) { const el = $("status"); el.textContent = msg; el.className = "status " + (kind || ""); }
async function run(label, fn) {
  if (!inPowerPoint.ok) return status("Preview only: open the Deli tab inside PowerPoint to use this.", "err");
  try { await PowerPoint.run(fn); status(label + " ✓", "ok"); } catch (e) { console.error(e); status(label + ": " + (e.message || e), "err"); }
}
const currentSlide = (ctx) => (has("1.5") ? ctx.presentation.getSelectedSlides().getItemAt(0) : ctx.presentation.slides.getItemAt(0));

// ================= Render primitives as native shapes =================
function styleText(shape, p) {
  const tf = shape.textFrame, f = p.font || {};
  tf.textRange.text = p.text;
  if (f.name) tf.textRange.font.name = f.name;
  if (f.size) tf.textRange.font.size = f.size;
  if (f.color) tf.textRange.font.color = f.color;
  tf.textRange.font.bold = !!f.bold;
  tf.textRange.font.italic = !!f.italic;
  tf.textRange.paragraphFormat.horizontalAlignment = p.align || "Left";
  tf.verticalAlignment = p.valign || "Top";
  tf.wordWrap = true;
  const m = p.kind === "text" ? 0 : 6;
  tf.leftMargin = m; tf.rightMargin = m; tf.topMargin = p.kind === "text" ? 0 : 3; tf.bottomMargin = p.kind === "text" ? 0 : 3;
}

async function drawPrims(shapes, prims, groupName, ctx) {
  const made = [];
  prims.forEach((p) => {
    let s;
    const box = { left: p.x, top: p.y, width: Math.max(p.w, 0.01), height: Math.max(p.h, 0.01) };
    if (p.kind === "line") {
      s = shapes.addLine("Straight", { left: p.x, top: p.y, width: p.w, height: p.h });
      s.lineFormat.color = p.line; s.lineFormat.weight = p.lineW || 1;
      if (p.dash) s.lineFormat.dashStyle = p.dash;
    } else if (p.kind === "text") {
      s = shapes.addTextBox(p.text || "", box);
      s.fill.clear(); s.lineFormat.visible = false;
      styleText(s, p);
    } else {
      s = shapes.addGeometricShape(p.geom, box);
      if (p.fill) { s.fill.setSolidColor(p.fill); if (p.alpha) s.fill.transparency = 1 - p.alpha; } else s.fill.clear();
      if (p.line) { s.lineFormat.visible = true; s.lineFormat.color = p.line; s.lineFormat.weight = p.lineW || 1; if (p.dash) s.lineFormat.dashStyle = p.dash; } else s.lineFormat.visible = false;
      if (p.text) styleText(s, p);
      if (p.rotation && has("1.10")) s.rotation = p.rotation;
    }
    s.name = "Deli " + p.name;
    made.push(s);
  });
  await ctx.sync();
  if (made.length > 1 && has("1.8")) {
    made.forEach((s) => s.load("id"));
    await ctx.sync();
    const g = shapes.addGroup(made.map((s) => s.id));
    g.name = "Deli " + groupName;
    await ctx.sync();
  }
}

async function insertElement(key, opts) {
  const el = EL.ELEMENTS[key];
  const prims = EL.build(key, opts);
  await run("Inserted " + el.label, async (ctx) => drawPrims(currentSlide(ctx).shapes, prims, el.label, ctx));
}

// ================= Elements gallery =================
const elOpts = {}; // remembers last-used parameters per element
function buildGallery(filter) {
  const host = $("el-gallery");
  host.innerHTML = "";
  const q = (filter || "").trim().toLowerCase();
  const groups = {};
  Object.entries(EL.ELEMENTS).forEach(([k, e]) => {
    if (q && !(e.label + " " + e.tags + " " + e.group).toLowerCase().includes(q)) return;
    (groups[e.group] = groups[e.group] || []).push([k, e]);
  });
  if (!Object.keys(groups).length) { host.innerHTML = '<p class="hint">No element matches that search.</p>'; return; }
  Object.entries(groups).forEach(([g, list]) => {
    const h = document.createElement("h3"); h.textContent = g; host.appendChild(h);
    const wrap = document.createElement("div"); wrap.className = "tiles";
    list.forEach(([k, e]) => {
      const b = document.createElement("button");
      b.className = "tile" + (e.params.length ? " has-opts" : "");
      b.title = e.label + (e.params.length ? " (click for options)" : "");
      b.setAttribute("aria-label", e.label);
      b.dataset.key = k;
      b.innerHTML = EL.thumbnail(k, elOpts[k] || {}, 64, 44);
      b.onclick = () => (e.params.length ? openDrawer(wrap, b, k) : insertElement(k, {}));
      wrap.appendChild(b);
    });
    host.appendChild(wrap);
  });
}

function openDrawer(wrap, tile, key) {
  const e = EL.ELEMENTS[key];
  wrap.querySelectorAll(".drawer").forEach((d) => d.remove());
  document.querySelectorAll(".tile.sel").forEach((t) => t.classList.remove("sel"));
  if (tile.dataset.open === "1") { tile.dataset.open = ""; return; }
  document.querySelectorAll(".tile").forEach((t) => (t.dataset.open = ""));
  tile.dataset.open = "1"; tile.classList.add("sel");
  const cur = EL.options(key, elOpts[key]);
  const d = document.createElement("div"); d.className = "drawer";
  d.innerHTML = `<strong>${e.label}</strong>`;
  e.params.forEach((p) => {
    const l = document.createElement("label"); l.textContent = p.label;
    let inp;
    if (p.type === "select") { inp = document.createElement("select"); p.options.forEach((o) => { const op = document.createElement("option"); op.value = o; op.textContent = o; inp.appendChild(op); }); }
    else { inp = document.createElement("input"); inp.type = p.type === "number" ? "number" : "text"; if (p.min != null) inp.min = p.min; if (p.max != null) inp.max = p.max; }
    inp.value = cur[p.key]; inp.dataset.key = p.key;
    inp.oninput = () => { elOpts[key] = readDrawer(d); tile.innerHTML = EL.thumbnail(key, elOpts[key], 64, 44); };
    l.appendChild(inp); d.appendChild(l);
  });
  const go = document.createElement("button"); go.className = "primary"; go.textContent = "Insert";
  go.onclick = () => { elOpts[key] = readDrawer(d); insertElement(key, elOpts[key]); };
  d.appendChild(go);
  // place the drawer on the row after the clicked tile
  const tiles = [...wrap.children].filter((c) => c.classList.contains("tile"));
  const perRow = Math.max(1, Math.round(wrap.clientWidth / (tile.offsetWidth + 6)) || 4);
  const idx = tiles.indexOf(tile), endOfRow = tiles[Math.min(tiles.length - 1, Math.floor(idx / perRow) * perRow + perRow - 1)];
  endOfRow.after(d);
}
function readDrawer(d) { const o = {}; d.querySelectorAll("[data-key]").forEach((i) => (o[i.dataset.key] = i.value)); return o; }

// ================= Images: logo + icons (inserted as pictures/SVG) =================
function insertImage(label, data, coercion, size) {
  if (!inPowerPoint.ok) return status("Preview only: open the Deli tab inside PowerPoint to use this.", "err");
  const opts = { coercionType: coercion, imageWidth: size.w };
  if (size.h) opts.imageHeight = size.h;
  opts.imageLeft = Math.round((D.SLIDE.w - size.w) / 2); opts.imageTop = 200;
  Office.context.document.setSelectedDataAsync(data, opts, (r) => {
    if (r.status === Office.AsyncResultStatus.Succeeded) status("Inserted " + label + " ✓", "ok");
    else status(label + ": " + (r.error && r.error.message), "err");
  });
}
async function fileBase64(path) {
  const buf = await (await fetch(path)).arrayBuffer();
  let bin = ""; const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
async function insertLogo(kind) {
  const b64 = await fileBase64(`assets/logo-${kind}.png`);
  insertImage("Deli logo", b64, Office.CoercionType.Image, { w: 110, h: 80 });
}

const ICON_COLORS = [["#17362E", "Deli green"], ["#C9A46A", "Gold"], ["#4E7A6C", "Teal"], ["#1F2A27", "Ink"], ["#6B7570", "Muted"], ["#A4463B", "Alert red"], ["#FFFFFF", "White"]];
const icState = { color: "#17362E", style: "plain", data: null, shown: 0, matches: [] };
function iconSVG(inner, color, style, px) {
  const glyph = `<g fill="none" stroke="${color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${inner}</g>`;
  if (style === "plain") return `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 24 24">${glyph}</svg>`;
  const shapeBg = style === "circle" ? `<circle cx="18" cy="18" r="17" fill="${color}"/>` : style === "ring" ? `<circle cx="18" cy="18" r="16.5" fill="#FFFFFF" stroke="${color}" stroke-width="1.5"/>` : `<rect x="1" y="1" width="34" height="34" rx="7" fill="${color}"/>`;
  const fg = style === "ring" ? color : (color.toUpperCase() === "#FFFFFF" ? "#17362E" : (color.toUpperCase() === "#C9A46A" ? "#17362E" : "#FFFFFF"));
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 36 36">${shapeBg}<g transform="translate(8 8) scale(0.8333)" fill="none" stroke="${fg}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${inner}</g></svg>`;
}
async function loadIcons() {
  if (icState.data) return;
  $("ic-count").textContent = "Loading icons…";
  icState.data = await (await fetch("assets/icons.json")).json();
  const sel = $("ic-cat");
  icState.data.cats.forEach((c) => { const o = document.createElement("option"); o.value = c; o.textContent = c; sel.appendChild(o); });
  filterIcons();
}
const FEATURED = "coffee bread cake chef-hat tools-kitchen-2 building-store truck-delivery basket shopping-cart cash coin wallet receipt chart-bar chart-line chart-pie trending-up trending-down target users user user-check school calendar clock map-pin building-warehouse settings bulb star heart check x alert-triangle flag leaf paw".split(" ");
function filterIcons() {
  const q = $("ic-search").value.trim().toLowerCase(), cat = $("ic-cat").value;
  const all = icState.data.icons;
  let m = all.filter((i) => (!cat || i[1] === cat) && (!q || i[0].includes(q) || i[2].toLowerCase().includes(q)));
  if (!q && !cat) { const by = new Map(all.map((i) => [i[0], i])), set = new Set(FEATURED); m = FEATURED.map((n) => by.get(n)).filter(Boolean).concat(all.filter((i) => !set.has(i[0]))); }
  icState.matches = m; icState.shown = 0; $("ic-grid").innerHTML = ""; showMoreIcons();
}
function showMoreIcons() {
  const grid = $("ic-grid"), next = icState.matches.slice(icState.shown, icState.shown + 120);
  next.forEach(([name, , , inner]) => {
    const b = document.createElement("button"); b.title = name.replace(/-/g, " "); b.setAttribute("aria-label", b.title);
    b.innerHTML = iconSVG(inner, icState.color === "#FFFFFF" ? "#17362E" : icState.color, icState.style, 24);
    b.onclick = () => { const px = Number($("ic-size").value); insertImage("icon " + b.title, iconSVG(inner, icState.color, icState.style, px * 4 / 3), Office.CoercionType.XmlSvg, { w: px, h: px }); };
    grid.appendChild(b);
  });
  icState.shown += next.length;
  $("ic-count").textContent = `${icState.matches.length.toLocaleString()} icons${icState.matches.length ? "" : " — try another word"}`;
  $("ic-more").hidden = icState.shown >= icState.matches.length;
}
function setupIcons() {
  ICON_COLORS.forEach(([c, n], i) => { const b = document.createElement("button"); b.style.background = c; b.title = n; if (!i) b.classList.add("on"); b.onclick = () => { icState.color = c; document.querySelectorAll("#ic-colors button").forEach((x) => x.classList.toggle("on", x === b)); filterIcons(); }; $("ic-colors").appendChild(b); });
  document.querySelectorAll("#ic-style button").forEach((b) => (b.onclick = () => { icState.style = b.dataset.style; document.querySelectorAll("#ic-style button").forEach((x) => x.classList.toggle("on", x === b)); filterIcons(); }));
  let t; $("ic-search").oninput = () => { clearTimeout(t); t = setTimeout(filterIcons, 150); };
  $("ic-cat").onchange = filterIcons;
  $("ic-more").onclick = showMoreIcons;
}

// ================= Layout tools =================
async function withSelection(label, min, fn) {
  if (!has("1.5")) return status("This PowerPoint can't read the selection (needs PowerPointApi 1.5).", "err");
  await run(label, async (ctx) => {
    const sel = ctx.presentation.getSelectedShapes();
    sel.load("items/id,items/name,items/left,items/top,items/width,items/height");
    await ctx.sync();
    if (sel.items.length < min) throw new Error(`select at least ${min} shape${min > 1 ? "s" : ""}`);
    await fn(sel.items, ctx, sel);
    await ctx.sync();
  });
}
const geom = (items) => items.map((s) => ({ left: s.left, top: s.top, width: s.width, height: s.height }));
const applyPos = (items, pos) => items.forEach((s, i) => { s.left = pos[i].left; s.top = pos[i].top; });
const gapPt = () => Math.max(0, Number($("gap").value) || 0) * D.IN;
const anchor = () => $("anchor").value;
let clip = null;

const G = (body) => `<svg viewBox="0 0 32 32" width="22" height="22">${body}</svg>`;
const bar = (x, y, w, h, c) => `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="1.5" fill="${c || "#17362E"}"/>`;
const vl = (x) => `<rect x="${x - 0.75}" y="3" width="1.5" height="26" fill="#C9A46A"/>`, hl = (y) => `<rect x="3" y="${y - 0.75}" width="26" height="1.5" fill="#C9A46A"/>`;
const LAYOUT = {
  "lay-align": [
    ["Align left", G(vl(5) + bar(7, 7, 18, 6) + bar(7, 18, 11, 6)), () => withSelection("Aligned left", 1, (it) => applyPos(it, D.align(geom(it), "left", anchor())))],
    ["Align center", G(vl(16) + bar(6, 7, 20, 6) + bar(10, 18, 12, 6)), () => withSelection("Centered", 1, (it) => applyPos(it, D.align(geom(it), "center", anchor())))],
    ["Align right", G(vl(27) + bar(7, 7, 18, 6) + bar(14, 18, 11, 6)), () => withSelection("Aligned right", 1, (it) => applyPos(it, D.align(geom(it), "right", anchor())))],
    ["Align top", G(hl(5) + bar(7, 7, 6, 18) + bar(18, 7, 6, 11)), () => withSelection("Aligned top", 1, (it) => applyPos(it, D.align(geom(it), "top", anchor())))],
    ["Align middle", G(hl(16) + bar(7, 6, 6, 20) + bar(18, 10, 6, 12)), () => withSelection("Aligned middle", 1, (it) => applyPos(it, D.align(geom(it), "middle", anchor())))],
    ["Align bottom", G(hl(27) + bar(7, 7, 6, 18) + bar(18, 14, 6, 11)), () => withSelection("Aligned bottom", 1, (it) => applyPos(it, D.align(geom(it), "bottom", anchor())))],
  ],
  "lay-space": [
    ["Distribute horizontally", G(bar(3, 9, 6, 14) + bar(13, 9, 6, 14) + bar(23, 9, 6, 14)), () => withSelection("Distributed", 3, (it) => applyPos(it, D.distribute(geom(it), "h")))],
    ["Distribute vertically", G(bar(9, 3, 14, 6) + bar(9, 13, 14, 6) + bar(9, 23, 14, 6)), () => withSelection("Distributed", 3, (it) => applyPos(it, D.distribute(geom(it), "v")))],
    ["Side by side with gap", G(bar(3, 9, 8, 14) + bar(14, 9, 8, 14) + `<path d="M24 16h5" stroke="#C9A46A" stroke-width="2"/>`), () => withSelection("Stacked across", 2, (it) => applyPos(it, D.stack(geom(it), "h", gapPt())))],
    ["Stacked with gap", G(bar(9, 3, 14, 8) + bar(9, 14, 14, 8) + `<path d="M16 24v5" stroke="#C9A46A" stroke-width="2"/>`), () => withSelection("Stacked down", 2, (it) => applyPos(it, D.stack(geom(it), "v", gapPt())))],
    ["Arrange in grid", G(bar(4, 4, 10, 10) + bar(18, 4, 10, 10) + bar(4, 18, 10, 10) + bar(18, 18, 10, 10, "#C9A46A")), () => withSelection("Arranged in grid", 2, (it) => applyPos(it, D.grid(geom(it), Math.max(1, Number($("gridcols").value) || 3), gapPt())))],
    ["Snap to grid", G(`<path d="M4 4h24M4 12h24M4 20h24M4 28h24M4 4v24M12 4v24M20 4v24M28 4v24" stroke="#DCE5DF" stroke-width="1"/>` + bar(12, 12, 8, 8)), () => withSelection("Snapped", 1, (it) => it.forEach((s) => { s.left = D.snap(s.left, 3.6); s.top = D.snap(s.top, 3.6); s.width = D.snap(s.width, 3.6); s.height = D.snap(s.height, 3.6); }))],
  ],
  "lay-size": [
    ["Same width (as first selected)", G(bar(4, 6, 24, 7) + bar(4, 19, 24, 7, "#C9A46A")), () => withSelection("Same width", 2, (it) => D.matchSize(geom(it), "width").forEach((z, i) => { it[i].width = z.width; }))],
    ["Same height (as first selected)", G(bar(6, 4, 7, 24) + bar(19, 4, 7, 24, "#C9A46A")), () => withSelection("Same height", 2, (it) => D.matchSize(geom(it), "height").forEach((z, i) => { it[i].height = z.height; }))],
    ["Same size (as first selected)", G(bar(4, 8, 11, 16) + bar(17, 8, 11, 16, "#C9A46A")), () => withSelection("Same size", 2, (it) => D.matchSize(geom(it), "size").forEach((z, i) => { it[i].width = z.width; it[i].height = z.height; }))],
    ["Fit to content width", G(`<path d="M3 16h26M3 10v12M29 10v12" stroke="#C9A46A" stroke-width="1.6"/>` + bar(8, 12, 16, 8)), () => withSelection("Fitted to content width", 1, (it) => { const b = D.bounds(geom(it)), k = (D.SLIDE.w - 2 * D.SLIDE.margin) / b.width; it.forEach((s) => { s.left = D.SLIDE.margin + (s.left - b.left) * k; s.width = s.width * k; }); })],
  ],
  "lay-arrange": [
    ["Swap two shapes", G(bar(3, 10, 9, 12) + bar(20, 10, 9, 12, "#C9A46A") + `<path d="M13 13h6M13 19h6" stroke="#17362E" stroke-width="1.6"/>`), () => withSelection("Swapped", 2, (it) => applyPos(it.slice(0, 2), D.swap(geom(it)[0], geom(it)[1])))],
    ["Copy position & size", G(bar(6, 6, 14, 14) + `<rect x="12" y="12" width="14" height="14" rx="1.5" fill="none" stroke="#C9A46A" stroke-width="1.6"/>`), () => withSelection("Copied position & size", 1, (it) => { clip = geom(it)[0]; })],
    ["Paste position & size", G(`<rect x="6" y="6" width="14" height="14" rx="1.5" fill="none" stroke="#C9A46A" stroke-width="1.6"/>` + bar(12, 12, 14, 14)), () => clip ? withSelection("Pasted position & size", 1, (it) => it.forEach((s) => Object.assign(s, clip))) : status("Copy a position first.", "err")],
    ["Bring to front", G(bar(4, 4, 14, 14, "#DCE5DF") + bar(12, 12, 16, 16)), () => has("1.8") ? withSelection("Brought to front", 1, (it) => it.forEach((s) => s.setZOrder("BringToFront"))) : status("Needs PowerPointApi 1.8.", "err")],
    ["Send to back", G(bar(12, 12, 16, 16) + bar(4, 4, 14, 14, "#DCE5DF")), () => has("1.8") ? withSelection("Sent to back", 1, (it) => it.forEach((s) => s.setZOrder("SendToBack"))) : status("Needs PowerPointApi 1.8.", "err")],
    ["Group", G(`<rect x="3" y="3" width="26" height="26" rx="3" fill="none" stroke="#C9A46A" stroke-dasharray="3 2"/>` + bar(7, 7, 8, 8) + bar(17, 17, 8, 8)), () => has("1.8") ? withSelection("Grouped", 2, async (it, ctx) => { ctx.presentation.getSelectedSlides().getItemAt(0).shapes.addGroup(it.map((s) => s.id)); }) : status("Needs PowerPointApi 1.8.", "err")],
    ["Ungroup", G(bar(5, 5, 8, 8) + bar(19, 19, 8, 8) + `<path d="M15 15l2 2" stroke="#C9A46A" stroke-width="2"/>`), () => has("1.8") ? withSelection("Ungrouped", 1, async (it, ctx) => { it.forEach((s) => s.load("type")); await ctx.sync(); it.filter((s) => s.type === "Group").forEach((s) => s.group.ungroup()); }) : status("Needs PowerPointApi 1.8.", "err")],
  ],
};

const TXT = [
  ["Deli fonts (Cambria titles, Calibri text)", G(`<text x="16" y="22" font-size="18" font-family="Cambria,serif" font-weight="700" text-anchor="middle" fill="#17362E">Aa</text>`), () => textEach("Brand fonts", (r) => r.load("font/size"), (r) => { r.font.name = (r.font.size || 0) >= 24 ? "Cambria" : "Calibri"; })],
  ["Bigger text", G(`<text x="16" y="22" font-size="17" font-family="Calibri,sans-serif" font-weight="700" text-anchor="middle" fill="#17362E">A+</text>`), () => textEach("Bigger text", (r) => r.load("font/size"), (r) => { r.font.size = Math.round((r.font.size || 12) + 1); })],
  ["Smaller text", G(`<text x="16" y="22" font-size="14" font-family="Calibri,sans-serif" font-weight="700" text-anchor="middle" fill="#17362E">A−</text>`), () => textEach("Smaller text", (r) => r.load("font/size"), (r) => { r.font.size = Math.max(8, Math.round((r.font.size || 12) - 1)); })],
  ["Deli green text", G(`<circle cx="16" cy="16" r="9" fill="#17362E"/>`), () => textEach("Green text", null, (r) => { r.font.color = "#17362E"; })],
  ["Gold text", G(`<circle cx="16" cy="16" r="9" fill="#C9A46A"/>`), () => textEach("Gold text", null, (r) => { r.font.color = "#C9A46A"; })],
  ["Bullets on/off", G(`<circle cx="7" cy="10" r="2" fill="#C9A46A"/><circle cx="7" cy="22" r="2" fill="#C9A46A"/>` + bar(12, 8, 16, 4) + bar(12, 20, 16, 4)), () => textEach("Bullets toggled", (r) => r.load("paragraphFormat/bulletFormat/visible"), (r) => { r.paragraphFormat.bulletFormat.visible = !r.paragraphFormat.bulletFormat.visible; })],
  ["Remove text margins", G(`<rect x="4" y="6" width="24" height="20" fill="none" stroke="#C9A46A" stroke-dasharray="3 2"/>` + bar(4, 10, 24, 3) + bar(4, 16, 18, 3)), () => withSelection("Margins removed", 1, (it) => it.forEach((s) => { const tf = s.textFrame; tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0; }))],
];
async function textEach(label, loader, apply) {
  await withSelection(label, 1, async (it, ctx) => {
    const rs = it.map((s) => s.textFrame.textRange);
    if (loader) { rs.forEach(loader); await ctx.sync(); }
    rs.forEach(apply);
  });
}

function buildIconRows() {
  Object.entries(LAYOUT).forEach(([host, list]) => list.forEach(([t, svg, fn]) => $(host).appendChild(iconBtn(t, svg, fn))));
  TXT.forEach(([t, svg, fn]) => $("txt-tools").appendChild(iconBtn(t, svg, fn)));
}
function iconBtn(title, svg, fn) { const b = document.createElement("button"); b.className = "tile"; b.title = title; b.setAttribute("aria-label", title); b.innerHTML = svg; b.onclick = fn; return b; }

// ================= Numbers =================
async function formatSelected(style) {
  await withSelection("Formatted numbers", 1, async (it, ctx) => {
    const rs = it.map((s) => { const r = s.textFrame.textRange; r.load("text"); return r; });
    await ctx.sync();
    rs.forEach((r) => { const out = r.text.replace(/(−|-)?(Rp)?\d[\d.,]*(bn|m|k|%)?/gi, (m) => D.formatNumber(m, style)); if (out !== r.text) r.text = out; });
  });
}

// ================= Tables & library slides =================
async function insertTable() {
  if (!has("1.8")) return status("Branded tables need a current Microsoft 365 PowerPoint (PowerPointApi 1.8).", "err");
  const rows = Math.max(2, Math.min(15, Number($("tbl-rows").value) || 5)), cols = Math.max(2, Math.min(8, Number($("tbl-cols").value) || 4));
  await run(`Inserted ${rows}×${cols} table`, async (ctx) => {
    const t = currentSlide(ctx).shapes.addTable(rows, cols, Object.assign({ left: D.SLIDE.margin, top: 150, width: D.SLIDE.w - 2 * D.SLIDE.margin }, D.tableSpec(rows, cols)));
    t.name = "Deli table"; await ctx.sync();
  });
}
let libraryB64 = null;
async function insertSlides(b64, ids, label) {
  await run(label, async (ctx) => {
    const opts = { formatting: "UseDestinationTheme" };
    if (ids) opts.sourceSlideIds = ids;
    if (has("1.5")) { const sel = ctx.presentation.getSelectedSlides(); sel.load("items/id"); await ctx.sync(); if (sel.items.length) opts.targetSlideId = sel.items[0].id; }
    ctx.presentation.insertSlidesFromBase64(b64, opts);
    await ctx.sync();
  });
}
async function insertLibrarySlide(id) {
  libraryB64 = libraryB64 || (await fileBase64("assets/Deli_Elements_Library.pptx"));
  await insertSlides(libraryB64, [id + "#"], "Inserted library slide");
}

// ================= Agenda wizard =================
async function makeAgenda() {
  const items = $("agenda-items").value.split("\n").map((s) => s.trim()).filter(Boolean).slice(0, 7);
  if (items.length < 2) return status("Add at least two agenda items, one per line.", "err");
  await run(`Created ${items.length} agenda slides`, async (ctx) => {
    const pres = ctx.presentation;
    const masters = pres.slideMasters; masters.load("items/id");
    await ctx.sync();
    const layouts = masters.items[0].layouts; layouts.load("items/id,items/name");
    await ctx.sync();
    const lay = layouts.items.find((l) => /deli blank/i.test(l.name)) || layouts.items.find((l) => /blank/i.test(l.name)) || layouts.items[0];
    let index = null;
    if (has("1.5")) { const sel = pres.getSelectedSlides(); sel.load("items/id"); const all = pres.slides; all.load("items/id"); await ctx.sync(); if (sel.items.length) index = all.items.findIndex((s) => s.id === sel.items[0].id) + 1; }
    for (let k = 0; k < items.length; k++) {
      const opt = { slideMasterId: masters.items[0].id, layoutId: lay.id };
      if (index !== null && has("1.8")) opt.index = index + k;
      pres.slides.add(opt);
      await ctx.sync();
      const slides = pres.slides; slides.load("items/id"); await ctx.sync();
      const slide = index !== null && has("1.8") ? slides.items[index + k] : slides.items[slides.items.length - 1];
      const title = slide.shapes.addTextBox("Agenda", { left: D.SLIDE.margin, top: 54, width: 600, height: 50 });
      styleText(title, { kind: "text", text: "Agenda", font: { name: "Cambria", size: 30, bold: true, color: "#17362E" }, align: "Left", valign: "Middle" });
      title.name = "Deli agenda title";
      const prims = EL.build("agenda", { items: items.length, current: k + 1, y: 140, x: D.SLIDE.margin }).map((p) => (/^text-/.test(p.name) ? Object.assign({}, p, { text: items[Number(p.name.split("-")[1]) - 1] }) : p));
      await drawPrims(slide.shapes, prims, "agenda", ctx);
    }
  });
}

// ================= My slides (personal library in this browser profile) =================
const MY_KEY = "deliMySlides";
const myLoad = () => { try { return JSON.parse(localStorage.getItem(MY_KEY) || "[]"); } catch (e) { return []; } };
const mySave = (list) => { try { localStorage.setItem(MY_KEY, JSON.stringify(list)); return true; } catch (e) { status("Not enough local storage for that slide. Delete an older one first.", "err"); return false; } };
function renderMy() {
  const host = $("my-list"); host.innerHTML = "";
  const list = myLoad();
  if (!list.length) { host.innerHTML = '<p class="hint">Nothing saved yet.</p>'; return; }
  list.forEach((m, i) => {
    const row = document.createElement("div"); row.className = "item";
    const ins = document.createElement("button"); ins.textContent = m.name; ins.title = "Insert after the selected slide"; ins.onclick = () => insertSlides(m.b64, null, "Inserted " + m.name);
    const del = document.createElement("button"); del.textContent = "✕"; del.title = "Remove from my slides"; del.onclick = () => { const l = myLoad(); l.splice(i, 1); mySave(l); renderMy(); };
    row.append(ins, del); host.appendChild(row);
  });
}
async function saveMySlide() {
  if (!has("1.8")) return status("Saving slides needs a current Microsoft 365 PowerPoint (PowerPointApi 1.8).", "err");
  const name = $("my-name").value.trim() || "My slide " + (myLoad().length + 1);
  await run("Saved " + name, async (ctx) => {
    const r = currentSlide(ctx).exportAsBase64(); await ctx.sync();
    const l = myLoad(); l.push({ name, b64: r.value }); if (mySave(l)) { $("my-name").value = ""; renderMy(); }
  });
}

// ================= Brand check =================
let lastIssues = [];
async function brandCheck() {
  await run("Brand check finished", async (ctx) => {
    const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
    const per = slides.items.map((sl) => { const sh = sl.shapes; sh.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height"); return sh; });
    await ctx.sync();
    const all = [], texty = [];
    per.forEach((sh, si) => sh.items.forEach((s) => {
      const rec = { slideIndex: si, slideId: slides.items[si].id, shapeId: s.id, name: s.name, left: s.left, top: s.top, width: s.width, height: s.height };
      all.push(rec);
      if (["GeometricShape", "TextBox", "Placeholder"].includes(s.type)) { const r = s.textFrame.textRange; r.load("text,font/name,font/size,font/color"); texty.push([rec, r]); }
    }));
    await ctx.sync();
    texty.forEach(([rec, r]) => { rec.text = r.text; rec.font = r.font.name; rec.size = r.font.size; rec.color = r.font.color; });
    lastIssues = D.checkDeck(all).map((iss) => Object.assign(iss, { rec: all.find((a) => a.slideIndex === iss.slideIndex && a.name === iss.shapeName) }));
    renderIssues();
  });
}
function renderIssues() {
  const box = $("issues"); box.innerHTML = "";
  if (!lastIssues.length) { box.innerHTML = '<p class="good">No issues found. The deck is on brand.</p>'; $("fix").disabled = true; return; }
  const fixable = lastIssues.filter((i) => i.fixable).length;
  const head = document.createElement("p"); head.textContent = `${lastIssues.length} issue${lastIssues.length > 1 ? "s" : ""} found` + (fixable ? `, ${fixable} fixable automatically.` : "."); box.appendChild(head);
  const ul = document.createElement("ul");
  lastIssues.forEach((i) => { const li = document.createElement("li"); li.textContent = `Slide ${i.slideIndex + 1} · ${i.shapeName || "shape"}: ${i.msg}`; ul.appendChild(li); });
  box.appendChild(ul); $("fix").disabled = !fixable;
}
async function fixIssues() {
  await run("Fixed fonts and colours", async (ctx) => {
    for (const i of lastIssues.filter((x) => x.fixable && x.rec)) {
      const font = ctx.presentation.slides.getItem(i.rec.slideId).shapes.getItem(i.rec.shapeId).textFrame.textRange.font;
      if (i.type === "font") font.name = (i.rec.size || 0) >= 24 ? "Cambria" : "Calibri";
      if (i.type === "color") font.color = D.nearestBrand(i.rec.color);
    }
    await ctx.sync();
  });
  await brandCheck();
}

// ================= Wiring =================
function wire() {
  buildGallery("");
  let t; $("el-search").oninput = (e) => { clearTimeout(t); t = setTimeout(() => buildGallery(e.target.value), 120); };
  setupIcons();
  buildIconRows();
  renderMy();
  document.querySelectorAll("[data-num]").forEach((b) => (b.onclick = () => formatSelected(b.dataset.num)));
  document.querySelectorAll("[data-lib]").forEach((b) => (b.onclick = () => insertLibrarySlide(b.dataset.lib)));
  document.querySelectorAll("[data-logo]").forEach((b) => (b.onclick = () => insertLogo(b.dataset.logo)));
  $("tbl-insert").onclick = insertTable;
  $("agenda-make").onclick = makeAgenda;
  $("my-save").onclick = saveMySlide;
  $("check").onclick = brandCheck;
  $("fix").onclick = fixIssues;
  document.querySelectorAll(".tab").forEach((tab) => (tab.onclick = () => {
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("on", x === tab));
    document.querySelectorAll(".panel").forEach((p) => p.classList.toggle("on", p.id === tab.dataset.panel));
    if (tab.dataset.panel === "p-ic") loadIcons();
  }));
}

Office.onReady((info) => {
  wire();
  inPowerPoint.ok = info.host === Office.HostType.PowerPoint;
  if (!inPowerPoint.ok) return status("Preview only: open the Deli tab inside PowerPoint to use these tools.", "err");
  const v = ["1.4", "1.5", "1.8", "1.10"].filter(has).pop() || "older";
  status(`Ready · PowerPoint API ${v}`, "ok");
});
