/* Runs taskpane.html + commands.js in jsdom against a mock of the PowerPoint JS API and clicks every button. */
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const { JSDOM } = require("jsdom");

const SRC = path.join(__dirname, "..", "src");
let ids = 1;
const log = [];

function mockShape(props) {
  const s = Object.assign({ id: String(ids++), name: "", type: "GeometricShape", left: 0, top: 0, width: 10, height: 10, load() {} }, props);
  s.fill = { setSolidColor(c) { assert(/^#[0-9A-F]{6}$/i.test(c), "bad color " + c); s._fill = c; }, clear() { s._fill = null; } };
  s.lineFormat = {};
  const font = { name: "Arial", size: 7, color: "#FF0000" };
  s.textFrame = { textRange: { text: props && props.text || "", font, paragraphFormat: {}, load() {} } };
  return s;
}
function mockSlide() {
  const items = [];
  const add = (type, box, extra) => { const sh = mockShape(Object.assign({ type }, box, extra)); items.push(sh); return sh; };
  return {
    id: "256#", shapes: {
      items, load() {},
      addGeometricShape(g, box) { assert(["Ellipse", "Pie", "FlowChartDelay", "RoundRectangle", "Rectangle", "Chevron"].includes(g), "geom " + g); log.push("geom:" + g); return add("GeometricShape", box); },
      addTextBox(t, box) { log.push("text"); return add("TextBox", box, { text: t }); },
      addLine(kind, box) { assert.strictEqual(kind, "Straight"); log.push("line"); return add("Line", box); },
      addGroup(list) { assert(list.length > 1 && list.every((x) => typeof x === "string"), "group ids"); log.push("group:" + list.length); return add("Group", {}); },
      addTable(r, c, o) { assert(o.values.length === r && o.values[0].length === c && o.specificCellProperties.length === r); log.push(`table:${r}x${c}`); return add("Table", {}); },
      getItem(id) { return items.find((x) => x.id === id); },
    },
  };
}

async function main() {
  const slide = mockSlide();
  // Pre-existing shapes for layout tools and brand check
  const a = slide.shapes.addTextBox("Owner: [name]", { left: 100, top: 50, width: 80, height: 40 });
  const b = slide.shapes.addTextBox("2843279950", { left: 300, top: 120, width: 60, height: 60 });
  const c = slide.shapes.addTextBox("0.431", { left: 200, top: 80, width: 40, height: 20 });
  log.length = 0;
  let selected = [a, b, c];
  const presentation = {
    slides: { items: [slide], load() {}, getItemAt: () => slide, getItem: () => slide },
    getSelectedSlides: () => ({ items: [slide], load() {}, getItemAt: () => slide }),
    getSelectedShapes: () => ({ items: selected, load() {} }),
    insertSlidesFromBase64(b64, opts) { assert(b64.length > 1000 && /^\d+#$/.test(opts.sourceSlideIds[0])); log.push("insertSlides:" + opts.sourceSlideIds[0]); },
  };
  const registry = {};
  const Office = {
    HostType: { PowerPoint: "PowerPoint" },
    context: { requirements: { isSetSupported: () => true } },
    onReady(cb) { cb({ host: "PowerPoint" }); return Promise.resolve(); },
    actions: { associate(n, f) { registry[n] = f; } },
  };
  const PowerPoint = { run: async (fn) => fn({ presentation, sync: async () => {} }) };

  const html = fs.readFileSync(path.join(SRC, "taskpane.html"), "utf8").replace(/<script[^>]*><\/script>/g, "");
  const dom = new JSDOM(html, { runScripts: "outside-only" });
  const w = dom.window;
  Object.assign(w, { Office, PowerPoint, btoa: (s) => Buffer.from(s, "binary").toString("base64"),
    fetch: async (u) => ({ arrayBuffer: async () => { const buf = fs.readFileSync(path.join(SRC, u)); return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); } }) });
  w.console.error = (e) => { throw e; };
  w.eval(fs.readFileSync(path.join(SRC, "core.js"), "utf8"));
  w.eval(fs.readFileSync(path.join(SRC, "taskpane.js"), "utf8"));
  w.eval(fs.readFileSync(path.join(SRC, "commands.js"), "utf8"));
  const doc = w.document, statusEl = doc.getElementById("status");
  const tick = () => new Promise((r) => setTimeout(r, 20));
  const click = async (el) => { el.click(); await tick(); await tick(); assert(!statusEl.className.includes("err"), `${el.textContent}: ${statusEl.textContent}`); };

  assert(statusEl.textContent.startsWith("Ready"), statusEl.textContent);
  // Insert every element
  const elButtons = [...doc.querySelectorAll("#elements button")];
  for (const btn of elButtons) await click(btn);
  assert(log.filter((x) => x.startsWith("group")).length > 5, "multi-part elements were grouped");
  // Library slides, layout tools, numbers, table
  for (const btn of doc.querySelectorAll("[data-lib]")) await click(btn);
  for (const btn of doc.querySelectorAll("[data-layout]")) { a.left = 100; b.left = 300; c.left = 200; await click(btn); }
  selected = [b, c];
  for (const btn of doc.querySelectorAll("[data-num]")) { b.textFrame.textRange.text = "2843279950"; c.textFrame.textRange.text = "0.431"; await click(btn); }
  b.textFrame.textRange.text = "2843279950"; await click(doc.querySelector('[data-num="rp"]'));
  assert.strictEqual(b.textFrame.textRange.text, "Rp2.8bn");
  await click(doc.getElementById("tbl-insert"));
  // Brand check finds the planted problems, then fixes fonts/colors
  await click(doc.getElementById("check"));
  const txt = doc.getElementById("issues").textContent;
  assert(/Off-brand font/.test(txt) && /placeholder|Placeholder/.test(txt) && /below 9pt/.test(txt), txt);
  assert(!doc.getElementById("fix").disabled);
  await click(doc.getElementById("fix"));
  assert.strictEqual(a.textFrame.textRange.font.name, "Calibri");
  assert.strictEqual(a.textFrame.textRange.font.color, "#A4463B");
  // Ribbon commands
  selected = [a, b, c];
  for (const [name, fn] of Object.entries(registry)) { let done = false; await fn({ completed: () => (done = true) }); assert(done, name + " completed"); }
  assert.strictEqual(Object.keys(registry).length, 10);
  console.log(`taskpane mock: ${elButtons.length} element buttons, ${doc.querySelectorAll("[data-lib]").length} library slides, ${doc.querySelectorAll("[data-layout]").length} layout tools, 10 ribbon commands — all ran without errors`);
}
main().catch((e) => { console.error(e); process.exit(1); });
