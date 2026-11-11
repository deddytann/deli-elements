/* Loads taskpane.html in jsdom with a mock PowerPoint API and presses every button. */
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const { JSDOM } = require(require.resolve("jsdom", { paths: [process.env.JSDOM_PATH || __dirname] }));

const ROOT = path.join(__dirname, "..", "..");
let ids = 1;
const log = [];
const GEOMS = ["Ellipse", "Pie", "FlowChartDelay", "RoundRectangle", "Rectangle", "Chevron", "HomePlate", "Diamond", "Star5", "UpArrow", "DownArrow", "RightArrow", "WedgeRoundRectCallout", "RightBrace"];

function mockShape(props) {
  const s = Object.assign({ id: String(ids++), name: "", type: "GeometricShape", left: 0, top: 0, width: 10, height: 10, load() {}, setZOrder(z) { assert(["BringToFront", "SendToBack"].includes(z)); log.push("z:" + z); } }, props);
  s.fill = { setSolidColor(c) { assert(/^#[0-9A-F]{6}$/i.test(c), "colour " + c); }, clear() {}, set transparency(v) { assert(v >= 0 && v <= 1); } };
  s.lineFormat = { set dashStyle(v) { assert(["Dash"].includes(v)); } };
  s.group = { ungroup() { log.push("ungroup"); } };
  const font = { name: "Arial", size: 7, color: "#FF0000" };
  s.textFrame = { textRange: { text: (props && props.text) || "", font, paragraphFormat: { bulletFormat: { visible: false } }, load() {} } };
  return s;
}
function mockSlide(id) {
  const items = [];
  const add = (type, box, extra) => { const sh = mockShape(Object.assign({ type }, box, extra)); items.push(sh); return sh; };
  return {
    id, exportAsBase64() { log.push("export"); return { value: "UEsDBBQ" + "A".repeat(2000) }; },
    shapes: {
      items, load() {},
      addGeometricShape(g, box) { assert(GEOMS.includes(g), "geom " + g); return add("GeometricShape", box); },
      addTextBox(t, box) { return add("TextBox", box, { text: t }); },
      addLine(kind, box) { assert.strictEqual(kind, "Straight"); return add("Line", box); },
      addGroup(list) { assert(list.length > 1 && list.every((x) => typeof x === "string")); log.push("group"); return add("Group", {}); },
      addTable(r, c, o) { assert(o.values.length === r && o.specificCellProperties[0].length === c); log.push("table"); return add("Table", {}); },
      getItem(sid) { return items.find((x) => x.id === sid); },
    },
  };
}

async function main() {
  const slides = [mockSlide("256#"), mockSlide("257#")];
  const slide = slides[0];
  const a = slide.shapes.addTextBox("Owner: [name]", { left: 100, top: 50, width: 80, height: 40 });
  const b = slide.shapes.addTextBox("2843279950", { left: 300, top: 120, width: 60, height: 60 });
  const c = slide.shapes.addTextBox("0.431", { left: 200, top: 80, width: 40, height: 20 });
  let selected = [a, b, c];
  const presentation = {
    slides: { get items() { return slides; }, load() {}, getItemAt: (i) => slides[i], getItem: (sid) => slides.find((s) => s.id === sid) || slides[0], add(o) { assert(o.layoutId && o.slideMasterId); const s = mockSlide(String(300 + slides.length) + "#"); if (o.index != null) slides.splice(o.index, 0, s); else slides.push(s); log.push("addSlide"); } },
    slideMasters: { items: [{ id: "m1", layouts: { load() {}, items: [{ id: "l1", name: "Deli Title Only" }, { id: "l2", name: "Deli Blank" }] } }], load() {} },
    getSelectedSlides: () => ({ items: [slide], load() {}, getItemAt: () => slide }),
    getSelectedShapes: () => ({ items: selected, load() {} }),
    insertSlidesFromBase64(b64, opts) { assert(b64.length > 1000); if (opts.sourceSlideIds) assert(/^\d+#$/.test(opts.sourceSlideIds[0])); log.push("insertSlides"); },
  };
  const registry = {}, images = [];
  const Office = {
    HostType: { PowerPoint: "PowerPoint" }, CoercionType: { Image: "image", XmlSvg: "xmlSvg" }, AsyncResultStatus: { Succeeded: "succeeded" },
    context: { requirements: { isSetSupported: () => true }, document: { setSelectedDataAsync(data, o, cb) { assert(o.imageWidth > 0); if (o.coercionType === "xmlSvg") assert(/^<svg[^>]+xmlns/.test(data) && !/undefined|NaN/.test(data)); images.push(o.coercionType); cb({ status: "succeeded" }); } } },
    onReady(cb) { cb({ host: "PowerPoint" }); return Promise.resolve(); },
    actions: { associate(n, f) { registry[n] = f; } },
  };
  const PowerPoint = { run: async (fn) => fn({ presentation, sync: async () => {} }) };

  const html = fs.readFileSync(path.join(ROOT, "taskpane.html"), "utf8").replace(/<script[^>]*><\/script>/g, "");
  const dom = new JSDOM(html, { runScripts: "outside-only", url: "https://deddytann.github.io/deli-elements/taskpane.html" });
  const w = dom.window;
  Object.assign(w, { Office, PowerPoint, btoa: (s) => Buffer.from(s, "binary").toString("base64"),
    fetch: async (u) => { const buf = fs.readFileSync(path.join(ROOT, u)); return { arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), json: async () => JSON.parse(buf.toString()) }; } });
  w.console.error = (e) => { throw e; };
  ["core.js", "elements.js", "taskpane.js", "commands.js"].forEach((f) => w.eval(fs.readFileSync(path.join(ROOT, f), "utf8")));
  const doc = w.document, st = doc.getElementById("status");
  const tick = () => new Promise((r) => setTimeout(r, 15));
  const click = async (el, label) => { el.click(); await tick(); await tick(); assert(!st.className.includes("err"), `${label || el.title || el.textContent}: ${st.textContent}`); };
  assert(st.textContent.startsWith("Ready"), st.textContent);

  // Elements: every tile; tiles with options open a drawer, then Insert
  const tiles = [...doc.querySelectorAll("#el-gallery .tile")];
  for (const t of tiles) {
    assert(t.querySelector("svg"), "thumbnail for " + t.title);
    await click(t);
    if (t.classList.contains("has-opts")) { const go = doc.querySelector(".drawer .primary"); assert(go, "drawer for " + t.title); const inp = doc.querySelector(".drawer input[type=number]"); if (inp) { inp.value = inp.max; inp.dispatchEvent(new w.Event("input")); } await click(go, "insert " + t.title); }
  }
  // search filters
  const s = doc.getElementById("el-search"); s.value = "harvey"; s.dispatchEvent(new w.Event("input")); await new Promise((r) => setTimeout(r, 200));
  assert.strictEqual(doc.querySelectorAll("#el-gallery .tile").length, 5);
  for (const btn of doc.querySelectorAll("[data-logo]")) await click(btn);
  for (const btn of doc.querySelectorAll("[data-lib]")) await click(btn);
  // Icons
  doc.querySelector('[data-panel="p-ic"]').click(); await tick(); await tick();
  assert(doc.querySelectorAll("#ic-grid button").length === 120, "first page of icons");
  assert.strictEqual(doc.querySelector("#ic-grid button").title, "coffee", "featured first");
  for (const st2 of doc.querySelectorAll("#ic-style button")) { await click(st2); await click(doc.querySelector("#ic-grid button"), "insert icon " + st2.textContent); }
  const q = doc.getElementById("ic-search"); q.value = "bread"; q.dispatchEvent(new w.Event("input")); await new Promise((r) => setTimeout(r, 250));
  assert(doc.querySelectorAll("#ic-grid button").length >= 1 && doc.querySelectorAll("#ic-grid button").length < 30, "icon search");
  await click(doc.getElementById("ic-colors").children[1]);
  // Layout + text tools (positions reset each time)
  for (const anchor of ["selection", "first", "slide"]) { doc.getElementById("anchor").value = anchor; for (const btn of doc.querySelectorAll("#lay-align .tile, #lay-space .tile, #lay-size .tile, #lay-arrange .tile")) { a.left = 100; b.left = 300; c.left = 200; a.type = "Group"; await click(btn); } }
  for (const btn of doc.querySelectorAll("#txt-tools .tile")) await click(btn);
  selected = [b, c];
  b.textFrame.textRange.text = "2843279950"; await click(doc.querySelector('[data-num="rp"]'));
  assert.strictEqual(b.textFrame.textRange.text, "Rp2.8bn");
  await click(doc.getElementById("tbl-insert"));
  // Agenda wizard
  const before = slides.length; await click(doc.getElementById("agenda-make"));
  assert.strictEqual(slides.length, before + 4, "4 agenda slides");
  assert(slides[1].shapes.items.some((x) => x.textFrame.textRange.text === "Where we are"), "agenda text filled");
  // My slides
  doc.getElementById("my-name").value = "Test slide"; await click(doc.getElementById("my-save"));
  assert.strictEqual(doc.querySelectorAll("#my-list .item").length, 1);
  await click(doc.querySelector("#my-list .item button"));
  await click(doc.querySelectorAll("#my-list .item button")[1]);
  // Brand check + fix
  selected = [a, b, c]; a.type = "TextBox"; a.textFrame.textRange.font.name = "Arial"; a.textFrame.textRange.font.color = "#FF0000";
  await click(doc.getElementById("check"));
  assert(/Off-brand font/.test(doc.getElementById("issues").textContent));
  await click(doc.getElementById("fix"));
  // Ribbon commands
  for (const [n, fn] of Object.entries(registry)) { let done = false; await fn({ completed: () => (done = true) }); assert(done, n); }
  console.log(`panel mock: ${tiles.length} element tiles, ${images.length} images/icons inserted, layout+text tools x3 anchors, table, agenda (${slides.length - before} slides), my slides, brand check, ${Object.keys(registry).length} ribbon commands — no errors`);
}
main().catch((e) => { console.error(e); process.exit(1); });
