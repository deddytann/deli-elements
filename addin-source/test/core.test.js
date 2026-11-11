const assert = require("assert");
const D = require("../src/core.js");

const shapes = [
  { left: 100, top: 50, width: 80, height: 40 },
  { left: 300, top: 120, width: 60, height: 60 },
  { left: 200, top: 80, width: 40, height: 20 },
];

// align
assert.deepStrictEqual(D.align(shapes, "left").map((p) => p.left), [100, 100, 100]);
assert.deepStrictEqual(D.align(shapes, "right").map((p) => p.left), [280, 300, 320]);
assert.deepStrictEqual(D.align(shapes, "top").map((p) => p.top), [50, 50, 50]);
assert.deepStrictEqual(D.align(shapes, "bottom").map((p) => p.top), [140, 120, 160]);
assert.deepStrictEqual(D.align([shapes[0]], "center"), [{ left: 440, top: 50 }]); // one shape centres on 960pt slide
// distribute: span 100..360 (260), widths 180 → gap 40; order by left: 0,2,1
const dh = D.distribute(shapes, "h");
assert.deepStrictEqual(dh.map((p) => p.left), [100, 300, 220]);
// stack with 14.4pt gap
assert.deepStrictEqual(D.stack(shapes, "h", 14.4).map((p) => p.left), [100, 248.8, 194.4]);
// size and swap
assert.deepStrictEqual(D.matchSize(shapes, "size")[2], { width: 80, height: 40 });
assert.deepStrictEqual(D.matchSize(shapes, "width")[1], { width: 80, height: 60 });
assert.deepStrictEqual(D.swap(shapes[0], shapes[1]), [{ left: 300, top: 120 }, { left: 100, top: 50 }]);
assert.strictEqual(D.snap(37, 3.6), 36);

// numbers
assert.strictEqual(D.parseNumber("Rp2.843.279.950"), 2843279950);
assert.strictEqual(D.parseNumber("2,843,279,950"), 2843279950);
assert.strictEqual(D.parseNumber("6.59bn"), 6.59e9);
assert.strictEqual(D.formatNumber("2843279950", "rp"), "Rp2.8bn");
assert.strictEqual(D.formatNumber("921442250", "rp"), "Rp921.4m");
assert.strictEqual(D.formatNumber("172416", "rp"), "Rp172k");
assert.strictEqual(D.formatNumber("1232.77", "thousands"), "1,233");
assert.strictEqual(D.formatNumber("0.431", "pct"), "43.1%");
assert.strictEqual(D.formatNumber("hello", "rp"), "hello");

// brand check
const issues = D.checkDeck([
  { slideIndex: 0, name: "ok", left: 10, top: 10, width: 100, height: 20, text: "Fine", font: "Calibri", size: 14, color: "#17362E" },
  { slideIndex: 0, name: "bad font", left: 10, top: 10, width: 100, height: 20, text: "Hi", font: "Arial", size: 14, color: "#17362E" },
  { slideIndex: 1, name: "bad color", left: 10, top: 10, width: 100, height: 20, text: "Hi", font: "Calibri", size: 8, color: "#FF0000" },
  { slideIndex: 2, name: "off", left: 900, top: 10, width: 100, height: 20 },
  { slideIndex: 2, name: "leftover", left: 10, top: 10, width: 100, height: 20, text: "Owner: [name]", font: "+mn-lt", size: 12, color: null },
]);
assert.deepStrictEqual(issues.map((i) => i.type).sort(), ["color", "font", "leftover", "offslide", "size"]);
assert.strictEqual(D.nearestBrand("#FF0000"), "#A4463B");

// every element builds, stays on the slide, and has valid primitives
const GEOMS = ["Ellipse", "Pie", "FlowChartDelay", "RoundRectangle", "Rectangle", "Chevron"];
Object.keys(D.ELEMENTS).forEach((k) => {
  const prims = D.buildElement(k, { steps: 6, current: 3, rows: 8, n: 7, text: "Confidential" });
  assert(prims.length > 0, k);
  prims.forEach((p) => {
    assert(["shape", "text", "line"].includes(p.kind), `${k} kind`);
    if (p.kind === "shape") assert(GEOMS.includes(p.geom), `${k} geom ${p.geom}`);
    assert(p.x >= 0 && p.y >= 0 && p.x + p.w <= 960.01 && p.y + p.h <= 540.01, `${k}/${p.name} off slide: ${JSON.stringify(p)}`);
    if (p.fill) assert(/^#[0-9A-F]{6}$/i.test(p.fill), `${k} fill`);
  });
});

// table spec
const t = D.tableSpec(3, 4);
assert.strictEqual(t.values.length, 3);
assert.strictEqual(t.specificCellProperties[0][0].fill.color, "#17362E");
assert.strictEqual(t.specificCellProperties[1][1].font.bold, false);

console.log("core.js: all tests passed (" + Object.keys(D.ELEMENTS).length + " elements)");
