const assert = require("assert");
const path = require("path");
const ROOT = path.join(__dirname, "..", "..");
const D = require(path.join(ROOT, "core.js"));
const EL = require(path.join(ROOT, "elements.js"));

const shapes = [{ left: 100, top: 50, width: 80, height: 40 }, { left: 300, top: 120, width: 60, height: 60 }, { left: 200, top: 80, width: 40, height: 20 }];
assert.deepStrictEqual(D.align(shapes, "left").map((p) => p.left), [100, 100, 100]);
assert.deepStrictEqual(D.align(shapes, "right").map((p) => p.left), [280, 300, 320]);
assert.deepStrictEqual(D.align(shapes, "left", "first").map((p) => p.left), [100, 100, 100]);
assert.deepStrictEqual(D.align(shapes, "right", "first").map((p) => p.left), [100, 120, 140]);
assert.deepStrictEqual(D.align(shapes, "top", "slide").map((p) => p.top), [0, 0, 0]);
assert.deepStrictEqual(D.align([shapes[0]], "center"), [{ left: 440, top: 50 }]);
assert.deepStrictEqual(D.distribute(shapes, "h").map((p) => p.left), [100, 300, 220]);
assert.deepStrictEqual(D.stack(shapes, "h", 14.4).map((p) => p.left), [100, 248.8, 194.4]);
assert.deepStrictEqual(D.grid(shapes, 2, 10), [{ left: 100, top: 50 }, { left: 100, top: 120 }, { left: 190, top: 50 }]);
assert.deepStrictEqual(D.matchSize(shapes, "size")[2], { width: 80, height: 40 });
assert.deepStrictEqual(D.swap(shapes[0], shapes[1]), [{ left: 300, top: 120 }, { left: 100, top: 50 }]);
assert.strictEqual(D.parseNumber("Rp2.843.279.950"), 2843279950);
assert.strictEqual(D.formatNumber("2843279950", "rp"), "Rp2.8bn");
assert.strictEqual(D.formatNumber("0.431", "pct"), "43.1%");
assert.strictEqual(D.formatNumber("1232.77", "thousands"), "1,233");
const issues = D.checkDeck([
  { slideIndex: 0, name: "a", left: 10, top: 10, width: 100, height: 20, text: "Hi", font: "Arial", size: 14, color: "#17362E" },
  { slideIndex: 1, name: "b", left: 10, top: 10, width: 100, height: 20, text: "Hi", font: "Calibri", size: 8, color: "#FF0000" },
  { slideIndex: 2, name: "c", left: 900, top: 10, width: 100, height: 20 },
  { slideIndex: 2, name: "d", left: 10, top: 10, width: 100, height: 20, text: "Owner: [name]", font: "+mn-lt", size: 12 },
]);
assert.deepStrictEqual(issues.map((i) => i.type).sort(), ["color", "font", "leftover", "offslide", "size"]);

// Every element, at default and at its extreme parameters, builds valid primitives that stay on the slide.
const GEOMS = ["Ellipse", "Pie", "FlowChartDelay", "RoundRectangle", "Rectangle", "Chevron", "HomePlate", "Diamond", "Star5", "UpArrow", "DownArrow", "RightArrow", "WedgeRoundRectCallout", "RightBrace"];
let checked = 0;
Object.entries(EL.ELEMENTS).forEach(([k, e]) => {
  const variants = [{}];
  e.params.filter((p) => p.type === "number").forEach((p) => { variants.push({ [p.key]: p.min }); variants.push({ [p.key]: p.max }); });
  e.params.filter((p) => p.type === "select").forEach((p) => p.options.forEach((o) => variants.push({ [p.key]: o })));
  if (e.params.find((p) => p.key === "steps") && e.params.find((p) => p.key === "current")) variants.push({ steps: 8, current: 8 });
  variants.forEach((v) => {
    const prims = EL.build(k, v);
    assert(prims.length > 0, k);
    prims.forEach((p) => {
      const where = `${k} ${JSON.stringify(v)} ${p.name}`;
      assert(["shape", "text", "line"].includes(p.kind), where);
      if (p.kind === "shape") assert(GEOMS.includes(p.geom), where + " geom " + p.geom);
      assert(p.x >= -0.5 && p.y >= -0.5 && p.x + p.w <= 960.5 && p.y + p.h <= 540.5, `${where} off slide ${JSON.stringify([p.x, p.y, p.w, p.h])}`);
      [p.fill, p.line, p.font && p.font.color].filter(Boolean).forEach((c) => assert(/^#[0-9A-F]{6}$/i.test(c), where + " colour " + c));
      if (p.text !== undefined) assert(typeof p.text === "string", where + " text");
    });
    const svg = EL.thumbnail(k, v, 64, 44);
    assert(/^<svg[^>]+viewBox="[-\d. ]+"/.test(svg) && !/NaN|undefined/.test(svg), k + " thumbnail");
    assert(!/(\s(fill|stroke)="[^"]*")[^>]*\1/.test(svg.replace(/\s(fill|stroke)="([^"]*)"([^>]*)\s\1="/g, " DUP")) , k + " duplicate attributes");
    checked++;
  });
});
// Smart elements respond to their parameters
assert.strictEqual(EL.build("evalDots", { score: 2 }).filter((p) => p.fill === EL.COLORS.green).length, 2);
assert.strictEqual(EL.build("chevrons", { steps: 6 }).length, 6);
assert.strictEqual(EL.build("agenda", { items: 4, current: 2 }).filter((p) => p.name.startsWith("item-") && p.fill === EL.COLORS.green).length, 1);
assert(EL.build("trafficV", { state: "red" }).some((p) => p.name === "lamp-red" && p.fill !== "#3A4440"));
assert.strictEqual(EL.build("percentBar", { pct: 150 }).find((p) => p.name === "value").text, "100%"); // clamped
console.log(`unit: layout/number/brand-check OK; ${Object.keys(EL.ELEMENTS).length} elements, ${checked} parameter variants built and thumbnailed`);
