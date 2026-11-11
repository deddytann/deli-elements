/*
 * Deli Elements: pure logic (no Office.js), shared by the task pane and the tests.
 * Units are PowerPoint points (72 pt = 1 inch). A 16:9 Deli slide is 960 x 540 pt.
 */
(function (root) {
  "use strict";

  const BRAND = {
    green: "#17362E", gold: "#C9A46A", teal: "#4E7A6C", goldLt: "#E8D6B5", paper: "#FAF8F4",
    white: "#FFFFFF", ink: "#1F2A27", muted: "#6B7570", line: "#E4DED3", sage: "#DCE5DF", red: "#A4463B",
  };
  const BRAND_FONTS = ["Cambria", "Calibri"];
  const SLIDE = { w: 960, h: 540, margin: 43.2 }; // 0.6 inch content margin
  const IN = 72;

  // ---------- Layout math ----------
  function bounds(shapes) {
    const l = Math.min(...shapes.map((s) => s.left));
    const t = Math.min(...shapes.map((s) => s.top));
    const r = Math.max(...shapes.map((s) => s.left + s.width));
    const b = Math.max(...shapes.map((s) => s.top + s.height));
    return { left: l, top: t, right: r, bottom: b, width: r - l, height: b - t };
  }

  /** Returns new {left, top} for each shape. With one shape, aligns to the slide. */
  function align(shapes, mode) {
    const ref = shapes.length === 1 ? { left: 0, top: 0, right: SLIDE.w, bottom: SLIDE.h, width: SLIDE.w, height: SLIDE.h } : bounds(shapes);
    return shapes.map((s) => {
      const p = { left: s.left, top: s.top };
      if (mode === "left") p.left = ref.left;
      if (mode === "center") p.left = ref.left + (ref.width - s.width) / 2;
      if (mode === "right") p.left = ref.right - s.width;
      if (mode === "top") p.top = ref.top;
      if (mode === "middle") p.top = ref.top + (ref.height - s.height) / 2;
      if (mode === "bottom") p.top = ref.bottom - s.height;
      return p;
    });
  }

  /** Equal gaps between shapes along an axis ("h" or "v"). Needs 3+ shapes. Order of result matches input. */
  function distribute(shapes, axis) {
    if (shapes.length < 3) return shapes.map((s) => ({ left: s.left, top: s.top }));
    const pos = axis === "h" ? "left" : "top", size = axis === "h" ? "width" : "height";
    const order = shapes.map((s, i) => i).sort((a, b) => shapes[a][pos] - shapes[b][pos]);
    const first = shapes[order[0]], last = shapes[order[order.length - 1]];
    const span = last[pos] + last[size] - first[pos];
    const total = shapes.reduce((acc, s) => acc + s[size], 0);
    const gap = (span - total) / (shapes.length - 1);
    const out = shapes.map((s) => ({ left: s.left, top: s.top }));
    let cursor = first[pos];
    order.forEach((i) => { out[i][pos] = cursor; cursor += shapes[i][size] + gap; });
    return out;
  }

  /** Places shapes side by side (or stacked) with a fixed gap, starting at the first shape's position. */
  function stack(shapes, axis, gap) {
    const pos = axis === "h" ? "left" : "top", size = axis === "h" ? "width" : "height";
    const order = shapes.map((s, i) => i).sort((a, b) => shapes[a][pos] - shapes[b][pos]);
    const out = shapes.map((s) => ({ left: s.left, top: s.top }));
    let cursor = shapes[order[0]][pos];
    order.forEach((i) => { out[i][pos] = cursor; cursor += shapes[i][size] + gap; });
    return out;
  }

  /** Same width/height/size as the FIRST selected shape. */
  function matchSize(shapes, mode) {
    const ref = shapes[0];
    return shapes.map((s) => ({
      width: mode === "height" ? s.width : ref.width,
      height: mode === "width" ? s.height : ref.height,
    }));
  }

  function swap(a, b) {
    return [{ left: b.left, top: b.top }, { left: a.left, top: a.top }];
  }

  function snap(v, grid) { return Math.round(v / grid) * grid; }

  // ---------- Number formatting ----------
  function parseNumber(text) {
    const m = String(text).replace(/\s/g, "").match(/^(-|−)?(Rp)?([\d.,]+)(%|bn|m|k)?$/i);
    if (!m) return null;
    let digits = m[3];
    // Indonesian style "1.234.567,8" vs English "1,234,567.8"
    // ("0.431" is always a decimal; "1.234" is read as Indonesian thousands, as ESB exports write it)
    if (/^[1-9]\d{0,2}(\.\d{3})+(,\d+)?$/.test(digits)) digits = digits.replace(/\./g, "").replace(",", ".");
    else digits = digits.replace(/,/g, "");
    const n = Number(digits);
    if (!isFinite(n)) return null;
    const mult = { bn: 1e9, m: 1e6, k: 1e3 }[(m[4] || "").toLowerCase()] || 1;
    return (m[1] ? -1 : 1) * n * mult;
  }

  function withSep(n, decimals) {
    const neg = n < 0;
    const s = Math.abs(n).toFixed(decimals);
    const [i, d] = s.split(".");
    const body = i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (d ? "." + d : "");
    return (neg ? "−" : "") + body;
  }

  function formatNumber(text, style) {
    const n = parseNumber(text);
    if (n === null) return text;
    if (style === "thousands") return withSep(n, 0);
    if (style === "rp") {
      const a = Math.abs(n);
      if (a >= 1e9) return (n < 0 ? "−" : "") + "Rp" + withSep(a / 1e9, 1) + "bn";
      if (a >= 1e6) return (n < 0 ? "−" : "") + "Rp" + withSep(a / 1e6, 1) + "m";
      if (a >= 1e3) return (n < 0 ? "−" : "") + "Rp" + withSep(a / 1e3, 0) + "k";
      return (n < 0 ? "−" : "") + "Rp" + withSep(a, 0);
    }
    if (style === "pct") return withSep(Math.abs(n) <= 1 && !/%/.test(text) ? n * 100 : n, 1) + "%";
    return text;
  }

  // ---------- Brand check ----------
  const PALETTE = Object.values(BRAND).map((c) => c.toUpperCase()).concat(["#000000"]);
  function normHex(c) { return c ? ("#" + String(c).replace("#", "")).toUpperCase() : c; }
  function isBrandColor(c) { return !c || PALETTE.includes(normHex(c)); }
  function isBrandFont(f) { return !f || f.startsWith("+") || BRAND_FONTS.includes(f); }
  const LEFTOVER = /\[[^\]]{0,40}\]|lorem ipsum|\bTODO\b|\bXXX\b|\bTBD\b/i;

  /** shapes: [{slideIndex, name, left, top, width, height, text?, font?, size?, color?}] → issues */
  function checkDeck(shapes) {
    const issues = [];
    shapes.forEach((s) => {
      const where = { slideIndex: s.slideIndex, shapeName: s.name };
      if (s.left < -1 || s.top < -1 || s.left + s.width > SLIDE.w + 1 || s.top + s.height > SLIDE.h + 1)
        issues.push(Object.assign({ type: "offslide", msg: "Shape sits partly off the slide" }, where));
      if (s.text === undefined || s.text === null || s.text === "") return;
      if (!isBrandFont(s.font)) issues.push(Object.assign({ type: "font", msg: `Off-brand font: ${s.font}`, fixable: true }, where));
      if (s.color && !isBrandColor(s.color)) issues.push(Object.assign({ type: "color", msg: `Off-brand text color: ${normHex(s.color)}`, fixable: true }, where));
      if (s.size && s.size < 9) issues.push(Object.assign({ type: "size", msg: `Text below 9pt (${s.size}pt)` }, where));
      if (LEFTOVER.test(s.text)) issues.push(Object.assign({ type: "leftover", msg: `Placeholder text left in: “${s.text.match(LEFTOVER)[0]}”` }, where));
    });
    return issues;
  }

  /** Nearest brand color, used by "Fix colors". */
  function nearestBrand(hex) {
    const h = normHex(hex).slice(1);
    const rgb = [0, 2, 4].map((i) => parseInt(h.substr(i, 2), 16));
    let best = BRAND.ink, bd = Infinity;
    Object.values(BRAND).forEach((c) => {
      const v = [1, 3, 5].map((i) => parseInt(c.substr(i, 2), 16));
      const d = v.reduce((a, x, i) => a + (x - rgb[i]) ** 2, 0);
      if (d < bd) { bd = d; best = c; }
    });
    return best;
  }

  // ---------- Element specs ----------
  // Each element is a list of primitives; the renderer turns them into PowerPoint shapes and groups them.
  // prim: {kind:"shape"|"text"|"line", geom, x, y, w, h, fill, line, lineW, dash, text, font:{name,size,bold,italic,color}, align, valign, name}
  const F = (size, color, extra) => Object.assign({ name: "Calibri", size, color }, extra || {});
  const C = BRAND;

  function harvey(pct, x, y, d) {
    const p = [];
    const ring = (fill) => ({ kind: "shape", geom: "Ellipse", x, y, w: d, h: d, fill, line: C.green, lineW: 1.25, name: "ring" });
    if (pct <= 0) p.push(ring(C.white));
    else if (pct >= 100) p.push(ring(C.green));
    else if (pct === 25) { p.push(ring(C.green)); p.push({ kind: "shape", geom: "Pie", x, y, w: d, h: d, fill: C.white, line: null, name: "mask" }); p.push(Object.assign(ring(null), { name: "outline" })); }
    else if (pct === 50) { p.push(ring(C.white)); p.push({ kind: "shape", geom: "FlowChartDelay", x: x + d / 2, y, w: d / 2, h: d, fill: C.green, line: null, name: "half" }); p.push(Object.assign(ring(null), { name: "outline" })); }
    else { p.push(ring(C.white)); p.push({ kind: "shape", geom: "Pie", x, y, w: d, h: d, fill: C.green, line: null, rotation: 270, name: "fill" }); p.push(Object.assign(ring(null), { name: "outline" })); }
    return p;
  }

  function pill(t, fill, color, x, y, w) {
    return [{ kind: "shape", geom: "RoundRectangle", x, y, w, h: 26, fill, line: null, text: t, font: F(11, color, { bold: true }), align: "Center", valign: "Middle", name: "pill" }];
  }

  const ELEMENTS = {
    harvey0: { label: "Harvey ball 0%", group: "Ratings", build: (x, y) => harvey(0, x, y, 28) },
    harvey25: { label: "Harvey ball 25%", group: "Ratings", build: (x, y) => harvey(25, x, y, 28) },
    harvey50: { label: "Harvey ball 50%", group: "Ratings", build: (x, y) => harvey(50, x, y, 28) },
    harvey75: { label: "Harvey ball 75%", group: "Ratings", build: (x, y) => harvey(75, x, y, 28) },
    harvey100: { label: "Harvey ball 100%", group: "Ratings", build: (x, y) => harvey(100, x, y, 28) },
    checkYes: { label: "Check ✓", group: "Ratings", build: (x, y) => [{ kind: "shape", geom: "Ellipse", x, y, w: 28, h: 28, fill: C.green, line: null, text: "✓", font: F(14, C.white, { bold: true }), align: "Center", valign: "Middle", name: "check" }] },
    checkNo: { label: "Cross ✗", group: "Ratings", build: (x, y) => [{ kind: "shape", geom: "Ellipse", x, y, w: 28, h: 28, fill: C.red, line: null, text: "✗", font: F(14, C.white, { bold: true }), align: "Center", valign: "Middle", name: "cross" }] },
    checkPart: { label: "Partly –", group: "Ratings", build: (x, y) => [{ kind: "shape", geom: "Ellipse", x, y, w: 28, h: 28, fill: C.muted, line: null, text: "–", font: F(14, C.white, { bold: true }), align: "Center", valign: "Middle", name: "partly" }] },
    lightGreen: { label: "Light: on track", group: "Ratings", build: (x, y) => [{ kind: "shape", geom: "Ellipse", x, y, w: 18, h: 18, fill: C.green, line: null, name: "light" }] },
    lightAmber: { label: "Light: at risk", group: "Ratings", build: (x, y) => [{ kind: "shape", geom: "Ellipse", x, y, w: 18, h: 18, fill: C.gold, line: null, name: "light" }] },
    lightRed: { label: "Light: off track", group: "Ratings", build: (x, y) => [{ kind: "shape", geom: "Ellipse", x, y, w: 18, h: 18, fill: C.red, line: null, name: "light" }] },
    pillOn: { label: "Pill: On track", group: "Status", build: (x, y) => pill("On track", C.green, C.white, x, y, 94) },
    pillRisk: { label: "Pill: At risk", group: "Status", build: (x, y) => pill("At risk", C.gold, C.green, x, y, 94) },
    pillOff: { label: "Pill: Off track", group: "Status", build: (x, y) => pill("Off track", C.red, C.white, x, y, 94) },
    pillDone: { label: "Pill: Done", group: "Status", build: (x, y) => pill("Done", C.sage, C.green, x, y, 94) },
    marker: { label: "Number marker", group: "Status", params: ["n"], build: (x, y, o) => [{ kind: "shape", geom: "Ellipse", x, y, w: 30, h: 30, fill: C.green, line: null, text: String((o && o.n) || 1), font: F(13, C.white, { bold: true }), align: "Center", valign: "Middle", name: "marker" }] },
    sticker: { label: "Sticker", group: "Callouts", params: ["text"], build: (x, y, o) => { const t = ((o && o.text) || "DRAFT").toUpperCase(); const red = /CONFIDENTIAL/.test(t); return [{ kind: "shape", geom: "RoundRectangle", x, y, w: 36 + t.length * 7.6, h: 27, fill: C.white, line: red ? C.red : C.gold, lineW: 1.25, text: t, font: F(10.5, red ? C.red : C.gold, { bold: true }), align: "Center", valign: "Middle", name: "sticker" }]; } },
    takeaway: { label: "Key takeaway box", group: "Callouts", build: (x, y) => [{ kind: "shape", geom: "RoundRectangle", x, y, w: 547, h: 76, fill: C.green, line: null, text: "So what: one or two sentences with the conclusion and its number.", font: F(13, C.white), align: "Left", valign: "Middle", name: "takeaway" }] },
    note: { label: "Note box", group: "Callouts", build: (x, y) => [{ kind: "shape", geom: "RoundRectangle", x, y, w: 547, h: 43, fill: C.goldLt, line: null, text: "Note: an assumption or caveat the reader must not miss.", font: F(12, C.ink), align: "Left", valign: "Middle", name: "note" }] },
    source: { label: "Source line", group: "Callouts", build: () => [{ kind: "text", x: SLIDE.margin, y: 488, w: 871, h: 14, text: "Source: system or report, period covered. Note on definitions.", font: F(9, C.muted, { italic: true }), align: "Left", valign: "Top", name: "source" }] },
    statDark: { label: "Stat card (dark)", group: "Numbers", build: (x, y) => [
      { kind: "shape", geom: "RoundRectangle", x, y, w: 281, h: 108, fill: C.green, line: null, name: "card" },
      { kind: "text", x: x + 22, y: y + 12, w: 240, h: 44, text: "~Rp109bn", font: F(30, C.gold, { name: "Cambria", bold: true }), align: "Left", valign: "Middle", name: "value" },
      { kind: "text", x: x + 22, y: y + 62, w: 240, h: 20, text: "Headline number", font: F(13, C.white, { bold: true }), align: "Left", valign: "Top", name: "label" },
      { kind: "text", x: x + 22, y: y + 82, w: 240, h: 18, text: "Context line", font: F(11, C.goldLt), align: "Left", valign: "Top", name: "context" }] },
    statLight: { label: "Stat card (light)", group: "Numbers", build: (x, y) => [
      { kind: "shape", geom: "RoundRectangle", x, y, w: 281, h: 108, fill: C.white, line: C.line, lineW: 0.75, name: "card" },
      { kind: "text", x: x + 22, y: y + 12, w: 240, h: 44, text: "Rp180k", font: F(30, C.green, { name: "Cambria", bold: true }), align: "Left", valign: "Middle", name: "value" },
      { kind: "text", x: x + 22, y: y + 62, w: 240, h: 20, text: "Supporting number", font: F(13, C.ink, { bold: true }), align: "Left", valign: "Top", name: "label" },
      { kind: "text", x: x + 22, y: y + 82, w: 240, h: 18, text: "Context line", font: F(11, C.muted), align: "Left", valign: "Top", name: "context" }] },
    process: { label: "Chevron process", group: "Structure", params: ["steps"], build: (x, y, o) => {
      const n = Math.max(2, Math.min(7, (o && o.steps) || 4)), total = SLIDE.w - 2 * SLIDE.margin, w = total / n, p = [];
      for (let i = 0; i < n; i++) p.push({ kind: "shape", geom: "Chevron", x: SLIDE.margin + i * w, y, w: w - 4, h: 46, fill: i === n - 1 ? C.gold : C.green, line: null, text: `Step ${i + 1}`, font: F(14, i === n - 1 ? C.green : C.white, { bold: true }), align: "Center", valign: "Middle", name: `step-${i + 1}` });
      return p; } },
    tracker: { label: "Section tracker", group: "Structure", params: ["steps", "current"], build: (x, y, o) => {
      const n = Math.max(2, Math.min(8, (o && o.steps) || 5)), cur = Math.max(1, Math.min(n, (o && o.current) || 1)), p = [];
      for (let i = 0; i < n; i++) p.push({ kind: "shape", geom: "Chevron", x: x + i * 58, y, w: 60, h: 24, fill: i + 1 === cur ? C.green : C.sage, line: null, name: `tracker-${i + 1}` });
      return p; } },
    matrix: { label: "2x2 matrix", group: "Structure", build: (x, y) => {
      const s = 300, p = [], names = ["Quick wins", "Big bets", "Fill-ins", "Avoid"];
      for (let i = 0; i < 4; i++) { const cx = x + (i % 2) * (s / 2), cy = y + Math.floor(i / 2) * (s / 2); p.push({ kind: "shape", geom: "Rectangle", x: cx, y: cy, w: s / 2 - 3, h: s / 2 - 3, fill: i === 0 ? C.sage : C.white, line: C.line, lineW: 0.75, text: names[i], font: F(12, C.green, { bold: true }), align: "Left", valign: "Top", name: `q${i + 1}` }); }
      p.push({ kind: "text", x, y: y + s + 4, w: s, h: 18, text: "Effort →", font: F(11, C.muted, { bold: true }), align: "Center", valign: "Top", name: "x-axis" });
      p.push({ kind: "text", x: x - 70, y: y + s / 2 - 9, w: 64, h: 18, text: "Impact ↑", font: F(11, C.muted, { bold: true }), align: "Right", valign: "Middle", name: "y-axis" });
      return p; } },
    gantt: { label: "Gantt skeleton (12 months)", group: "Structure", params: ["rows"], build: (x, y, o) => {
      const rows = Math.max(1, Math.min(8, (o && o.rows) || 4)), gx = 245, cw = (SLIDE.w - SLIDE.margin - gx) / 12, p = [];
      const mo = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      mo.forEach((m, i) => { p.push({ kind: "text", x: gx + i * cw, y, w: cw, h: 18, text: m, font: F(10.5, C.muted, { bold: true }), align: "Center", valign: "Top", name: `month-${i + 1}` }); p.push({ kind: "line", x: gx + i * cw, y: y + 22, w: 0, h: rows * 46, line: C.line, lineW: 0.75, name: `grid-${i + 1}` }); });
      for (let r = 0; r < rows; r++) { const ry = y + 34 + r * 46; p.push({ kind: "text", x: SLIDE.margin, y: ry, w: gx - SLIDE.margin - 8, h: 28, text: `Workstream ${r + 1}`, font: F(13, C.ink, { bold: true }), align: "Left", valign: "Middle", name: `row-${r + 1}` }); p.push({ kind: "shape", geom: "RoundRectangle", x: gx + (r * 2 % 10) * cw + 3, y: ry + 5, w: 3 * cw - 6, h: 18, fill: C.green, line: null, name: `bar-${r + 1}` }); }
      return p; } },
  };

  /** Place an element: default position centres it on the slide content area. */
  function buildElement(key, opts) {
    const el = ELEMENTS[key];
    if (!el) throw new Error("Unknown element " + key);
    const probe = el.build(0, 0, opts);
    const b = bounds(probe.map((p) => ({ left: p.x, top: p.y, width: p.w, height: p.h })));
    const x = (opts && opts.x != null) ? opts.x : Math.round((SLIDE.w - b.width) / 2);
    const y = (opts && opts.y != null) ? opts.y : Math.round(150);
    return el.build(x, y, opts);
  }

  // Branded table cell properties for PowerPoint.ShapeCollection.addTable (PowerPointApi 1.8)
  function tableSpec(rows, cols) {
    const values = [], cells = [];
    for (let r = 0; r < rows; r++) {
      const vr = [], cr = [];
      for (let c = 0; c < cols; c++) {
        vr.push(r === 0 ? (c === 0 ? "Item" : `Column ${c}`) : (c === 0 ? `Row ${r}` : ""));
        cr.push({ fill: { color: r === 0 ? BRAND.green : BRAND.white }, font: { name: "Calibri", size: 12, bold: r === 0 || c === 0, color: r === 0 ? BRAND.white : BRAND.ink } });
      }
      values.push(vr); cells.push(cr);
    }
    return { values, specificCellProperties: cells };
  }

  const api = { BRAND, BRAND_FONTS, SLIDE, IN, bounds, align, distribute, stack, matchSize, swap, snap, parseNumber, formatNumber, isBrandColor, isBrandFont, checkDeck, nearestBrand, ELEMENTS, buildElement, tableSpec };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.DeliCore = api;
})(typeof self !== "undefined" ? self : this);
