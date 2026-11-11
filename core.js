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
  // anchor: "selection" (default), "first" (first selected shape) or "slide". One shape always aligns to the slide.
  function align(shapes, mode, anchor) {
    const slide = { left: 0, top: 0, right: SLIDE.w, bottom: SLIDE.h, width: SLIDE.w, height: SLIDE.h };
    const f = shapes[0];
    const ref = shapes.length === 1 || anchor === "slide" ? slide
      : anchor === "first" ? { left: f.left, top: f.top, right: f.left + f.width, bottom: f.top + f.height, width: f.width, height: f.height }
      : bounds(shapes);
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

  /** Arrange shapes in a grid of `cols` columns (reading order by position), keeping the first shape's corner. */
  function grid(shapes, cols, gap) {
    const order = shapes.map((s, i) => i).sort((a, b) => (shapes[a].top - shapes[b].top) || (shapes[a].left - shapes[b].left));
    const cw = Math.max(...shapes.map((s) => s.width)), rh = Math.max(...shapes.map((s) => s.height));
    const x0 = Math.min(...shapes.map((s) => s.left)), y0 = Math.min(...shapes.map((s) => s.top));
    const out = shapes.map((s) => ({ left: s.left, top: s.top }));
    order.forEach((i, k) => { out[i] = { left: x0 + (k % cols) * (cw + gap), top: y0 + Math.floor(k / cols) * (rh + gap) }; });
    return out;
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

  const api = { BRAND, BRAND_FONTS, SLIDE, IN, bounds, align, distribute, stack, grid, matchSize, swap, snap, parseNumber, formatNumber, isBrandColor, isBrandFont, checkDeck, nearestBrand, tableSpec };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.DeliCore = api;
})(typeof self !== "undefined" ? self : this);
