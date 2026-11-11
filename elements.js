/*
 * Deli Elements: element definitions + thumbnail renderer (pure JS, no Office.js).
 * Every element returns "primitives" in points (960 x 540 slide):
 *   {kind:"shape", geom, x,y,w,h, fill, alpha, line, lineW, dash, text, font, align, valign, rotation, name}
 *   {kind:"text",  x,y,w,h, text, font, align, valign, name}
 *   {kind:"line",  x,y,w,h, line, lineW, dash, name}
 * geom uses PowerPoint GeometricShapeType names.
 */
(function (root) {
  "use strict";
  const C = { green: "#17362E", gold: "#C9A46A", teal: "#4E7A6C", goldLt: "#E8D6B5", paper: "#FAF8F4", white: "#FFFFFF", ink: "#1F2A27", muted: "#6B7570", line: "#E4DED3", sage: "#DCE5DF", red: "#A4463B", grey: "#C9CCC7" };
  const W = 960, M = 43.2, CW = W - 2 * M;
  const F = (size, color, x) => Object.assign({ name: "Calibri", size, color }, x || {});
  const FH = (size, color, x) => Object.assign({ name: "Cambria", size, color, bold: true }, x || {});
  const S = (geom, x, y, w, h, o) => Object.assign({ kind: "shape", geom, x, y, w, h, name: geom }, o || {});
  const T = (x, y, w, h, text, font, o) => Object.assign({ kind: "text", x, y, w, h, text, font, align: "Left", valign: "Top", name: "text" }, o || {});
  const L = (x, y, w, h, color, weight, o) => Object.assign({ kind: "line", x, y, w, h, line: color || C.muted, lineW: weight || 1, name: "line" }, o || {});
  const lbl = (text, size, color, extra) => ({ text, font: F(size, color, extra), align: "Center", valign: "Middle" });
  const clamp = (v, a, b, d) => { v = Number(v); return isFinite(v) ? Math.max(a, Math.min(b, Math.round(v))) : d; };

  // ---- builders ----
  function harvey(pct, x, y, d) {
    const ring = (fill, name) => S("Ellipse", x, y, d, d, { fill, line: C.green, lineW: 1.25, name });
    if (pct <= 0) return [ring(C.white, "ring")];
    if (pct >= 100) return [ring(C.green, "ring")];
    if (pct === 25) return [ring(C.green, "base"), S("Pie", x, y, d, d, { fill: C.white, line: null, name: "mask" }), ring(null, "outline")];
    if (pct === 50) return [ring(C.white, "base"), S("FlowChartDelay", x + d / 2, y, d / 2, d, { fill: C.green, line: null, name: "half" }), ring(null, "outline")];
    return [ring(C.white, "base"), S("Pie", x, y, d, d, { fill: C.green, line: null, rotation: 270, name: "fill" }), ring(null, "outline")];
  }
  const circleGlyph = (glyph, fill) => (x, y) => [S("Ellipse", x, y, 30, 30, Object.assign({ fill, line: null, name: "badge" }, lbl(glyph, 15, C.white, { bold: true })))];
  const pill = (t, fill, color) => (x, y) => [S("RoundRectangle", x, y, 96, 26, Object.assign({ fill, line: null, name: "pill" }, lbl(t, 11, color, { bold: true })))];

  function traffic(state, vertical) {
    return (x, y) => {
      const on = { green: "#3FA66B", amber: "#E3B04B", red: "#D2503F" };
      const order = ["red", "amber", "green"], d = 20, pad = 6, len = 3 * d + 4 * pad;
      const p = [S("RoundRectangle", x, y, vertical ? d + 2 * pad : len, vertical ? len : d + 2 * pad, { fill: C.ink, line: null, name: "housing" })];
      order.forEach((k, i) => {
        const cx = vertical ? x + pad : x + pad + i * (d + pad), cy = vertical ? y + pad + i * (d + pad) : y + pad;
        p.push(S("Ellipse", cx, cy, d, d, { fill: k === state ? on[k] : "#3A4440", line: null, name: "lamp-" + k }));
      });
      return p;
    };
  }

  function percentBar(pct) {
    return (x, y) => {
      const w = 220, h = 14, f = Math.max(0, Math.min(100, pct)) / 100;
      const p = [S("RoundRectangle", x, y, w, h, { fill: C.sage, line: null, name: "track" })];
      if (f > 0) p.push(S("RoundRectangle", x, y, Math.max(h, w * f), h, { fill: C.green, line: null, name: "fill" }));
      p.push(T(x + w + 8, y - 3, 50, 20, pct + "%", F(13, C.green, { bold: true }), { name: "value" }));
      return p;
    };
  }

  function thermometer(pct) {
    return (x, y) => {
      const tw = 16, th = 110, bulb = 30, f = Math.max(0, Math.min(100, pct)) / 100, cx = x + (bulb - tw) / 2;
      const fillH = (th - 8) * f;
      const p = [S("RoundRectangle", cx, y, tw, th, { fill: C.white, line: C.green, lineW: 1.25, name: "tube" })];
      if (fillH > 0) p.push(S("Rectangle", cx + 4, y + th - fillH, tw - 8, fillH + 4, { fill: C.red, line: null, name: "level" }));
      p.push(S("Ellipse", x, y + th - 6, bulb, bulb, { fill: C.red, line: C.green, lineW: 1.25, name: "bulb" }));
      p.push(T(x + bulb + 8, y + th - 30 - fillH * 0.0, 50, 20, pct + "%", F(13, C.green, { bold: true }), { name: "value" }));
      return p;
    };
  }

  function evaluation(n, star) {
    return (x, y) => {
      const p = [];
      for (let i = 0; i < 5; i++) p.push(S(star ? "Star5" : "Ellipse", x + i * 26, y, 20, 20, { fill: i < n ? (star ? C.gold : C.green) : (star ? C.white : C.sage), line: star ? C.gold : null, lineW: 1, name: "mark-" + (i + 1) }));
      return p;
    };
  }

  function stat(dark) {
    return (x, y) => [
      S("RoundRectangle", x, y, 281, 108, { fill: dark ? C.green : C.white, line: dark ? null : C.line, lineW: 0.75, name: "card" }),
      T(x + 22, y + 12, 240, 44, dark ? "~Rp109bn" : "Rp180k", FH(30, dark ? C.gold : C.green), { valign: "Middle", name: "value" }),
      T(x + 22, y + 62, 240, 20, dark ? "Headline number" : "Supporting number", F(13, dark ? C.white : C.ink, { bold: true }), { name: "label" }),
      T(x + 22, y + 82, 240, 18, "Context line", F(11, dark ? C.goldLt : C.muted), { name: "context" })];
  }

  function kpi(up) {
    return (x, y) => [
      S("RoundRectangle", x, y, 210, 96, { fill: C.white, line: C.line, lineW: 0.75, name: "card" }),
      T(x + 16, y + 10, 180, 18, "Metric name", F(12, C.muted), { name: "label" }),
      T(x + 16, y + 30, 120, 36, "Rp6.6bn", FH(24, C.green), { valign: "Middle", name: "value" }),
      S(up ? "UpArrow" : "DownArrow", x + 140, y + 38, 12, 18, { fill: up ? C.green : C.red, line: null, name: "arrow" }),
      T(x + 156, y + 36, 50, 22, up ? "+8%" : "−4%", F(12, up ? C.green : C.red, { bold: true }), { valign: "Middle", name: "delta" }),
      T(x + 16, y + 70, 180, 16, "vs last month", F(10, C.muted), { name: "basis" })];
  }

  function chevrons(n, home) {
    return (x, y) => { const w = CW / n, p = []; for (let i = 0; i < n; i++) p.push(S(home && i === 0 ? "HomePlate" : home ? "HomePlate" : "Chevron", M + i * w, y, w - 4, 46, Object.assign({ fill: i === n - 1 ? C.gold : C.green, line: null, name: "step-" + (i + 1) }, lbl("Step " + (i + 1), 14, i === n - 1 ? C.green : C.white, { bold: true })))); return p; };
  }

  function numberedSteps(n) {
    return (x, y) => {
      const gap = CW / n, p = [L(M + gap / 2, y + 16, gap * (n - 1), 0, C.line, 2, { name: "rail" })];
      for (let i = 0; i < n; i++) {
        const cx = M + gap * i + gap / 2;
        p.push(S("Ellipse", cx - 16, y, 32, 32, Object.assign({ fill: i === 0 ? C.gold : C.green, line: null, name: "num-" + (i + 1) }, lbl(String(i + 1), 13, i === 0 ? C.green : C.white, { bold: true }))));
        p.push(T(cx - gap / 2 + 6, y + 40, gap - 12, 20, "Step " + (i + 1), F(13, C.green, { bold: true }), { align: "Center", name: "title-" + (i + 1) }));
        p.push(T(cx - gap / 2 + 6, y + 60, gap - 12, 30, "Short description", F(11, C.muted), { align: "Center", name: "desc-" + (i + 1) }));
      }
      return p;
    };
  }

  function verticalFlow(n) {
    return (x, y) => { const p = []; for (let i = 0; i < n; i++) { const yy = y + i * 62; p.push(S("RoundRectangle", x, yy, 260, 42, Object.assign({ fill: i === n - 1 ? C.green : C.white, line: i === n - 1 ? null : C.line, lineW: 0.75, name: "box-" + (i + 1) }, lbl("Step " + (i + 1), 13, i === n - 1 ? C.white : C.green, { bold: true })))); if (i < n - 1) p.push(S("DownArrow", x + 122, yy + 44, 16, 16, { fill: C.gold, line: null, name: "arrow-" + (i + 1) })); } return p; };
  }

  function tracker(n, cur, dots) {
    return (x, y) => {
      const p = [];
      for (let i = 0; i < n; i++) {
        if (dots) { if (i < n - 1) p.push(L(x + i * 40 + 14, y + 7, 40, 0, C.line, 1.5, { name: "link-" + (i + 1) })); p.push(S("Ellipse", x + i * 40, y, 14, 14, { fill: i + 1 < cur ? C.green : i + 1 === cur ? C.gold : C.line, line: null, name: "dot-" + (i + 1) })); }
        else p.push(S("Chevron", x + i * 58, y, 60, 22, { fill: i + 1 === cur ? C.green : C.sage, line: null, name: "tab-" + (i + 1) }));
      }
      return p;
    };
  }

  function agenda(n, cur) {
    return (x, y) => {
      const p = [];
      for (let i = 0; i < n; i++) {
        const yy = y + i * 50, on = i + 1 === cur;
        p.push(S("RoundRectangle", x, yy, 540, 42, { fill: on ? C.green : C.white, line: on ? null : C.line, lineW: 0.75, name: "item-" + (i + 1) }));
        p.push(T(x + 16, yy, 40, 42, String(i + 1).padStart(2, "0"), FH(18, on ? C.gold : C.goldLt), { valign: "Middle", name: "num-" + (i + 1) }));
        p.push(T(x + 66, yy, 450, 42, "Agenda item " + (i + 1), F(15, on ? C.white : C.ink, { bold: on }), { valign: "Middle", name: "text-" + (i + 1) }));
      }
      return p;
    };
  }

  function timeline(n) {
    return (x, y) => {
      const gap = CW / n, p = [L(M, y + 40, CW, 0, C.green, 2, { name: "axis" })];
      for (let i = 0; i < n; i++) {
        const cx = M + gap * i + gap / 2;
        p.push(T(cx - gap / 2, y, gap, 18, "Q" + ((i % 4) + 1) + " 2027", F(11, C.gold, { bold: true }), { align: "Center", name: "date-" + (i + 1) }));
        p.push(S("Diamond", cx - 9, y + 31, 18, 18, { fill: i === n - 1 ? C.gold : C.green, line: null, name: "milestone-" + (i + 1) }));
        p.push(T(cx - gap / 2 + 4, y + 56, gap - 8, 34, "Milestone " + (i + 1), F(12, C.ink), { align: "Center", name: "label-" + (i + 1) }));
      }
      return p;
    };
  }

  function gantt(rows) {
    return (x, y) => {
      const gx = 245, cw = (W - M - gx) / 12, p = [];
      ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].forEach((m, i) => {
        p.push(T(gx + i * cw, y, cw, 16, m, F(10.5, C.muted, { bold: true }), { align: "Center", name: "month-" + (i + 1) }));
        p.push(L(gx + i * cw, y + 20, 0, rows * 40 + 6, C.line, 0.75, { name: "grid-" + (i + 1) }));
      });
      for (let r = 0; r < rows; r++) {
        const ry = y + 28 + r * 40;
        p.push(T(M, ry, gx - M - 8, 26, "Workstream " + (r + 1), F(13, C.ink, { bold: true }), { valign: "Middle", name: "row-" + (r + 1) }));
        const start = (r * 2) % 9;
        p.push(S("RoundRectangle", gx + start * cw + 3, ry + 5, 3 * cw - 6, 16, { fill: r === 1 ? C.gold : C.green, line: null, name: "bar-" + (r + 1) }));
      }
      return p;
    };
  }

  function matrix(n) {
    return (x, y) => {
      const s = n === 2 ? 300 : 330, c = s / n, p = [], names2 = ["Quick wins", "Big bets", "Fill-ins", "Avoid"];
      for (let i = 0; i < n * n; i++) {
        const r = Math.floor(i / n), k = i % n;
        p.push(S("Rectangle", x + k * c, y + r * c, c - 3, c - 3, Object.assign({ fill: (n === 2 && i === 0) || (n === 3 && i === 2) ? C.sage : C.white, line: C.line, lineW: 0.75, name: "cell-" + (i + 1) }, { text: n === 2 ? names2[i] : "", font: F(12, C.green, { bold: true }), align: "Left", valign: "Top" })));
      }
      p.push(T(x, y + s + 4, s, 16, n === 2 ? "Effort →" : "Low  →  High", F(11, C.muted, { bold: true }), { align: "Center", name: "x-axis" }));
      p.push(T(x - 72, y + s / 2 - 8, 66, 16, "Impact ↑", F(11, C.muted, { bold: true }), { align: "Right", name: "y-axis" }));
      return p;
    };
  }

  function pyramid(n) {
    return (x, y) => { const p = [], h = 40, maxW = 380; for (let i = 0; i < n; i++) { const w = maxW * (i + 1) / n, cx = (W - w) / 2; p.push(S("Rectangle", cx, y + i * (h + 4), w, h, Object.assign({ fill: [C.gold, C.green, C.teal, "#2F5247", "#6E8F84"][i], line: null, name: "level-" + (i + 1) }, lbl("Level " + (i + 1), 13, i === 0 ? C.green : C.white, { bold: true })))); } return p; };
  }

  function funnel(n) {
    return (x, y) => { const p = [], h = 38, maxW = 460; for (let i = 0; i < n; i++) { const w = maxW * (1 - i * 0.6 / n), cx = (W - w) / 2; p.push(S("RoundRectangle", cx, y + i * (h + 5), w, h, Object.assign({ fill: i === n - 1 ? C.gold : [C.green, "#2F5247", C.teal, "#6E8F84", "#8FA89F"][i % 5], line: null, name: "stage-" + (i + 1) }, lbl("Stage " + (i + 1) + "  ·  1,000", 13, i === n - 1 ? C.green : C.white, { bold: true })))); } return p; };
  }

  function venn(n) {
    return (x, y) => {
      const d = 170, pos = n === 2 ? [[0, 0], [110, 0]] : [[0, 0], [110, 0], [55, 95]], cols = [C.green, C.gold, C.teal], p = [];
      pos.forEach((q, i) => p.push(S("Ellipse", x + q[0], y + q[1], d, d, { fill: cols[i], alpha: 0.35, line: cols[i], lineW: 1.25, name: "set-" + (i + 1) })));
      pos.forEach((q, i) => p.push(T(x + q[0] + (i === 1 ? 90 : i === 2 ? 45 : 10), y + q[1] + (i === 2 ? 110 : 70), 70, 20, "Set " + "ABC"[i], F(13, C.green, { bold: true }), { align: "Center", name: "label-" + (i + 1) })));
      return p;
    };
  }

  function cycle(n) {
    return (x, y) => {
      const R = 110, cx = x + 140, cy = y + 140, d = 74, p = [S("Ellipse", cx - R, cy - R, 2 * R, 2 * R, { fill: null, line: C.line, lineW: 2, dash: "Dash", name: "ring" })];
      for (let i = 0; i < n; i++) { const a = -Math.PI / 2 + i * 2 * Math.PI / n; p.push(S("Ellipse", cx + R * Math.cos(a) - d / 2, cy + R * Math.sin(a) - d / 2, d, d, Object.assign({ fill: i === 0 ? C.gold : C.green, line: null, name: "node-" + (i + 1) }, lbl("Phase " + (i + 1), 12, i === 0 ? C.green : C.white, { bold: true })))); }
      return p;
    };
  }

  function orgChart(n) {
    return (x, y) => {
      const bw = 150, bh = 46, gap = Math.min(30, (CW - n * bw) / Math.max(1, n - 1)), total = n * bw + (n - 1) * gap, sx = (W - total) / 2, rx = (W - bw) / 2, p = [];
      p.push(S("RoundRectangle", rx, y, bw, bh, Object.assign({ fill: C.green, line: null, name: "head" }, lbl("Head", 13, C.white, { bold: true }))));
      p.push(L(W / 2, y + bh, 0, 18, C.muted, 1, { name: "stem" }));
      if (n > 1) p.push(L(sx + bw / 2, y + bh + 18, total - bw, 0, C.muted, 1, { name: "bus" }));
      for (let i = 0; i < n; i++) { const bx = sx + i * (bw + gap); p.push(L(bx + bw / 2, y + bh + 18, 0, 18, C.muted, 1, { name: "drop-" + (i + 1) })); p.push(S("RoundRectangle", bx, y + bh + 36, bw, bh, Object.assign({ fill: C.white, line: C.line, lineW: 0.75, name: "unit-" + (i + 1) }, lbl("Team " + (i + 1), 13, C.green, { bold: true })))); }
      return p;
    };
  }

  function issueTree() {
    return (x, y) => {
      const bw = 170, bh = 40, p = [], node = (nx, ny, t, root, name) => p.push(S("RoundRectangle", nx, ny, bw, bh, Object.assign({ fill: root ? C.green : C.white, line: root ? null : C.line, lineW: 0.75, name }, lbl(t, 12, root ? C.white : C.green, { bold: true }))));
      const elbow = (px, py, cx, cys, k) => { const mid = px + 18; p.push(L(px, py, 18, 0, C.muted, 1, { name: k + "-a" })); p.push(L(mid, Math.min(...cys), 0, Math.max(...cys) - Math.min(...cys), C.muted, 1, { name: k + "-v" })); cys.forEach((c, i) => p.push(L(mid, c, cx - mid, 0, C.muted, 1, { name: k + "-" + i }))); };
      const x1 = x, x2 = x + bw + 40, x3 = x2 + bw + 40, ys = [0, 56, 112, 168].map((v) => y + v);
      elbow(x1 + bw, y + 84 + bh / 2, x2, [y + 28 + bh / 2, y + 140 + bh / 2], "e1");
      elbow(x2 + bw, y + 28 + bh / 2, x3, [ys[0] + bh / 2, ys[1] + bh / 2], "e2");
      elbow(x2 + bw, y + 140 + bh / 2, x3, [ys[2] + bh / 2, ys[3] + bh / 2], "e3");
      node(x1, y + 84, "Key question", true, "root"); node(x2, y + 28, "Branch A", false, "a"); node(x2, y + 140, "Branch B", false, "b");
      ys.forEach((yy, i) => node(x3, yy, "Driver " + "A1 A2 B1 B2".split(" ")[i], false, "leaf-" + (i + 1)));
      return p;
    };
  }

  function swimlanes(n) {
    return (x, y) => { const p = [], h = 52; for (let i = 0; i < n; i++) { const yy = y + i * (h + 4); p.push(S("Rectangle", M, yy, 130, h, Object.assign({ fill: C.green, line: null, name: "lane-label-" + (i + 1) }, lbl("Team " + (i + 1), 13, C.white, { bold: true })))); p.push(S("Rectangle", M + 134, yy, CW - 134, h, { fill: i % 2 ? C.white : C.sage, line: null, name: "lane-" + (i + 1) })); p.push(S("RoundRectangle", M + 160 + i * 90, yy + 12, 180, 28, Object.assign({ fill: C.white, line: C.green, lineW: 1, name: "task-" + (i + 1) }, lbl("Activity", 12, C.green, { bold: true })))); } return p; };
  }

  function options(n) {
    return (x, y) => { const gap = 14, w = (CW - gap * (n - 1)) / n, p = []; for (let i = 0; i < n; i++) { const rec = i === Math.floor((n - 1) / 2), xx = M + i * (w + gap); p.push(S("RoundRectangle", xx, y, w, 170, { fill: rec ? C.green : C.white, line: rec ? null : C.line, lineW: 0.75, name: "card-" + (i + 1) })); p.push(T(xx + 16, y + 14, w - 32, 22, "Option " + "ABCDE"[i], F(16, rec ? C.white : C.green, { bold: true }), { name: "title-" + (i + 1) })); p.push(T(xx + 16, y + 46, w - 32, 100, "What it is\nCost and time\nMain risk", F(12, rec ? C.goldLt : C.ink), { name: "body-" + (i + 1) })); if (rec) p.push(S("RoundRectangle", xx + w - 104, y + 14, 90, 22, Object.assign({ fill: C.gold, line: null, name: "badge" }, lbl("Recommended", 10, C.green, { bold: true })))); } return p; };
  }

  function columns(n) {
    return (x, y) => { const gap = 14, w = (CW - gap * (n - 1)) / n, p = []; for (let i = 0; i < n; i++) { const xx = M + i * (w + gap); p.push(S("Rectangle", xx, y, w, 34, Object.assign({ fill: C.green, line: null, name: "head-" + (i + 1) }, lbl("Heading " + (i + 1), 13, C.white, { bold: true })))); p.push(S("Rectangle", xx, y + 34, w, 150, Object.assign({ fill: C.white, line: C.line, lineW: 0.75, name: "body-" + (i + 1) }, { text: "Point one\nPoint two\nPoint three", font: F(12, C.ink), align: "Left", valign: "Top" }))); } return p; };
  }

  function numberedList(n) {
    return (x, y) => { const p = []; for (let i = 0; i < n; i++) { const yy = y + i * 44; p.push(S("Ellipse", x, yy, 30, 30, Object.assign({ fill: C.green, line: null, name: "num-" + (i + 1) }, lbl(String(i + 1), 13, C.white, { bold: true })))); p.push(T(x + 42, yy - 2, 460, 18, "Point " + (i + 1) + " headline", F(14, C.green, { bold: true }), { name: "head-" + (i + 1) })); p.push(T(x + 42, yy + 16, 460, 16, "One line of supporting detail", F(11.5, C.muted), { name: "detail-" + (i + 1) })); } return p; };
  }

  // ---- catalogue ----
  // params: [{key,label,type:"number"|"select"|"text",min,max,options,def}]
  const P = (key, label, type, def, extra) => Object.assign({ key, label, type, def }, extra || {});
  const steps = (def, max) => P("steps", "Steps", "number", def, { min: 2, max: max || 7 });
  const E = {};
  const add = (key, group, label, build, params, tags) => { E[key] = { group, label, build, params: params || [], tags: tags || "" }; };

  // Ratings & smart elements
  [0, 25, 50, 75, 100].forEach((p) => add("harvey" + p, "Ratings", `Harvey ball ${p}%`, (x, y) => harvey(p, x, y, 28), null, "harvey ball rating quarter"));
  add("trafficV", "Ratings", "Traffic light (vertical)", (x, y, o) => traffic(o.state, true)(x, y), [P("state", "Lit", "select", "green", { options: ["green", "amber", "red"] })], "traffic light status rag");
  add("trafficH", "Ratings", "Traffic light (horizontal)", (x, y, o) => traffic(o.state, false)(x, y), [P("state", "Lit", "select", "amber", { options: ["green", "amber", "red"] })], "traffic light status rag");
  add("dotGreen", "Ratings", "Status dot: on track", (x, y) => [S("Ellipse", x, y, 18, 18, { fill: C.green, line: null, name: "dot" })], null, "rag dot green");
  add("dotAmber", "Ratings", "Status dot: at risk", (x, y) => [S("Ellipse", x, y, 18, 18, { fill: C.gold, line: null, name: "dot" })], null, "rag dot amber");
  add("dotRed", "Ratings", "Status dot: off track", (x, y) => [S("Ellipse", x, y, 18, 18, { fill: C.red, line: null, name: "dot" })], null, "rag dot red");
  add("percentBar", "Ratings", "Percent bar", (x, y, o) => percentBar(o.pct)(x, y), [P("pct", "Percent", "number", 65, { min: 0, max: 100 })], "progress completion");
  add("thermometer", "Ratings", "Thermometer", (x, y, o) => thermometer(o.pct)(x, y), [P("pct", "Percent", "number", 70, { min: 0, max: 100 })], "level gauge");
  add("evalDots", "Ratings", "Evaluation dots", (x, y, o) => evaluation(o.score, false)(x, y), [P("score", "Score (0–5)", "number", 3, { min: 0, max: 5 })], "score rating");
  add("stars", "Ratings", "Star rating", (x, y, o) => evaluation(o.score, true)(x, y), [P("score", "Stars (0–5)", "number", 4, { min: 0, max: 5 })], "review rating star");
  add("check", "Ratings", "Check", circleGlyph("✓", C.green), null, "yes tick done");
  add("cross", "Ratings", "Cross", circleGlyph("✗", C.red), null, "no x fail");
  add("partly", "Ratings", "Partly", circleGlyph("–", C.muted), null, "partial");
  add("question", "Ratings", "Unknown", circleGlyph("?", C.gold), null, "question open");
  add("trendUp", "Ratings", "Trend up", (x, y) => [S("UpArrow", x, y, 20, 28, { fill: C.green, line: null, name: "up" })], null, "arrow increase");
  add("trendDown", "Ratings", "Trend down", (x, y) => [S("DownArrow", x, y, 20, 28, { fill: C.red, line: null, name: "down" })], null, "arrow decrease");
  add("trendFlat", "Ratings", "Trend flat", (x, y) => [S("RightArrow", x, y, 28, 20, { fill: C.muted, line: null, name: "flat" })], null, "arrow stable");

  // Status & labels
  add("pillOn", "Status", "Pill: On track", pill("On track", C.green, C.white), null, "status");
  add("pillRisk", "Status", "Pill: At risk", pill("At risk", C.gold, C.green), null, "status");
  add("pillOff", "Status", "Pill: Off track", pill("Off track", C.red, C.white), null, "status");
  add("pillDone", "Status", "Pill: Done", pill("Done", C.sage, C.green), null, "status complete");
  add("marker", "Status", "Number marker", (x, y, o) => [S("Ellipse", x, y, 30, 30, Object.assign({ fill: C.green, line: null, name: "marker" }, lbl(String(o.n), 13, C.white, { bold: true })))], [P("n", "Number", "number", 1, { min: 1, max: 99 })], "number badge");
  add("letter", "Status", "Letter marker", (x, y, o) => [S("Ellipse", x, y, 30, 30, Object.assign({ fill: C.gold, line: null, name: "marker" }, lbl(String(o.letter || "A").slice(0, 2).toUpperCase(), 13, C.green, { bold: true })))], [P("letter", "Letter", "text", "A")], "letter badge");
  add("sticker", "Status", "Sticker", (x, y, o) => { const t = String(o.text || "DRAFT").toUpperCase(), red = /CONFIDENTIAL/.test(t); return [S("RoundRectangle", x, y, 36 + t.length * 7.6, 27, Object.assign({ fill: C.white, line: red ? C.red : C.gold, lineW: 1.25, name: "sticker" }, lbl(t, 10.5, red ? C.red : C.gold, { bold: true })))]; }, [P("text", "Text", "select", "DRAFT", { options: ["DRAFT", "FOR DISCUSSION", "PRELIMINARY", "ILLUSTRATIVE", "CONFIDENTIAL", "BACKUP"] })], "draft confidential stamp");
  add("newBadge", "Status", "NEW badge", (x, y) => [S("Star5", x, y, 46, 46, Object.assign({ fill: C.gold, line: null, name: "burst" }, lbl("NEW", 9, C.green, { bold: true })))], null, "new star");
  add("tag", "Status", "Label tag", (x, y, o) => [S("Rectangle", x, y, 24 + String(o.text).length * 7, 22, Object.assign({ fill: C.sage, line: null, name: "tag" }, lbl(String(o.text), 11, C.green, { bold: true })))], [P("text", "Text", "text", "Label")], "tag chip");

  // Callouts
  add("takeaway", "Callouts", "Key takeaway", (x, y) => [S("RoundRectangle", M, y, CW, 64, { fill: C.green, line: null, text: "So what: one or two sentences with the conclusion and its number.", font: F(14, C.white), align: "Left", valign: "Middle", name: "takeaway" })], null, "so what message conclusion");
  add("note", "Callouts", "Note", (x, y) => [S("RoundRectangle", M, y, CW, 40, { fill: C.goldLt, line: null, text: "Note: an assumption or caveat the reader must not miss.", font: F(12, C.ink), align: "Left", valign: "Middle", name: "note" })], null, "caveat assumption");
  add("source", "Callouts", "Source line", () => [T(M, 488, CW, 14, "Source: system or report, period covered. Note on definitions.", F(9, C.muted, { italic: true }), { name: "source" })], null, "footnote source");
  add("footnote", "Callouts", "Footnote", () => [T(M, 472, CW, 14, "1. Footnote text explaining a figure on this slide.", F(9, C.muted), { name: "footnote" })], null, "footnote");
  add("bubble", "Callouts", "Speech bubble", (x, y) => [S("WedgeRoundRectCallout", x, y, 220, 70, { fill: C.white, line: C.green, lineW: 1.25, text: "Comment or quote", font: F(13, C.green), align: "Center", valign: "Middle", name: "bubble" })], null, "comment callout");
  add("quote", "Callouts", "Quote card", (x, y) => [S("RoundRectangle", x, y, 300, 130, { fill: C.white, line: C.line, lineW: 0.75, name: "card" }), T(x + 16, y + 4, 40, 44, "“", FH(44, C.gold), { name: "mark" }), T(x + 20, y + 46, 260, 44, "A guest or team quote that proves the point.", F(14, C.green, { name: "Cambria", italic: true }), { name: "quote" }), T(x + 20, y + 100, 260, 16, "— Name, role", F(11, C.muted), { name: "by" })], null, "testimonial");
  add("highlight", "Callouts", "Highlight frame", (x, y) => [S("RoundRectangle", x, y, 240, 120, { fill: null, line: C.red, lineW: 2, dash: "Dash", name: "frame" })], null, "focus circle box");
  add("braceRight", "Callouts", "Brace with label", (x, y, o) => [S("RightBrace", x, y, 18, 120, { fill: null, line: C.green, lineW: 1.5, name: "brace" }), T(x + 26, y + 46, 150, 28, o.text, F(13, C.green, { bold: true }), { valign: "Middle", name: "label" })], [P("text", "Label", "text", "Total: Rp109bn")], "bracket sum");
  add("arrowLabel", "Callouts", "Arrow with text", (x, y) => [S("RightArrow", x, y, 220, 56, { fill: C.gold, line: null, text: "Implication", font: F(13, C.green, { bold: true }), align: "Center", valign: "Middle", name: "arrow" })], null, "arrow so what");

  // Numbers
  add("statDark", "Numbers", "Stat card (dark)", stat(true), null, "number kpi headline");
  add("statLight", "Numbers", "Stat card (light)", stat(false), null, "number kpi");
  add("kpiUp", "Numbers", "KPI tile (up)", kpi(true), null, "metric delta increase");
  add("kpiDown", "Numbers", "KPI tile (down)", kpi(false), null, "metric delta decrease");
  add("bigNumber", "Numbers", "Big number", (x, y) => [T(x, y, 300, 70, "43%", FH(60, C.green), { name: "number" }), T(x, y + 72, 300, 20, "of October group sales", F(13, C.muted), { name: "caption" })], null, "hero number");

  // Text & boxes
  add("card", "Boxes", "Card with heading", (x, y) => [S("RoundRectangle", x, y, 280, 150, { fill: C.white, line: C.line, lineW: 0.75, name: "card" }), T(x + 18, y + 16, 244, 22, "Card heading", F(15, C.green, { bold: true }), { name: "heading" }), T(x + 18, y + 46, 244, 90, "Two or three lines of supporting text.", F(12, C.ink), { name: "body" })], null, "box text");
  add("darkCard", "Boxes", "Highlight card", (x, y) => [S("RoundRectangle", x, y, 280, 150, { fill: C.green, line: null, name: "card" }), T(x + 18, y + 16, 244, 22, "Card heading", F(15, C.gold, { bold: true }), { name: "heading" }), T(x + 18, y + 46, 244, 90, "Two or three lines of supporting text.", F(12, C.white), { name: "body" })], null, "box text dark");
  add("numberedList", "Boxes", "Numbered list", (x, y, o) => numberedList(o.rows)(x, y), [P("rows", "Rows", "number", 4, { min: 2, max: 7 })], "list points");
  add("columns", "Boxes", "Columns with headers", (x, y, o) => columns(o.cols)(x, y), [P("cols", "Columns", "number", 3, { min: 2, max: 5 })], "columns pillars");
  add("prosCons", "Boxes", "Pros and cons", (x, y) => [S("RoundRectangle", M, y, CW / 2 - 7, 170, { fill: C.white, line: C.line, lineW: 0.75, name: "pros" }), T(M + 18, y + 14, 300, 22, "✓  Pros", F(15, C.green, { bold: true }), { name: "pros-head" }), T(M + 18, y + 44, 380, 110, "Point one\nPoint two", F(12, C.ink), { name: "pros-body" }), S("RoundRectangle", M + CW / 2 + 7, y, CW / 2 - 7, 170, { fill: C.white, line: C.line, lineW: 0.75, name: "cons" }), T(M + CW / 2 + 25, y + 14, 300, 22, "✗  Cons", F(15, C.red, { bold: true }), { name: "cons-head" }), T(M + CW / 2 + 25, y + 44, 380, 110, "Point one\nPoint two", F(12, C.ink), { name: "cons-body" })], null, "comparison");
  add("beforeAfter", "Boxes", "Before → after", (x, y) => [S("RoundRectangle", M, y, 400, 120, { fill: C.white, line: C.line, lineW: 0.75, text: "BEFORE\nState in one or two lines", font: F(13, C.ink), align: "Left", valign: "Top", name: "before" }), S("RightArrow", M + 410, y + 44, 50, 32, { fill: C.gold, line: null, name: "arrow" }), S("RoundRectangle", M + 470, y, CW - 470, 120, { fill: C.green, line: null, text: "AFTER\nState in one or two lines", font: F(13, C.white), align: "Left", valign: "Top", name: "after" })], null, "change from to");
  add("options", "Boxes", "Option cards", (x, y, o) => options(o.n)(x, y), [P("n", "Options", "number", 3, { min: 2, max: 4 })], "alternatives recommendation");

  // Structure & diagrams
  add("chevrons", "Diagrams", "Chevron process", (x, y, o) => chevrons(o.steps, false)(x, y), [steps(4)], "process flow steps");
  add("arrows", "Diagrams", "Arrow process", (x, y, o) => chevrons(o.steps, true)(x, y), [steps(4)], "process pentagon");
  add("numberedSteps", "Diagrams", "Numbered steps", (x, y, o) => numberedSteps(o.steps)(x, y), [steps(4, 6)], "process stages");
  add("verticalFlow", "Diagrams", "Vertical flow", (x, y, o) => verticalFlow(o.steps)(x, y), [steps(3, 5)], "flow down");
  add("trackerTabs", "Diagrams", "Section tracker", (x, y, o) => tracker(o.steps, o.current, false)(x, y), [steps(5, 8), P("current", "Current", "number", 2, { min: 1, max: 8 })], "agenda progress");
  add("trackerDots", "Diagrams", "Progress dots", (x, y, o) => tracker(o.steps, o.current, true)(x, y), [steps(5, 10), P("current", "Current", "number", 3, { min: 1, max: 10 })], "progress");
  add("agenda", "Diagrams", "Agenda list", (x, y, o) => agenda(o.items, o.current)(x, y), [P("items", "Items", "number", 5, { min: 2, max: 7 }), P("current", "Highlight", "number", 1, { min: 0, max: 7 })], "agenda contents");
  add("timeline", "Diagrams", "Timeline", (x, y, o) => timeline(o.steps)(x, y), [P("steps", "Milestones", "number", 5, { min: 2, max: 8 })], "milestones roadmap");
  add("gantt", "Diagrams", "Gantt (12 months)", (x, y, o) => gantt(o.rows)(x, y), [P("rows", "Rows", "number", 5, { min: 1, max: 8 })], "project plan schedule");
  add("matrix2", "Diagrams", "2x2 matrix", (x, y) => matrix(2)(x, y), null, "quadrant prioritisation");
  add("matrix3", "Diagrams", "3x3 matrix", (x, y) => matrix(3)(x, y), null, "nine box grid");
  add("pyramid", "Diagrams", "Pyramid", (x, y, o) => pyramid(o.levels)(x, y), [P("levels", "Levels", "number", 3, { min: 2, max: 5 })], "hierarchy");
  add("funnel", "Diagrams", "Funnel", (x, y, o) => funnel(o.stages)(x, y), [P("stages", "Stages", "number", 4, { min: 2, max: 6 })], "conversion pipeline");
  add("venn2", "Diagrams", "Venn (2)", (x, y) => venn(2)(x, y), null, "overlap");
  add("venn3", "Diagrams", "Venn (3)", (x, y) => venn(3)(x, y), null, "overlap");
  add("cycle", "Diagrams", "Cycle", (x, y, o) => cycle(o.n)(x, y), [P("n", "Phases", "number", 4, { min: 3, max: 6 })], "loop circular");
  add("orgChart", "Diagrams", "Org chart", (x, y, o) => orgChart(o.n)(x, y), [P("n", "Teams", "number", 4, { min: 1, max: 5 })], "organisation structure");
  add("issueTree", "Diagrams", "Issue / driver tree", (x, y) => issueTree()(x, y), null, "logic tree mece");
  add("swimlanes", "Diagrams", "Swimlanes", (x, y, o) => swimlanes(o.lanes)(x, y), [P("lanes", "Lanes", "number", 3, { min: 2, max: 6 })], "process roles");

  // Placeholders
  add("imageBox", "Placeholders", "Image placeholder", (x, y) => [S("Rectangle", x, y, 280, 170, { fill: C.sage, line: null, text: "Image", font: F(13, C.muted), align: "Center", valign: "Middle", name: "image" })], null, "photo picture");
  add("chartBox", "Placeholders", "Chart placeholder", (x, y) => [S("Rectangle", x, y, 400, 220, { fill: C.white, line: C.line, lineW: 0.75, dash: "Dash", text: "Chart goes here", font: F(13, C.muted), align: "Center", valign: "Middle", name: "chart" })], null, "graph");

  function options_(key, o) { const d = {}; (E[key].params || []).forEach((p) => { let v = o && o[p.key] != null && o[p.key] !== "" ? o[p.key] : p.def; if (p.type === "number") v = clamp(v, p.min, p.max, p.def); d[p.key] = v; }); return d; }

  function bounds(prims) {
    const l = Math.min(...prims.map((p) => p.x)), t = Math.min(...prims.map((p) => p.y));
    const r = Math.max(...prims.map((p) => p.x + p.w)), b = Math.max(...prims.map((p) => p.y + p.h));
    return { x: l, y: t, w: r - l, h: b - t };
  }

  /** Build an element positioned on the slide (centered horizontally, near the top of the content area). */
  function build(key, o) {
    const el = E[key];
    if (!el) throw new Error("Unknown element " + key);
    const opts = options_(key, o);
    const probe = el.build(0, 0, opts), b = bounds(probe);
    const x = o && o.x != null ? o.x : Math.round((W - b.w) / 2) - b.x;
    const y = o && o.y != null ? o.y : Math.max(140, Math.round((540 - b.h) / 2)) - b.y;
    return el.build(x, y, opts);
  }

  // ---------- Thumbnail renderer (SVG) ----------
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  function geomSVG(p) {
    const { x, y, w, h } = p, cx = x + w / 2, cy = y + h / 2, rx = Math.min(w, h) * 0.17;
    switch (p.geom) {
      case "Ellipse": return `<ellipse cx="${cx}" cy="${cy}" rx="${w / 2}" ry="${h / 2}"/>`;
      case "RoundRectangle": return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${rx}"/>`;
      case "Chevron": { const n = Math.min(h / 2, w / 2); return `<polygon points="${x},${y} ${x + w - n},${y} ${x + w},${cy} ${x + w - n},${y + h} ${x},${y + h} ${x + n},${cy}"/>`; }
      case "HomePlate": { const n = Math.min(h / 2, w / 2); return `<polygon points="${x},${y} ${x + w - n},${y} ${x + w},${cy} ${x + w - n},${y + h} ${x},${y + h}"/>`; }
      case "Diamond": return `<polygon points="${cx},${y} ${x + w},${cy} ${cx},${y + h} ${x},${cy}"/>`;
      case "Pie": return `<path transform="rotate(${p.rotation || 0} ${cx} ${cy})" d="M${cx} ${cy}L${x + w} ${cy}A${w / 2} ${h / 2} 0 1 1 ${cx} ${y}Z"/>`;
      case "FlowChartDelay": return `<path d="M${x} ${y}H${x + w - h / 2}A${h / 2} ${h / 2} 0 0 1 ${x + w - h / 2} ${y + h}H${x}Z"/>`;
      case "Star5": { const pts = []; for (let i = 0; i < 10; i++) { const r = i % 2 ? 0.4 : 1, a = -Math.PI / 2 + i * Math.PI / 5; pts.push(`${cx + Math.cos(a) * r * w / 2},${cy + Math.sin(a) * r * h / 2}`); } return `<polygon points="${pts.join(" ")}"/>`; }
      case "UpArrow": return `<polygon points="${cx},${y} ${x + w},${y + h * 0.45} ${x + w * 0.72},${y + h * 0.45} ${x + w * 0.72},${y + h} ${x + w * 0.28},${y + h} ${x + w * 0.28},${y + h * 0.45} ${x},${y + h * 0.45}"/>`;
      case "DownArrow": return `<polygon points="${cx},${y + h} ${x + w},${y + h * 0.55} ${x + w * 0.72},${y + h * 0.55} ${x + w * 0.72},${y} ${x + w * 0.28},${y} ${x + w * 0.28},${y + h * 0.55} ${x},${y + h * 0.55}"/>`;
      case "RightArrow": { const hw = Math.min(w * 0.45, h); return `<polygon points="${x},${y + h * 0.25} ${x + w - hw},${y + h * 0.25} ${x + w - hw},${y} ${x + w},${cy} ${x + w - hw},${y + h} ${x + w - hw},${y + h * 0.75} ${x},${y + h * 0.75}"/>`; }
      case "WedgeRoundRectCallout": return `<path d="M${x + rx} ${y}H${x + w - rx}Q${x + w} ${y} ${x + w} ${y + rx}V${y + h - rx}Q${x + w} ${y + h} ${x + w - rx} ${y + h}H${x + w * 0.4}L${x + w * 0.2} ${y + h * 1.3}L${x + w * 0.25} ${y + h}H${x + rx}Q${x} ${y + h} ${x} ${y + h - rx}V${y + rx}Q${x} ${y} ${x + rx} ${y}Z"/>`;
      case "RightBrace": return `<path d="M${x} ${y}Q${cx} ${y} ${cx} ${y + h * 0.1}V${cy - h * 0.05}Q${cx} ${cy} ${x + w} ${cy}Q${cx} ${cy} ${cx} ${cy + h * 0.05}V${y + h * 0.9}Q${cx} ${y + h} ${x} ${y + h}"/>`;
      default: return `<rect x="${x}" y="${y}" width="${w}" height="${h}"/>`;
    }
  }
  function primSVG(p, scale) {
    const minPx = 3.2;
    const textSVG = () => {
      if (!p.text) return "";
      const fs = (p.font && p.font.size || 12), col = (p.font && p.font.color) || C.ink, lines = String(p.text).split("\n");
      if (fs * scale < minPx) { // too small to read: draw grey text bars
        const bw = Math.min(p.w * 0.7, String(lines[0]).length * fs * 0.5), bx = p.align === "Center" ? p.x + (p.w - bw) / 2 : p.x + (p.kind === "text" ? 0 : 8);
        return lines.slice(0, 3).map((_, i) => `<rect x="${bx}" y="${p.y + (p.valign === "Middle" ? p.h / 2 - fs * 0.3 : fs * 0.4) + i * fs * 1.3}" width="${bw * (i ? 0.7 : 1)}" height="${fs * 0.6}" rx="${fs * 0.3}" fill="${col}" opacity="0.55"/>`).join("");
      }
      const anchor = p.align === "Center" ? "middle" : p.align === "Right" ? "end" : "start";
      const tx = p.align === "Center" ? p.x + p.w / 2 : p.align === "Right" ? p.x + p.w : p.x + (p.kind === "text" ? 0 : 8);
      const ty = p.valign === "Middle" ? p.y + p.h / 2 + fs * 0.35 : p.y + fs;
      return `<text x="${tx}" y="${ty}" font-size="${fs}" font-family="${p.font && p.font.name === "Cambria" ? "Cambria,Georgia,serif" : "Calibri,Segoe UI,sans-serif"}" font-weight="${p.font && p.font.bold ? 700 : 400}" fill="${col}" text-anchor="${anchor}">${esc(lines[0])}</text>`;
    };
    if (p.kind === "line") return `<line x1="${p.x}" y1="${p.y}" x2="${p.x + p.w}" y2="${p.y + p.h}" stroke="${p.line}" stroke-width="${p.lineW || 1}"${p.dash ? ' stroke-dasharray="4 3"' : ""}/>`;
    if (p.kind === "text") return textSVG();
    const fill = p.fill || "none", stroke = p.line || "none";
    const g = geomSVG(p).replace(/^<(\w+)/, `<$1 fill="${p.geom === "RightBrace" ? "none" : fill}"${p.alpha ? ` fill-opacity="${p.alpha}"` : ""} stroke="${stroke}" stroke-width="${p.lineW || 1}"${p.dash ? ' stroke-dasharray="6 4"' : ""}`);
    return g + textSVG();
  }
  /** SVG thumbnail of an element, fitted into a box of the given pixel size. */
  function thumbnail(key, o, boxW, boxH) {
    const prims = build(key, o || {}), b = bounds(prims), pad = Math.max(b.w, b.h) * 0.08;
    const vb = [b.x - pad, b.y - pad, b.w + 2 * pad, b.h + 2 * pad];
    const scale = Math.min((boxW || 64) / vb[2], (boxH || 40) / vb[3]);
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${vb.join(" ")}" width="${boxW || 64}" height="${boxH || 40}" preserveAspectRatio="xMidYMid meet">${prims.map((p) => primSVG(p, scale)).join("")}</svg>`;
  }

  const api = { ELEMENTS: E, build, bounds, thumbnail, options: options_, COLORS: C };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.DeliElements = api;
})(typeof self !== "undefined" ? self : this);
