/* Ribbon buttons that run without opening the task pane (align, distribute, size, swap). */
/* global Office, PowerPoint, DeliCore */
"use strict";

(function () {
  const D = DeliCore;

  async function onSelection(min, apply, event) {
    try {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) throw new Error("needs PowerPointApi 1.5");
      await PowerPoint.run(async (context) => {
        const sel = context.presentation.getSelectedShapes();
        sel.load("items/left,items/top,items/width,items/height");
        await context.sync();
        if (sel.items.length >= min) {
          apply(sel.items, sel.items.map((s) => ({ left: s.left, top: s.top, width: s.width, height: s.height })));
          await context.sync();
        }
      });
    } catch (e) {
      console.error(e);
    }
    event.completed();
  }
  const pos = (items, p) => items.forEach((s, i) => { s.left = p[i].left; s.top = p[i].top; });

  const actions = {
    alignLeft: (e) => onSelection(1, (it, g) => pos(it, D.align(g, "left")), e),
    alignCenter: (e) => onSelection(1, (it, g) => pos(it, D.align(g, "center")), e),
    alignRight: (e) => onSelection(1, (it, g) => pos(it, D.align(g, "right")), e),
    alignTop: (e) => onSelection(1, (it, g) => pos(it, D.align(g, "top")), e),
    alignMiddle: (e) => onSelection(1, (it, g) => pos(it, D.align(g, "middle")), e),
    alignBottom: (e) => onSelection(1, (it, g) => pos(it, D.align(g, "bottom")), e),
    distH: (e) => onSelection(3, (it, g) => pos(it, D.distribute(g, "h")), e),
    distV: (e) => onSelection(3, (it, g) => pos(it, D.distribute(g, "v")), e),
    sameSize: (e) => onSelection(2, (it, g) => D.matchSize(g, "size").forEach((z, i) => { it[i].width = z.width; it[i].height = z.height; }), e),
    swapShapes: (e) => onSelection(2, (it, g) => pos(it.slice(0, 2), D.swap(g[0], g[1])), e),
  };

  Office.onReady(() => {
    Object.entries(actions).forEach(([name, fn]) => Office.actions.associate(name, fn));
  });
})();
