/**
 * Place series nodes from the precomputed starter layout (layout.json).
 *
 * The export pipeline runs excel-grapher's clustered force layout and writes
 * positions in layout units, where `linkDistance` is the spring rest length.
 * placeNodes maps one link distance to the mean node box footprint, then
 * pushes overlapping boxes apart vertically. x is the input -> output flow, so
 * it is never changed. Scaling alone until no boxes overlap would keep the
 * exact shape, but wide label boxes make that several times the viewport.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.SeriesGraphForceLayout = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const MAX_PASSES = 1000;
  const EPS = 1e-6;

  function need(o, a, b) {
    return {
      x: (o.sizes[a].width + o.sizes[b].width) / 2 + o.gapX,
      y: (o.sizes[a].height + o.sizes[b].height) / 2 + o.gapY,
    };
  }

  function baseScale(o, ids) {
    let width = 0;
    let height = 0;
    for (const id of ids) {
      width += o.sizes[id].width / ids.length;
      height += o.sizes[id].height / ids.length;
    }
    return Math.sqrt((width + o.gapX) * (height + o.gapY)) / o.linkDistance;
  }

  /** One sweep over all pairs; returns whether any box moved. */
  function separatePass(o, ids, placed) {
    let moved = false;
    for (let i = 0; i < ids.length; i += 1) {
      for (let j = i + 1; j < ids.length; j += 1) {
        const p = placed[ids[i]];
        const q = placed[ids[j]];
        const n = need(o, ids[i], ids[j]);
        const dy = q.y - p.y;
        const overlapX = n.x - Math.abs(q.x - p.x);
        const overlapY = n.y - Math.abs(dy);
        if (overlapX <= EPS || overlapY <= EPS) continue;
        moved = true;
        // Level boxes split in id order, so the result is deterministic.
        const d = ((dy < 0 ? -1 : 1) * (overlapY + EPS)) / 2;
        p.y -= d;
        q.y += d;
      }
    }
    return moved;
  }

  /** Last resort if passes do not converge: stretch y until every pair clears. */
  function stretchApart(o, ids, placed) {
    let scale = 1;
    for (let i = 0; i < ids.length; i += 1) {
      for (let j = i + 1; j < ids.length; j += 1) {
        const n = need(o, ids[i], ids[j]);
        const dx = Math.abs(placed[ids[i]].x - placed[ids[j]].x);
        const dy = Math.abs(placed[ids[i]].y - placed[ids[j]].y);
        // Passes split level pairs first, so an x-overlapping pair has dy > 0.
        if (dx < n.x) scale = Math.max(scale, n.y / dy);
      }
    }
    for (const id of ids) placed[id].y *= scale;
  }

  /**
   * @param {{positions: Record<string, number[]>,
   *   sizes: Record<string, {width: number, height: number}>,
   *   linkDistance: number, gapX: number, gapY: number}} o
   * @returns {Record<string, {x: number, y: number}> | null} null when the
   *   layout lacks a series in `sizes` (stale layout.json).
   */
  function placeNodes(o) {
    const ids = Object.keys(o.sizes);
    if (!ids.every((id) => Array.isArray(o.positions[id]))) return null;

    const scale = baseScale(o, ids);
    const placed = {};
    for (const id of ids) {
      placed[id] = { x: scale * o.positions[id][0], y: scale * o.positions[id][1] };
    }
    let passes = 0;
    while (passes < MAX_PASSES && separatePass(o, ids, placed)) passes += 1;
    if (passes === MAX_PASSES) stretchApart(o, ids, placed);
    return placed;
  }

  return { placeNodes };
});
