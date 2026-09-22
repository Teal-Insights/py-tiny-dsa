/**
 * Tiny DSA interactive dependency graph.
 * Loads series topology from GET /api/graph and recomputes via
 * POST /api/evaluate with backend=formula_evaluator (excel-grapher).
 * Fall back: ./bootstrap.json for static docs preview when the API is offline.
 */
(function () {
  "use strict";

  const BACKEND = "formula_evaluator";
  const LAYER_ORIGIN_X = 140;
  const LAYER_GAP_X = 280;
  const LAYER_ROW = 78;
  const LAYER_TOP = 70;

  const PREVIEW =
    typeof document !== "undefined" &&
    (document.body.classList.contains("preview") ||
      new URLSearchParams(window.location.search).get("preview") === "1" ||
      new URLSearchParams(window.location.search).get("preview") === "true");

  /** @type {any[]} */
  let SERIES = [];
  /** @type {string[][]} */
  let SERIES_EDGES = [];
  /** @type {Record<string, any>} */
  let SERIES_BY_ID = {};
  /** @type {Record<string, any>} */
  let DEFAULTS = {};
  /** @type {Record<string, any>} */
  let inputs = {};
  /** @type {Record<string, any>} */
  let values = {};
  let selectedId = null;
  let cy = null;
  const positionUndo = [];
  let dragOrigin = null;
  let apiBase = "";

  function trimSlash(url) {
    return String(url || "").replace(/\/+$/, "");
  }

  /**
   * Resolve remote FormulaEvaluator API base (e.g. Railway).
   * Order: ?api=… → meta[name=tiny-dsa-graph-api] → window.TINY_DSA_GRAPH_API
   * → config.js TINY_DSA_GRAPH_API. Same-origin /api is tried next by loadBootstrap.
   */
  function configuredApiBase() {
    const params = new URLSearchParams(window.location.search);
    const fromQuery = params.get("api");
    if (fromQuery) return trimSlash(fromQuery);
    const meta = document.querySelector('meta[name="tiny-dsa-graph-api"]');
    if (meta && meta.content) return trimSlash(meta.content);
    if (typeof window.TINY_DSA_GRAPH_API === "string" && window.TINY_DSA_GRAPH_API) {
      return trimSlash(window.TINY_DSA_GRAPH_API);
    }
    return "";
  }

  function apiUrl(path) {
    const base = trimSlash(apiBase);
    const suffix = path.startsWith("/") ? path : `/${path}`;
    return base ? `${base}${suffix}` : suffix;
  }

  function cloneMap(value) {
    if (value && typeof value === "object" && !Array.isArray(value)) {
      return { ...value };
    }
    return value;
  }

  function cloneDefaults() {
    const next = {};
    for (const [key, value] of Object.entries(DEFAULTS)) {
      next[key] = cloneMap(value);
    }
    return next;
  }

  /** Coerce JSON string keys back to series key types (years are ints). */
  function normalizeSeriesValue(series, raw) {
    if (raw == null) return raw;
    if (!series || series.keys.length === 1 && series.keys[0] == null) {
      return raw;
    }
    const out = {};
    for (const key of series.keys) {
      if (Object.prototype.hasOwnProperty.call(raw, key)) {
        out[key] = raw[key];
      } else if (Object.prototype.hasOwnProperty.call(raw, String(key))) {
        out[key] = raw[String(key)];
      }
    }
    return out;
  }

  function normalizeValuesPayload(rawValues) {
    const out = {};
    for (const [id, raw] of Object.entries(rawValues || {})) {
      const series = SERIES_BY_ID[id];
      out[id] = normalizeSeriesValue(series, raw);
    }
    return out;
  }

  function normalizeInputsPayload(rawInputs) {
    const out = {};
    for (const [id, raw] of Object.entries(rawInputs || {})) {
      const series = SERIES_BY_ID[id];
      out[id] = normalizeSeriesValue(series, raw);
    }
    return out;
  }

  async function fetchJson(url, options) {
    const response = await fetch(url, options);
    let payload = null;
    try {
      payload = await response.json();
    } catch (_err) {
      payload = null;
    }
    if (!response.ok) {
      const message =
        (payload && (payload.error || payload.message)) ||
        `HTTP ${response.status}`;
      const err = new Error(message);
      err.status = response.status;
      err.payload = payload;
      throw err;
    }
    return payload;
  }

  async function loadBootstrap() {
    const configured = configuredApiBase();
    const candidates = [];
    if (configured) {
      candidates.push({
        href: `${configured}/api/graph?backend=${encodeURIComponent(BACKEND)}`,
        base: configured,
      });
    }
    candidates.push(
      {
        href: new URL(
          `/api/graph?backend=${encodeURIComponent(BACKEND)}`,
          window.location.origin
        ).href,
        base: window.location.origin,
      },
      {
        href: new URL(
          `./api/graph?backend=${encodeURIComponent(BACKEND)}`,
          window.location.href
        ).href,
        base: null,
      },
      { href: new URL("./bootstrap.json", window.location.href).href, base: "" }
    );
    let lastError = null;
    for (const candidate of candidates) {
      try {
        const data = await fetchJson(candidate.href);
        if (candidate.base === null) {
          const absolute = new URL(candidate.href);
          apiBase = trimSlash(absolute.href.replace(/\/api\/graph\?.*$/, ""));
        } else {
          apiBase = candidate.base;
        }
        return data;
      } catch (err) {
        lastError = err;
      }
    }
    throw lastError || new Error("Could not load graph bootstrap");
  }

  async function evaluateRemote(nextInputs) {
    if (!apiBase && !window.location.pathname.includes("api")) {
      // Static bootstrap-only mode (docs preview): no live recompute.
      throw new Error(
        "Live recompute needs the FormulaEvaluator API. Deploy to Railway and set GRAPH_API_BASE, or run: uv run python scripts/serve_graph_api.py"
      );
    }
    const url = apiUrl(`/api/evaluate`);
    const payload = await fetchJson(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ backend: BACKEND, inputs: nextInputs }),
    });
    return normalizeValuesPayload(payload.values);
  }

  function applyBootstrap(data) {
    SERIES = data.nodes || [];
    SERIES_EDGES = data.edges || [];
    SERIES_BY_ID = Object.fromEntries(SERIES.map((s) => [s.id, s]));
    DEFAULTS = normalizeInputsPayload(data.defaults || {});
    inputs = cloneDefaults();
    values = normalizeValuesPayload(data.values || {});
  }

  function formatValue(value) {
    if (typeof value === "string") return value;
    if (typeof value === "number" && Number.isFinite(value)) {
      if (Number.isInteger(value)) return String(value);
      return value.toFixed(2);
    }
    return String(value);
  }

  function valuesText(series, allValues) {
    const raw = normalizeSeriesValue(series, allValues[series.id]);
    if (series.keys.length === 1 && series.keys[0] == null) {
      return formatValue(raw);
    }
    return series.keys.map((key) => formatValue(raw[key])).join(", ");
  }

  function addressText(series) {
    if (series.address) return series.address;
    if (series.addresses) {
      return series.keys.map((key) => {
        const addresses = series.addresses;
        return addresses[key] ?? addresses[String(key)];
      }).join(", ");
    }
    return "";
  }

  function nodeSize(valuesStr) {
    const width = Math.min(360, Math.max(168, 28 + valuesStr.length * 8.2));
    return { width, height: 64 };
  }

  function buildElements(allValues) {
    const elements = [];
    for (const series of SERIES) {
      const vals = valuesText(series, allValues);
      const size = nodeSize(vals);
      elements.push({
        data: {
          id: series.id,
          name: series.label,
          valuesText: vals,
          label: `${vals}\n${series.label}`,
          role: series.role,
          sheet: series.sheet,
          kind: series.kind,
          address: addressText(series),
          editable: series.role === "input",
          width: size.width,
          height: size.height,
        },
        classes: `series ${series.role}${series.role === "input" ? " editable" : ""}`,
      });
    }
    for (const [source, target] of SERIES_EDGES) {
      elements.push({
        data: { id: `${source}->${target}`, source, target },
        classes: "dep",
      });
    }
    return elements;
  }

  function computeLayers() {
    const ids = SERIES.map((s) => s.id);
    const preds = Object.fromEntries(ids.map((id) => [id, []]));
    const succs = Object.fromEntries(ids.map((id) => [id, []]));
    for (const [source, target] of SERIES_EDGES) {
      preds[target].push(source);
      succs[source].push(target);
    }

    const layer = Object.fromEntries(ids.map((id) => [id, 0]));
    const indegree = Object.fromEntries(ids.map((id) => [id, preds[id].length]));
    const queue = ids.filter((id) => indegree[id] === 0);
    let seen = 0;
    while (queue.length) {
      const u = queue.shift();
      seen += 1;
      for (const v of succs[u]) {
        layer[v] = Math.max(layer[v], layer[u] + 1);
        indegree[v] -= 1;
        if (indegree[v] === 0) queue.push(v);
      }
    }
    if (seen !== ids.length) {
      console.warn("Series graph has a cycle; falling back to role columns");
      return null;
    }
    return layer;
  }

  function orderWithinLayers(layersById) {
    const maxLayer = Math.max(...Object.values(layersById));
    const columns = Array.from({ length: maxLayer + 1 }, () => []);
    for (const series of SERIES) {
      columns[layersById[series.id]].push(series.id);
    }

    const succs = Object.fromEntries(SERIES.map((s) => [s.id, []]));
    const preds = Object.fromEntries(SERIES.map((s) => [s.id, []]));
    for (const [source, target] of SERIES_EDGES) {
      succs[source].push(target);
      preds[target].push(source);
    }

    const rank = {};
    columns.forEach((col) => {
      col.forEach((id, index) => {
        rank[id] = index;
      });
    });

    for (let sweep = 0; sweep < 2; sweep += 1) {
      for (let L = 1; L <= maxLayer; L += 1) {
        columns[L].sort((a, b) => {
          const bary = (id) => {
            const neighbors = preds[id];
            if (!neighbors.length) return rank[id];
            return neighbors.reduce((sum, n) => sum + rank[n], 0) / neighbors.length;
          };
          return bary(a) - bary(b);
        });
        columns[L].forEach((id, index) => {
          rank[id] = index;
        });
      }
      for (let L = maxLayer - 1; L >= 0; L -= 1) {
        columns[L].sort((a, b) => {
          const bary = (id) => {
            const neighbors = succs[id];
            if (!neighbors.length) return rank[id];
            return neighbors.reduce((sum, n) => sum + rank[n], 0) / neighbors.length;
          };
          return bary(a) - bary(b);
        });
        columns[L].forEach((id, index) => {
          rank[id] = index;
        });
      }
    }
    return columns;
  }

  function applyNeuralLayout(cyInstance) {
    const layersById = computeLayers();
    if (!layersById) {
      const columns = { input: [], internal: [], output: [] };
      for (const series of SERIES) columns[series.role].push(series.id);
      const roleX = { input: 140, internal: 520, output: 900 };
      for (const role of ["input", "internal", "output"]) {
        columns[role].forEach((id, index) => {
          cyInstance.$id(id).position({
            x: roleX[role],
            y: LAYER_TOP + index * LAYER_ROW,
          });
        });
      }
      return;
    }

    const columns = orderWithinLayers(layersById);
    columns.forEach((col, layerIndex) => {
      col.forEach((id, rowIndex) => {
        cyInstance.$id(id).position({
          x: LAYER_ORIGIN_X + layerIndex * LAYER_GAP_X,
          y: LAYER_TOP + rowIndex * LAYER_ROW,
        });
      });
    });
  }

  function patchValues(cyInstance, allValues, changedIds) {
    for (const series of SERIES) {
      const node = cyInstance.$id(series.id);
      if (node.empty()) continue;
      const vals = valuesText(series, allValues);
      const prev = node.data("valuesText");
      const size = nodeSize(vals);
      node.data("valuesText", vals);
      node.data("label", `${vals}\n${series.label}`);
      node.data("width", size.width);
      node.data("height", size.height);
      if (changedIds && changedIds.has(series.id) && prev !== vals) {
        node.addClass("changed");
        setTimeout(() => node.removeClass("changed"), 700);
      }
    }
    cyInstance.nodes().forEach((node) => node.trigger("position"));
  }

  function toast(message) {
    const el = document.getElementById("toast");
    el.textContent = message;
    el.classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.remove("show"), 2200);
  }

  function downstreamOf(seriesId) {
    const out = new Set([seriesId]);
    let grew = true;
    while (grew) {
      grew = false;
      for (const [s, t] of SERIES_EDGES) {
        if (out.has(s) && !out.has(t)) {
          out.add(t);
          grew = true;
        }
      }
    }
    return out;
  }

  async function recompute(changedSeriesId) {
    try {
      values = await evaluateRemote(inputs);
      patchValues(cy, values, changedSeriesId ? downstreamOf(changedSeriesId) : null);
      if (selectedId) renderSide(selectedId);
    } catch (err) {
      console.error(err);
      toast(err.message || "Evaluate failed");
    }
  }

  function optionLabel(series, option) {
    const labels = series.optionLabels || {};
    return labels[option] ?? labels[String(option)] ?? String(option);
  }

  function renderSide(nodeId) {
    const panel = document.getElementById("side");
    if (!nodeId) {
      panel.innerHTML =
        '<p class="empty">Select a series node. Amber inputs are editable (comma-separated values). Drag nodes to rearrange; Ctrl+Z undoes a move. Values come from excel-grapher FormulaEvaluator.</p>';
      selectedId = null;
      return;
    }
    const node = cy.$id(nodeId);
    if (node.empty()) {
      renderSide(null);
      return;
    }
    selectedId = nodeId;
    const series = SERIES_BY_ID[nodeId];
    const editable = series.role === "input";
    const vals = valuesText(series, values);
    const keysHint =
      series.keys[0] == null ? "scalar" : series.keys.join(", ");

    let editor = "";
    if (editable) {
      if (series.kind === "enum") {
        editor = `<label for="edit-value">Value</label><select id="edit-value">${series.options
          .map(
            (o) =>
              `<option value="${o}" ${o === inputs.country_name ? "selected" : ""}>${o}</option>`
          )
          .join("")}</select>`;
      } else if (series.kind === "enum_int") {
        editor = `<label for="edit-value">Value</label><select id="edit-value">${series.options
          .map(
            (o) =>
              `<option value="${o}" ${Number(o) === Number(inputs.shock_type) ? "selected" : ""}>${optionLabel(
                series,
                o
              )}</option>`
          )
          .join("")}</select>`;
      } else if (series.kind === "int") {
        editor = `<label for="edit-value">Value (integer ${series.domain.min}–${series.domain.max})</label><input id="edit-value" type="number" step="1" min="${series.domain.min}" max="${series.domain.max}" value="${inputs.shock_year}" />`;
      } else {
        editor = `<label for="edit-value">Values (comma-separated · ${keysHint})</label><input id="edit-value" type="text" value="${vals}" />`;
      }
      editor += `<button class="primary" type="button" id="apply-edit">Apply</button>`;
    } else {
      editor = `<p class="hint">Read-only ${series.role} series. Edit an amber input upstream to change these values.</p>`;
    }

    panel.innerHTML = `
      <h2>${series.label}</h2>
      <div class="meta">${addressText(series) || "—"} · ${series.role} · ${series.sheet}</div>
      <div class="values-display">${vals}</div>
      <div style="margin-top:0.75rem">${editor}</div>
      <p class="hint" style="margin-top:0.85rem">Keys: ${keysHint}. Double-click an input node to focus the editor. Ctrl+Z undoes node moves.</p>
    `;

    const apply = document.getElementById("apply-edit");
    if (apply) {
      apply.addEventListener("click", async () => {
        const raw = document.getElementById("edit-value").value;
        if (!commitEdit(series, raw)) return;
        apply.disabled = true;
        await recompute(series.id);
        apply.disabled = false;
        toast(`Updated ${series.label}`);
      });
    }
  }

  function commitEdit(series, raw) {
    try {
      if (series.kind === "enum") {
        if (!series.options.includes(raw)) throw new Error("Invalid country");
        inputs.country_name = raw;
        return true;
      }
      if (series.kind === "enum_int") {
        const n = Number(raw);
        if (!series.options.map(Number).includes(n)) throw new Error("Invalid shock type");
        inputs.shock_type = n;
        return true;
      }
      if (series.kind === "int") {
        const n = Number(raw);
        if (!Number.isInteger(n) || n < series.domain.min || n > series.domain.max) {
          throw new Error("Out of range");
        }
        inputs.shock_year = n;
        return true;
      }
      const parts = String(raw)
        .split(",")
        .map((part) => part.trim())
        .filter((part) => part.length > 0);
      if (parts.length !== series.keys.length) {
        throw new Error(`Expected ${series.keys.length} values`);
      }
      const next = {};
      series.keys.forEach((key, index) => {
        const n = Number(parts[index]);
        if (!Number.isFinite(n) || n < series.domain.min || n > series.domain.max) {
          throw new Error(`Out of range at ${key}`);
        }
        next[key] = n;
      });
      inputs[series.id] = next;
      return true;
    } catch (err) {
      toast(err.message || "Invalid value");
      return false;
    }
  }

  function undoMove() {
    const last = positionUndo.pop();
    if (!last) {
      toast("Nothing to undo");
      return;
    }
    const node = cy.$id(last.id);
    if (!node.empty()) {
      node.position({ x: last.x, y: last.y });
      toast("Undo move");
    }
  }

  function attachHtmlLabels(cyInstance) {
    if (typeof cyInstance.nodeHtmlLabel !== "function") return;
    cyInstance.nodeHtmlLabel([
      {
        query: "node.series",
        halign: "center",
        valign: "center",
        halignBox: "center",
        valignBox: "center",
        tpl(data) {
          return (
            `<div class="cy-html-node">` +
            `<div class="cy-html-values">${data.valuesText}</div>` +
            `<div class="cy-html-name">${data.name}</div>` +
            `</div>`
          );
        },
      },
    ]);
    cyInstance
      .style()
      .selector("node.series")
      .style({
        label: "",
        "text-opacity": 0,
      })
      .update();
  }

  function initCy() {
    cy = cytoscape({
      container: document.getElementById("cy"),
      elements: buildElements(values),
      autoungrabify: PREVIEW,
      userPanningEnabled: !PREVIEW,
      userZoomingEnabled: !PREVIEW,
      boxSelectionEnabled: false,
      style: [
        {
          selector: "node.series",
          style: {
            shape: "round-rectangle",
            width: "data(width)",
            height: "data(height)",
            label: "data(label)",
            "text-wrap": "wrap",
            "text-max-width": 340,
            "text-valign": "center",
            "text-halign": "center",
            "font-size": 15,
            "font-weight": "bold",
            "border-width": 2,
            color: "#1c1917",
            "background-opacity": 1,
            "z-index": 10,
          },
        },
        {
          selector: "node.series.input",
          style: {
            "background-color": "#fef3c7",
            "border-color": "#d97706",
          },
        },
        {
          selector: "node.series.internal",
          style: {
            "background-color": "#e0e7ff",
            "border-color": "#6366f1",
          },
        },
        {
          selector: "node.series.output",
          style: {
            "background-color": "#d1fae5",
            "border-color": "#059669",
          },
        },
        {
          selector: "node.series.changed",
          style: {
            "border-color": "#f43f5e",
            "border-width": 3,
          },
        },
        {
          selector: "node.series:selected",
          style: {
            "border-width": 3,
            "border-color": "#0f766e",
          },
        },
        {
          selector: "edge.dep",
          style: {
            width: 2,
            "curve-style": "bezier",
            "source-endpoint": "50% 0",
            "target-endpoint": "-50% 0",
            "target-arrow-shape": "triangle",
            "target-arrow-color": "#a8a29e",
            "line-color": "#a8a29e",
            "arrow-scale": 1,
            opacity: 0.85,
            "z-index": 1,
          },
        },
      ],
      layout: { name: "preset" },
      wheelSensitivity: 0.25,
      minZoom: 0.3,
      maxZoom: 2.5,
    });

    applyNeuralLayout(cy);
    try {
      attachHtmlLabels(cy);
    } catch (err) {
      console.warn("HTML labels unavailable; using native labels", err);
    }
    cy.fit(undefined, PREVIEW ? 28 : 48);

    if (PREVIEW) return;

    cy.on("tap", "node.series", (evt) => {
      renderSide(evt.target.id());
    });
    cy.on("tap", (evt) => {
      if (evt.target === cy) renderSide(null);
    });
    cy.on("dbltap", "node.series.editable", (evt) => {
      renderSide(evt.target.id());
      const input = document.getElementById("edit-value");
      if (input) input.focus();
    });

    cy.on("grab", "node.series", (evt) => {
      const node = evt.target;
      dragOrigin = {
        id: node.id(),
        x: node.position("x"),
        y: node.position("y"),
      };
    });
    cy.on("dragfree", "node.series", (evt) => {
      if (!dragOrigin || dragOrigin.id !== evt.target.id()) return;
      const pos = evt.target.position();
      if (pos.x !== dragOrigin.x || pos.y !== dragOrigin.y) {
        positionUndo.push({ ...dragOrigin });
      }
      dragOrigin = null;
    });
  }

  async function resetAll() {
    inputs = cloneDefaults();
    positionUndo.length = 0;
    try {
      values = await evaluateRemote(inputs);
    } catch (err) {
      console.error(err);
      toast(err.message || "Reset evaluate failed");
      return;
    }
    cy.elements().remove();
    cy.add(buildElements(values));
    applyNeuralLayout(cy);
    try {
      attachHtmlLabels(cy);
    } catch (err) {
      console.warn("HTML labels unavailable; using native labels", err);
    }
    cy.fit(undefined, 48);
    renderSide(null);
    toast("Reset to workbook defaults");
  }

  function fitGraph() {
    cy.fit(undefined, 40);
  }

  const root = typeof globalThis !== "undefined" ? globalThis : window;
  root.TinyDsaGraph = {
    backend: BACKEND,
    get SERIES() {
      return SERIES;
    },
    get DEFAULTS() {
      return DEFAULTS;
    },
    cloneDefaults,
    evaluateRemote,
  };

  async function main() {
    const cyEl = typeof document !== "undefined" ? document.getElementById("cy") : null;
    if (!cyEl || typeof cytoscape !== "function") return;

    cyEl.innerHTML =
      '<p style="padding:1rem;font:14px system-ui;color:#57534e;">Loading FormulaEvaluator graph…</p>';

    try {
      const data = await loadBootstrap();
      applyBootstrap(data);
      cyEl.innerHTML = "";
      if (!PREVIEW) {
        document.getElementById("btn-reset").addEventListener("click", () => {
          resetAll();
        });
        document.getElementById("btn-fit").addEventListener("click", fitGraph);
        document.addEventListener("keydown", (event) => {
          const key = event.key.toLowerCase();
          if (!(event.ctrlKey || event.metaKey) || key !== "z" || event.shiftKey) return;
          const tag = (event.target && event.target.tagName) || "";
          if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
          event.preventDefault();
          undoMove();
        });
      }
      initCy();
      if (!PREVIEW) renderSide(null);
      if (!apiBase && !PREVIEW) {
        toast("API offline — edits need serve_graph_api.py");
      }
    } catch (err) {
      console.error("Tiny DSA graph failed to initialize", err);
      cyEl.innerHTML =
        '<p style="padding:1rem;font:14px system-ui;color:#b91c1c;">Graph failed to load. Serve with <code>uv run python scripts/serve_graph_api.py</code> (FormulaEvaluator) or provide <code>bootstrap.json</code>.</p>';
    }
  }

  main();
})();
