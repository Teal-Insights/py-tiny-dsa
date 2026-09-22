/**
 * Tiny DSA interactive dependency graph.
 * Browser-side evaluator mirrors tiny_dsa.model.Model / internals for the
 * canonical defaults (validated against Python goldens in scripts/check_graph_eval.py).
 */
(function () {
  "use strict";

  const YEARS = [1, 2, 3, 4, 5];
  const COUNTRIES = ["Borvelia", "Litellia", "Aurelium"];
  const SHOCK_PARAMS = ["Growth", "Interest", "Primary balance"];

  const LAYER_ORIGIN_X = 140;
  const LAYER_GAP_X = 280;
  const LAYER_ROW = 78;
  const LAYER_TOP = 70;

  const DEFAULTS = {
    country_name: "Borvelia",
    country_initial_debt: { Borvelia: 60, Litellia: 80, Aurelium: 40 },
    growth_baseline: { 1: 3.5, 2: 3.5, 3: 3.5, 4: 3.5, 5: 3.5 },
    interest_baseline: { 1: 4, 2: 4, 3: 4, 4: 4, 5: 4 },
    primary_balance_baseline: { 1: -1, 2: -0.5, 3: 0, 4: 0.5, 5: 1 },
    shock_year: 2,
    shock_type: 1,
    shock_magnitudes: { Growth: -2, Interest: 2, "Primary balance": -1 },
  };

  /** Series-level dependency edges (producer → consumer). */
  const SERIES_EDGES = [
    ["country_name", "initial_debt_resolved"],
    ["country_initial_debt", "initial_debt_resolved"],
    ["initial_debt_resolved", "engine_initial_debt_baseline"],
    ["initial_debt_resolved", "engine_initial_debt_shocked"],
    ["shock_type", "shock_magnitude_resolved"],
    ["shock_magnitudes", "shock_magnitude_resolved"],
    ["shock_year", "shock_active"],
    ["growth_baseline", "shocked_growth"],
    ["shock_type", "shocked_growth"],
    ["shock_magnitude_resolved", "shocked_growth"],
    ["shock_active", "shocked_growth"],
    ["interest_baseline", "shocked_interest"],
    ["shock_type", "shocked_interest"],
    ["shock_magnitude_resolved", "shocked_interest"],
    ["shock_active", "shocked_interest"],
    ["primary_balance_baseline", "shocked_primary_balance"],
    ["shock_type", "shocked_primary_balance"],
    ["shock_magnitude_resolved", "shocked_primary_balance"],
    ["shock_active", "shocked_primary_balance"],
    ["engine_initial_debt_baseline", "baseline_path_internal"],
    ["growth_baseline", "baseline_path_internal"],
    ["interest_baseline", "baseline_path_internal"],
    ["primary_balance_baseline", "baseline_path_internal"],
    ["engine_initial_debt_shocked", "shocked_path_internal"],
    ["shocked_growth", "shocked_path_internal"],
    ["shocked_interest", "shocked_path_internal"],
    ["shocked_primary_balance", "shocked_path_internal"],
    ["baseline_path_internal", "output_baseline"],
    ["shocked_path_internal", "output_shocked"],
    ["output_baseline", "output_delta"],
    ["output_shocked", "output_delta"],
  ];

  const SERIES = [
    {
      id: "country_name",
      role: "input",
      label: "country_name",
      sheet: "Inputs",
      kind: "enum",
      address: "Inputs!B5",
      options: COUNTRIES,
      keys: [null],
    },
    {
      id: "country_initial_debt",
      role: "input",
      label: "country_initial_debt",
      sheet: "Inputs",
      kind: "country_map",
      domain: { min: 0, max: 200 },
      keys: COUNTRIES,
      addresses: { Borvelia: "Inputs!B10", Litellia: "Inputs!B11", Aurelium: "Inputs!B12" },
    },
    {
      id: "growth_baseline",
      role: "input",
      label: "growth_baseline",
      sheet: "Inputs",
      kind: "year_map",
      domain: { min: -10, max: 15 },
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Inputs!${"CDEFG"[i]}16`])),
    },
    {
      id: "interest_baseline",
      role: "input",
      label: "interest_baseline",
      sheet: "Inputs",
      kind: "year_map",
      domain: { min: 0, max: 20 },
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Inputs!${"CDEFG"[i]}17`])),
    },
    {
      id: "primary_balance_baseline",
      role: "input",
      label: "primary_balance_baseline",
      sheet: "Inputs",
      kind: "year_map",
      domain: { min: -15, max: 15 },
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Inputs!${"CDEFG"[i]}18`])),
    },
    {
      id: "shock_year",
      role: "input",
      label: "shock_year",
      sheet: "Inputs",
      kind: "int",
      address: "Inputs!B21",
      domain: { min: 1, max: 5 },
      keys: [null],
    },
    {
      id: "shock_type",
      role: "input",
      label: "shock_type",
      sheet: "Inputs",
      kind: "enum_int",
      address: "Inputs!B22",
      options: [1, 2, 3],
      optionLabels: { 1: "1 · growth", 2: "2 · interest", 3: "3 · primary balance" },
      keys: [null],
    },
    {
      id: "shock_magnitudes",
      role: "input",
      label: "shock_magnitudes",
      sheet: "Inputs",
      kind: "shock_map",
      domain: { min: -30, max: 30 },
      keys: SHOCK_PARAMS,
      addresses: {
        Growth: "Inputs!B26",
        Interest: "Inputs!C26",
        "Primary balance": "Inputs!D26",
      },
    },
    {
      id: "initial_debt_resolved",
      role: "internal",
      label: "initial_debt_resolved",
      sheet: "Inputs",
      kind: "scalar",
      address: "Inputs!B6",
      keys: [null],
    },
    {
      id: "engine_initial_debt_baseline",
      role: "internal",
      label: "engine_initial_debt_baseline",
      sheet: "Engine",
      kind: "scalar",
      address: "Engine!B6",
      keys: [null],
    },
    {
      id: "engine_initial_debt_shocked",
      role: "internal",
      label: "engine_initial_debt_shocked",
      sheet: "Engine",
      kind: "scalar",
      address: "Engine!B20",
      keys: [null],
    },
    {
      id: "shock_magnitude_resolved",
      role: "internal",
      label: "shock_magnitude_resolved",
      sheet: "Engine",
      kind: "scalar",
      address: "Engine!B9",
      keys: [null],
    },
    {
      id: "shock_active",
      role: "internal",
      label: "shock_active",
      sheet: "Engine",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Engine!${"CDEFG"[i]}10`])),
    },
    {
      id: "shocked_growth",
      role: "internal",
      label: "shocked_growth",
      sheet: "Engine",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Engine!${"CDEFG"[i]}14`])),
    },
    {
      id: "shocked_interest",
      role: "internal",
      label: "shocked_interest",
      sheet: "Engine",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Engine!${"CDEFG"[i]}15`])),
    },
    {
      id: "shocked_primary_balance",
      role: "internal",
      label: "shocked_primary_balance",
      sheet: "Engine",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Engine!${"CDEFG"[i]}16`])),
    },
    {
      id: "baseline_path_internal",
      role: "internal",
      label: "baseline_path_internal",
      sheet: "Engine",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Engine!${"CDEFG"[i]}6`])),
    },
    {
      id: "shocked_path_internal",
      role: "internal",
      label: "shocked_path_internal",
      sheet: "Engine",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Engine!${"CDEFG"[i]}20`])),
    },
    {
      id: "output_baseline",
      role: "output",
      label: "output_baseline",
      sheet: "Outputs",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Outputs!${"BCDEF"[i]}12`])),
    },
    {
      id: "output_shocked",
      role: "output",
      label: "output_shocked",
      sheet: "Outputs",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Outputs!${"BCDEF"[i]}13`])),
    },
    {
      id: "output_delta",
      role: "output",
      label: "output_delta",
      sheet: "Outputs",
      kind: "year_map",
      keys: YEARS,
      addresses: Object.fromEntries(YEARS.map((y, i) => [y, `Outputs!${"BCDEF"[i]}14`])),
    },
  ];

  const SERIES_BY_ID = Object.fromEntries(SERIES.map((s) => [s.id, s]));

  function cloneDefaults() {
    return {
      country_name: DEFAULTS.country_name,
      country_initial_debt: { ...DEFAULTS.country_initial_debt },
      growth_baseline: { ...DEFAULTS.growth_baseline },
      interest_baseline: { ...DEFAULTS.interest_baseline },
      primary_balance_baseline: { ...DEFAULTS.primary_balance_baseline },
      shock_year: DEFAULTS.shock_year,
      shock_type: DEFAULTS.shock_type,
      shock_magnitudes: { ...DEFAULTS.shock_magnitudes },
    };
  }

  function choose(shockType, a, b, c) {
    if (shockType === 1) return a;
    if (shockType === 2) return b;
    return c;
  }

  function debtStep(prev, r, g, pb) {
    return (prev * (1 + r / 100)) / (1 + g / 100) - pb;
  }

  /** Evaluate the full model; returns map seriesId → scalar | {key: value}. */
  function evaluate(inputs) {
    const initial_debt_resolved = inputs.country_initial_debt[inputs.country_name];
    const engine_initial_debt_baseline = initial_debt_resolved;
    const engine_initial_debt_shocked = initial_debt_resolved;
    const shockKey = SHOCK_PARAMS[inputs.shock_type - 1];
    const shock_magnitude_resolved = inputs.shock_magnitudes[shockKey];

    const shock_active = {};
    const shocked_growth = {};
    const shocked_interest = {};
    const shocked_primary_balance = {};
    const baseline_path_internal = {};
    const shocked_path_internal = {};

    for (const y of YEARS) {
      const active = y >= inputs.shock_year ? 1 : 0;
      shock_active[y] = active;
      shocked_growth[y] =
        inputs.growth_baseline[y] +
        choose(inputs.shock_type, shock_magnitude_resolved, 0, 0) * active;
      shocked_interest[y] =
        inputs.interest_baseline[y] +
        choose(inputs.shock_type, 0, shock_magnitude_resolved, 0) * active;
      shocked_primary_balance[y] =
        inputs.primary_balance_baseline[y] +
        choose(inputs.shock_type, 0, 0, shock_magnitude_resolved) * active;

      if (y === 1) {
        baseline_path_internal[y] = debtStep(
          engine_initial_debt_baseline,
          inputs.interest_baseline[y],
          inputs.growth_baseline[y],
          inputs.primary_balance_baseline[y]
        );
        shocked_path_internal[y] = debtStep(
          engine_initial_debt_shocked,
          shocked_interest[y],
          shocked_growth[y],
          shocked_primary_balance[y]
        );
      } else {
        baseline_path_internal[y] = debtStep(
          baseline_path_internal[y - 1],
          inputs.interest_baseline[y],
          inputs.growth_baseline[y],
          inputs.primary_balance_baseline[y]
        );
        shocked_path_internal[y] = debtStep(
          shocked_path_internal[y - 1],
          shocked_interest[y],
          shocked_growth[y],
          shocked_primary_balance[y]
        );
      }
    }

    const output_baseline = { ...baseline_path_internal };
    const output_shocked = { ...shocked_path_internal };
    const output_delta = {};
    for (const y of YEARS) {
      output_delta[y] = output_shocked[y] - output_baseline[y];
    }

    return {
      country_name: inputs.country_name,
      country_initial_debt: { ...inputs.country_initial_debt },
      growth_baseline: { ...inputs.growth_baseline },
      interest_baseline: { ...inputs.interest_baseline },
      primary_balance_baseline: { ...inputs.primary_balance_baseline },
      shock_year: inputs.shock_year,
      shock_type: inputs.shock_type,
      shock_magnitudes: { ...inputs.shock_magnitudes },
      initial_debt_resolved,
      engine_initial_debt_baseline,
      engine_initial_debt_shocked,
      shock_magnitude_resolved,
      shock_active,
      shocked_growth,
      shocked_interest,
      shocked_primary_balance,
      baseline_path_internal,
      shocked_path_internal,
      output_baseline,
      output_shocked,
      output_delta,
    };
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
    const raw = allValues[series.id];
    if (series.keys.length === 1 && series.keys[0] == null) {
      return formatValue(raw);
    }
    return series.keys.map((key) => formatValue(raw[key])).join(", ");
  }

  function addressText(series) {
    if (series.address) return series.address;
    if (series.addresses) {
      return series.keys.map((key) => series.addresses[key]).join(", ");
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

  /**
   * Longest-path layering on the series DAG.
   * layer(v) = 0 for sources; otherwise 1 + max(layer(u)) over edges u→v.
   * Every edge then goes strictly forward (layer(u) < layer(v)), so no
   * same-column edges — unlike the semantic input/internal/output buckets.
   */
  function computeLayers() {
    const ids = SERIES.map((s) => s.id);
    const preds = Object.fromEntries(ids.map((id) => [id, []]));
    const succs = Object.fromEntries(ids.map((id) => [id, []]));
    for (const [source, target] of SERIES_EDGES) {
      preds[target].push(source);
      succs[source].push(target);
    }

    const layer = Object.fromEntries(ids.map((id) => [id, 0]));
    const indegree = Object.fromEntries(
      ids.map((id) => [id, preds[id].length])
    );
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

  /** Order nodes in a layer by average neighbor layer-index × rank (barycenter). */
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

    // Two sweeps: left→right then right→left, sorting by neighbor barycenters.
    for (let sweep = 0; sweep < 2; sweep += 1) {
      for (let L = 1; L <= maxLayer; L += 1) {
        columns[L].sort((a, b) => {
          const bary = (id) => {
            const neighbors = preds[id];
            if (!neighbors.length) return rank[id];
            return (
              neighbors.reduce((sum, n) => sum + rank[n], 0) / neighbors.length
            );
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
            return (
              neighbors.reduce((sum, n) => sum + rank[n], 0) / neighbors.length
            );
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
      // Cycle fallback: semantic role columns (may retain same-layer edges).
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
    // Refresh HTML value overlays after data changes.
    cyInstance.nodes().forEach((node) => node.trigger("position"));
  }

  let inputs = cloneDefaults();
  let values = evaluate(inputs);
  let selectedId = null;
  let cy = null;
  const positionUndo = [];
  let dragOrigin = null;

  function toast(message) {
    const el = document.getElementById("toast");
    el.textContent = message;
    el.classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.remove("show"), 1800);
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

  function recompute(changedSeriesId) {
    values = evaluate(inputs);
    patchValues(cy, values, changedSeriesId ? downstreamOf(changedSeriesId) : null);
    if (selectedId) renderSide(selectedId);
  }

  function renderSide(nodeId) {
    const panel = document.getElementById("side");
    if (!nodeId) {
      panel.innerHTML =
        '<p class="empty">Select a series node. Amber inputs are editable (comma-separated values). Drag nodes to rearrange; Ctrl+Z undoes a move.</p>';
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
      series.keys[0] == null
        ? "scalar"
        : series.keys.join(", ");

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
              `<option value="${o}" ${o === inputs.shock_type ? "selected" : ""}>${
                series.optionLabels[o]
              }</option>`
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
      apply.addEventListener("click", () => {
        const raw = document.getElementById("edit-value").value;
        if (!commitEdit(series, raw)) return;
        recompute(series.id);
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
        if (!series.options.includes(n)) throw new Error("Invalid shock type");
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
            // Relative to node centre: 50% x = right edge, -50% x = left edge.
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
    cy.fit(undefined, 48);

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

  function resetAll() {
    inputs = cloneDefaults();
    values = evaluate(inputs);
    positionUndo.length = 0;
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
    evaluate,
    DEFAULTS,
    cloneDefaults,
    _goldens: evaluate(cloneDefaults()),
  };

  const cyEl = typeof document !== "undefined" ? document.getElementById("cy") : null;
  if (cyEl && typeof cytoscape === "function") {
    document.getElementById("btn-reset").addEventListener("click", resetAll);
    document.getElementById("btn-fit").addEventListener("click", fitGraph);
    document.addEventListener("keydown", (event) => {
      const key = event.key.toLowerCase();
      if (!(event.ctrlKey || event.metaKey) || key !== "z" || event.shiftKey) return;
      const tag = (event.target && event.target.tagName) || "";
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
      event.preventDefault();
      undoMove();
    });
    try {
      initCy();
      renderSide(null);
    } catch (err) {
      console.error("Tiny DSA graph failed to initialize", err);
      cyEl.innerHTML =
        '<p style="padding:1rem;font:14px system-ui;color:#b91c1c;">Graph failed to load. Check the browser console for details.</p>';
    }
  }
})();
