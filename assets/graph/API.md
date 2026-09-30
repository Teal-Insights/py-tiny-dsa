# Series Graph — HTTP API

Version 0.1 (draft). Consumer: [`app.js`](app.js).

The HTTP contract the interactive dependency-graph UI expects. The reference backend is `{package}/graph_api.py` (`bootstrap()`, `evaluate()`, `normalize_inputs()`) served by [`scripts/serve_graph_api.py`](../../scripts/serve_graph_api.py). Any backend that meets this contract works: the local stdlib server, a hosted service, or a static snapshot.

The graph is an author-curated **series DAG**. Nodes are input, internal, and output series; edges run producer → consumer. The server recomputes values from the exported `Model` or the excel-grapher `FormulaEvaluator`. The browser never evaluates formulas.

## 1. Conventions

- Every API response is `application/json; charset=utf-8`.
- Numbers are finite JSON numbers (`json.dumps(..., allow_nan=False)`). Integer-valued floats SHOULD be emitted as integers.
- Payloads are **flat**: a scalar series is one JSON value; a keyed series is an object `key → scalar`. Keys are JSON strings (e.g. `"2024"`); the UI coerces them back using the node's `keys`. No tuple coordinates.
- CORS: send `Access-Control-Allow-Origin: *` on every response and allow
  `GET, POST, OPTIONS` with `Content-Type`. `OPTIONS` returns `204`. Hosted backends MUST do this because docs embeds are cross-origin.
- Clients ignore unknown fields; backends MAY add them.

## 2. Base URL discovery

The configured API base is the first non-empty value of:

1. the `?api=<base>` query parameter;
2. `<meta name="series-graph-api" content="…">` in [`index.html`](index.html);
3. `window.SERIES_GRAPH_API`, set in [`config.js`](config.js).

The UI then fetches the bootstrap from the first candidate that succeeds:

1. `<configured>/api/graph?backend=…` (if a base is configured)
2. `<origin>/api/graph?backend=…`
3. `./api/graph?backend=…`, relative to the page (the base is derived from this URL)
4. `./bootstrap.json` (static snapshot, §6)

Trailing slashes are stripped before joining.

## 3. Endpoints

### `GET /api/health`

```json
{ "ok": true, "backends": ["export", "formula_evaluator"] }
```

`backends` lists the available backends (§4). Safe to poll.

### `GET /api/graph?backend=<name>`

Returns topology, defaults, and evaluated default values in one round trip. `backend` is optional: it defaults to `formula_evaluator` when available, otherwise `export`. Unknown values return `400`.

```json
{
  "axes": { "years": [2024, 2025] },
  "defaults": {
    "mode": 1,
    "rate": { "2024": 4.1, "2025": 4.5 }
  },
  "nodes": [
    {
      "id": "rate",
      "role": "input",
      "label": "rate",
      "sheet": "Inputs",
      "kind": "year_map",
      "keys": [2024, 2025],
      "addresses": { "2024": "Inputs!C10", "2025": "Inputs!D10" },
      "domain": { "min": -10.0, "max": 15.0 }
    }
  ],
  "edges": [["rate", "adjusted_rate"], ["adjusted_rate", "total"]],
  "values": {
    "mode": 1,
    "rate": { "2024": 4.1, "2025": 4.5 },
    "adjusted_rate": { "2024": 4.1, "2025": 4.5 },
    "total": 8.6
  },
  "backend": "formula_evaluator",
  "backends": ["export", "formula_evaluator"]
}
```

| Field | Description |
|-------|-------------|
| `axes` | Named axis domains, display only. May be `{}`. |
| `defaults` | Workbook inputs, flat. Keys are input node ids. Seeds the editor and Reset. |
| `nodes` | One per series, in display order. |
| `edges` | `[producer, consumer]` pairs forming a DAG. |
| `values` | Evaluated series for `defaults`, flat, covering every node. Inputs echo their defaults. |
| `backend` / `backends` | Backend used; backends available. |

All fields are required.

#### Node object

| Field | Notes |
|-------|-------|
| `id` | Unique; also the key in `values` and `defaults`. |
| `role` | `input` (editable) \| `internal` \| `output`. |
| `label`, `sheet` | Display text and workbook sheet. |
| `kind` | Editor/value kind; see below. |
| `keys` | Ordered axis members, or `[null]` for a scalar. |
| `address` \| `addresses` | Exactly one: a single cell, or `key → cell`. Display-only provenance. |
| `domain` | Optional `{min, max}` numeric range. |
| `options`, `optionLabels` | Allowed values and their labels, for enum kinds. |

| `kind` | Editor | Validation |
|--------|--------|------------|
| `scalar` | read-only | — |
| `enum` | select over `options` | value ∈ `options` |
| `enum_int` | select over `options` | numeric value ∈ `options` |
| `int` | number input | integer within `domain` |
| any other kind with `domain` (e.g. `year_map`) | comma-separated text | one finite number per key, within `domain` |

### `POST /api/evaluate`

Recomputes every series. The UI calls it on each edit and on Reset, always sending the full merged input map, never a diff.

```json
{ "backend": "formula_evaluator", "inputs": { "mode": 1, "rate": { "2024": 4.1, "2025": 5.0 } } }
```

- `backend` is optional; when omitted, the server default applies.
- `inputs` is optional; when omitted, the defaults apply. It MUST be an object if present. Missing ids fall back to defaults. Unknown ids return `400` with a per-key `errors` map. The reference backend requires a supplied keyed series to include every key in its default.

```json
{
  "backend": "formula_evaluator",
  "values": { "mode": 1, "rate": { "2024": 4.1, "2025": 5.0 }, "adjusted_rate": { "...": 0 }, "total": 9.1 }
}
```

`values` covers every node: inputs echo the normalized inputs, and the server computes the rest.

## 4. Backends

| Name | Engine | Available when |
|------|--------|----------------|
| `export` | exported `Model` (pure Python) | always |
| `formula_evaluator` | excel-grapher `FormulaEvaluator` over the workbook fixture | `excel-grapher` and the `.xlsx` fixture are present |

An unrecognized name returns `400`; an unavailable backend returns `503`.

## 5. Errors

```json
{ "error": "human-readable message", "errors": { "series_id": "detail" } }
```

`errors` is optional and holds per-id details.

| Status | When |
|--------|------|
| `400` | malformed JSON, non-object body or `inputs`, unknown backend, unknown or invalid inputs |
| `404` | unknown route (`api/` paths never fall through to static files) |
| `500` | unexpected server error |
| `503` | requested backend unavailable |

The UI toasts `error || message || "HTTP <status>"` and does not retry.

## 6. Static snapshot (`bootstrap.json`)

With no live backend, the UI loads `./bootstrap.json`, which is exactly the `GET /api/graph` response body, written by [`scripts/write_graph_bootstrap.py`](../../scripts/write_graph_bootstrap.py). The graph renders read-only, and the first edit reports that live recompute needs the API. A backend MAY serve only this file and skip `/api`.

## 7. Static hosting

The reference server also serves this directory at `/`. It resolves directories to `index.html`, sets MIME types by extension, and blocks path traversal.

## 8. Conformance checklist

1. Every `nodes[].id` is a key in `GET /api/graph` `values`.
2. Every edge endpoint is a known node id.
3. `edges` form a DAG. On a cycle, the UI warns and falls back to role columns.
4. Every `defaults` key is an `input` node, and every `input` node is in `defaults`.
5. For keyed nodes, `values[id]` has every member of `keys`.
6. `POST /api/evaluate` with `{"inputs": defaults}` returns the same `values` as `GET /api/graph` on the same backend.
7. Unknown input ids return `400` with `errors`.
8. `OPTIONS` returns `204` with the CORS headers.
9. No response contains `NaN`, `Infinity`, or tuple-keyed objects.
