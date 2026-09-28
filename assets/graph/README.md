# Interactive dependency graph

Fullscreen Cytoscape explorer. The docs homepage shows a **read-only preview**
(`index.html?preview=1`) that links here for editing, pan, and zoom.

| File | Role |
|------|------|
| `index.html` | Shell UI (toolbar, `#cy`, side panel); `?preview=1` hides chrome |
| `config.js` | `TINY_DSA_GRAPH_API` base URL (Railway); empty = same-origin |
| `app.js` | Cytoscape wiring; loads schema/values from the graph API |
| `bootstrap.json` | Static snapshot for offline / docs paint before API responds |
| `style.css` | Layout, role colors, HTML node value typography |

Values are computed by **excel-grapher FormulaEvaluator**. Layout uses
topological layers (longest path).

## Architecture (GitHub Pages + Railway)

GitHub Pages cannot run FormulaEvaluator. Split hosting like this:

| Layer | Where | URL users see |
|-------|--------|----------------|
| Docs + graph UI | GitHub Pages | `https://teal-insights.github.io/py-tiny-dsa/…` |
| FormulaEvaluator API | Railway (or similar) | `https://….up.railway.app` (XHR only) |

The address bar stays on **github.io**. The browser calls Railway for
`/api/graph` and `/api/evaluate` (CORS is open).

### 1. Deploy the API to Railway

From this repo root (Dockerfile + `railway.toml` included):

```bash
# Install Railway CLI, then:
railway login
railway init    # or link an existing project
railway up
railway domain  # copy the public HTTPS URL
```

Health check: `GET /api/health`.

### 2. Point Pages at Railway

In the GitHub repo settings → **Secrets and variables → Actions**, add:

- Variable or secret `GRAPH_API_BASE` = `https://your-service.up.railway.app`
  (no trailing slash)

Docs deploy (`.github/workflows/deploy-docs.yml`) writes that into
`assets/graph/config.js` before publishing Pages.

### 3. Local overrides

```text
?api=https://your-service.up.railway.app
```

or edit `config.js` / the `tiny-dsa-graph-api` meta tag.

## Run locally (API + UI together)

```bash
uv sync --group graph
uv run python scripts/serve_graph_api.py
# http://127.0.0.1:8765/
```

## Static snapshot (docs paint without API)

```bash
uv run python scripts/write_graph_bootstrap.py
```

## Parity check

```bash
uv run python scripts/check_graph_eval.py
```
