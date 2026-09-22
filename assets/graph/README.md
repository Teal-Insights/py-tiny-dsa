# Interactive dependency graph

Fullscreen Cytoscape explorer. The docs homepage shows a **read-only preview**
(`index.html?preview=1`) that links here for editing, pan, and zoom.

| File | Role |
|------|------|
| `index.html` | Shell UI (toolbar, `#cy`, side panel); `?preview=1` hides chrome |
| `app.js` | Series topology, browser-side `Model` mirror, Cytoscape wiring |
| `style.css` | Layout, role colors, HTML node value typography |

Layout uses **topological layers** (longest path). Each series is a single
node; values are a comma-separated list. Drag nodes to rearrange; **Ctrl+Z**
undoes a move.

```bash
python -m http.server 8000 --directory assets/graph
# interactive: http://localhost:8000/
# preview:     http://localhost:8000/?preview=1
```

Keep `app.js` `evaluate()` aligned with `tiny_dsa.model.Model`:

```bash
uv run python scripts/check_graph_eval.py
```
