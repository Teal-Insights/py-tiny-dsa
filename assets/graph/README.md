# Interactive dependency graph

Static Cytoscape explorer embedded on the docs homepage (`user_guide/index.qmd`).

| File | Role |
|------|------|
| `index.html` | Shell UI (toolbar, `#cy`, side panel) |
| `app.js` | Series topology, browser-side `Model` mirror, Cytoscape wiring |
| `style.css` | Layout, role colors, HTML node value typography |

Layout is a fixed three-column neural net: **inputs → internals → outputs**.
Each series is a single node; cell values are shown as a comma-separated list
with larger type than the series name. Drag nodes to rearrange; **Ctrl+Z**
undoes a move.

Great Docs / Quarto copies `assets/**` into the site. Preview locally:

```bash
python -m http.server 8000 --directory assets/graph
# open http://localhost:8000/
```

Keep `app.js` `evaluate()` aligned with `tiny_dsa.model.Model`. Check with:

```bash
uv run python scripts/check_graph_eval.py
```
