#!/usr/bin/env python3
"""Write assets/graph/bootstrap.json for static docs preview (no live API).

Prefers FormulaEvaluator when available; otherwise the exported Model.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets" / "graph" / "bootstrap.json"


def main() -> int:
    sys.path.insert(0, str(ROOT))
    from tiny_dsa.graph_api import available_backends, bootstrap

    backends = available_backends()
    backend = "formula_evaluator" if "formula_evaluator" in backends else "export"
    payload = bootstrap(backend=backend)
    OUT.write_text(json.dumps(payload, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    print(f"wrote {OUT} (backend={backend})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
