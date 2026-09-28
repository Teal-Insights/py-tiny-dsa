#!/usr/bin/env python3
"""Serve the Tiny DSA series-graph API (stdlib HTTP).

Endpoints:
  GET  /api/graph              bootstrap (schema + defaults + values)
  POST /api/evaluate           recompute; body {"backend":"...","inputs":{...}}
  GET  /                       static assets/graph/ (viz UI)

Default backend is FormulaEvaluator when available (``uv sync --group graph``).
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tiny_dsa.graph_api import (  # noqa: E402
    GraphApiError,
    available_backends,
    bootstrap,
    evaluate,
)
from tiny_dsa.graph_schema import BackendName  # noqa: E402

STATIC_ROOT = ROOT / "assets" / "graph"


def _json_response(handler: BaseHTTPRequestHandler, status: int, payload: Any) -> None:
    body = json.dumps(payload, allow_nan=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    handler.send_header("Access-Control-Allow-Headers", "Content-Type")
    handler.end_headers()
    handler.wfile.write(body)


def _default_backend() -> BackendName:
    backends = available_backends()
    if "formula_evaluator" in backends:
        return "formula_evaluator"
    return "export"


def _parse_backend(raw: Any) -> BackendName:
    if raw is None or raw == "":
        return _default_backend()
    if raw in ("export", "formula_evaluator"):
        return raw  # type: ignore[return-value]
    raise GraphApiError(f"unknown backend: {raw!r}", status=400)


class GraphApiHandler(BaseHTTPRequestHandler):
    server_version = "TinyDsaGraphAPI/0.1"

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A003
        sys.stderr.write("%s - %s\n" % (self.address_string(), format % args))

    def do_OPTIONS(self) -> None:  # noqa: N802
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path in ("/api/health", "/api/health/"):
            from tiny_dsa.graph_api import available_backends

            _json_response(
                self,
                200,
                {
                    "ok": True,
                    "backends": available_backends(),
                },
            )
            return
        if parsed.path in ("/api/graph", "/api/graph/"):
            try:
                qs = parse_qs(parsed.query)
                backend = _parse_backend((qs.get("backend") or [""])[0])
                _json_response(self, 200, bootstrap(backend=backend))
            except GraphApiError as exc:
                _json_response(self, exc.status, exc.as_dict())
            except Exception as exc:  # pragma: no cover
                _json_response(self, 500, {"error": str(exc)})
            return
        self._serve_static(parsed.path)
        return

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path not in ("/api/evaluate", "/api/evaluate/"):
            _json_response(self, 404, {"error": "not found"})
            return
        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length) if length else b"{}"
        try:
            payload = json.loads(raw_body.decode("utf-8") or "{}")
            if not isinstance(payload, dict):
                raise GraphApiError("JSON body must be an object", status=400)
            backend = _parse_backend(payload.get("backend"))
            inputs = payload.get("inputs")
            if inputs is not None and not isinstance(inputs, dict):
                raise GraphApiError("inputs must be an object", status=400)
            values = evaluate(inputs, backend=backend)
            _json_response(self, 200, {"backend": backend, "values": values})
        except GraphApiError as exc:
            _json_response(self, exc.status, exc.as_dict())
        except json.JSONDecodeError as exc:
            _json_response(self, 400, {"error": f"invalid JSON: {exc}"})
        except Exception as exc:  # pragma: no cover
            _json_response(self, 500, {"error": str(exc)})

    def _serve_static(self, path: str) -> None:
        rel = path.lstrip("/") or "index.html"
        if rel.startswith("api/"):
            _json_response(self, 404, {"error": "not found"})
            return
        candidate = (STATIC_ROOT / rel).resolve()
        try:
            candidate.relative_to(STATIC_ROOT.resolve())
        except ValueError:
            self.send_error(404)
            return
        if candidate.is_dir():
            candidate = candidate / "index.html"
        if not candidate.is_file():
            self.send_error(404)
            return
        data = candidate.read_bytes()
        ctype, _ = mimetypes.guess_type(str(candidate))
        self.send_response(200)
        self.send_header("Content-Type", ctype or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(data)


def main(argv: list[str] | None = None) -> int:
    import os

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--host",
        default=os.environ.get("HOST", "127.0.0.1"),
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.environ.get("PORT", "8765")),
    )
    args = parser.parse_args(argv)
    server = ThreadingHTTPServer((args.host, args.port), GraphApiHandler)
    print(f"Tiny DSA graph API on http://{args.host}:{args.port}/")
    print("  GET  /api/health")
    print("  GET  /api/graph")
    print("  POST /api/evaluate")
    print(f"  static {STATIC_ROOT} at /")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nshutting down")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
