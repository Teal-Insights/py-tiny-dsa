# Tiny DSA FormulaEvaluator graph API (Railway / container)

FROM python:3.13-slim

COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

WORKDIR /app

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    PYTHONUNBUFFERED=1

COPY pyproject.toml uv.lock README.md ./
COPY tiny_dsa ./tiny_dsa
COPY bindings ./bindings
COPY tests/fixtures ./tests/fixtures
COPY assets/graph ./assets/graph
COPY scripts/serve_graph_api.py ./scripts/serve_graph_api.py

RUN uv sync --frozen --no-dev --group graph \
    && uv run python -c "from tiny_dsa.graph_formula_evaluator import is_available; assert is_available(), 'FormulaEvaluator unavailable'"

ENV HOST=0.0.0.0 \
    PORT=8765

EXPOSE 8765

CMD ["uv", "run", "python", "scripts/serve_graph_api.py", "--host", "0.0.0.0"]
