#!/usr/bin/env bash
# Idempotent bootstrap for the tiny-dsa Cloud Agent environment.
# Installs the uv toolchain, the Quarto CLI (required by great-docs to render
# the documentation site), the project's pinned Python interpreter, and all
# runtime + dev dependencies from the committed uv.lock.
set -euo pipefail

QUARTO_VERSION="1.6.42"

if ! command -v uv >/dev/null 2>&1; then
  curl -LsSf https://astral.sh/uv/install.sh | sh
fi
export PATH="$HOME/.local/bin:$PATH"

if [ "$(quarto --version 2>/dev/null || true)" != "$QUARTO_VERSION" ]; then
  tmp="$(mktemp -d)"
  curl -LsSf -o "$tmp/quarto.tar.gz" \
    "https://github.com/quarto-dev/quarto-cli/releases/download/v${QUARTO_VERSION}/quarto-${QUARTO_VERSION}-linux-amd64.tar.gz"
  sudo rm -rf /opt/quarto
  sudo mkdir -p /opt/quarto
  sudo tar -xzf "$tmp/quarto.tar.gz" -C /opt/quarto --strip-components=1
  sudo ln -sf /opt/quarto/bin/quarto /usr/local/bin/quarto
  rm -rf "$tmp"
fi

uv python install
uv sync --group dev
