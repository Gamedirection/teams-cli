#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${repo_root}/term.everything"

if ! command -v go >/dev/null 2>&1; then
  echo "Go is required to build term.everything." >&2
  exit 1
fi

make build
