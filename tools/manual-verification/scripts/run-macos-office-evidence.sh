#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../../.." && pwd)"

version="$(
  sed -n 's:.*<VersionPrefix>\(.*\)</VersionPrefix>.*:\1:p' \
    "${REPO_ROOT}/src/DotnetPoi.Core/DotnetPoi.Core.csproj" | head -n 1
)"
if [ -z "${version}" ]; then
  version="unknown"
fi

revision="$(git -C "${REPO_ROOT}" rev-parse --short HEAD 2>/dev/null || true)"
if [ -z "${revision}" ]; then
  revision="nogit"
fi

DOTNETPOI_VERSION="${version}" \
DOTNETPOI_REVISION="${revision}" \
DOTNETPOI_EVIDENCE_ID="v${version}-${revision}-macos" \
python3 "${SCRIPT_DIR}/run_macos_office_evidence.py"
