#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PHASE11_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
REPO_ROOT="$(cd "${PHASE11_DIR}/../.." && pwd)"
COMPOSE_FILE="${PHASE11_DIR}/docker-compose.yml"

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

evidence_id="v${version}-${revision}"

docker compose -f "${COMPOSE_FILE}" up -d
docker compose -f "${COMPOSE_FILE}" exec -T \
  -e DOTNETPOI_VERSION="${version}" \
  -e DOTNETPOI_REVISION="${revision}" \
  -e DOTNETPOI_EVIDENCE_ID="${evidence_id}" \
  libreoffice \
  python3 /workspace/tools/manual-verification/scripts/run_linux_evidence.py

printf "Evidence written to: tools/manual-verification/evidence/%s/INDEX.md\n" "${evidence_id}"
