#!/usr/bin/env bash
set -euo pipefail

PHASE11_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${PHASE11_DIR}/../.." && pwd)"

cd "${REPO_ROOT}"

echo "==> Generating manual verification documents"
dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj

echo "==> Running Linux LibreOffice manual open/store/reopen check"
tools/manual-verification/scripts/run-linux-manual-check.sh

echo "==> Running Linux LibreOffice evidence export"
tools/manual-verification/scripts/run-linux-evidence.sh

echo "==> Linux manual verification complete"
