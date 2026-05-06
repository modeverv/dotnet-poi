#!/usr/bin/env bash
set -euo pipefail

PHASE11_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${PHASE11_DIR}/../.." && pwd)"

cd "${REPO_ROOT}"

echo "==> Generating manual verification documents"
dotnet run --project tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj

echo "==> Running macOS Microsoft Office evidence"
tools/manual-verification/scripts/run-macos-office-evidence.sh

echo "==> macOS manual verification complete"
