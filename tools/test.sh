#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$repo_root"

require_cmd() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "Missing required command: $1" >&2
    exit 1
  fi
}

require_cmd dotnet
require_cmd mvn
require_cmd java

if [ ! -d "$repo_root/poi" ]; then
  echo "Missing poi/ submodule. Run: git submodule update --init --recursive" >&2
  exit 1
fi

configuration="Debug"

printf '\n== Build solution ==\n'
dotnet build DotnetPOI.sln --no-incremental -c "$configuration"

printf '\n== Interop [A] Java writes fixtures ==\n'
mvn test \
  -f tests/DotnetPoi.Interop.Tests/java/pom.xml \
  -Dtest=WriteForDotnetTest \
  -B --no-transfer-progress

printf '\n== Interop [A] C# reads Java fixtures ==\n'
dotnet test tests/DotnetPoi.Interop.Tests/ \
  --no-build -c "$configuration" \
  --filter "Category=ReadFromPoi"

printf '\n== Interop [B] C# writes fixtures ==\n'
dotnet test tests/DotnetPoi.Interop.Tests/ \
  --no-build -c "$configuration" \
  --filter "Category=WriteForPoi"

printf '\n== Interop [B] Java reads C# fixtures ==\n'
mvn test \
  -f tests/DotnetPoi.Interop.Tests/java/pom.xml \
  -Dtest=ReadFromDotnetTest \
  -B --no-transfer-progress

printf '\n== Done ==\n'

