#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
PACKAGE_SOURCE="${1:-/tmp/nuget-publish}"

if [ ! -d "$PACKAGE_SOURCE" ]; then
  echo "Package source directory does not exist: $PACKAGE_SOURCE" >&2
  exit 1
fi

cd "$ROOT_DIR"

msbuild_property() {
  local project="$1"
  local property="$2"
  dotnet msbuild "$project" -nologo -getProperty:"$property"
}

PACKAGES=(
  "DotnetPoi.Common:src/DotnetPoi.Common/DotnetPoi.Common.csproj:DotnetPoi.SS.UserModel.CellType"
  "DotnetPoi.POIFS:src/DotnetPoi.POIFS/DotnetPoi.POIFS.csproj:DotnetPoi.POIFS.FileSystem.FileMagic"
  "DotnetPoi.Legacy:src/DotnetPoi.Legacy/DotnetPoi.Legacy.csproj:DotnetPoi.HSSF.UserModel.HSSFWorkbook"
  "DotnetPoi.Formula:src/DotnetPoi.Formula/DotnetPoi.Formula.csproj:DotnetPoi.Formula.FormulaEvaluator"
  "DotnetPoi.Ooxml:src/DotnetPoi.Ooxml/DotnetPoi.Ooxml.csproj:DotnetPoi.XSSF.UserModel.XSSFWorkbook"
  "DotnetPoi.All:src/DotnetPoi.All/DotnetPoi.All.csproj:DotnetPoi.XWPF.UserModel.XWPFDocument"
)

TMP_DIR="$(mktemp -d)"
trap 'rm -rf "$TMP_DIR"' EXIT

for entry in "${PACKAGES[@]}"; do
  package_id="${entry%%:*}"
  rest="${entry#*:}"
  project="${rest%%:*}"
  type_name="${rest#*:}"
  version="$(msbuild_property "$project" PackageVersion)"

  if [ ! -f "$PACKAGE_SOURCE/$package_id.$version.nupkg" ]; then
    echo "Expected package not found: $PACKAGE_SOURCE/$package_id.$version.nupkg" >&2
    exit 1
  fi

  smoke_dir="$TMP_DIR/$package_id"
  project_name="${package_id}.InstallSmoke"
  dotnet new console --name "$project_name" --output "$smoke_dir" >/dev/null

  cat > "$smoke_dir/nuget.config" <<EOF
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="local-dotnet-poi" value="$PACKAGE_SOURCE" />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
  </packageSources>
</configuration>
EOF

  cat > "$smoke_dir/Program.cs" <<EOF
using System;

Console.WriteLine(typeof($type_name).Assembly.GetName().Name);
EOF

  dotnet add "$smoke_dir/$project_name.csproj" package "$package_id" --version "$version" --source "$PACKAGE_SOURCE" >/dev/null
  dotnet restore "$smoke_dir/$project_name.csproj" --configfile "$smoke_dir/nuget.config" >/dev/null
  dotnet run --project "$smoke_dir/$project_name.csproj" --no-restore >/dev/null
  echo "NuGet install smoke OK: $package_id $version"
done

echo "NuGet install smoke OK for all packages."
