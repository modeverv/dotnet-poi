#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT_DIR"

PROJECTS=(
  "Common:src/DotnetPoi.Common/DotnetPoi.Common.csproj"
  "POIFS:src/DotnetPoi.POIFS/DotnetPoi.POIFS.csproj"
  "Legacy:src/DotnetPoi.Legacy/DotnetPoi.Legacy.csproj"
  "Formula:src/DotnetPoi.Formula/DotnetPoi.Formula.csproj"
  "Ooxml:src/DotnetPoi.Ooxml/DotnetPoi.Ooxml.csproj"
  "All:src/DotnetPoi.All/DotnetPoi.All.csproj"
)

msbuild_property() {
  local project="$1"
  local property="$2"
  dotnet msbuild "$project" -nologo -getProperty:"$property"
}

COMMON_ID=""
COMMON_VERSION=""
POIFS_ID=""
POIFS_VERSION=""
OOXML_ID=""
OOXML_VERSION=""
LEGACY_ID=""
LEGACY_VERSION=""
FORMULA_ID=""
FORMULA_VERSION=""
ALL_ID=""
ALL_VERSION=""

echo "Release package order:"
for entry in "${PROJECTS[@]}"; do
  name="${entry%%:*}"
  project="${entry#*:}"
  id="$(msbuild_property "$project" PackageId)"
  version="$(msbuild_property "$project" PackageVersion)"
  readme="$(msbuild_property "$project" PackageReadmeFile)"

  case "$name" in
    Common)
      COMMON_ID="$id"
      COMMON_VERSION="$version"
      ;;
    POIFS)
      POIFS_ID="$id"
      POIFS_VERSION="$version"
      ;;
    Ooxml)
      OOXML_ID="$id"
      OOXML_VERSION="$version"
      ;;
    Legacy)
      LEGACY_ID="$id"
      LEGACY_VERSION="$version"
      ;;
    Formula)
      FORMULA_ID="$id"
      FORMULA_VERSION="$version"
      ;;
    All)
      ALL_ID="$id"
      ALL_VERSION="$version"
      ;;
  esac

  echo "  - $id $version ($project)"

  if [ -z "$id" ] || [ -z "$version" ]; then
    echo "Missing PackageId/PackageVersion in $project" >&2
    exit 1
  fi

  if [ "$readme" != "README.md" ]; then
    echo "$project must set <PackageReadmeFile>README.md</PackageReadmeFile>" >&2
    exit 1
  fi

  project_dir="$(dirname "$project")"
  if [ ! -f "$project_dir/$readme" ]; then
    echo "Missing package README: $project_dir/$readme" >&2
    exit 1
  fi

  if ! grep -q "$id" "$project_dir/$readme"; then
    echo "$project_dir/$readme should mention $id" >&2
    exit 1
  fi
done

if [ "$COMMON_ID" != "DotnetPoi.Common" ] ||
   [ "$POIFS_ID" != "DotnetPoi.POIFS" ] ||
   [ "$OOXML_ID" != "DotnetPoi.Ooxml" ] ||
   [ "$LEGACY_ID" != "DotnetPoi.Legacy" ] ||
   [ "$FORMULA_ID" != "DotnetPoi.Formula" ] ||
   [ "$ALL_ID" != "DotnetPoi.All" ]; then
  echo "Unexpected PackageId. Release package IDs are fixed." >&2
  exit 1
fi

if ! grep -q "DotnetPoi.Common" src/DotnetPoi.All/DotnetPoi.All.csproj ||
   ! grep -q "DotnetPoi.POIFS" src/DotnetPoi.All/DotnetPoi.All.csproj ||
   ! grep -q "DotnetPoi.Ooxml" src/DotnetPoi.All/DotnetPoi.All.csproj ||
   ! grep -q "DotnetPoi.Legacy" src/DotnetPoi.All/DotnetPoi.All.csproj ||
   ! grep -q "DotnetPoi.Formula" src/DotnetPoi.All/DotnetPoi.All.csproj; then
  echo "DotnetPoi.All must reference Common, POIFS, Ooxml, Legacy, and Formula." >&2
  exit 1
fi

for doc in README.md NOW.md src/DotnetPoi.Common/README.md src/DotnetPoi.POIFS/README.md src/DotnetPoi.Ooxml/README.md src/DotnetPoi.Legacy/README.md src/DotnetPoi.Formula/README.md src/DotnetPoi.All/README.md; do
  if [ ! -f "$doc" ]; then
    echo "Missing release-facing document: $doc" >&2
    exit 1
  fi
done

for id in DotnetPoi.Common DotnetPoi.POIFS DotnetPoi.Ooxml DotnetPoi.Legacy DotnetPoi.Formula DotnetPoi.All; do
  if ! grep -q "$id" README.md; then
    echo "README.md should mention $id" >&2
    exit 1
  fi
done

if ! grep -q "NOW.md" README.md; then
  echo "README.md should link to NOW.md so release status updates have a canonical source." >&2
  exit 1
fi

if [ -n "${GITHUB_REF_NAME:-}" ] && [[ "$GITHUB_REF_NAME" == v* ]]; then
  tag_version="${GITHUB_REF_NAME#v}"
  if [ "$tag_version" != "$OOXML_VERSION" ] || [ "$tag_version" != "$ALL_VERSION" ]; then
    echo "Release tag $GITHUB_REF_NAME must match DotnetPoi.Ooxml and DotnetPoi.All package versions." >&2
    echo "  DotnetPoi.Ooxml: $OOXML_VERSION" >&2
    echo "  DotnetPoi.All:   $ALL_VERSION" >&2
    exit 1
  fi
fi

echo "Package hygiene OK."
