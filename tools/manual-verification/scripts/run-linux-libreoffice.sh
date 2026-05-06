#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PHASE11_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
COMPOSE_FILE="${PHASE11_DIR}/docker-compose.yml"

mkdir -p "${PHASE11_DIR}/evidence"

docker compose -f "${COMPOSE_FILE}" up -d
docker compose -f "${COMPOSE_FILE}" ps

printf "\nLibreOffice Linux desktop is starting.\n"
printf "Web desktop:       http://localhost:3110\n"
printf "HTTPS web desktop: https://localhost:3111\n"
printf "Workspace path:    /workspace\n"
printf "Evidence path:     /workspace/tools/manual-verification/evidence\n"
