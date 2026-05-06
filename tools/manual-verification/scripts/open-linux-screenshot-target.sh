#!/usr/bin/env bash
set -euo pipefail

if [ "$#" -ne 1 ]; then
  echo "usage: $0 examples/output/<file>" >&2
  exit 2
fi

TARGET="$1"
CONTAINER_PATH="/workspace/${TARGET#./}"

docker exec dotnet-poi-phase11-libreoffice bash -lc "
  pkill -x soffice.bin || true
  pkill -x oosplash || true
  rm -f /tmp/phase11-open.log
  su -s /bin/bash abc -c 'DISPLAY=:0 XDG_RUNTIME_DIR=/config/.XDG WAYLAND_DISPLAY=wayland-1 SAL_USE_VCLPLUGIN=gtk3 nohup libreoffice --norestore --nofirststartwizard --view \"${CONTAINER_PATH}\" >/tmp/phase11-open.log 2>&1 &'
"
