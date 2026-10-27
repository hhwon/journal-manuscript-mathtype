#!/usr/bin/env bash
# Install the Microsoft Times New Roman family WITHOUT bundling font binaries.
# --apt: use OS package installer (interactive Microsoft font EULA).
# --from-dir: copy four user-supplied, properly licensed font faces locally.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON="${PYTHON:-python3}"

usage() {
  echo "Usage: bash scripts/install_tnr.sh --apt | --from-dir /path/to/licensed-fonts | --check"
  echo "  --apt: requires Debian contrib/Ubuntu multiverse, Internet and license agreement"
  echo "  --from-dir: install your four licensed Times New Roman *.ttf files for current user"
  echo "  --check: verify all four faces using their internal font metadata"
}

if [[ $# -eq 1 && "$1" == "--check" ]]; then
  exec "$PYTHON" "$HERE/check_tnr.py" --strict --summary
elif [[ $# -eq 1 && "$1" == "--apt" ]]; then
  if ! command -v apt-get >/dev/null 2>&1; then
    echo 'ERROR: apt-get is unavailable; use --from-dir instead.' >&2; exit 2
  fi
  if ! apt-cache policy ttf-mscorefonts-installer | grep -Eq 'Candidate: [^ (]+'; then
    echo 'ERROR: ttf-mscorefonts-installer unavailable. Enable Debian contrib or Ubuntu multiverse,' >&2
    echo '       run apt-get update, and retry; do not install a fake TNR alias.' >&2
    exit 2
  fi
  echo 'The installer may request acceptance of Microsoft font license terms.'
  echo 'Review the terms yourself; this script does NOT pre-accept any EULA.'
  if [[ "$(id -u)" -eq 0 ]]; then
    apt-get install ttf-mscorefonts-installer
  else
    sudo apt-get install ttf-mscorefonts-installer
  fi
  fc-cache -f
  "$PYTHON" "$HERE/check_tnr.py" --strict --summary
elif [[ $# -eq 2 && "$1" == "--from-dir" ]]; then
  FONT_SRC="$2"
  "$PYTHON" "$HERE/check_tnr.py" --font-dir "$FONT_SRC" --strict --summary
  DEST="${XDG_DATA_HOME:-$HOME/.local/share}/fonts/TimesNewRoman"
  mkdir -p "$DEST"
  "$PYTHON" - "$HERE" "$FONT_SRC" "$DEST" <<'PY'
import sys
from pathlib import Path
import shutil
sys.path.insert(0,sys.argv[1])
from check_tnr import metadata, find_paths
source, target = sys.argv[2:]
found = {}
for path in find_paths(source):
    info = metadata(path)
    if info and info['role'] not in found:
        found[info['role']] = path
for role in ('regular','italic','bold','bold_italic'):
    shutil.copy2(found[role],Path(target)/(role+'.ttf'))
print('Installed four licensed Times New Roman faces in:',target)
PY
  fc-cache -f
  echo 'For Matplotlib, restart Python after installation or clear its old font cache.'
  "$PYTHON" "$HERE/check_tnr.py" --strict --summary
else
  usage
  exit 2
fi
