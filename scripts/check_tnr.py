#!/usr/bin/env python3
"""Verify four *genuine named* Times New Roman faces, not fontconfig aliases.

This intentionally never installs, downloads, or redistributes Microsoft fonts.
It can also verify a directory of licensed fonts before a local installation.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import subprocess
from typing import Any

ROLES = ('regular', 'italic', 'bold', 'bold_italic')


def metadata(path: Path) -> dict[str, Any] | None:
    """Read family/style from a font's own OpenType name and style tables."""
    try:
        from fontTools.ttLib import TTFont
        with TTFont(str(path), lazy=True) as face:
            names = face['name']
            families = set()
            for record in names.names:
                if record.nameID in (1, 16):
                    try:
                        families.add(record.toUnicode().strip().casefold())
                    except Exception:
                        pass
            if 'times new roman' not in families:
                return None
            head_style = face['head'].macStyle
            fs_selection = face['OS/2'].fsSelection if 'OS/2' in face else 0
            bold = bool(head_style & 0x01 or fs_selection & 0x20)
            italic = bool(head_style & 0x02 or fs_selection & 0x01)
            role = 'bold_italic' if bold and italic else 'bold' if bold else 'italic' if italic else 'regular'
            return {'path': str(path.resolve()), 'role': role, 'family': 'Times New Roman'}
    except Exception:
        return None


def find_paths(directory: str | None = None) -> list[Path]:
    """Use fc-list's actual registered files, never a substituted fc-match result."""
    if directory:
        directory_path = Path(directory).expanduser()
        if not directory_path.is_dir():
            raise FileNotFoundError(f'Font directory not found: {directory_path}')
        return sorted(x for x in directory_path.rglob('*') if x.is_file() and x.suffix.casefold() in ('.ttf', '.otf', '.ttc'))
    if not shutil.which('fc-list'):
        raise RuntimeError('fontconfig (fc-list) is not installed')
    p = subprocess.run(['fc-list', '--format', '%{file}\n'], text=True, capture_output=True, check=True)
    return sorted({Path(x.strip()) for x in p.stdout.splitlines() if x.strip()})


def inventory(directory: str | None = None) -> dict[str, Any]:
    found: dict[str, list[str]] = {k: [] for k in ROLES}
    error: str | None = None
    try:
        import fontTools.ttLib  # must not silently report 'missing fonts' if parser is missing
        paths = find_paths(directory)
        for path in paths:
            info = metadata(path)
            if info:
                found[info['role']].append(info['path'])
    except (ImportError, FileNotFoundError, OSError, RuntimeError, subprocess.CalledProcessError) as exc:
        error = str(exc)
    missing = [k for k, v in found.items() if not v]
    return {'ready': not missing and error is None, 'family': 'Times New Roman',
            'faces': found, 'missing_faces': missing, 'error': error,
            'source': str(Path(directory).expanduser()) if directory else 'fontconfig',
            'warning': 'Tests only real internal font family/style metadata, not fontconfig aliases. ' +
                       'MathType MTEF typography and Symbol/MT Extra fonts are separate.'}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--font-dir', help='Inspect licensed local TTF/OTF files before installation')
    ap.add_argument('--strict', action='store_true', help='Return failure status unless all four faces found')
    ap.add_argument('--summary', action='store_true', help='Print concise status instead of JSON')
    args = ap.parse_args()
    result = inventory(args.font_dir)
    if args.summary:
        print('Times New Roman: '+('ALL FOUR FACES VERIFIED' if result['ready'] else
                                  'MISSING '+', '.join(result['missing_faces'])))
        if result['error']:
            print('Error:', result['error'])
        for key, paths in result['faces'].items():
            for item in paths:
                print(key, item)
    else:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0 if result['ready'] or not args.strict else 2


if __name__ == '__main__':
    raise SystemExit(main())
