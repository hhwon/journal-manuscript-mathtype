#!/usr/bin/env python3
"""Diagnose MathType backend, version, font availability and Linux renderer."""
from __future__ import annotations
import argparse
import importlib
import importlib.metadata
import json
import shutil
import subprocess
import sys

EXPECTED = '0.3.3'


def check():
    report = {'status': 'ready', 'production_mathtype_ready': True, 'checks': {}}
    modules = {
        'python-docx': 'docx', 'latex2mathml': 'latex2mathml',
        'docx-equation': 'docx_equation', 'Pillow': 'PIL', 'lxml': 'lxml',
        'olefile': 'olefile',
    }
    for name, module in modules.items():
        item = {}
        try:
            item['version'] = importlib.metadata.version(name)
            importlib.import_module(module)
            item['status'] = 'ok'
        except (ModuleNotFoundError, importlib.metadata.PackageNotFoundError) as exc:
            item = {'status': 'missing', 'error': str(exc)}
            report['production_mathtype_ready'] = False
        report['checks'][name] = item
    eq = report['checks']['docx-equation']
    if eq.get('status') == 'ok':
        if eq['version'] != EXPECTED:
            eq['status'] = 'wrong_version'
            eq['required'] = EXPECTED
            report['production_mathtype_ready'] = False
        else:
            try:
                mod = importlib.import_module('docx_equation')
                for symbol in ('EquationRegistry', 'EquationStyle', 'ExportOptions',
                               'MathTypeOptions', 'NumberingOptions'):
                    if not hasattr(mod, symbol):
                        raise RuntimeError(f'Missing API symbol {symbol}')
                eq['registry_api'] = 'present'
            except (ImportError, RuntimeError) as exc:
                eq['status'] = 'incompatible_api'
                eq['error'] = str(exc)
                report['production_mathtype_ready'] = False
    # previews: Playwright's Chromium + Times New Roman, TeX Gyre Termes Math, URW Standard Symbols PS (preview_math.py)
    try:
        sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
        import preview_math
        preview_math.check()
        report['checks']['equation_previews'] = {'status': 'ok', 'math_font': preview_math.termes_math(),
                                                 'symbol_font': preview_math.urw_symbol()}
    except Exception as exc:                       # missing font, Playwright or browser
        report['checks']['equation_previews'] = {'status': 'missing', 'error': str(exc)[:300],
                                                 'fix': 'pip install playwright && python -m playwright install chromium; '
                                                        'fonts-texgyre-math (or texlive), fonts-urw-base35, Times New Roman'}
        report['production_mathtype_ready'] = False
    fonts = shutil.which('fc-list')
    report['checks']['fontconfig'] = {'status': 'available' if fonts else 'not_found'}
    if fonts:
        try:
            output = subprocess.run([fonts, ':', 'family'], capture_output=True, text=True,
                                    timeout=4, check=False).stdout.lower()
            report['checks']['fontconfig']['math_fonts_detected'] = any(x in output
                        for x in ('stix', 'liberation serif', 'dejavu math', 'noto serif'))
        except (OSError, subprocess.TimeoutExpired):
            pass
    report['status'] = 'ready' if report['production_mathtype_ready'] else 'not_ready'
    report['windows_math_type_edit_test'] = 'not_performed'
    return report


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--strict', action='store_true', help='Exit nonzero if native MathType backend is not ready')
    args = ap.parse_args()
    report = check()
    print(json.dumps(report, indent=2, ensure_ascii=False))
    return 0 if report['production_mathtype_ready'] or not args.strict else 2

if __name__ == '__main__':
    sys.exit(main())
