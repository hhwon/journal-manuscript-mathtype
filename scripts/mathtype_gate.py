"""Delivery gate for generated MathType objects (run on every DOCX before it leaves the machine).

Fails unless: every 'Equation Native' parses as MTEF 5 (mtef_check), every OLE compound file has a valid directory
tree (cfb_treecheck), settings carry no w:updateFields, and no Symbol-style character lacks its Symbol position.

    python mathtype_gate.py manuscript.docx
"""
import json
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import cfb_treecheck  # noqa: E402
import mtef_check  # noqa: E402


def gate(docx, unmapped_symbols=()):
    report = {'objects': 0, 'malformed_mtef': [], 'invalid_tree': [], 'update_fields': False,
              'symbol_chars_without_position': sorted(unmapped_symbols)}
    for name, mtef, _ in mtef_check.equations(docx):
        report['objects'] += 1
        try:
            mtef_check.check(mtef)
        except Exception as exc:
            report['malformed_mtef'].append(f'{name}: {exc}')
    z = zipfile.ZipFile(docx)
    for n in z.namelist():
        if n.startswith('word/embeddings/') and n.endswith('.bin'):
            errs = cfb_treecheck.check(z.read(n))
            if errs:
                report['invalid_tree'].append(f'{n}: {errs[:2]}')
    if 'word/settings.xml' in z.namelist():
        report['update_fields'] = bool(re.search(rb'<w:updateFields', z.read('word/settings.xml')))
    report['pass'] = not (report['malformed_mtef'] or report['invalid_tree'] or report['update_fields']
                          or report['symbol_chars_without_position'])
    return report


if __name__ == '__main__':
    r = gate(sys.argv[1])
    print(json.dumps(r, indent=1, ensure_ascii=False))
    sys.exit(0 if r['pass'] else 1)
