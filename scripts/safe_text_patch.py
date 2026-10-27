#!/usr/bin/env python3
"""Edit one-run text in an existing DOCX without reserialising MathType OLE parts.

Input JSON: [{"old":"literal text", "new":"replacement", "expected":1}, ...]
Only changes plain Word w:t elements. Refuses cross-run text replacement and
ambiguous matches; leaves all original DOCX parts except document.xml intact.
Use only for local prose corrections, not formula, field, numbering or layout edits.
"""
from __future__ import annotations
import argparse
from copy import copy
import json
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def _plain_text_nodes(root):
    """Limit edits to direct paragraph runs, excluding complex Word fields.

    Field results contain w:t too, but are *not* ordinary prose. They must stay
    under Word's SEQ/REF/PAGE control. This refuses touching these cached values.
    """
    for p in root.xpath('.//w:body//w:p', namespaces=NS):
        field_depth = 0
        for element in p:
            if element.tag != f'{{{W}}}r':
                continue
            controls = element.findall(f'{{{W}}}fldChar')
            start = sum(c.get(f'{{{W}}}fldCharType') == 'begin' for c in controls)
            end = sum(c.get(f'{{{W}}}fldCharType') == 'end' for c in controls)
            if field_depth == 0 and not controls and element.find(f'{{{W}}}instrText') is None:
                if element.find(f'{{{W}}}object') is None and element.find(f'{{{W}}}drawing') is None:
                    for txt in element.findall(f'{{{W}}}t'):
                        yield txt
            field_depth += start - end
            if field_depth < 0:
                field_depth = 0


def patch_text(input_docx, output_docx, replacements):
    source = Path(input_docx).resolve()
    output = Path(output_docx).resolve()
    if source == output:
        raise ValueError('Refusing in-place patch: always keep the original DOCX')
    if not isinstance(replacements, list) or not replacements:
        raise ValueError('Replacement JSON must be a nonempty list')
    for item in replacements:
        if not isinstance(item, dict) or not isinstance(item.get('old'), str) or not item['old']:
            raise ValueError('Each replacement requires nonempty string old and new')
        if not isinstance(item.get('new'), str):
            raise ValueError('Each replacement requires string new')
        if item['old'] == item['new']:
            raise ValueError('Old and new text are identical')
        if int(item.get('expected', 1)) < 1:
            raise ValueError('Expected occurrence count must be positive')
    with ZipFile(source, 'r') as zin:
        if zin.testzip():
            raise ValueError('Input DOCX has an invalid ZIP entry')
        root = etree.fromstring(zin.read('word/document.xml'))
        # Protect text in field instructions / Word formulas / OLE object wrappers.
        nodes = list(_plain_text_nodes(root))
        applied = []
        for item in replacements:
            old, new = item['old'], item['new']
            exact = bool(item.get('exact', False))
            matches = [node for node in nodes if ((node.text or '') == old if exact else old in (node.text or ''))]
            actual = len(matches) if exact else sum((node.text or '').count(old) for node in matches)
            expected = int(item.get('expected', 1))
            if actual != expected:
                raise ValueError(f'Replacement {old!r}: expected {expected} plain text occurrences; got {actual}. '
                                 'The text may be split across runs. Do not alter MathType/field XML blindly.')
            for node in matches:
                node.text = new if exact else node.text.replace(old, new)
                if node.text and (node.text[0].isspace() or node.text[-1].isspace()):
                    node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            applied.append({'old': old, 'new': new, 'count': actual})
        xml = etree.tostring(root, encoding='utf-8', xml_declaration=True, standalone=True)
        output.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(prefix='.patch-', suffix='.docx', dir=output.parent, delete=False) as f:
            tmp = Path(f.name)
        try:
            with ZipFile(tmp, 'w', ZIP_DEFLATED) as zout:
                for entry in zin.infolist():
                    zout.writestr(copy(entry), xml if entry.filename == 'word/document.xml'
                                  else zin.read(entry.filename))
            with ZipFile(tmp) as check:
                if check.testzip():
                    raise ValueError('Output DOCX failed ZIP CRC audit')
                for name in zin.namelist():
                    if name != 'word/document.xml' and zin.read(name) != check.read(name):
                        raise ValueError(f'Unexpected change to protected DOCX part: {name}')
            tmp.replace(output)
        except Exception:
            tmp.unlink(missing_ok=True)
            raise
    return {'input': str(source), 'output': str(output),
            'changed_parts': ['word/document.xml'], 'replacements': applied,
            'ole_binaries_preserved_byte_for_byte': True,
            'windows_mathtype_editability_verified': False}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('input_docx')
    ap.add_argument('changes_json')
    ap.add_argument('-o', '--output', required=True)
    args = ap.parse_args()
    result = patch_text(args.input_docx, args.output,
                        json.loads(Path(args.changes_json).read_text(encoding='utf-8')))
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
