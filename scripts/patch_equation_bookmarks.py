#!/usr/bin/env python3
"""Add stable Word bookmarks to MathType equation labels after DOCX embedding.

The upstream EquationRegistry replaces complete display paragraphs. This patch
adds the user's logical identifier to the backend-generated equation paragraph
*without* modifying the OLE object, its relationships, or its MTEF content.
"""
from __future__ import annotations
from pathlib import Path
from copy import copy
from tempfile import NamedTemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
O = 'urn:schemas-microsoft-com:office:office'
NS = {'w': W, 'o': O}


def patch(path, equation_anchors):
    path = Path(path)
    anchors = list(equation_anchors)
    if not anchors:
        return 0
    with ZipFile(path) as z:
        root = etree.fromstring(z.read('word/document.xml'))
        candidates = []
        for p in root.xpath('.//w:body//w:p', namespaces=NS):
            if p.xpath('.//o:OLEObject', namespaces=NS) and p.xpath(
                    './w:pPr/w:tabs/w:tab[@w:val="center"]', namespaces=NS):
                candidates.append(p)
        if len(candidates) != len(anchors):
            raise ValueError(f'Expected {len(anchors)} display equation paragraphs; found {len(candidates)}')
        for p, anchor in zip(candidates, anchors):
            tabs = p.xpath('./w:r[w:tab]', namespaces=NS)
            if len(tabs) < 2:
                raise ValueError(f'Missing right-aligned equation number tab for {anchor.identifier}')
            right_tab = tabs[-1]
            after = [child for child in p if p.index(child) > p.index(right_tab)]
            if not after:
                raise ValueError(f'Missing equation number for {anchor.identifier}')
            visible = ''.join(child.text or '' for node in after
                              for child in node.xpath('.//w:t', namespaces=NS))
            if f'({anchor.number})' != visible:
                raise ValueError(f'Equation label mismatch for {anchor.identifier}: {visible!r} != ({anchor.number})')
            start = etree.Element(f'{{{W}}}bookmarkStart')
            start.set(f'{{{W}}}id', str(anchor.bookmark_id))
            start.set(f'{{{W}}}name', anchor.bookmark)
            end = etree.Element(f'{{{W}}}bookmarkEnd')
            end.set(f'{{{W}}}id', str(anchor.bookmark_id))
            p.insert(p.index(right_tab) + 1, start)
            p.append(end)
        xml = etree.tostring(root, encoding='utf-8', xml_declaration=True, standalone=True)
        with NamedTemporaryFile(dir=path.parent, prefix='.eqrefs-', suffix='.docx', delete=False) as tmp:
            temp = Path(tmp.name)
        try:
            with ZipFile(temp, 'w', ZIP_DEFLATED) as out:
                for member in z.infolist():
                    out.writestr(copy(member), xml if member.filename == 'word/document.xml'
                                 else z.read(member.filename))
        except Exception:
            temp.unlink(missing_ok=True)
            raise
    temp.replace(path)
    return len(anchors)


if __name__ == '__main__':
    raise SystemExit('Import patch(path, equation_anchors) from the manuscript builder.')
