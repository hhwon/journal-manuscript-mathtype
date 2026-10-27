#!/usr/bin/env python3
"""Offline regression test for experimental v4 explicit fonts and MT numbering.

Proves package structure and OOXML MathType-field preservation only; it DOES
NOT certify Word+MathType UI behavior or pixel-perfect Windows typography.
"""
from __future__ import annotations
import argparse
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from ole_editor import FormulaDocument
from ole_native import CompoundFile
from mtef_font_audit import read_font_profile
from mtef_explicit import FONT_STYLES
from mtef_subset import validate_subset_mtef

NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def test(docx, reference):
    f=FormulaDocument(docx)
    if len(f.formulas)!=9:raise AssertionError('Unexpected formula count')
    with ZipFile(docx) as z,ZipFile(reference) as zr:
        root=etree.fromstring(z.read('word/document.xml'))
        styles=etree.fromstring(z.read('word/styles.xml'))
        instr=' '.join(root.xpath('.//w:instrText/text()',namespaces=NS))
        for code in ('MACROBUTTON MTPlaceRef','SEQ MTEqn \\h','SEQ MTEqn \\c'):
            assert instr.count(code)==9,(code,instr.count(code))
        assert len(root.xpath('.//w:p[w:pPr/w:pStyle[@w:val="MTDisplayEquation"]]',namespaces=NS))==9
        assert len(styles.xpath('.//w:style[@w:styleId="MTDisplayEquation"]',namespaces=NS))==1
        for i in range(1,9):
            mt=CompoundFile(z.read(f.by_index(i).object_part)).read('Equation Native')[28:]
            profile=read_font_profile(mt)
            if i%2==0:
                assert len(profile['explicit_font_styles'])==5
                a=profile['explicit_font_styles']
                assert [(v['font'],v['bold'],v['italic']) for v in a]==[
                    ('Times New Roman',False,True),('Times New Roman',False,False),
                    ('Times New Roman',True,False),('Symbol',False,True),('Symbol',False,False)
                ]
            else:
                assert profile['explicit_font_styles']==[]
                assert validate_subset_mtef(mt)['glyphs']>0
        assert z.read(f.by_index(9).object_part) in [zr.read(x) for x in zr.namelist() if x.startswith('word/embeddings/')]
    print('PASS: 9 true OLE, 4 semantic/explicit pairs, 9 native MTEqn fields, 1 unchanged original control')

if __name__=='__main__':
    p=argparse.ArgumentParser()
    p.add_argument('docx',type=Path)
    p.add_argument('--reference',required=True,type=Path)
    x=p.parse_args();test(x.docx,x.reference)
