#!/usr/bin/env python3
"""Regression tests for MTEF *native* font roles (not just Word PNG previews)."""
from pathlib import Path
from zipfile import ZipFile
import argparse
from mtef_subset import make_mtef, validate_subset_mtef
from ole_editor import FormulaDocument
from ole_native import CompoundFile
from mtef_font_audit import read_font_profile

TESTS = {
    r'E=kT': ['variable','symbol','variable','variable'],
    r'\sigma=E\varepsilon': ['lcgreek','symbol','variable','lcgreek'],
    r'\Gamma=\Sigma+E': ['ucgreek','symbol','ucgreek','symbol','variable'],
    r'\sin(\theta)+\cos(\phi)': ['function','function','function','symbol','lcgreek','symbol','symbol','function','function','function','symbol','lcgreek','symbol'],
    r'\mathrm{Re}=\mathbf{u}': ['text','text','symbol','vector'],
    r'\operatorname{det}=x': ['function','function','function','symbol','variable'],
    r'\mathbf{u}_i': ['vector','variable'],
    r'\frac{1}{2}': ['number','number'],
    r'\sqrt{x^{2}+y^{2}}': ['variable','number','symbol','variable','number'],
}

def run(source: Path):
    doc=FormulaDocument(source)
    seed=next(f for f in doc.formulas if f.prog_id=='Equation.DSMT4' and f.mtef_version==5)
    with ZipFile(source) as z:
        original=CompoundFile(z.read(seed.object_part)).read('Equation Native')[28:]
    ref=read_font_profile(original)
    s={item['name']:item for item in ref['styles']}
    for name,font,italic,bold in [
       ('Text','Times New Roman',False,False),
       ('Function','Times New Roman',False,False),
       ('Variable','Times New Roman',True,False),
       ('L.C. Greek','Symbol',True,False),
       ('U.C. Greek','Symbol',False,False),
       ('Symbol','Symbol',False,False),
       ('Vector','Times New Roman',False,True),
       ('Number','Times New Roman',False,False),
    ]:
        r=s[name]
        assert (r['font'],r['italic'],r['bold']) == (font,italic,bold),(name,r)
    assert ref['full_size']=={'value':'12','unit':'pt'}
    for latex,roles in TESTS.items():
        actual=make_mtef(latex,original)
        assert read_font_profile(actual)==ref
        audit=validate_subset_mtef(actual)
        got=[g['role'] for g in audit['characters']]
        assert got==roles,(latex,got,roles)
        if '\\sin' in latex or '\\operatorname' in latex:
            assert audit['function_starts'] >= 1
        print('PASS',latex,'roles:',','.join(got))
    for unsupported in (r'\epsilon',r'\varphi',r'\boldsymbol{\sigma}',r'\int_0^1',r'\text{a b}'):
        try:make_mtef(unsupported,original)
        except ValueError:pass
        else:raise AssertionError('Unsupported variant silently accepted: '+unsupported)
    print('PASS strict rejection of unsupported / ambiguous math styles')
    print('PASS font preferences identical to original MathType 7 12pt donor')

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--source',required=True)
    args=p.parse_args();run(Path(args.source))
