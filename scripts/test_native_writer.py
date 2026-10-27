#!/usr/bin/env python3
"""Native MTEF+CFB engine integration tests using user-provided MathType DOCX.

Run from Skill root:
 python scripts/test_native_writer.py --source path/to/real-mathtype.docx
No Word/MathType is required for Linux structural checks.
"""
import argparse
from pathlib import Path
from zipfile import ZipFile
import hashlib
import tempfile
import sys

from ole_editor import FormulaDocument, replace_from_latex
from ole_native import CompoundFile
from ole_cfb_writer import repack_stream
from mtef_subset import make_mtef, validate_subset_mtef
from build_native_ole_test import build

CASES=[r'E=kT',r'\sigma=E\varepsilon',
       r'\sigma_{ij}=C_{ijkl}\varepsilon_{kl}',
       r'E=mc^{2}',r'\frac{1}{2}',r'\sqrt{x^{2}+y^{2}}',
       r'\sigma_{ij}=C_{ijkl}\varepsilon_{kl}+D_{ijklmn}\varepsilon_{kl,mn}']


def run(source):
    source=Path(source)
    d=FormulaDocument(source)
    f=next((item for item in d.formulas if item.prog_id=='Equation.DSMT4' and item.mtef_version==5),None)
    if not f:raise ValueError('No Equation.DSMT4/MTEF v5 seed')
    with ZipFile(source) as z:
        raw=z.read(f.object_part)
        native=CompoundFile(raw).read('Equation Native')
        seed=native[28:]
        cf=CompoundFile(raw)
        for n in [len(native),50,1800,4095,4100,9000,33000]:
            value=bytes(i % 251 for i in range(n))
            rewritten=repack_stream(raw,'Equation Native',value)
            trial=CompoundFile(rewritten)
            assert trial.read('Equation Native')==value
            assert all(trial.read(name)==cf.read(name)
                       for name in cf.entries if name!='Equation Native')
    print('PASS CFB regular/mini-stream roundtrips across 4096-byte boundary')
    for latex in CASES:
        v=validate_subset_mtef(make_mtef(latex,seed))
        assert v['glyphs']>0
    print('PASS 7 MTEF generated structures roundtrip/validate')
    for value in [r'\int_0^1 x dx',r'\mathcal{C}',r'\frac{1}{2',r'x__2']:
        try:make_mtef(value,seed)
        except ValueError:pass
        else:raise AssertionError('Unsupported input accepted: '+value)
    print('PASS unsupported macro or malformed LaTeX refused')
    with tempfile.TemporaryDirectory() as t:
        folder=Path(t)
        original=folder/'native.docx'; edited=folder/'edited.docx'
        build(source,original)
        baseline=FormulaDocument(original)
        assert len(baseline.formulas)==9
        edited_metadata=replace_from_latex(original,3,CASES[-1],edited)
        revised=FormulaDocument(edited)
        assert len(revised.formulas)==len(baseline.formulas)
        assert edited_metadata['unchanged_equations_verified']==8
        with ZipFile(original) as a,ZipFile(edited) as b:
            for item in baseline.formulas:
                if item.index==3:continue
                assert a.read(item.object_part)==b.read(item.object_part)
                assert a.read(item.preview_part)==b.read(item.preview_part)
        assert revised.by_index(3).object_sha256!=baseline.by_index(3).object_sha256
        print('PASS new DOCX with 9 OLE objects + in-place target edit on a copy')
        try:replace_from_latex(original,3,CASES[0],original)
        except ValueError:pass
        else:raise AssertionError('In-place edit permitted unexpectedly')
        print('PASS refused original overwrite')
        try:replace_from_latex(original,3,r'\begin{matrix}x\end{matrix}',folder/'bad.docx')
        except ValueError:pass
        else:raise AssertionError('Unsupported LaTeX was accepted')
        assert not (folder/'bad.docx').exists()
        print('PASS unsupported expressions fail without output')
    print('All native tests passed; Windows MathType editing is NOT tested.')


def main():
    p=argparse.ArgumentParser();p.add_argument('--source',required=True)
    a=p.parse_args();run(a.source)
if __name__=='__main__':main()
