#!/usr/bin/env python3
"""Generate side-by-side MathType semantic-vs-explicit font diagnostic DOCX."""
from pathlib import Path
import argparse
from build_native_ole_test import build
from mathtype_numbering import upgrade

PAIRS=[
('A1','Semantic Variable','E=kT','synthesized'),
('A2','Explicit Times italic Variable','E=kT','synthesized-explicit'),
('A3','Semantic Greek tensor',r'\sigma_{ij}=C_{ijkl}\varepsilon_{kl}','synthesized'),
('A4','Explicit Times/Symbol tensor',r'\sigma_{ij}=C_{ijkl}\varepsilon_{kl}','synthesized-explicit'),
('A5','Semantic Function',r'\sin(\theta)+\cos(\phi)','synthesized'),
('A6','Explicit Times/Symbol Function',r'\sin(\theta)+\cos(\phi)','synthesized-explicit'),
('A7','Semantic bold vector',r'\mathrm{Re}=\mathbf{u}','synthesized'),
('A8','Explicit bold Times vector',r'\mathrm{Re}=\mathbf{u}','synthesized-explicit'),
('C1','Unmodified original MathType control',None,'original'),
]

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--source',required=True,type=Path)
    ap.add_argument('-o','--output',required=True,type=Path)
    x=ap.parse_args()
    raw=x.output.with_name('_mtef_pairs_unumbered.docx')
    build(x.source,raw,formulas=PAIRS)
    upgrade(raw,x.source,x.output)
    raw.unlink(missing_ok=True)
    mp=raw.with_suffix('.manifest.json')
    if mp.exists():
        mp.rename(x.output.with_suffix('.manifest.json'))

if __name__=='__main__':main()
