#!/usr/bin/env python3
"""Generate a Word/MathType *native* font semantic role regression suite.

No Office / MathType is required for Linux generation. Windows application
rendering and editing are a separate manual acceptance step.
"""
from pathlib import Path
import argparse
from build_native_ole_test import build

FORMULAS_FONT=[
 ('A1','Latin variables (italic)',r'E=kT','synthesized'),
 ('A2','Greek variables (Symbol italic)',r'\sigma=E\varepsilon','synthesized'),
 ('A3','Tensor subscripts (Variable)',r'\sigma_{ij}=C_{ijkl}\varepsilon_{kl}','synthesized'),
 ('A4','Raised indices and digits (Number)',r'E=mc^{2}','synthesized'),
 ('A5','Function names (Function upright)',r'\sin(\theta)+\cos(\phi)','synthesized'),
 ('A6','Function-start boundary (Function upright)',r'\exp(x)=\ln(y)','synthesized'),
 ('A7','Roman characters and bold vector',r'\mathrm{Re}=\mathbf{u}','synthesized'),
 ('A8','Uppercase Greek (Symbol upright)',r'\Gamma=\Sigma+E','synthesized'),
 ('A9','High-order tensor expression',r'\sigma_{ij}=C_{ijkl}\varepsilon_{kl}+D_{ijklmn}\varepsilon_{kl,mn}','synthesized'),
 ('A10','Fraction template',r'\frac{1}{2}','synthesized'),
 ('A11','Root template',r'\sqrt{x^{2}+y^{2}}','synthesized'),
 ('A12','Explicit math operator (Function)',r'\operatorname{det}=x','synthesized'),
 ('A13','Bold variable with index',r'\mathbf{u}_i','synthesized'),
 ('B1','Native CHAR record rewrite',r'\tau=E\varepsilon','patched'),
 ('C1','Original unchanged MathType control',None,'original'),
]

if __name__=='__main__':
    ap=argparse.ArgumentParser()
    ap.add_argument('--source',required=True)
    ap.add_argument('-o','--output',required=True)
    args=ap.parse_args()
    build(Path(args.source),Path(args.output),formulas=FORMULAS_FONT)
