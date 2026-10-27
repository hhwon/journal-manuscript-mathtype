#!/usr/bin/env python3
"""Run a corpus of complex formulas through the installed native MathType backend.

The generated file is a regression fixture and must not be submitted as a paper.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path

from build_paper import make_doc

ROOT=Path(__file__).resolve().parent.parent


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('-o','--output',default='mathtype-regression.docx')
    ap.add_argument('--chrome-path')
    args=ap.parse_args()
    corpus=json.loads((ROOT/'examples/math-regression.json').read_text())
    spec={
      'title':'MathType equation regression suite - not for journal submission',
      'authors':[],
      'abstract':'This technical file checks editable OLE conversion of diverse mathematical structures.',
      'keywords':['Regression'],
      'sections':[{'heading':'1. Equation regression',
        'blocks':[x for obj in corpus['equations'] for x in (
           {'type':'paragraph','text':obj['name']},
           {'type':'equation','latex':obj['latex'],'id':'eq:'+obj['name']}
        )]}]
    }
    result=make_doc(spec,Path(args.output).resolve(),False,args.chrome_path,True)
    print(json.dumps(result,indent=2))

if __name__=='__main__':main()
