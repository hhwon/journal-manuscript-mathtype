#!/usr/bin/env python3
"""Compare a DOCX's key paragraph/style settings to the user's reference manuscript."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path
from docx import Document


def extract(path):
    doc=Document(path)
    sec=doc.sections[0]
    result={
        'paper_cm':[round(sec.page_width.cm,2),round(sec.page_height.cm,2)],
        'margins_cm':[round(x.cm,2) for x in (sec.left_margin,sec.right_margin,sec.top_margin,sec.bottom_margin)],
        'styles':{},
    }
    for name in ('MS Title','MS Body','Heading 1','Heading 2','MS Figure Caption',
                 'MS Table Caption','MS Table Cell','MS References'):
        if name not in doc.styles:
            result['styles'][name]={'error':'missing'}
            continue
        st=doc.styles[name]
        fmt=st.paragraph_format
        result['styles'][name]={
          'font':st.font.name,
          'font_size_pt':round(st.font.size.pt,2) if st.font.size else None,
          'bold':st.font.bold,
          'italic':st.font.italic,
          'line_spacing':fmt.line_spacing,
          'align':str(fmt.alignment),
          'first_indent_cm':round(fmt.first_line_indent.cm,2) if fmt.first_line_indent else 0,
        }
    return result


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('docx')
    ap.add_argument('--json')
    args=ap.parse_args()
    result=extract(args.docx)
    issues=[]
    if result['paper_cm']!=[21.59,27.94]:issues.append('Page size differs from reference Letter')
    if result['margins_cm']!=[2.54]*4:issues.append('Margins differ from reference 2.54 cm')
    if result['styles']['MS Body'].get('font')!='Times New Roman':issues.append('Body font is not Times New Roman')
    if result['styles']['MS Body'].get('font_size_pt')!=12:issues.append('Body font is not 12 pt')
    result['warnings']=issues
    print(json.dumps(result,ensure_ascii=False,indent=2))
    if args.json:Path(args.json).write_text(json.dumps(result,indent=2),encoding='utf-8')
    return 0 if not issues else 2

if __name__=='__main__':sys.exit(main())
