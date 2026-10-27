#!/usr/bin/env python3
"""Unset inherited body indents on numbered MathType display equations.

Docx-equation's numbered paragraph helper emits tab stops but not w:ind. A
first-line indent on the document Normal style must not shift equation centres.
"""
from pathlib import Path
from copy import copy
from tempfile import NamedTemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
O='urn:schemas-microsoft-com:office:office'
NS={'w':W,'o':O}

def repair(path):
    path=Path(path)
    with ZipFile(path) as zin:
        root=etree.fromstring(zin.read('word/document.xml'))
        changed=0
        for p in root.xpath('.//w:p',namespaces=NS):
            if not p.xpath('.//o:OLEObject',namespaces=NS):
                continue
            if not p.xpath('./w:pPr/w:tabs/w:tab[@w:val="center"]',namespaces=NS):
                continue
            ppr=p.find(f'{{{W}}}pPr')
            pstyle=ppr.find(f'{{{W}}}pStyle')
            if pstyle is None:
                pstyle=etree.Element(f'{{{W}}}pStyle')
                ppr.insert(0,pstyle)
            pstyle.set(f'{{{W}}}val','MS Equation')
            ind=ppr.find(f'{{{W}}}ind')
            if ind is None:
                ind=etree.SubElement(ppr,f'{{{W}}}ind')
            ind.set(f'{{{W}}}left','0')
            ind.set(f"{{{W}}}firstLine","0")
            changed+=1
        if not changed:return 0
        xml=etree.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
        with NamedTemporaryFile(dir=path.parent,prefix='.mathfix-',suffix='.docx',delete=False) as tmp:
            temp=Path(tmp.name)
        try:
            with ZipFile(temp,'w',compression=ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    zout.writestr(copy(item),xml if item.filename=='word/document.xml' else zin.read(item.filename))
        except Exception:
            temp.unlink(missing_ok=True)
            raise
    temp.replace(path)
    return changed

if __name__=='__main__':
    import sys
    print('Adjusted display equations:',repair(sys.argv[1]))
