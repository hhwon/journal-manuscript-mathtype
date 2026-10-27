#!/usr/bin/env python3
"""Recreate MathType 7 native numbering in an existing native-OLE test DOCX.

Extracts the exact numbering field tree from a user-provided Word/MathType
reference, including MACROBUTTON MTPlaceRef and SEQ MTEqn records. The native
Word field structure is not equivalent to manually typed `(n)` or SEQ Equation.

Experimental: MathType Word add-in interaction (insert reference, renumber)
can only be accepted by running actual Microsoft Word + MathType 7 on Windows.
"""
from __future__ import annotations
import argparse
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W,'o':'urn:schemas-microsoft-com:office:office'}
Q=lambda x:f'{{{W}}}{x}'


def serialise(root):
    return etree.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)


def number_field_template(ref_root):
    paragraphs=ref_root.xpath('//w:p[.//w:instrText[contains(text(),"MTPlaceRef")]]', namespaces=NS)
    for para in paragraphs:
        runs=para.xpath('./w:r', namespaces=NS)
        start=next((i for i,r in enumerate(runs) if 'MACROBUTTON MTPlaceRef' in ''.join(r.xpath('.//w:instrText/text()',namespaces=NS))),None)
        if start is None or start==0:continue
        start-=1 # field-begin run
        number=0
        end=None
        for i in range(start,len(runs)):
            if runs[i].xpath('./w:fldChar[@w:fldCharType="begin"]',namespaces=NS):number+=1
            if runs[i].xpath('./w:fldChar[@w:fldCharType="end"]',namespaces=NS):
                number-=1
                if number==0:
                    end=i+1
                    break
        if end is None:continue
        chosen=runs[start:end]
        codes=' '.join(i.text or '' for r in chosen for i in r.xpath('.//w:instrText',namespaces=NS))
        if 'SEQ MTEqn \\h' not in codes or 'SEQ MTEqn \\c' not in codes:
            continue
        return chosen
    raise RuntimeError('No authentic MathType MTPlaceRef+MTEqn compound field in reference')


def _word_run(tag=None,text=None):
    run=etree.Element(Q('r'))
    if tag=='field':
        etree.SubElement(run,Q('fldChar'),{Q('fldCharType'):text})
    elif tag=='instr':
        el=etree.SubElement(run,Q('instrText'))
        el.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        el.text=text
    elif tag=='text':
        el=etree.SubElement(run,Q('t'));el.text=text
    return run


def cloned_number_runs(template,num,mode='exact'):
    """Canonical complex fields, with same MathType instructions as source.

    The reference DOCX has legacy w:instrText caches for numeric results,
    which LibreOffice displays as field-code corruption. Reconstruct proper
    OOXML field separators and w:t caches rather than propagating that defect.
    """
    # Validate the reference contains MathType's actual macro/SEQ instructions.
    codes=' '.join(e.text or '' for r in template for e in r.xpath('.//w:instrText',namespaces=NS))
    for token in ('MACROBUTTON MTPlaceRef','SEQ MTEqn \\h','SEQ MTEqn \\c'):
        if token not in codes:raise ValueError('Not the MathType reference sequence: '+token)
    if mode=='exact':
        # Preserve all native field runs byte-for-byte except cached label,
        # using exactly what the real MathType Word ribbon produced.
        result=[deepcopy(r) for r in template]
        updated=0
        for r in result:
            for instr in r.xpath('.//w:instrText',namespaces=NS):
                if (instr.text or '').strip()=='1':
                    instr.text=str(num)
                    updated+=1
        if updated!=1:raise RuntimeError('Unexpected native field-number cache')
        return result
    if mode!='canonical':raise ValueError('field mode must be exact or canonical')
    result=[
        _word_run('field','begin'),
        _word_run('instr',' MACROBUTTON MTPlaceRef \\* MERGEFORMAT '),
        _word_run('field','separate'),
        _word_run('field','begin'),
        _word_run('instr',' SEQ MTEqn \\h \\* MERGEFORMAT '),
        _word_run('field','separate'),
        _word_run('field','end'),
        _word_run('text','('),
        _word_run('field','begin'),
        _word_run('instr',' SEQ MTEqn \\c \\* Arabic \\* MERGEFORMAT '),
        _word_run('field','separate'),
        _word_run('text',str(num)),
        _word_run('field','end'),
        _word_run('text',')'),
        _word_run('field','end'),
    ]
    return result


def import_styles(target,reference):
    existing={s.get(Q('styleId')) for s in target if s.tag==Q('style')}
    copied=[]
    for sid in ('MTDisplayEquation','MTDisplayEquationChar','MTEquationSection'):
        rows=reference.xpath('./w:style[@w:styleId="%s"]'%sid,namespaces=NS)
        if rows and sid not in existing:
            dest=deepcopy(rows[0])
            target.append(dest)
            copied.append(sid)
    style=target.xpath('./w:style[@w:styleId="MTDisplayEquation"]',namespaces=NS)
    if len(style)!=1:raise RuntimeError('Reference missing MTDisplayEquation Word style')
    return copied


def make_native_number_para(row,template,num,mode):
    # existing row contains a full real OLE w:object with a valid v:imagedata and
    # o:OLEObject rel. Copy, do NOT regenerate or change the OLE binary/image.
    center=row.xpath('./w:tr/w:tc[2]/w:p',namespaces=NS)[0]
    objects=center.xpath('./w:r[w:object]',namespaces=NS)
    if len(objects)!=1:raise RuntimeError('Expected one native OLE run in formula row')
    para=etree.Element(Q('p'))
    prop=etree.SubElement(para,Q('pPr'))
    etree.SubElement(prop,Q('pStyle'),{Q('val'):'MTDisplayEquation'})
    etree.SubElement(prop,Q('spacing'),{Q('before'):'100',Q('after'):'140'})
    etree.SubElement(prop,Q('keepTogether'))
    tabrun=etree.SubElement(para,Q('r'))
    etree.SubElement(tabrun,Q('tab'))
    para.append(deepcopy(objects[0]))
    tabrun=etree.SubElement(para,Q('r'))
    etree.SubElement(tabrun,Q('tab'))
    for run in cloned_number_runs(template,num,mode=mode):
        para.append(run)
    return para


def update_tabs(styles_xml,document_xml):
    sect=document_xml.xpath('./w:body/w:sectPr',namespaces=NS)
    if not sect:raise RuntimeError('Document has no section properties')
    s=sect[0]
    pg=s.xpath('./w:pgSz',namespaces=NS)[0]
    mar=s.xpath('./w:pgMar',namespaces=NS)[0]
    width=int(pg.get(Q('w')))-int(mar.get(Q('left')))-int(mar.get(Q('right')))
    style=styles_xml.xpath('./w:style[@w:styleId="MTDisplayEquation"]',namespaces=NS)[0]
    ppr=style.find(Q('pPr'))
    if ppr is None:ppr=etree.SubElement(style,Q('pPr'))
    tabs=ppr.find(Q('tabs'))
    if tabs is None:tabs=etree.SubElement(ppr,Q('tabs'))
    for old in list(tabs):tabs.remove(old)
    etree.SubElement(tabs,Q('tab'),{Q('val'):'center',Q('pos'):str(width//2)})
    etree.SubElement(tabs,Q('tab'),{Q('val'):'right',Q('pos'):str(width)})
    return {'text_width_twips':width,'center_tab_twips':width//2,'right_tab_twips':width}


def upgrade(input_docx,reference_docx,output_docx,mode="exact"):
    input_docx,reference_docx,output_docx=map(Path,(input_docx,reference_docx,output_docx))
    with ZipFile(reference_docx) as z:
        refdoc=etree.fromstring(z.read('word/document.xml'))
        refstyles=etree.fromstring(z.read('word/styles.xml'))
    tmpl=number_field_template(refdoc)
    with ZipFile(input_docx) as z:
        doc=etree.fromstring(z.read('word/document.xml'))
        styles=etree.fromstring(z.read('word/styles.xml'))
        imported=import_styles(styles,refstyles)
        tabs=update_tabs(styles,doc)
        # Source test document stores all OLE in one-row formula tables.
        # Replace *only* those tables with MathType-numbered paragraphs.
        tables=doc.xpath('./w:body/w:tbl',namespaces=NS)
        targets=[]
        for table in tables:
            if len(table.xpath('.//o:OLEObject',namespaces=NS))==1:
                targets.append(table)
        if not targets:raise RuntimeError('No target OLE formula tables found')
        for num,row in enumerate(targets,1):
            new_para=make_native_number_para(row,tmpl,num,mode)
            row.getparent().replace(row,new_para)
        changed={'word/document.xml':serialise(doc),'word/styles.xml':serialise(styles)}
        output_docx.parent.mkdir(parents=True,exist_ok=True)
        with ZipFile(output_docx,'w',ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item,changed.get(item.filename,z.read(item.filename)))
    # Fail closed on binary changes and non-native numbering.
    with ZipFile(input_docx) as a,ZipFile(output_docx) as b:
        for name in a.namelist():
            if name.startswith('word/embeddings/') or name.startswith('word/media/'):
                if a.read(name)!=b.read(name):raise AssertionError('Unexpected binary change: '+name)
        text=b.read('word/document.xml').decode('utf-8')
        for expected in ('MACROBUTTON MTPlaceRef','SEQ MTEqn \\h','SEQ MTEqn \\c'):
            if text.count(expected)!=len(targets):
                raise AssertionError(f'Native numbering count mismatch {expected}')
    print(f'OK {output_docx}: {len(targets)} authentic MathType numbering fields, styles {imported}, tabs {tabs}')


def main():
    p=argparse.ArgumentParser()
    p.add_argument('input',type=Path)
    p.add_argument('--reference',required=True,type=Path)
    p.add_argument('-o','--output',required=True,type=Path)
    p.add_argument('--mode',choices=['exact','canonical'],default='exact')
    x=p.parse_args()
    upgrade(x.input,x.reference,x.output,mode=x.mode)

if __name__=='__main__':main()
