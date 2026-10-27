#!/usr/bin/env python3
"""Integration tests against real user-supplied MathType DOCX fixtures.

Run: python scripts/test_ole_editor.py --source old.docx --donor other.docx
The manuscripts are intentionally NOT bundled with the Skill.
"""
from __future__ import annotations
import argparse
from copy import copy
from io import BytesIO
from pathlib import Path
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
from PIL import Image
from ole_editor import (FormulaDocument, replace_from_donor, extract, DOC, RELS, CTYPES,
                        R, NS, PR, CT)


def altered_docx(input_docx, output, action, index=2):
    """Only for tests: create path-sharing or mixed-preview-format fixtures."""
    src = FormulaDocument(input_docx)
    root = src.document
    rels = src.rels
    idx = index-1
    changes = {}
    newparts = {}
    if action=='shared_ole':
        a, b = src.objects[idx], src.objects[idx+1]
        a.set('{%s}id'%R,b.get('{%s}id'%R))
    elif action=='shared_preview':
        a = src.objects[idx].getparent().xpath('.//v:imagedata',namespaces=NS)[0]
        b = src.objects[idx+1].getparent().xpath('.//v:imagedata',namespaces=NS)[0]
        a.set('{%s}id'%R,b.get('{%s}id'%R))
    elif action=='png_donor':
        a=src.objects[idx+1].getparent().xpath('.//v:imagedata',namespaces=NS)[0]
        rid='rId2000000'
        if rid in src.rel_map: raise AssertionError('Synthetic RID collision')
        pic=Image.new('RGB',(200,48),'white')
        output_png=BytesIO()
        pic.save(output_png,format='PNG')
        part='word/media/ole-editor-test.png'
        if part in src.names: raise AssertionError('Synthetic part collision')
        newparts[part]=output_png.getvalue()
        etree.SubElement(rels,'{%s}Relationship'%PR,Id=rid,
            Type=src.rel_map[src.by_index(index+1).preview_rel].get('Type'),
            Target='media/ole-editor-test.png')
        a.set('{%s}id'%R,rid)
        changes[RELS]=etree.tostring(rels,encoding='UTF-8',xml_declaration=True,standalone=True)
        with ZipFile(input_docx) as z:
            c=etree.fromstring(z.read(CTYPES))
        if not c.xpath('./ct:Default[@Extension="png"]',namespaces={'ct':CT}):
            etree.SubElement(c,'{%s}Default'%CT, Extension='png', ContentType='image/png')
        changes[CTYPES]=etree.tostring(c,encoding='UTF-8',xml_declaration=True,standalone=True)
    else: raise ValueError(action)
    changes[DOC]=etree.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
    with ZipFile(input_docx) as z, ZipFile(output,'w',ZIP_DEFLATED) as out:
        for item in z.infolist():
            out.writestr(copy(item),changes.get(item.filename,z.read(item.filename)))
        for key,data in newparts.items(): out.writestr(key,data)
    return output


def run(source,donor):
    orig=FormulaDocument(source)
    alternative=FormulaDocument(donor)
    assert len(orig.formulas)>8 and len(alternative.formulas)>5
    assert all(f.mtef_version in (3,5) for f in orig.formulas)
    print('PASS list',len(orig.formulas),'formula objects, legacy v3:',
          sum(f.mtef_version==3 for f in orig.formulas))
    with tempfile.TemporaryDirectory() as t:
        t=Path(t)
        # Same-document v5 replacement.
        output=t/'same.docx'
        r=replace_from_donor(source,2,source,3,output)
        assert r['unchanged_equations_verified']==len(orig.formulas)-1
        assert len(r['changed_entries'])==3
        print('PASS same DOCX donor, untouched OLE objects exact')
        # Extraction gives a valid native stream payload.
        files=extract(source,2,t/'extract')
        assert (t/'extract'/'equation.mtef').read_bytes()[0] in (3,5)
        print('PASS direct MTEF extraction')
        # Legacy formula test if present (must retain same ProgID).
        legacy=[f.index for f in orig.formulas if f.prog_id=='Equation.3']
        if len(legacy)>1:
            r=replace_from_donor(source,legacy[0],source,legacy[1],t/'old.docx')
            assert r['mtef_version']==3
            print('PASS Equation.3 / MTEF v3 legacy branch')
        # Cross-document donor should preserve everything outside target.
        other=next((f for f in alternative.formulas if f.prog_id==orig.by_index(2).prog_id),None)
        if other:
            r=replace_from_donor(source,2,donor,other.index,t/'cross.docx')
            assert r['unchanged_equations_verified']==len(orig.formulas)-1
            print('PASS cross-document donor formula')
        # Shared old object must be cloned instead of overwriting the peer.
        shared=altered_docx(source,t/'shared.docx','shared_ole')
        r=replace_from_donor(shared,2,source,4,t/'shared-out.docx')
        assert any('mathEdit' in n for n in r['added_entries'])
        print('PASS shared OLE relationship isolation')
        sharedp=altered_docx(source,t/'shared-preview.docx','shared_preview')
        r=replace_from_donor(sharedp,2,source,4,t/'shared-preview-out.docx')
        assert any('mathEdit' in n for n in r['added_entries'])
        print('PASS shared preview relationship isolation')
        # PNG branch exercises adding image part, relationship and content type.
        png=altered_docx(source,t/'png-donor.docx','png_donor')
        r=replace_from_donor(source,2,png,3,t/'png-out.docx')
        assert r['preview_part'].endswith('.png')
        assert any(n.endswith('.png') for n in r['added_entries'])
        print('PASS changed preview media type: WMF -> PNG')
        if legacy:
            try: replace_from_donor(source,legacy[0],source,2,t/'wrong.docx')
            except ValueError as err: assert 'ProgID mismatch' in str(err)
            else: raise AssertionError('MTEF3/5 ProgID mixing should fail')
            print('PASS incompatible ProgID refused')
        try: replace_from_donor(source,2,source,3,source)
        except ValueError as err: assert 'in-place' in str(err)
        else: raise AssertionError('Should refuse in-place patch')
        print('PASS in-place overwrite refused')
    print('All integration checks passed. Windows editing remains unverified.')


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--source',required=True)
    p.add_argument('--donor',required=True)
    a=p.parse_args()
    run(a.source,a.donor)

if __name__=='__main__': main()
