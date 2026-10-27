#!/usr/bin/env python3
"""Self-owned surgical DOCX OLE MathType editor (prototype v0.1).

This module understands Word's relationships/preview layers and MathType OLE2
CFB streams. It NEVER requires Word, MathType, EquationLab, or docx-equation.
Currently it can index, extract and replace an existing MathType object with a
real editable MathType donor from a DOCX. It cannot yet synthesize arbitrary
MTEF from LaTeX or perform syntax-aware edits inside MTEF v5.

Safety: fail closed, never in-place; untouched ZIP entries remain byte-identical
in uncompressed payload; validate complete output before atomic commit.
"""
from __future__ import annotations
import argparse
from copy import copy
from dataclasses import dataclass
import hashlib
import json
import posixpath
from pathlib import Path
import re
from tempfile import NamedTemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree
from ole_native import equation_native_summary

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
O = 'urn:schemas-microsoft-com:office:office'
V = 'urn:schemas-microsoft-com:vml'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'w': W, 'o': O, 'v': V, 'r': R, 'pr': PR}
DOC = 'word/document.xml'
RELS = 'word/_rels/document.xml.rels'
CTYPES = '[Content_Types].xml'
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'
CONTENT_TYPES = {
    '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
    '.wmf': 'image/x-wmf', '.emf': 'image/x-emf', '.gif': 'image/gif',
    '.tif': 'image/tiff', '.tiff': 'image/tiff', '.svg': 'image/svg+xml',
    '.bmp': 'image/bmp',
}


@dataclass(frozen=True)
class Formula:
    index: int
    paragraph: int
    object_rel: str
    object_part: str
    preview_rel: str
    preview_part: str
    prog_id: str
    shape_id: str
    shape_style: str
    context: str
    mtef_version: int
    object_sha256: str
    preview_sha256: str

    def as_dict(self):
        return {key: getattr(self, key) for key in self.__dataclass_fields__}


class FormulaDocument:
    def __init__(self, path: str | Path):
        self.path = Path(path).resolve()
        if self.path.suffix.lower() != '.docx':
            raise ValueError('Only unencrypted .docx is supported')
        with ZipFile(self.path) as z:
            if z.testzip():
                raise ValueError('Invalid input DOCX ZIP CRC')
            self.names = set(z.namelist())
            if any(n.startswith('_xmlsignatures/') for n in self.names):
                raise ValueError('Signed DOCX cannot be surgically changed without invalidating signature')
            for item in (DOC, RELS, CTYPES):
                if item not in self.names:
                    raise ValueError(f'Missing package part {item}')
            self.document = etree.fromstring(z.read(DOC))
            self.rels = etree.fromstring(z.read(RELS))
            self.rel_map = {e.get('Id'): e for e in self.rels}
            self.objects = list(self.document.xpath('.//o:OLEObject', namespaces=NS))
            self.formulas = []
            # Main story paragraph sequence. Header/footer objects are not currently edited.
            # Python lxml wrappers are not guaranteed to preserve id across XPath, so
            # use an XPath parent walk index based on structurally stable XML root.
            paras = self.document.xpath('.//w:body//w:p', namespaces=NS)
            for i, obj in enumerate(self.objects, 1):
                wrap = obj.getparent()
                if wrap is None or wrap.tag != '{%s}object' % W:
                    raise ValueError(f'Formula #{i} is not inside supported w:object wrapper')
                if len(wrap.xpath('./o:OLEObject', namespaces=NS)) != 1:
                    raise ValueError('Ambiguous MathType object grouping')
                images = wrap.xpath('.//v:imagedata', namespaces=NS)
                shapes = wrap.xpath('.//v:shape', namespaces=NS)
                if len(images) != 1 or len(shapes) != 1:
                    raise ValueError(f'Formula #{i} has no unique VML preview')
                obj_id = obj.get('{%s}id' % R)
                image_id = images[0].get('{%s}id' % R)
                if not obj_id or not image_id:
                    raise ValueError(f'Formula #{i} missing OLE/preview relationship')
                obj_part = self._resolve(obj_id)
                preview_part = self._resolve(image_id)
                if obj_part not in self.names or preview_part not in self.names:
                    raise ValueError(f'Formula #{i} target package part not present')
                prog_id = obj.get('ProgID') or ''
                if not prog_id.lower().startswith('equation.'):
                    raise ValueError(f'Formula #{i} not an Equation MathType OLE: {prog_id!r}')
                b = z.read(obj_part)
                native = equation_native_summary(b)
                if native['mtef_version'] not in (3, 5):
                    raise ValueError(f'Formula #{i} unsupported MTEF version: {native["mtef_version"]}')
                prev = z.read(preview_part)
                p = wrap
                while p is not None and p.tag != '{%s}p' % W:
                    p = p.getparent()
                number = next((j+1 for j, e in enumerate(paras) if e is p), 0)
                context = '' if p is None else ' '.join(''.join(p.xpath('.//w:t/text()', namespaces=NS)).split())[:200]
                self.formulas.append(Formula(i, number, obj_id, obj_part,
                                             image_id, preview_part, prog_id,
                                             shapes[0].get('id',''), shapes[0].get('style',''),
                                             context, native['mtef_version'],
                                             hashlib.sha256(b).hexdigest(),
                                             hashlib.sha256(prev).hexdigest()))

    def _resolve(self, relid: str) -> str:
        rel = self.rel_map.get(relid)
        if rel is None or rel.get('TargetMode') == 'External':
            raise ValueError(f'External or missing relationship {relid}')
        part = posixpath.normpath(posixpath.join('word', rel.get('Target','')))
        if not part.startswith('word/') or part.startswith('word/../') or part == 'word/':
            raise ValueError('Unsafe relationship target path')
        return part

    def by_index(self, index: int) -> Formula:
        if not (1 <= index <= len(self.formulas)):
            raise IndexError(f'Equation #{index} not found (total {len(self.formulas)})')
        return self.formulas[index-1]

    def _parts_used(self, target: str):
        # Some legitimate document relationships target ../customXml,
        # external URLs, etc.; only the matching internal part matters.
        return {rid for rid, rel in self.rel_map.items()
                if rel.get('TargetMode') != 'External' and
                posixpath.normpath(posixpath.join('word', rel.get('Target', ''))) == target}

    def _rid_references(self, rid: str):
        return len(self.document.xpath('.//*[@r:id="%s"]' % rid, namespaces=NS))

    def _preview_uses(self, part: str):
        return sum(self._rid_references(r) for r in self._parts_used(part))


def _serialize(root):
    return etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def _new_part_name(existing: set[str], directory: str, stem: str, extension: str):
    for k in range(1, 100000):
        name = f'{directory}/{stem}{k}{extension}'
        if name not in existing:
            existing.add(name)
            return name
    raise ValueError('Cannot allocate unused part name')


def _new_rid(rels):
    used = {e.get('Id') for e in rels}
    for k in range(1, 10000000):
        rid = f'rId{k}'
        if rid not in used:
            return rid
    raise ValueError('Cannot allocate unused relationship Id')


def _new_rel(rels, rid: str, original, part: str):
    rel = etree.SubElement(rels, '{%s}Relationship' % PR)
    rel.set('Id', rid)
    rel.set('Type', original.get('Type'))
    rel.set('Target', posixpath.relpath(part, 'word'))
    return rel


def _copy_display_dimensions(dst_object, src_object):
    # Preserve destination shape/object identity, replace only geometric sizing.
    for tag in ('dxaOrig', 'dyaOrig'):
        key = '{%s}%s' % (W, tag)
        if src_object.get(key) is not None:
            dst_object.set(key, src_object.get(key))
    dst_shape = dst_object.xpath('.//v:shape', namespaces=NS)[0]
    src_shape = src_object.xpath('.//v:shape', namespaces=NS)[0]
    a = dst_shape.get('style') or ''
    b = src_shape.get('style') or ''
    for prop in ('width', 'height'):
        pattern = re.compile(r'(?i)(?:^|;)\s*'+re.escape(prop)+r'\s*:\s*[^;]+')
        found = pattern.search(b)
        if not found:
            raise ValueError('Donor formula preview shape lacks width/height')
        dim = found.group().lstrip(';').strip().split(':',1)[1].strip()
        target_pattern = re.compile(r'(?i)(^|;)\s*'+re.escape(prop)+r'\s*:\s*[^;]+')
        if target_pattern.search(a):
            a = target_pattern.sub(lambda match: match.group(1) + prop+':'+dim, a, count=1)
        else:
            a = (a.rstrip(';')+';'+prop+':'+dim).lstrip(';')
    dst_shape.set('style',a)
    # Word represents the OLE equation's baseline via the surrounding run.
    # Carry over only this property, not the donor's full paragraph/run style.
    dst_run = dst_object.getparent()
    src_run = src_object.getparent()
    if dst_run is not None and src_run is not None and (
            dst_run.tag == '{%s}r' % W and src_run.tag == '{%s}r' % W):
        src_pr = src_run.find('{%s}rPr' % W)
        dst_pr = dst_run.find('{%s}rPr' % W)
        src_baseline = None if src_pr is None else src_pr.find('{%s}position' % W)
        if src_baseline is not None:
            if dst_pr is None:
                dst_pr = etree.Element('{%s}rPr' % W)
                dst_run.insert(0, dst_pr)
            dst_baseline = dst_pr.find('{%s}position' % W)
            if dst_baseline is None:
                dst_baseline = etree.SubElement(dst_pr, '{%s}position' % W)
            dst_baseline.set('{%s}val' % W, src_baseline.get('{%s}val' % W, '0'))
        elif dst_pr is not None:
            old = dst_pr.find('{%s}position' % W)
            if old is not None:
                dst_pr.remove(old)


def _ensure_content_type(root, suffix: str):
    ext = suffix.lower().lstrip('.')
    media_type = CONTENT_TYPES.get('.'+ext)
    if not media_type:
        raise ValueError(f'Unsupported donor preview extension: {suffix}')
    already = root.xpath('./ct:Default[@Extension="%s"]' % ext, namespaces={'ct':CT})
    if not already:
        etree.SubElement(root,'{%s}Default' % CT, Extension=ext, ContentType=media_type)


def replace_from_donor(input_docx, index, donor_docx, donor_index, output_docx, fit='donor'):
    """Object-level MathType replacement; uses a genuine existing MathType donor.

    Identical document roles are preserved (ObjectID, ShapeID, paragraph, captions,
    references); only the selected OLE bytes and preview are replaced or relinked.
    No arbitrary MathType formula synthesis happens in this function.
    """
    source = FormulaDocument(input_docx)
    donor = FormulaDocument(donor_docx)
    src = source.by_index(index)
    rep = donor.by_index(donor_index)
    if source.path == Path(output_docx).resolve() or donor.path == Path(output_docx).resolve():
        raise ValueError('Refusing in-place modification; output must be a new file')
    if fit not in ('donor','original'):
        raise ValueError('fit must be donor or original')
    if src.prog_id != rep.prog_id:
        raise ValueError(f'ProgID mismatch: target {src.prog_id}, donor {rep.prog_id}; refuse cross-version substitution')
    dst_ole_obj = source.objects[index-1]
    donor_ole_obj = donor.objects[donor_index-1]
    dst_wrap = dst_ole_obj.getparent()
    donor_wrap = donor_ole_obj.getparent()
    doc_root = source.document
    rels = source.rels
    ctypes = None
    changed = {}
    added = {}
    with ZipFile(source.path) as zin, ZipFile(donor.path) as dzip:
        if zin.testzip() or dzip.testzip():
            raise ValueError('Input/donor ZIP CRC check failed')
        new_ole = dzip.read(rep.object_part)
        new_preview = dzip.read(rep.preview_part)
        if equation_native_summary(new_ole)['mtef_version'] not in (3, 5):
            raise ValueError('Donor is not a supported MTEF v3/v5 formula')
        # Replace object bytes in place unless this specific package part is
        # referenced by multiple Word instances, in which case clone the part.
        existing = set(zin.namelist())
        parts_linked = source._parts_used(src.object_part)
        ref_count = sum(source._rid_references(rid) for rid in parts_linked)
        if ref_count == 1:
            changed[src.object_part] = new_ole
            selected_ole_part = src.object_part
        else:
            selected_ole_part = _new_part_name(existing,'word/embeddings','mathEdit', '.bin')
            added[selected_ole_part] = new_ole
            new_id = _new_rid(rels)
            _new_rel(rels, new_id, source.rel_map[src.object_rel], selected_ole_part)
            dst_ole_obj.set('{%s}id' % R,new_id)
        # If compatible and unique, overwrite original preview; otherwise
        # attach new media part and relationship, preserving any other media.
        ext_target = Path(src.preview_part).suffix.lower()
        ext_donor = Path(rep.preview_part).suffix.lower()
        if ext_target == ext_donor and source._preview_uses(src.preview_part) == 1:
            changed[src.preview_part] = new_preview
            selected_preview = src.preview_part
        else:
            selected_preview = _new_part_name(existing,'word/media','mathEdit',ext_donor)
            added[selected_preview] = new_preview
            new_id = _new_rid(rels)
            _new_rel(rels,new_id,source.rel_map[src.preview_rel],selected_preview)
            image = dst_wrap.xpath('.//v:imagedata', namespaces=NS)[0]
            image.set('{%s}id' % R, new_id)
            ctypes = etree.fromstring(zin.read(CTYPES))
            _ensure_content_type(ctypes,ext_donor)
        if fit == 'donor':
            _copy_display_dimensions(dst_wrap, donor_wrap)
        changed[DOC] = _serialize(doc_root)
        if selected_ole_part != src.object_part or selected_preview != src.preview_part:
            changed[RELS] = _serialize(rels)
        if ctypes is not None:
            changed[CTYPES] = _serialize(ctypes)
        out = Path(output_docx).resolve()
        out.parent.mkdir(parents=True,exist_ok=True)
        with NamedTemporaryFile(prefix='.ole-edit-', suffix='.docx', dir=out.parent, delete=False) as f:
            temp = Path(f.name)
        try:
            with ZipFile(temp,'w',ZIP_DEFLATED) as zout:
                for entry in zin.infolist():
                    zout.writestr(copy(entry), changed.get(entry.filename, zin.read(entry.filename)))
                for name, data in added.items():
                    zout.writestr(name,data)
            with ZipFile(temp) as trial:
                if trial.testzip():
                    raise ValueError('Output CRC audit failed')
                for name in zin.namelist():
                    if name not in changed and zin.read(name) != trial.read(name):
                        raise ValueError(f'Protected ZIP entry changed: {name}')
                # Exact audits for every OLE/preview not associated with target.
                original_entries = source.formulas
                for item in original_entries:
                    if item.index == index:
                        continue
                    if zin.read(item.object_part) != trial.read(item.object_part):
                        raise ValueError(f'Unedited MathType binary changed: #{item.index}')
                    if zin.read(item.preview_part) != trial.read(item.preview_part):
                        raise ValueError(f'Unedited MathType preview changed: #{item.index}')
            out_doc = FormulaDocument(temp)
            if len(out_doc.formulas) != len(source.formulas):
                raise ValueError('MathType formula count changed unexpectedly')
            result = out_doc.by_index(index)
            if result.object_sha256 != hashlib.sha256(new_ole).hexdigest():
                raise ValueError('Replacement MathType CFB does not match donor')
            if result.preview_sha256 != hashlib.sha256(new_preview).hexdigest():
                raise ValueError('Replacement preview does not match donor')
            temp.replace(out)
        except Exception:
            temp.unlink(missing_ok=True)
            raise
    return {
        'source': str(source.path), 'output':str(out), 'source_formula_count':len(source.formulas),
        'edited_index':index, 'donor':str(donor.path), 'donor_index':donor_index,
        'changed_entries':sorted(changed), 'added_entries':sorted(added),
        'ole_part':selected_ole_part, 'preview_part':selected_preview,
        'unchanged_equations_verified':len(source.formulas)-1,
        'mtef_version':rep.mtef_version,'windows_mathtype_editing_verified':False,
        'note':'A real MathType donor OLE was substituted; no LaTeX-to-MTEF conversion was performed.'
    }


def extract(input_docx, index, output_dir):
    doc=FormulaDocument(input_docx)
    item=doc.by_index(index)
    dest=Path(output_dir)
    dest.mkdir(parents=True,exist_ok=True)
    with ZipFile(doc.path) as z:
        (dest/'equation.ole').write_bytes(z.read(item.object_part))
        (dest/('preview'+Path(item.preview_part).suffix)).write_bytes(z.read(item.preview_part))
        # Extract exact MTEF as found in original native stream, not a new encoding.
        from ole_native import CompoundFile
        native = CompoundFile(z.read(item.object_part)).read('Equation Native')
        if native[:4] != b'\x1c\x00\x00\x00':
            raise ValueError('Unexpected EQNOLEFILEHDR length, cannot extract MTEF')
        (dest/'equation.mtef').write_bytes(native[28:])
    metadata=item.as_dict()
    (dest/'manifest.json').write_text(json.dumps(metadata,indent=2,ensure_ascii=False),encoding='utf-8')
    return {'output_dir':str(dest),'files':[p.name for p in dest.iterdir()]}


def replace_from_latex(input_docx, index, latex, output_docx):
    """Experimental surgical rewrite from a supported LaTeX subset.

    This rewrites real MTEF v5 mathematical records and rebuilds the CFB with
    our own writer; the target's existing style preamble and OLE non-equation
    streams are preserved. The VML preview is regenerated independently.
    Unsupported constructs and legacy v3 formulas fail closed.
    """
    from ole_cfb_writer import repack_stream
    from mtef_subset import make_mtef, validate_subset_mtef
    from native_preview import preview_png
    import struct

    source=FormulaDocument(input_docx)
    target=source.by_index(index)
    if target.prog_id != 'Equation.DSMT4' or target.mtef_version!=5:
        raise ValueError('Native LaTeX rewrite supports Equation.DSMT4 / MTEF v5 only')
    destination=Path(output_docx).resolve()
    if source.path==destination:
        raise ValueError('Refusing in-place editing of original DOCX')
    with ZipFile(source.path) as zin:
        old_ole=zin.read(target.object_part)
        from ole_native import CompoundFile
        original=CompoundFile(old_ole)
        native=original.read('Equation Native')
        style_mtef=native[28:]
        app_end=style_mtef.index(b'\x00',5)
        inline=bool(style_mtef[app_end+1]&1)
        encoded=make_mtef(latex,style_mtef,inline=inline)
        structure=validate_subset_mtef(encoded)
        header=bytearray(native[:28])
        struct.pack_into('<I',header,8,len(encoded))
        new_ole=repack_stream(old_ole,'Equation Native',bytes(header)+encoded)
        new_image,width_pt,height_pt=preview_png(latex)
        # Ensure *all* non-native OLE streams still contain their original data.
        updated=CompoundFile(new_ole)
        for name in original.entries:
            if name!='Equation Native' and original.read(name)!=updated.read(name):
                raise AssertionError(f'Protected CFB stream altered: {name!r}')
        if updated.read('Equation Native')[28:]!=encoded:
            raise AssertionError('MTEF CFB round-trip failed')

        changed={};added={}
        existing=set(zin.namelist())
        rels=source.rels
        dst=source.objects[index-1]
        wrap=dst.getparent()
        parts_linked=source._parts_used(target.object_part)
        ref_count=sum(source._rid_references(rid) for rid in parts_linked)
        if ref_count == 1:
            changed[target.object_part]=new_ole
            object_part=target.object_part
        else:
            object_part=_new_part_name(existing,'word/embeddings','nativeEdited','.bin')
            added[object_part]=new_ole
            rid=_new_rid(rels)
            _new_rel(rels,rid,source.rel_map[target.object_rel],object_part)
            dst.set('{%s}id'%R,rid)
        # If original preview is the same PNG type and is unique, overwrite.
        # Otherwise allocate a new relationship and PNG image.
        if target.preview_part.endswith('.png') and source._preview_uses(target.preview_part)==1:
            image_part=target.preview_part
            changed[image_part]=new_image
        else:
            image_part=_new_part_name(existing,'word/media','nativeEdited','.png')
            added[image_part]=new_image
            rid=_new_rid(rels)
            _new_rel(rels,rid,source.rel_map[target.preview_rel],image_part)
            image=wrap.xpath('.//v:imagedata',namespaces=NS)[0]
            image.set('{%s}id'%R,rid)
            types=etree.fromstring(zin.read(CTYPES))
            _ensure_content_type(types,'.png')
            changed[CTYPES]=_serialize(types)
        shape=wrap.xpath('.//v:shape',namespaces=NS)[0]
        style=shape.get('style') or ''
        for k,value in [('width',width_pt),('height',height_pt)]:
            repl=f'{k}:{value:.2f}pt'
            rx=re.compile(r'(?i)(^|;)\s*'+k+r'\s*:[^;]+')
            style=rx.sub(lambda m:m.group(1)+repl,style,count=1) if rx.search(style) else (style+';'+repl)
        shape.set('style',style)
        wrap.set('{%s}dxaOrig'%W,str(int(width_pt*20)))
        wrap.set('{%s}dyaOrig'%W,str(int(height_pt*20)))
        changed[DOC]=_serialize(source.document)
        if object_part!=target.object_part or image_part!=target.preview_part:
            changed[RELS]=_serialize(rels)
        destination.parent.mkdir(parents=True,exist_ok=True)
        with NamedTemporaryFile(prefix='.math-native-',suffix='.docx',dir=destination.parent,delete=False) as handle:
            tmp=Path(handle.name)
        try:
            with ZipFile(tmp,'w',ZIP_DEFLATED) as out:
                for entry in zin.infolist():
                    out.writestr(copy(entry),changed.get(entry.filename,zin.read(entry.filename)))
                for name,data in added.items():out.writestr(name,data)
            with ZipFile(tmp) as trial:
                if trial.testzip():raise ValueError('Output ZIP CRC error')
                for name in zin.namelist():
                    if name not in changed and trial.read(name)!=zin.read(name):
                        raise ValueError('Protected original part altered: '+name)
                for f in source.formulas:
                    if f.index==index:continue
                    if trial.read(f.object_part)!=zin.read(f.object_part):
                        raise ValueError(f'Unedited OLE modified: #{f.index}')
                    if trial.read(f.preview_part)!=zin.read(f.preview_part):
                        raise ValueError(f'Unedited preview modified: #{f.index}')
            checked=FormulaDocument(tmp)
            if len(checked.formulas)!=len(source.formulas):
                raise ValueError('Number of MathType objects changed')
            updated_f=checked.by_index(index)
            if updated_f.object_sha256!=hashlib.sha256(new_ole).hexdigest():
                raise ValueError('Generated OLE object failed integrity check')
            if updated_f.preview_sha256!=hashlib.sha256(new_image).hexdigest():
                raise ValueError('Generated formula preview failed integrity check')
            tmp.replace(destination)
        except Exception:
            tmp.unlink(missing_ok=True)
            raise
    return {'input':str(source.path),'output':str(destination),'formula_index':index,
            'latex':latex,'changed_ole':object_part,'changed_preview':image_part,
            'mtef_glyphs':structure['glyphs'],'mtef_templates':structure['templates'],
            'unchanged_equations_verified':len(source.formulas)-1,
            'windows_activation_verified':False}

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    sub=ap.add_subparsers(dest='command', required=True)
    ls=sub.add_parser('list',help='List genuine MathType OLE objects')
    ls.add_argument('docx')
    ls.add_argument('--out',help='Write JSON inventory')
    ex=sub.add_parser('extract',help='Export an OLE object, preview and MTEF bytes')
    ex.add_argument('docx'); ex.add_argument('--index',type=int,required=True)
    ex.add_argument('--out-dir',required=True)
    rp=sub.add_parser('replace',help='Replace one OLE equation using existing DOCX donor equation')
    rp.add_argument('docx'); rp.add_argument('--index',type=int,required=True)
    rp.add_argument('--donor',required=True); rp.add_argument('--donor-index',type=int,required=True)
    rp.add_argument('-o','--output',required=True)
    rp.add_argument('--fit',choices=('donor','original'),default='donor')
    lp=sub.add_parser('replace-latex',help='EXPERIMENTAL: synthesize MTEF v5 from a safe LaTeX subset')
    lp.add_argument('docx'); lp.add_argument('--index',type=int,required=True)
    lp.add_argument('--latex',required=True); lp.add_argument('-o','--output',required=True)
    args=ap.parse_args()
    if args.command=='list':
        d=FormulaDocument(args.docx)
        result={'document':str(d.path),'count':len(d.formulas),
                'formulas':[f.as_dict() for f in d.formulas]}
        if args.out:
            Path(args.out).write_text(json.dumps(result,indent=2,ensure_ascii=False),encoding='utf-8')
            print(f'Listed {len(d.formulas)} MathType objects -> {args.out}')
        else:
            print(json.dumps(result,indent=2,ensure_ascii=False))
    elif args.command=='extract':
        result=extract(args.docx,args.index,args.out_dir)
        print(json.dumps(result,indent=2,ensure_ascii=False))
    elif args.command=='replace-latex':
        result=replace_from_latex(args.docx,args.index,args.latex,args.output)
        print(json.dumps(result,indent=2,ensure_ascii=False))
    else:
        result=replace_from_donor(args.docx,args.index,args.donor,args.donor_index,args.output,args.fit)
        print(json.dumps(result,indent=2,ensure_ascii=False))

if __name__=='__main__':
    main()
