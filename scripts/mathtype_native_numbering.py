#!/usr/bin/env python3
"""Turn the builder's Word equation numbers into MathType's own numbering fields.

The builder numbers display equations with Word fields: ``(SEQ Eq)`` or a fixed
label such as ``(A.1)``, each inside a bookmark, and cites them with
``REF <bookmark> \\h``. MathType's add-in does not recognise these. This pass
rewrites them in the form MathType's "Insert Right-Numbered Equation" and
"Insert Equation Reference" write (read from Word documents numbered by MathType):

* number:  MACROBUTTON MTPlaceRef { SEQ MTEqn \\h } [ZEqnNum bookmark] ( {SEQ MTEqn \\c} )
  with ``{SEQ MTChap \\c \\* Alphabetic}.`` before the counter for a lettered label;
  the display text of a MACROBUTTON is its instruction text, so the parentheses and
  cached numbers are ``w:instrText`` and there is no field separator;
* reference: GOTOBUTTON ZEqnNum { REF ZEqnNum \\* Charformat \\! } with the cached ``(n)``;
* section break (hidden, style MTEquationSection): MACROBUTTON MTEditEquationSection2
  "Equation Chapter c Section 1" { SEQ MTEqn \\r \\h } { SEQ MTSec \\r 1 \\h } { SEQ MTChap \\r c \\h },
  one at the start of the document and one before the first equation of each lettered group.

Display paragraphs get the style MTDisplayEquation (a copy of their current style,
so the layout does not change). OLE parts are not touched.
"""
from __future__ import annotations

import hashlib
import re
from copy import copy, deepcopy
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
O = 'urn:schemas-microsoft-com:office:office'
NS = {'w': W, 'o': O}
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def q(tag):
    return f'{{{W}}}{tag}'


def _run(rpr, kind, value=None, no_proof=False, rstyle=None):
    r = etree.Element(q('r'))
    if rpr is not None or no_proof or rstyle:
        p = deepcopy(rpr) if rpr is not None else etree.Element(q('rPr'))
        if rstyle:
            etree.SubElement(p, q('rStyle')).set(q('val'), rstyle)
            p.insert(0, p[-1])
        if no_proof:
            etree.SubElement(p, q('noProof'))
        r.append(p)
    if kind == 'fld':
        etree.SubElement(r, q('fldChar')).set(q('fldCharType'), value)
    elif kind == 'instr':
        e = etree.SubElement(r, q('instrText'))
        e.set(XML_SPACE, 'preserve')
        e.text = value
    elif kind == 'tab':
        etree.SubElement(r, q('tab'))
    return r


def _simple(rpr, instr, cached):
    f = etree.Element(q('fldSimple'))
    f.set(q('instr'), instr)
    f.append(_run(rpr, 'instr', cached, no_proof=True))
    return f


def _bookmark(tag, bid, name=None):
    b = etree.Element(q(tag))
    b.set(q('id'), bid)
    if name:
        b.set(q('name'), name)
    return b


def number_nodes(rpr, bookmark, bid, letter, number):
    nodes = [_run(rpr, 'fld', 'begin'),
             _run(rpr, 'instr', ' MACROBUTTON MTPlaceRef \\* MERGEFORMAT '),
             _run(rpr, 'fld', 'begin'),
             _run(rpr, 'instr', ' SEQ MTEqn \\h \\* MERGEFORMAT '),
             _run(rpr, 'fld', 'end'),
             _bookmark('bookmarkStart', bid, bookmark),
             _run(rpr, 'instr', '(')]
    if letter:
        nodes += [_simple(rpr, ' SEQ MTChap \\c \\* Alphabetic \\* MERGEFORMAT ', letter),
                  _run(rpr, 'instr', '.')]
    nodes += [_simple(rpr, ' SEQ MTEqn \\c \\* Arabic \\* MERGEFORMAT ', str(number)),
              _run(rpr, 'instr', ')'),
              _bookmark('bookmarkEnd', bid),
              _run(rpr, 'fld', 'end')]
    return nodes


def section_nodes(chapter):
    out = [_run(None, 'fld', 'begin'),
           _run(None, 'instr', ' MACROBUTTON MTEditEquationSection2 '),
           _run(None, 'instr', f'Equation Chapter {chapter} Section 1', rstyle='MTEquationSection')]
    for code in ('SEQ MTEqn \\r \\h \\* MERGEFORMAT', 'SEQ MTSec \\r 1 \\h \\* MERGEFORMAT',
                 f'SEQ MTChap \\r {chapter} \\h \\* MERGEFORMAT'):
        out += [_run(None, 'fld', 'begin'), _run(None, 'instr', f' {code} '), _run(None, 'fld', 'end')]
    out.append(_run(None, 'fld', 'end'))
    return out


def reference_nodes(rpr, bookmark, cached):
    return [_run(rpr, 'fld', 'begin'),
            _run(rpr, 'instr', f' GOTOBUTTON {bookmark}  \\* MERGEFORMAT '),
            _run(rpr, 'fld', 'begin'),
            _run(rpr, 'instr', f' REF {bookmark} \\* Charformat \\! \\* MERGEFORMAT '),
            _run(rpr, 'fld', 'separate'),
            _run(rpr, 'instr', cached),
            _run(rpr, 'fld', 'end'),
            _run(rpr, 'fld', 'end')]


def _field_runs(paragraph, begin_run):
    """Runs of the complex field that starts at begin_run, up to its matching end."""
    runs, depth = [], 0
    node = begin_run
    while node is not None:
        if node.tag == q('r'):
            runs.append(node)
            for fc in node.xpath('./w:fldChar/@w:fldCharType', namespaces=NS):
                depth += {'begin': 1, 'end': -1}.get(fc, 0)
            if depth == 0:
                return runs
        node = node.getnext()
    raise ValueError('Unterminated field in paragraph')


def _rpr_of(run):
    p = run.find(q('rPr'))
    return deepcopy(p) if p is not None else None


def _ensure_styles(styles, source_style_id, layout_ppr):
    ids = {s.get(q('styleId')): s for s in styles.findall(q('style'))}
    if 'MTDisplayEquation' not in ids:
        src = ids.get(source_style_id)
        st = deepcopy(src) if src is not None else etree.SubElement(styles, q('style'))
        if src is None:
            st.set(q('type'), 'paragraph')
            etree.SubElement(st, q('basedOn')).set(q('val'), 'Normal')
        else:
            styles.append(st)
        st.set(q('styleId'), 'MTDisplayEquation')
        st.set(q('customStyle'), '1')
        name = st.find(q('name'))
        if name is None:
            name = etree.Element(q('name'))
            st.insert(0, name)
        name.set(q('val'), 'MTDisplayEquation')
        nxt = st.find(q('next'))
        if nxt is None:
            nxt = etree.Element(q('next'))
            st.insert(2, nxt)
        nxt.set(q('val'), 'Normal')
        # the tab stops, alignment and indent of the existing displays, so that an equation
        # inserted later with MathType's ribbon is laid out like the others
        ppr = st.find(q('pPr'))
        if ppr is None:
            ppr = etree.SubElement(st, q('pPr'))
        for tag in ('tabs', 'jc', 'ind'):
            src_el = layout_ppr.find(q(tag))
            if src_el is None:
                continue
            old_el = ppr.find(q(tag))
            if old_el is not None:
                ppr.remove(old_el)
            ppr.append(deepcopy(src_el))
        order = ['keepNext', 'tabs', 'spacing', 'ind', 'jc']
        for el in sorted(ppr, key=lambda e: order.index(etree.QName(e).localname)
                         if etree.QName(e).localname in order else len(order)):
            ppr.append(el)
    if 'MTEquationSection' not in ids:
        st = etree.SubElement(styles, q('style'))
        st.set(q('type'), 'character')
        st.set(q('customStyle'), '1')
        st.set(q('styleId'), 'MTEquationSection')
        etree.SubElement(st, q('name')).set(q('val'), 'MTEquationSection')
        etree.SubElement(st, q('basedOn')).set(q('val'), 'DefaultParagraphFont')
        rpr = etree.SubElement(st, q('rPr'))
        etree.SubElement(rpr, q('vanish'))
        etree.SubElement(rpr, q('color')).set(q('val'), 'FF0000')


LABEL = re.compile(r'^\((?:([A-Z])\.)?(\d+)\)$')


def convert(path):
    path = Path(path)
    with ZipFile(path) as z:
        root = etree.fromstring(z.read('word/document.xml'))
        styles = etree.fromstring(z.read('word/styles.xml'))
        body = root.find(q('body'))
        displays = []
        for p in body.iter(q('p')):
            if not p.xpath('.//o:OLEObject', namespaces=NS):
                continue
            tabs = p.xpath('./w:r[w:tab]', namespaces=NS)
            starts = p.xpath('./w:bookmarkStart', namespaces=NS)
            if len(tabs) < 2 or not starts:
                continue
            tail = [c for c in p if p.index(c) > p.index(tabs[-1])]
            text = ''.join(t.text or '' for c in tail for t in c.xpath('.//w:t', namespaces=NS))
            m = LABEL.match(text)
            if not m or 'MTPlaceRef' in ''.join(p.xpath('.//w:instrText/text()', namespaces=NS)):
                continue
            displays.append((p, tabs[-1], tail, starts[-1], m.group(1), int(m.group(2))))
        if not displays:
            return {'displays': 0, 'references': 0}
        used, rename = set(), {}
        style_id = None
        prev_letter = None
        for k, (p, tab, tail, start, letter, number) in enumerate(displays):
            old, bid = start.get(q('name')), start.get(q('id'))
            h = int(hashlib.sha1(old.encode()).hexdigest(), 16)
            while (name := f'ZEqnNum{100000 + h % 900000}') in used:
                h += 1
            used.add(name)
            rename[old] = (name, text_of(letter, number))
            rpr = next((_rpr_of(r) for r in tail if r.tag == q('r') and r.find(q('rPr')) is not None), None)
            for c in tail:
                p.remove(c)
            for node in number_nodes(rpr, name, bid, letter, number):
                p.append(node)
            ps = p.find(q('pPr')).find(q('pStyle'))
            if style_id is None:
                style_id = ps.get(q('val')).replace(' ', '')
            ps.set(q('val'), 'MTDisplayEquation')
            if k == 0 or letter != prev_letter:
                if k == 0 and letter:
                    raise ValueError('First display equation carries a lettered label')
                chapter = ord(letter) - 64 if letter else 1
                host = body.find(q('p')) if k == 0 else _section_host(p)
                ppr = host.find(q('pPr'))
                pos = host.index(ppr) + 1 if ppr is not None else 0
                for node in reversed(section_nodes(chapter)):
                    host.insert(pos, node)
            prev_letter = letter
        nref = 0
        for instr in list(body.iter(q('instrText'))):
            m = re.match(r'\s*REF\s+(\S+)\s+\\h\s*$', instr.text or '')
            if not m or m.group(1) not in rename:
                continue
            begin = instr.getparent().getprevious()
            runs = _field_runs(begin.getparent(), begin)
            name, label = rename[m.group(1)]
            cached = ''.join(t.text or '' for r in runs for t in r.xpath('./w:t', namespaces=NS))
            if cached != label:
                raise ValueError(f'Cached reference {cached!r} differs from the label {label!r}')
            parent, pos = begin.getparent(), begin.getparent().index(begin)
            rpr = _rpr_of(runs[0])
            for r in runs:
                parent.remove(r)
            for i, node in enumerate(reference_nodes(rpr, name, cached)):
                parent.insert(pos + i, node)
            nref += 1
        _ensure_styles(styles, style_id, displays[0][0].find(q('pPr')))
        changed = {'word/document.xml': etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True),
                   'word/styles.xml': etree.tostring(styles, encoding='UTF-8', xml_declaration=True, standalone=True)}
        with NamedTemporaryFile(dir=path.parent, prefix='.mtnum-', suffix='.docx', delete=False) as tmp:
            temp = Path(tmp.name)
        try:
            with ZipFile(temp, 'w', ZIP_DEFLATED) as out:
                for member in z.infolist():
                    out.writestr(copy(member), changed.get(member.filename, z.read(member.filename)))
        except Exception:
            temp.unlink(missing_ok=True)
            raise
    temp.replace(path)
    report = check(path)
    if report['displays'] != len(displays) or report['references'] != nref:
        raise AssertionError(f'MathType numbering check failed: {report}')
    return report


def text_of(letter, number):
    return f'({letter}.{number})' if letter else f'({number})'


def _section_host(p):
    """The heading paragraph that opens the group of p (or p itself)."""
    node = p.getprevious()
    while node is not None:
        st = node.xpath('./w:pPr/w:pStyle/@w:val', namespaces=NS)
        if st and 'Heading' in st[0]:
            return node
        node = node.getprevious()
    return p


def check(path):
    """Counts of MathType numbering fields; raises if they are inconsistent."""
    with ZipFile(path) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    instr = root.xpath('.//w:instrText/text()', namespaces=NS) + root.xpath('.//w:fldSimple/@w:instr', namespaces=NS)
    joined = '\n'.join(instr)
    marks = {b.get(q('name')) for b in root.iter(q('bookmarkStart'))}
    zeqn = sorted(m for m in marks if m.startswith('ZEqnNum'))
    places = joined.count('MACROBUTTON MTPlaceRef')
    refs = re.findall(r'GOTOBUTTON (ZEqnNum\d+) ', joined)
    rep = {'displays': places, 'references': len(refs),
           'lettered': joined.count('SEQ MTChap \\c'),
           'sections': joined.count('MTEditEquationSection2'),
           'word_seq_eq': len(re.findall(r'SEQ Eq\b', joined))}
    if not (places == joined.count('SEQ MTEqn \\h') == joined.count('SEQ MTEqn \\c') == len(zeqn)):
        raise AssertionError(f'Inconsistent MathType number fields: {rep}')
    missing = sorted(set(refs) - marks)
    if missing or rep['word_seq_eq']:
        raise AssertionError(f'Unresolved references {missing} or Word SEQ Eq left: {rep}')
    return rep


if __name__ == '__main__':
    import argparse
    import json
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('docx', type=Path)
    ap.add_argument('--check', action='store_true', help='only report the MathType numbering fields')
    a = ap.parse_args()
    print(json.dumps(check(a.docx) if a.check else convert(a.docx)))
