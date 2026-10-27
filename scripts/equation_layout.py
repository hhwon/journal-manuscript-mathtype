"""Place MathType objects in a generated DOCX as Word places genuine ones.

- Baseline: Word puts the bottom of an OLE preview on the text baseline; genuine MathType objects carry
  w:position = -(depth of the preview below the formula baseline). The depth comes from preview_math.DEPTHS (measured
  while rendering); objects without a measurement keep their position.
- No VML vertical-align on the shapes (genuine objects have none).
- Display equations: equation number right-aligned by a tab 10 twips inside the margin (otherwise it can wrap);
  display alignment 'center' (default) or 'left' at a given indent; the paragraph indents of docx-equation's numbered
  paragraph are reset with w:firstLine (w:first, which some code writes, is ignored by Word).
- Optional scale for objects inside table cells.

    from equation_layout import fix_layout; fix_layout('out.docx', display_align='left', indent_twips=482)
"""
import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
V = 'urn:schemas-microsoft-com:vml'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'w': W, 'v': V}
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow',
             'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern']


def fix_layout(path, display_align='center', indent_twips=0, table_scale=1.0, depths=None):
    if depths is None:
        import preview_math
        depths = preview_math.DEPTHS
    path = Path(path)
    with ZipFile(path) as z:
        infos = z.infolist()
        parts = {i.filename: z.read(i.filename) for i in infos}
    root = etree.fromstring(parts['word/document.xml'])
    rels = {r.get('Id'): r.get('Target') for r in etree.fromstring(parts['word/_rels/document.xml.rels'])}
    placed = 0
    for shape in root.iter(f'{{{V}}}shape'):
        st = shape.get('style', '').replace(';vertical-align:middle', '')
        img = shape.find(f'{{{V}}}imagedata')
        run = shape.getparent().getparent()
        if img is None or run.tag != f'{{{W}}}r':
            continue
        m = re.search(r'mathtype_preview_(\d+)\.png', rels.get(img.get(f'{{{R}}}id'), ''))
        if not m:
            continue
        h = float(re.search(r'height:([\d.]+)pt', st).group(1))
        if table_scale != 1.0 and list(run.iterancestors(f'{{{W}}}tc')):
            w = float(re.search(r'width:([\d.]+)pt', st).group(1)) * table_scale
            h *= table_scale
            st = re.sub(r'width:[\d.]+pt', f'width:{w:.1f}pt', st)
            st = re.sub(r'height:[\d.]+pt', f'height:{h:.1f}pt', st)
        shape.set('style', st)
        d = depths.get(int(m.group(1)))
        if d is None:
            continue
        depth_px, height_px = d
        rpr = run.find(f'{{{W}}}rPr')
        if rpr is None:
            rpr = etree.Element(f'{{{W}}}rPr')
            run.insert(0, rpr)
        pos = rpr.find(f'{{{W}}}position')
        if pos is None:
            pos = etree.Element(f'{{{W}}}position')
            k = 0
            for i, ch in enumerate(rpr):
                if etree.QName(ch).localname in RPR_ORDER:
                    k = i + 1
            rpr.insert(k, pos)
        pos.set(f'{{{W}}}val', str(-max(0, round(2 * h * depth_px / height_px))))
        placed += 1
    sect = root.find(f'.//{{{W}}}sectPr')
    width = None
    if sect is not None and sect.find(f'{{{W}}}pgSz') is not None:
        pg, mar = sect.find(f'{{{W}}}pgSz'), sect.find(f'{{{W}}}pgMar')
        width = int(pg.get(f'{{{W}}}w')) - int(mar.get(f'{{{W}}}left')) - int(mar.get(f'{{{W}}}right'))
    for p in root.xpath('.//w:p[.//w:tabs/w:tab[@w:val="center"]][.//o:OLEObject]',
                        namespaces={**NS, 'o': 'urn:schemas-microsoft-com:office:office'}):
        tabs = p.find(f'{{{W}}}pPr/{{{W}}}tabs')
        for tab in tabs:
            if tab.get(f'{{{W}}}val') == 'right' and width and int(tab.get(f'{{{W}}}pos')) >= width - 10:
                tab.set(f'{{{W}}}pos', str(width - 10))
            if tab.get(f'{{{W}}}val') == 'center' and display_align == 'left':
                tab.set(f'{{{W}}}val', 'left')
                tab.set(f'{{{W}}}pos', str(indent_twips))
        ind = p.find(f'{{{W}}}pPr/{{{W}}}ind')
        if ind is not None:
            for a in list(ind.attrib):
                del ind.attrib[a]
            ind.set(f'{{{W}}}left', '0')
            ind.set(f'{{{W}}}firstLine', '0')
    parts['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    tmp = path.with_suffix('.tmp')
    with ZipFile(tmp, 'w', ZIP_DEFLATED) as zo:
        for i in infos:
            zo.writestr(i.filename, parts[i.filename])
    tmp.replace(path)
    return placed
