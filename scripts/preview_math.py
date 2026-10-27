"""Equation previews that look like MathType's own (factory 'Times+Symbol 12' profile), with exact baselines.

The PNG preview is what Word shows until MathType redraws an object. docx-equation renders it with Chromium's MathML
layout (TeX-like: scripts 74 %, text-style fractions, operator spacing everywhere); MathType differs. This module
replaces the renderer and adjusts the MathML so that the preview follows MathType's rules (evidence in
references/mathtype-facts.md: glyph positions of 467 genuine MathType WMF previews, EquationLab's reconstruction of
MathType's layout, and Windows renderings):

- maths font: a renamed copy of TeX Gyre Termes Math whose MATH constants carry MathType's spacing table - scripts
  58/42 %, superscript +0.45 em, subscript -0.25 em, fraction numerator >= axis + t/2 + 0.35 em, denominator >= 1 em
  below the axis, minimum gap 8 %, scripts following tall fences; fractions keep full size in tables and fractions;
- operator spacing at full size: relations and the solidus 0.21 em, binary + and - 0.18 em, 0.14 em after a comma,
  none inside scripts;
- Symbol-style characters (operators, relations, integral; Greek in a Symbol profile) drawn with a Unicode-mapped copy
  of URW's metric clone of Symbol (Standard Symbols PS);
- integral: Symbol sign at 150 % centred on the axis, lower limit on its bottom edge, upper limit at its top, both after
  its advance plus 8 %;
- bar accent: 0.5 em rule 0.04 em thick, 0.13 em above the letter, starting 0.13 em (italic) / 0.09 em (upright)
  after the letter's origin; fraction bar overhang 8 %; 0.1 em inside an opening bracket or brace;
- crop = ink + 2 pt margin; the depth below the baseline is recorded in DEPTHS (preview index -> (depth, height) in
  pixels) for Word's baseline shift w:position = -depth (equation_layout.py).

Fonts are found on the host (kpsewhich/fc-match) and modified copies are written to ~/.cache/journal-manuscript-mathtype;
no font file is bundled. Needs Playwright (python -m pip install playwright; python -m playwright install chromium).
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import mtef_profile  # noqa: E402

CACHE = Path(os.environ.get('MATHTYPE_PREVIEW_CACHE', Path.home() / '.cache' / 'journal-manuscript-mathtype'))
MATH_FAMILY, SYMBOL_FAMILY = 'MTPreview Math', 'MTPreview Symbol'
DEPTHS: dict = {}
MARGIN_PT = 2.0
AXIS, RULE = 250, 50
CONSTANTS = {
    'ScriptPercentScaleDown': 58, 'ScriptScriptPercentScaleDown': 42,
    'AxisHeight': AXIS, 'FractionRuleThickness': RULE,
    'FractionNumeratorShiftUp': AXIS + RULE // 2 + 350, 'FractionNumeratorDisplayStyleShiftUp': AXIS + RULE // 2 + 350,
    'FractionDenominatorShiftDown': 1000 - AXIS + RULE // 2, 'FractionDenominatorDisplayStyleShiftDown': 1000 - AXIS + RULE // 2,
    'FractionNumeratorGapMin': 80, 'FractionNumDisplayStyleGapMin': 80,
    'FractionDenominatorGapMin': 80, 'FractionDenomDisplayStyleGapMin': 80,
    # scripts on a tall base follow its edges: superscript baseline ~0.35 script-em below the top of a tall fence
    'SubscriptShiftDown': 250, 'SubscriptTopMax': 2000, 'SubscriptBaselineDropMin': 100,
    'SuperscriptShiftUp': 450, 'SuperscriptShiftUpCramped': 450, 'SuperscriptBottomMin': 0,
    'SuperscriptBaselineDropMax': 210, 'SubSuperscriptGapMin': 80, 'SpaceAfterScript': 0,
    'OverbarVerticalGap': 110, 'OverbarExtraAscender': 0,
    'UpperLimitGapMin': 80, 'UpperLimitBaselineRiseMin': 250, 'LowerLimitGapMin': 80, 'LowerLimitBaselineDropMin': 1000,
}


# ---------------------------------------------------------------------------------------------------------------------
# Fonts
# ---------------------------------------------------------------------------------------------------------------------
def _find(cmd_list, patterns=()):
    for cmd in cmd_list:
        try:
            out = subprocess.run(cmd, capture_output=True, text=True, timeout=30).stdout.strip()
        except (OSError, subprocess.TimeoutExpired):
            continue
        if out and Path(out.splitlines()[0]).exists():
            return out.splitlines()[0]
    for p in patterns:
        if Path(p).exists():
            return p
    return None


def termes_math():
    p = os.environ.get('MATHTYPE_TERMES_MATH') or _find(
        [['kpsewhich', 'texgyretermes-math.otf'], ['fc-match', '-f', '%{file}', 'TeX Gyre Termes Math']],
        ['/usr/share/texmf/fonts/opentype/public/tex-gyre-math/texgyretermes-math.otf',
         '/usr/share/fonts/opentype/tex-gyre-math/texgyretermes-math.otf'])
    if not p or 'termes' not in p.lower():
        raise FileNotFoundError('TeX Gyre Termes Math not found (Debian/Ubuntu: fonts-texgyre-math or texlive-fonts-extra); '
                                'or set MATHTYPE_TERMES_MATH')
    return p


def urw_symbol():
    p = os.environ.get('MATHTYPE_SYMBOL_FONT') or _find([['fc-match', '-f', '%{file}', 'Standard Symbols PS']],
                                                       ['/usr/share/fonts/type1/urw-base35/StandardSymbolsPS.t1'])
    if not p or 'symbol' not in p.lower():
        raise FileNotFoundError('URW Standard Symbols PS not found (Debian/Ubuntu: fonts-urw-base35); or set MATHTYPE_SYMBOL_FONT')
    return p


def font_path():
    """Renamed Termes Math with MathType's MATH constants (cached by content; GUST Font License: renamed copy)."""
    key = hashlib.sha1(json.dumps(CONSTANTS, sort_keys=True).encode()).hexdigest()[:10]
    out = CACHE / f'mtpreview-math-{key}.otf'
    if out.exists():
        return out
    from fontTools.ttLib import TTFont
    f = TTFont(termes_math())
    mc = f['MATH'].table.MathConstants
    for k, v in CONSTANTS.items():
        cur = getattr(mc, k)
        if hasattr(cur, 'Value'):
            cur.Value = v
        else:
            setattr(mc, k, v)
    for rec in f['name'].names:
        if rec.nameID in (1, 4, 16, 18):
            rec.string = MATH_FAMILY
        elif rec.nameID == 6:
            rec.string = MATH_FAMILY.replace(' ', '')
    CACHE.mkdir(parents=True, exist_ok=True)
    f.save(out)
    return out


SYMBOL_GLYPHS = {'=': 'equal', '+': 'plus', '−': 'minus', '-': 'minus', '<': 'less', '>': 'greater', '×': 'multiply',
                 '±': 'plusminus', '∂': 'partialdiff', '∇': 'gradient', '∈': 'element', '∉': 'notelement',
                 '⊂': 'propersubset', '⊃': 'propersuperset', '⊆': 'reflexsubset', '⊇': 'reflexsuperset',
                 '→': 'arrowright', '←': 'arrowleft', '↔': 'arrowboth', '⇒': 'arrowdblright', '⇔': 'arrowdblboth',
                 '≈': 'approxequal', '≠': 'notequal', '≤': 'lessequal', '≥': 'greaterequal', '≡': 'equivalence',
                 '∞': 'infinity', '∗': 'asteriskmath', '∝': 'proportional', '⋅': 'dotmath', '·': 'dotmath',
                 '∼': 'similar', '∫': 'integral', '∑': 'summation', '∏': 'product', '∀': 'universal', '∃': 'existential',
                 '∩': 'intersection', '∪': 'union', '∅': 'emptyset', '…': 'ellipsis', '′': 'minute', '÷': 'divide'}
_GREEK = ('Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa Lambda Mu Nu Xi Omicron Pi Rho _ Sigma Tau Upsilon '
          'Phi Chi Psi Omega').split()
SYMBOL_GLYPHS.update({chr(0x391 + i): n for i, n in enumerate(_GREEK) if n != '_'})
SYMBOL_GLYPHS.update({chr(0x3B1 + i): n.lower() for i, n in enumerate(_GREEK) if n != '_'})
# MathType writes U+03D5 at Symbol 0x66 ('phi') and U+03C6 at 0x6A ('phi1')
SYMBOL_GLYPHS.update({'ς': 'sigma1', 'ϑ': 'theta1', 'ϕ': 'phi', 'φ': 'phi1', 'ϖ': 'omega1'})


def symbol_font_path():
    """Unicode-mapped CFF font from URW Standard Symbols PS (metric clone of Adobe Symbol; its own cmap is
    Symbol-encoded, and Chromium does not load Type 1)."""
    out = CACHE / 'mtpreview-symbol-v2.otf'
    if out.exists():
        return out
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.boundsPen import BoundsPen
    from fontTools.pens.t2CharStringPen import T2CharStringPen
    src_path = urw_symbol()
    if src_path.lower().endswith(('.t1', '.pfb', '.pfa')):
        from fontTools import t1Lib
        t1 = t1Lib.T1Font(src_path)
        t1.parse()
        gs = t1.getGlyphSet()
    else:
        from fontTools.ttLib import TTFont
        gs = TTFont(src_path).getGlyphSet()
    names = ['.notdef'] + sorted({n for n in SYMBOL_GLYPHS.values() if n in gs} | {'space'})
    charstrings, metrics = {}, {}
    for n in names:
        src = gs['space']
        b = BoundsPen(gs)
        if n != '.notdef':
            src = gs[n]
            src.draw(b)
        width = getattr(src, 'width', None)
        if width is None:
            src.draw(b)
            width = src.width
        pen = T2CharStringPen(width, gs)
        if n != '.notdef':
            src.draw(pen)
        charstrings[n] = pen.getCharString()
        metrics[n] = (width, b.bounds[0] if (b.bounds and n != '.notdef') else 0)
    fb = FontBuilder(1000, isTTF=False)
    fb.setupGlyphOrder(names)
    fb.setupCharacterMap({ord(c): n for c, n in SYMBOL_GLYPHS.items() if n in names} | {0x20: 'space'})
    fb.setupCFF(SYMBOL_FAMILY.replace(' ', ''), {'FullName': SYMBOL_FAMILY}, charstrings, {})
    fb.setupHorizontalMetrics(metrics)
    fb.setupHorizontalHeader(ascent=1010, descent=-293)
    fb.setupNameTable({'familyName': SYMBOL_FAMILY, 'styleName': 'Regular'})
    fb.setupOS2(sTypoAscender=1010, sTypoDescender=-293, usWinAscent=1010, usWinDescent=293)
    fb.setupPost()
    CACHE.mkdir(parents=True, exist_ok=True)
    fb.save(out)
    return out


_TNR = {}


def tnr_advance(ch, italic, bold):
    """Advance width (em) of a Times New Roman glyph."""
    from fontTools.ttLib import TTFont
    style = ('Bold ' if bold else '') + ('Italic' if italic else '') or 'Regular'
    style = style.strip() or 'Regular'
    if style not in _TNR:
        p = _find([['fc-match', '-f', '%{file}', f'Times New Roman:style={style}']])
        _TNR[style] = TTFont(p) if p else None
    f = _TNR[style]
    if f is None:
        return None
    g = f.getBestCmap().get(ord(ch))
    return f['hmtx'][g][0] / f['head'].unitsPerEm if g else None


# ---------------------------------------------------------------------------------------------------------------------
# MathML adjustments
# ---------------------------------------------------------------------------------------------------------------------
REL = set('=<>≤≥≈≠≡∈∉⊂⊃⊆⊇→←↔⇒⇔∼∝')
BIN = set('+−-±×⋅·∗÷∪∩')
PUNCT = set(',;')
FENCE_CHARS = set('()[]{}|‖⟨⟩〈〉⟦⟧')
SCRIPTS = {'msub', 'msup', 'msubsup', 'munder', 'mover', 'munderover', 'mmultiscripts'}
INTEGRALS = set('∫∬∭∮')
BARS = {'¯', '‾', '̅', '_'}
M = 'http://www.w3.org/1998/Math/MathML'


def _ln(e):
    return e.tag.split('}')[-1] if isinstance(e.tag, str) else ''


def _in_script(e):
    while e.getparent() is not None:
        p = e.getparent()
        if _ln(p) in SCRIPTS and p.index(e) > 0:
            return True
        e = p
    return False


def _el(tag, attrs=None, children=(), text=None):
    from lxml import etree
    x = etree.Element(f'{{{M}}}{tag}', {k: str(v) for k, v in (attrs or {}).items()})
    if text is not None:
        x.text = text
    for c in children:
        x.append(c)
    return x


def _sign(t):
    return _el('mo', {'data-sym': 1, 'lspace': '0em', 'rspace': '0em', 'stretchy': 'false', 'largeop': 'false',
                      'movablelimits': 'false', 'mathsize': '150%'}, text=t)


def transform(mml):
    """MathML -> MathML with MathType's spacing, integral, bar and Symbol-style marks (preview only)."""
    from lxml import etree
    root = etree.fromstring(mml.encode('utf-8'))
    greek_symbol = mtef_profile.PROFILE.greek == 'symbol'
    for s in list(root.iter(f'{{{M}}}msub', f'{{{M}}}msubsup', f'{{{M}}}munder', f'{{{M}}}munderover')):
        op = s[0]
        if _ln(op) != 'mo' or (op.text or '').strip() not in INTEGRALS:
            continue
        lower, upper = s[1], (s[2] if len(s) > 2 else None)
        lo = _el('mpadded', {'voffset': '-0.517em', **({'width': '0em'} if upper is not None else {})},
                 [_el('mstyle', {'scriptlevel': '+1'}, [lower])])
        parts = [_el('mpadded', {'voffset': '-0.357em'}, [_sign(op.text.strip())]), _el('mspace', {'width': '0.08em'}), lo]
        if upper is not None:
            parts.append(_el('mpadded', {'voffset': '0.62em'}, [_el('mstyle', {'scriptlevel': '+1'}, [upper])]))
        parts.append(_el('mspace', {'width': '0.1em'}))
        new = _el('mrow', {}, parts)
        new.tail = s.tail
        s.getparent().replace(s, new)
    for op in list(root.iter(f'{{{M}}}mo')):
        t = (op.text or '').strip()
        if t in INTEGRALS and op.get('mathsize') is None:
            new = _el('mrow', {}, [_el('mpadded', {'voffset': '-0.357em'}, [_sign(t)]), _el('mspace', {'width': '0.1em'})])
            new.tail = op.tail
            op.getparent().replace(op, new)
    for mv in list(root.iter(f'{{{M}}}mover')):
        if len(mv) != 2 or _ln(mv[1]) != 'mo' or (mv[1].text or '').strip() not in BARS:
            continue
        base = mv[0]
        while _ln(base) == 'mrow' and len(base) == 1:
            base = base[0]
        ch = (base.text or '').strip()
        italic = _ln(base) == 'mi' and base.get('mathvariant') is None and len(ch) == 1
        adv = tnr_advance(ch, italic, (base.get('mathvariant') or '').startswith('bold')) if len(ch) == 1 else None
        if adv is None:
            continue
        bar = _el('mstyle', {'scriptlevel': '0'}, [_el('mpadded', {'width': f'{adv:.3f}em', 'lspace': '0.13em' if italic else '0.09em'},
                  [_el('mspace', {'width': '0.5em', 'height': '0.04em', 'mathbackground': 'black'})])])
        mv.replace(mv[1], bar)
        mv.set('accent', 'false')
    for op in root.iter(f'{{{M}}}mo'):
        t = (op.text or '').strip()
        if t not in FENCE_CHARS:
            continue
        nxt, prv = op.getnext(), op.getprevious()
        if op.get('stretchy') == 'true' or (nxt is not None and _ln(nxt) == 'mtable') or (prv is not None and _ln(prv) == 'mtable'):
            op.set('stretchy', 'true')
            if t in '[{⟦':
                op.set('rspace', '0.1em')
    for op in root.iter(f'{{{M}}}mo', f'{{{M}}}mi'):
        t = (op.text or '').strip()
        if not t or t in FENCE_CHARS or op.get('stretchy') == 'true' or op.get('data-sym') is not None:
            continue
        if len(t) == 1 and t in SYMBOL_GLYPHS:
            greek = mtef_profile.greek_face(t)
            if not greek:
                op.set('data-sym', '1')
            elif greek_symbol and op.get('mathvariant') not in ('bold', 'bold-italic'):
                op.set('data-sym', 'i' if greek == mtef_profile.LC_GREEK and op.get('mathvariant') != 'normal' else '1')
        if _ln(op) != 'mo':
            continue
        if _in_script(op):
            l = r = 0
        elif t in REL or t in (':=', '/'):
            l = r = 0.21
        elif t in BIN:
            prev = op.getprevious()
            pt = (prev.text or '').strip() if prev is not None and _ln(prev) == 'mo' else None
            unary = prev is None or (pt is not None and (pt in REL or pt in BIN or pt in PUNCT or pt in '([{⟨〈'))
            l = r = 0 if unary else 0.18
        elif t in PUNCT:
            l, r = 0, 0.14
        else:
            l = r = 0
        op.set('lspace', f'{l}em')
        op.set('rspace', f'{r}em')
    return etree.tostring(root, encoding='unicode')


# ---------------------------------------------------------------------------------------------------------------------
# HTML and rendering
# ---------------------------------------------------------------------------------------------------------------------
# Chromium's MathML layout starts every table cell at the left and ignores text-align/columnalign: shift the contents of
# each cell by its slack (half for centre, all for right), keeping the column width.
ALIGN_JS = """(align) => { for (const e of document.querySelectorAll('mtd')) {
  const d = e.getBoundingClientRect(); let l = 1e9, r = -1e9;
  for (const c of e.querySelectorAll('mi,mn,mo,mtext,mspace')) { const b = c.getBoundingClientRect(); if (b.width) { l = Math.min(l, b.left); r = Math.max(r, b.right); } }
  if (r < l) continue;
  const cs = getComputedStyle(e), pl = parseFloat(cs.paddingLeft), pr = parseFloat(cs.paddingRight);
  const slack = (d.right - r) - (l - d.left); if (slack <= 0.5) continue;
  const shift = align === 'center' ? slack / 2 : slack;
  e.style.paddingLeft = (pl + shift) + 'px'; e.style.paddingRight = Math.max(0, pr - shift) + 'px'; } }"""
JS = """() => { const m = document.querySelector('math'); const r = m.getBoundingClientRect();
  const b = document.getElementById('baseline').getBoundingClientRect().top; return [r.left, r.top, r.right, r.bottom, b]; }"""


def html_for(mml_text, font_px):
    from docx_equation.shared import mathml as MM
    m = MM._mathml_for_html(mml_text)
    m = re.sub(r'<mi>([^<]{2,})</mi>', lambda x: x.group(0) if x.group(1).startswith('&#x') and x.group(1).count(';') == 1
               else f'<mi mathvariant="normal">{x.group(1)}</mi>', m)          # multi-letter identifiers upright
    m = transform(m)
    # only stretched fences and large operators use the maths font
    m = re.sub(r'<mo([^>]*stretchy="true"[^>]*)>', r'<mo data-tx="1"\1>', m)
    m = re.sub(r'<mo([^>]*)>([∑∏⟨⟩⟦⟧])</mo>', r'<mo data-tx="1"\1>\2</mo>', m)
    h = MM._html(m, font_px)
    css = f"""
@font-face {{ font-family: "{MATH_FAMILY}"; src: url("{font_path().as_uri()}"); }}
@font-face {{ font-family: "{SYMBOL_FAMILY}"; src: url("{symbol_font_path().as_uri()}"); }}
math {{ font-family: "{MATH_FAMILY}", serif !important; display: inline math !important; }}
mi, mn, mtext, mo {{ font-family: "Times New Roman"; text-transform: none; }}
mi {{ font-style: italic; }}
mi[mathvariant="normal"], mn, mtext {{ font-style: normal; }}
mi[mathvariant="bold"], mn[mathvariant="bold"] {{ font-weight: bold; font-style: normal; }}
mi[mathvariant="bold-italic"] {{ font-weight: bold; font-style: italic; }}
mo[data-tx] {{ font-family: "{MATH_FAMILY}"; }}
[data-sym] {{ font-family: "{SYMBOL_FAMILY}" !important; font-style: normal !important; }}
[data-sym="i"] {{ font-style: italic !important; }}
mtd {{ text-align: left; }}
mtable, mtd, mfrac > * {{ math-style: normal !important; math-depth: auto !important; }}
mfrac > * {{ padding-left: 0.08em; padding-right: 0.08em; }}
"""
    h = h.replace('</style>', css + '</style>', 1)
    h = h.replace('<div class="box">', '<div style="display:inline-block;padding:60px;background:#fff">'
                  '<span id="baseline" style="display:inline-block;width:0;height:0"></span>', 1)
    return h


def render_mathml_files(mathml_dir, output_dir, chrome_path=None, font_px=38):
    """Replacement for docx_equation's renderer: same files in, same PNGs out, plus DEPTHS."""
    import io
    import numpy as np
    from PIL import Image
    from playwright.sync_api import sync_playwright
    files = sorted(Path(mathml_dir).glob('equation_*.mml'))
    if not files:
        raise FileNotFoundError(f'No MathML files found in: {mathml_dir}')
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    exe = chrome_path or os.environ.get('MATHTYPE_CHROMIUM')
    scale = 2
    margin = round(MARGIN_PT / 12 * font_px * scale)
    dump = os.environ.get('PREVIEW_DUMP_MML')
    with sync_playwright() as pw:
        browser = pw.chromium.launch(**({'executable_path': os.path.realpath(exe)} if exe else {}))
        page = browser.new_page(device_scale_factor=scale, viewport={'width': 2600, 'height': 900})
        tmp = out / '_page.html'                         # file:// page: about:blank pages may not load file fonts
        for src in files:
            if dump:
                Path(dump).mkdir(parents=True, exist_ok=True)
                shutil.copy(src, Path(dump) / src.name)
            tmp.write_text(html_for(src.read_text(encoding='utf-8'), font_px), encoding='utf-8')
            page.goto(tmp.as_uri())
            page.evaluate('document.fonts.ready')
            if mtef_profile.PROFILE.matrix_align != 'left':
                page.evaluate(ALIGN_JS, mtef_profile.PROFILE.matrix_align)
            left, top, right, bottom, base = page.evaluate(JS)
            pad = 50
            clip = {'x': max(0, left - pad), 'y': max(0, top - pad), 'width': right - left + 2 * pad, 'height': bottom - top + 2 * pad}
            img = Image.open(io.BytesIO(page.screenshot(clip=clip))).convert('RGB')
            ink = np.asarray(img.convert('L')) < 243
            if not ink.any():
                raise RuntimeError(f'empty preview: {src.name}')
            ys, xs = np.where(ink.any(1))[0], np.where(ink.any(0))[0]
            y0, y1 = max(0, ys[0] - margin), min(img.height, ys[-1] + 1 + margin)
            x0, x1 = max(0, xs[0] - margin), min(img.width, xs[-1] + 1 + margin)
            img.crop((x0, y0, x1, y1)).save(out / f'{src.stem}.png')
            index = int(re.search(r'(\d+)$', src.stem).group(1))
            DEPTHS[index] = (float(y1 - (base - clip['y']) * scale), float(y1 - y0))
        browser.close()
        tmp.unlink()
    return len(files)


def install():
    """Route docx-equation's preview rendering through this module (call before registry.save)."""
    from docx_equation.mathtype import embed, legacy
    from docx_equation.shared import mathml as MM
    DEPTHS.clear()
    for mod in (embed, legacy, MM):
        mod.render_mathml_files = render_mathml_files


def check():
    """Raise if a font or the browser is missing."""
    termes_math(), urw_symbol(), font_path(), symbol_font_path()
    if tnr_advance('A', False, False) is None:
        raise FileNotFoundError('Times New Roman not found (see references/true-times-new-roman.md)')
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        exe = os.environ.get('MATHTYPE_CHROMIUM')
        pw.chromium.launch(**({'executable_path': os.path.realpath(exe)} if exe else {})).close()
    return True


if __name__ == '__main__':
    print('preview fonts and browser OK' if check() else 'failed')
