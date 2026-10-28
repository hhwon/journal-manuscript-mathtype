"""Make the MathType equations written by docx-equation 0.3.3 equivalent to genuine MathType objects.

Import this module before saving with docx-equation (``import mtef_profile; mtef_profile.configure(...)``). It patches
the backend's MathML parser and MTEF writer in place; the evidence for every rule is in references/mathtype-facts.md.

Profile (EQN_PREFS): a bare ``configure()`` gives MathType's factory "Times+Symbol 12" - fonts Times New Roman, Symbol, Courier New,
MT Extra; Variable Times italic, Vector Times bold, Number/Function/Text Times upright, Greek and the Symbol style in the
Symbol font; sizes 12 pt, 58, 42, 150, 100 %; factory spacing table; centred matrices. build_paper.py passes the
skill's house style instead (configure(greek='times', matrix_align='left')). Options: Greek in Times New Roman, matrix
column alignment, or a JSON profile file.

Record layouts (all as MathType writes them):
- an empty slot is a null LINE (01 01) without END record;
- subscript/superscript templates: TMPL 27/28/29, SUB size, slots, END, FULL size;
- big operators: TMPL, LINE body, SUB, lower, upper, SYM size (0x0d), CHAR operator (Symbol style), END; variation
  tvBO_LOWER 0x10, tvBO_UPPER 0x20, tvBO_SUM 0x40, integrals tvINT_1..3;
- fences: TMPL selector, LINE body, FULL, fence characters with typeface 22 (fnEXPAND); bar EC07/EC08, double bar
  EC09/EC0A, white brackets 301A/301B, angle 2329/232A;
- Symbol-style characters carry their 8-bit Symbol font position (mtefOPT_CHAR_ENC_CHAR_8), Times characters none;
- spaces are fnSPACE characters (typeface 24): EF02 thin, EF03 medium, EF04 thick, EF05 em, EF06 2-em, EF08 1-point;
- accents on one character: CHAR with an EMBELL record (bar 17, tilde 8, hat 9, arrow 11);
- MATRIX h_just: 0 left, 1 centre, 2 right (genuine centred matrices carry 1; the public specification's list says
  1 = left, which does not match MathType's behaviour).
"""
from __future__ import annotations

import json
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from docx_equation.mathtype import mtef as MT
from docx_equation.shared import mathml as MM

sys.path.insert(0, str(Path(__file__).resolve().parent))
from mtef_font_audit import _dimensions  # noqa: E402

# typefaces: 128 + MathType style
TEXT, FUNCTION, VARIABLE, LC_GREEK, UC_GREEK, SYMBOL, VECTOR, NUMBER = 0x81, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88
EXPAND, SPACE_TF = 0x96, 0x98
FULL, SUB, SYM = b'\x0a', b'\x0b', b'\x0d'

# ---------------------------------------------------------------------------------------------------------------------
# Profile
# ---------------------------------------------------------------------------------------------------------------------
# MathType equation preferences, in the order of the MTEF EQN_PREFS record and of MathType's .eqp files
# ("; MathType INI v1: Equation Preferences", sections [Styles], [Sizes], [Spacing]).
STYLE_KEYS = ['Text', 'Function', 'Variable', 'LCGreek', 'UCGreek', 'Symbol', 'Vector', 'Number', 'User1', 'User2', 'MTExtra',
              'TextFE']                              # Text (FE): text style for East Asian characters, optional
SIZE_KEYS = ['Full', 'Script', 'ScriptScript', 'Symbol', 'SubSymbol', 'User1', 'User2', 'SmallLargeIncr']
SPACING_KEYS = ['LineSpacing', 'MatrixRowSpacing', 'MatrixColSpacing', 'SuperscriptHeight', 'SubscriptDepth', 'SubSupGap',
                'LimHeight', 'LimDepth', 'LimLineSpacing', 'NumerHeight', 'DenomDepth', 'FractBarOver', 'FractBarThick',
                'SubFractBarThick', 'FractGap', 'FenceOver', 'OperSpacing', 'NonOperSpacing', 'CharWidth', 'MinGap',
                'VertRadGap', 'HorizRadGap', 'RadWidth', 'EmbellGap', 'PrimeHeight', 'BoxStrokeThick', 'StikeThruThick',
                'MatrixLineThick', 'RadStrokeThick', 'HorizFenceGap']
# MathType's factory 'Times+Symbol 12' (identical styles/spacing in 467 genuine equations; EquationLab new-formula.mjs)
FACTORY_STYLES = [('Times New Roman', 0), ('Times New Roman', 0), ('Times New Roman', 2), ('Symbol', 2), ('Symbol', 0),
                  ('Symbol', 0), ('Times New Roman', 1), ('Times New Roman', 0), ('Courier New', 0), ('Times New Roman', 0),
                  ('MT Extra', 0)]
FACTORY_SIZES = [('12', 'pt'), ('58', '%'), ('42', '%'), ('150', '%'), ('100', '%'), ('75', '%'), ('150', '%'), ('1', 'pt')]
FACTORY_SPACING = [(v, '%') for v in ('150 150 100 45 25 8 25 100 100 35 100 8 5 2.5 8 8 100 100 0 8 17 8 100 12.5 45 '
                                      '5 5 5 5 10').split()]
UNIT_CODES = {'in': 0, 'cm': 1, 'pt': 2, 'pc': 3, '%': 4}
ENCODINGS = {'Symbol': 3, 'MT Extra': 4}             # predefined MTEF encodings; other fonts: WinAllBasicCodePages (5)
MATRIX_ALIGN = {'left': 0, 'center': 1, 'right': 2}


def encode_dimensions(dims):
    """MTEF dimension array: count, then per value a unit nibble, digit nibbles (. = A, - = B) and F."""
    nib = []
    for value, unit in dims:
        nib.append(UNIT_CODES[unit])
        nib += [10 if c == '.' else 11 if c == '-' else int(c) for c in value]
        nib.append(15)
    if len(nib) % 2:
        nib.append(0)
    return bytes([len(dims)]) + bytes(nib[i] << 4 | nib[i + 1] for i in range(0, len(nib), 2))


def read_eqp(path):
    """MathType preference file (.eqp, 'MathType INI v1') -> (styles, sizes, spacing, notes). Keys missing from the
    file keep MathType's factory values (listed in notes)."""
    text = Path(path).read_text(encoding='utf-8-sig', errors='strict')
    if not re.search(r'^\s*;\s*MathType INI v1\s*:\s*Equation Preferences', text, re.M | re.I):
        raise ValueError(f'{path}: not a MathType INI v1 equation preference file')
    styles, sizes, spacing, notes = list(FACTORY_STYLES) + [None], list(FACTORY_SIZES), list(FACTORY_SPACING), []
    seen, section = set(), ''
    keys = {'styles': STYLE_KEYS, 'sizes': SIZE_KEYS, 'spacing': SPACING_KEYS}
    for n, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line or line.startswith(';'):
            continue
        m = re.match(r'^\[([^\]]+)\]$', line)
        if m:
            section = m.group(1).strip().lower()
            continue
        if '=' not in line or section not in keys:
            notes.append(f'line {n} ignored: {line[:60]}')
            continue
        key, value = (x.strip() for x in line.split('=', 1))
        key = next((k for k in keys[section] if k.lower() == re.sub(r'[\s()]', '', key).lower()), None)
        if key is None:
            notes.append(f'line {n} ignored (unknown key): {line[:60]}')
            continue
        seen.add((section, key))
        i = keys[section].index(key)
        if section == 'styles':
            font, _, flag = value.partition(',')
            flag = flag.strip().upper()
            if flag not in ('', 'B', 'I'):
                raise ValueError(f'{path} line {n}: unsupported style flag {flag!r}')
            styles[i] = (font.strip(), {'': 0, 'B': 1, 'I': 2}[flag])
        else:
            m = re.match(r'^(-?(?:\d+(?:\.\d*)?|\.\d+))\s*(pt|%|cm|in|pc)$', value)
            if not m:
                raise ValueError(f'{path} line {n}: unsupported dimension {value!r}')
            (sizes if section == 'sizes' else spacing)[i] = (m.group(1), m.group(2))
    for section, ks in keys.items():
        missing = [k for k in ks if (section, k) not in seen and k != 'TextFE']
        if missing:
            notes.append(f'[{section.title()}] not in the file, factory value used: {", ".join(missing)}')
    if sizes[0][1] != 'pt':
        raise ValueError(f'{path}: the full size must be given in pt')
    return styles, sizes, spacing, notes


@dataclass
class Profile:
    """MathType equation preferences: 11 styles (font, bits: 1 bold, 2 italic), 8 sizes, 30 spacing values, plus the
    matrix column alignment (not part of MathType's preferences)."""
    styles: list = field(default_factory=lambda: list(FACTORY_STYLES))
    sizes: list = field(default_factory=lambda: list(FACTORY_SIZES))
    spacing: list = field(default_factory=lambda: list(FACTORY_SPACING))
    matrix_align: str = 'center'
    source: str = 'MathType factory (Times+Symbol 12)'
    notes: list = field(default_factory=list)

    def __post_init__(self):
        if self.matrix_align not in MATRIX_ALIGN:
            raise ValueError(f'matrix_align must be one of {sorted(MATRIX_ALIGN)}')
        if len(self.styles) == 11:
            self.styles = list(self.styles) + [None]           # Text (FE) unused
        if len(self.styles) != 12 or len(self.sizes) != 8 or len(self.spacing) != 30:
            raise ValueError('a profile needs 11 styles (+ optional Text FE), 8 sizes and 30 spacing values')
        if self.styles[5][0] != 'Symbol':
            self.notes.append('Symbol style not in the Symbol font: MathType cannot build tall brackets and integrals '
                              'from it when it redraws an equation')

    @property
    def greek(self):                                # 'symbol' if Greek is in the Symbol font (Symbol positions)
        return 'symbol' if self.styles[3][0] == 'Symbol' else 'times'

    @property
    def symbol_positions(self):
        return self.styles[5][0] == 'Symbol'

    @property
    def full_pt(self):
        return float(self.sizes[0][0])

    def pct(self, key):
        """A spacing value in percent of the full size."""
        return self._pct(self.spacing[SPACING_KEYS.index(key)])

    def size_pct(self, key):
        """A size in percent of the full size."""
        return self._pct(self.sizes[SIZE_KEYS.index(key)])

    def _pct(self, dim):
        v, unit = float(dim[0]), dim[1]
        return v if unit == '%' else v * {'pt': 1, 'pc': 12, 'in': 72, 'cm': 72 / 2.54}[unit] / self.full_pt * 100

    def font_defs(self):
        fonts = []
        for st in self.styles:
            if st and st[0] not in fonts:
                fonts.append(st[0])
        return fonts


PROFILE = Profile()
UNMAPPED_SYMBOLS: set = set()        # Symbol-style characters without a Symbol position (report them)


def configure(greek: str | None = None, matrix_align: str = 'center', profile_json: str | None = None,
              eqp: str | None = None) -> Profile:
    """Set the equation profile.

    eqp: a MathType preference file (.eqp) - all styles, sizes and spacing are taken from it; profile_json may also
    name a .eqp file, or a JSON file {"eqp": ..., "greek": ..., "matrix_align": ...}. greek ('symbol' | 'times')
    overrides the Greek styles; without any argument the profile is MathType's factory 'Times+Symbol 12'."""
    global PROFILE
    if profile_json and str(profile_json).lower().endswith('.eqp'):
        eqp, profile_json = profile_json, None
    if profile_json:
        data = json.loads(Path(profile_json).read_text())
        eqp = data.get('eqp', eqp)
        greek, matrix_align = data.get('greek', greek), data.get('matrix_align', matrix_align)
    if eqp:
        styles, sizes, spacing, notes = read_eqp(eqp)
        source = str(eqp)
    else:
        styles, sizes, spacing, notes = list(FACTORY_STYLES), list(FACTORY_SIZES), list(FACTORY_SPACING), []
        source = 'MathType factory (Times+Symbol 12)'
    if greek is not None:
        if greek not in ('symbol', 'times'):
            raise ValueError("greek must be 'symbol' or 'times'")
        font = 'Symbol' if greek == 'symbol' else styles[2][0]       # Times Greek: the Variable style's font
        styles[3], styles[4] = (font, 2), (font, 0)
    PROFILE = Profile(styles, sizes, spacing, matrix_align, source, notes)
    return PROFILE


def _eqn_prefs():
    fonts = PROFILE.font_defs()
    out = bytearray([0x12, 0x00]) + encode_dimensions(PROFILE.sizes) + encode_dimensions(PROFILE.spacing)
    out.append(12)                                   # 11 styles + Text FE
    for st in PROFILE.styles:
        out += bytes([fonts.index(st[0]) + 1, st[1]]) if st else b'\x00'
    return bytes(out)


def _font_name_bytes(name):
    """Font names as MathType writes them: Latin names as is; East Asian names in the ANSI code page (GBK for
    Chinese, e.g. SimSun = cb ce cc e5) with the WinAllCodePages encoding."""
    try:
        return name.encode('latin-1'), False
    except UnicodeEncodeError:
        return name.encode('gbk'), True


def _preamble(mathtype_version):
    version = mathtype_version.upper()
    assert version in {'DSMT4', 'DSMT6'}
    out = b'\x05\x01\x00\x06\x09' + version.encode('ascii') + b'\x00\x01' + b'\x13WinAllBasicCodePages\x00'
    extended = False
    for f in PROFILE.font_defs():
        raw, wide = _font_name_bytes(f)
        if wide and not extended:
            out += b'\x13WinAllCodePages\x00'           # second ENCODING_DEF: index 6
            extended = True
        enc = 6 if wide else ENCODINGS.get(f, 5)
        out += b'\x11' + bytes([enc]) + raw + b'\x00'
    return out + _eqn_prefs()


MT._preamble = _preamble

# ---------------------------------------------------------------------------------------------------------------------
# Characters
# ---------------------------------------------------------------------------------------------------------------------
# Adobe Symbol encoding (font positions); MTCode = the key's code point unless remapped in MTCODE
SYMBOL_POS = {' ': 0x20, '!': 0x21, '∀': 0x22, '#': 0x23, '∃': 0x24, '%': 0x25, '&': 0x26, '∋': 0x27, '(': 0x28, ')': 0x29,
              '∗': 0x2A, '*': 0x2A, '+': 0x2B, ',': 0x2C, '−': 0x2D, '-': 0x2D, '.': 0x2E, '/': 0x2F, ':': 0x3A, ';': 0x3B,
              '<': 0x3C, '=': 0x3D, '>': 0x3E, '?': 0x3F, '≅': 0x40, '[': 0x5B, '∴': 0x5C, ']': 0x5D, '⊥': 0x5E, '_': 0x5F,
              '{': 0x7B, '|': 0x7C, '}': 0x7D, '∼': 0x7E, '~': 0x7E, '′': 0xA2, '≤': 0xA3, '⁄': 0xA4, '∞': 0xA5, '↔': 0xAB,
              '←': 0xAC, '↑': 0xAD, '→': 0xAE, '↓': 0xAF, '°': 0xB0, '±': 0xB1, '″': 0xB2, '≥': 0xB3, '×': 0xB4, '∝': 0xB5,
              '∂': 0xB6, '•': 0xB7, '÷': 0xB8, '≠': 0xB9, '≡': 0xBA, '≈': 0xBB, '…': 0xBC, 'ℵ': 0xC0, 'ℑ': 0xC1, 'ℜ': 0xC2,
              '℘': 0xC3, '⊗': 0xC4, '⊕': 0xC5, '∅': 0xC6, '∩': 0xC7, '∪': 0xC8, '⊃': 0xC9, '⊇': 0xCA, '⊄': 0xCB, '⊂': 0xCC,
              '⊆': 0xCD, '∈': 0xCE, '∉': 0xCF, '∠': 0xD0, '∇': 0xD1, '∏': 0xD5, '√': 0xD6, '⋅': 0xD7, '·': 0xD7, '¬': 0xD8,
              '∧': 0xD9, '∨': 0xDA, '⇔': 0xDB, '⇐': 0xDC, '⇑': 0xDD, '⇒': 0xDE, '⇓': 0xDF, '◊': 0xE0, '〈': 0xE1, '⟨': 0xE1,
              '∑': 0xE5, '〉': 0xF1, '⟩': 0xF1, '∫': 0xF2}
GREEK_POS = dict(zip('ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ', range(0x41, 0x5B))) | \
    dict(zip('αβχδεφγηιϕκλμνοπθρστυϖωξψζ', range(0x61, 0x7B)))
GREEK_POS.update({'φ': 0x6A, 'ϕ': 0x66, 'ϵ': 0x65, 'ϱ': 0x72, 'ϒ': 0xA1})          # as in genuine equations
MTCODE = {'-': 0x2212, '*': 0x2217, '⟨': 0x2329, '⟩': 0x232A, '\u3008': 0x2329, '\u3009': 0x232A, '·': 0x22C5, '~': 0x223C}
FUNCTION_CHARS = set('()[]{}.,;:/')                 # Times (Function style) in genuine equations
SYMBOL_STYLE = {c for c in SYMBOL_POS if c not in FUNCTION_CHARS and not c.isalnum() and c != ' '}


def greek_face(ch, default=None):
    if 'α' <= ch <= 'ω' or ch in 'ϑϕϵϖϱ':
        return LC_GREEK
    if 'Α' <= ch <= 'Ω':
        return UC_GREEK
    return default


def char_record(ch, typeface, embell=b''):
    """CHAR record; Greek in a Symbol-font profile and Symbol-style characters get their Symbol position."""
    opt = 0x01 if embell else 0x00
    code = MTCODE.get(ch, ord(ch)) if typeface == SYMBOL else ord(ch)
    pos = None
    if typeface == SYMBOL and PROFILE.symbol_positions:
        pos = SYMBOL_POS.get(ch)
        if pos is None:
            UNMAPPED_SYMBOLS.add(ch)
    elif typeface in (LC_GREEK, UC_GREEK) and PROFILE.greek == 'symbol':
        pos = GREEK_POS.get(ch)
    if pos is not None:
        opt |= 0x04
    return (bytes([0x02, opt, typeface]) + code.to_bytes(2, 'little') + (bytes([pos]) if pos is not None else b'')
            + embell)


def space_char(code):
    return bytes([0x02, 0x00, SPACE_TF]) + code.to_bytes(2, 'little')


_orig_char = MT._encode_char


def encode_char(char):
    if len(char) != 1:
        return b''.join(encode_char(c) for c in char)
    if char in '  ':
        return space_char(0xEF02)
    if char in FUNCTION_CHARS:
        return char_record(char, FUNCTION)
    g = greek_face(char)
    if g:
        return char_record(char, g)
    if char in SYMBOL_STYLE:
        return char_record(char, SYMBOL)
    rec = _orig_char(char)
    if rec[0] == 0x02 and rec[2] == SYMBOL:
        return char_record(char, SYMBOL)
    return rec


MT._encode_char = encode_char


@dataclass(frozen=True)
class StyledText:
    """Characters of one MathType style: Vector (bold), Function (upright) or Text."""
    value: str
    typeface: int

    def encode(self) -> bytes:
        if self.typeface == TEXT:
            if self.value.strip('  ') == '':                 # \, -> thin space(s)
                return b''.join(space_char(0xEF02) for _ in self.value)
            return b''.join(bytes([0x02, 0x00, TEXT]) + (0x20 if c == ' ' else ord(c)).to_bytes(2, 'little')
                            for c in self.value)
        out = []
        for ch in self.value:
            if ch.isalpha():
                face = self.typeface if self.typeface == VECTOR else greek_face(ch, self.typeface)
                out.append(char_record(ch, face))
            else:
                out.append(encode_char(ch))
        return b''.join(out)


@dataclass(frozen=True)
class SpaceChar:
    code: int

    def encode(self) -> bytes:
        return space_char(self.code)


def _space_code(width):
    m = re.match(r'\s*([\d.]+)\s*em\s*$', width or '')
    w = float(m.group(1)) if m else 0.1667
    for limit, code in ((0.12, 0xEF08), (0.2, 0xEF02), (0.25, 0xEF03), (0.5, 0xEF04), (1.5, 0xEF05)):
        if w <= limit:
            return code
    return 0xEF06


# ---------------------------------------------------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------------------------------------------------
BIG = {'∫': (0x0F, 0x01, False), '∬': (0x0F, 0x02, False), '∭': (0x0F, 0x03, False),
       '∑': (0x10, 0x00, True), '∏': (0x11, 0x00, True)}
FENCE = {('(', ')'): (0x01, '(', ')'), ('[', ']'): (0x03, '[', ']'), ('{', '}'): (0x02, '{', '}'),
         ('|', '|'): (0x04, '', ''), ('‖', '‖'): (0x05, '', ''),
         ('\u27e8', '\u27e9'): (0x00, '\u2329', '\u232a'), ('\u3008', '\u3009'): (0x00, '\u2329', '\u232a'),
         ('\u2329', '\u232a'): (0x00, '\u2329', '\u232a'), ('⟦', '⟧'): (0x08, '〚', '〛')}
STOP = {'=', ',', '+', '-', '−', '≤', '≥', '<', '>', '≈', ';', ':=', '∈', '→'}
EMBELL = {'¯': 17, '‾': 17, '̅': 17, '_': 17, '~': 8, '˜': 8, '̃': 8, '^': 9, 'ˆ': 9, '̂': 9, '→': 11, '⃗': 11}


def _fence_char(ch):
    return bytes([0x02, 0x00, EXPAND]) + ord(ch).to_bytes(2, 'little')


@dataclass(frozen=True)
class BigOp:
    op: str
    body: object
    lower: object = None
    upper: object = None
    sum_style: bool = False

    def encode(self) -> bytes:
        sel, base, _ = BIG[self.op]
        var = base | (0x10 if self.lower is not None else 0) | (0x20 if self.upper is not None else 0)
        var |= 0x40 if self.sum_style else 0
        return (MT._template(sel, var, 0) + MT._line(self.body) + SUB + MT._maybe_line(self.lower)
                + MT._maybe_line(self.upper) + SYM + char_record(self.op, SYMBOL) + b'\x00' + FULL)


# one-sided fences (\left. X \right| for an evaluation bar, cases braces): selector and character of each side
OPEN = {k[0]: (v[0], v[1]) for k, v in FENCE.items()}
CLOSE = {k[1]: (v[0], v[2]) for k, v in FENCE.items()}


@dataclass(frozen=True)
class FenceT:
    body: object
    left: str
    right: str

    def encode(self) -> bytes:
        """Variation 0x01 = left fence present, 0x02 = right fence present; only the present sides carry a fence
        character (genuine MathType writes an evaluation bar as selector 4, variation 0x02, one EC08 character)."""
        if self.left and self.right:
            sel, lc, rc = FENCE[(self.left, self.right)]
        elif self.left:
            sel, lc = OPEN[self.left]; rc = None
        else:
            sel, rc = CLOSE[self.right]; lc = None
        var = (0x01 if lc else 0) | (0x02 if rc else 0)
        out = MT._template(sel, var, 0) + MT._line(self.body) + FULL
        for ch in (lc, rc):
            if ch:
                out += _fence_char(ch)
        return out + b'\x00'


@dataclass(frozen=True)
class EmbChar:
    ch: str
    typeface: int
    embell: int

    def encode(self) -> bytes:
        return char_record(self.ch, self.typeface, embell=bytes([0x06, 0x00, self.embell, 0x00]))


MT._empty_line = lambda: b'\x01\x01'                    # null LINE has no END record


def _sub_encode(self):
    return (self.base.encode() + b'\x03\x00\x1b\x00\x00' + SUB + MT._line(self.subscript) + MT._empty_line()
            + b'\x00' + FULL)


MT.Subscript.encode = _sub_encode
_orig_matrix_encode = MT.Matrix.encode


def _matrix_encode(self):
    b = bytearray(_orig_matrix_encode(self))
    assert b[0] == 0x05 and b[1] == 0x00
    b[2], b[3], b[4] = 0x01, MATRIX_ALIGN[PROFILE.matrix_align], 0x01      # valign, h_just, v_just
    return bytes(b)


MT.Matrix.encode = _matrix_encode

# ---------------------------------------------------------------------------------------------------------------------
# MathML parsing
# ---------------------------------------------------------------------------------------------------------------------


def _ln(el):
    from lxml import etree
    return etree.QName(el).localname


def _txt(el):
    return MM._operator_text(el) if _ln(el) in {'mo', 'mi', 'mn', 'mtext'} or len(el) == 1 else ''


def _accent(element):
    """EmbChar for <mover>(one character)(accent)</mover>, else None."""
    if _ln(element) != 'mover' or len(element) < 2:
        return None
    acc = MM._operator_text(element[1])
    base = element[0]
    while _ln(base) == 'mrow' and len(base) == 1:
        base = base[0]
    if acc not in EMBELL or _ln(base) not in {'mi', 'mn'}:
        return None
    text = MM._normalize_text(''.join(base.itertext()))
    if len(text) != 1:
        return None
    var = base.get('mathvariant')
    if var in ('bold', 'bold-italic'):
        tf = VECTOR
    elif _ln(base) == 'mn':
        tf = NUMBER
    else:
        tf = greek_face(text, FUNCTION if var == 'normal' else VARIABLE)
    return EmbChar(text, tf, EMBELL[acc])


_orig_parse_element = MM._parse_element


def _parse_element(element):
    tag = _ln(element)
    if tag == 'mspace':
        if element.get('linebreak'):
            raise ValueError('line break inside a formula: use an array instead')
        return SpaceChar(_space_code(element.get('width')))
    acc = _accent(element)
    if acc is not None:
        return acc
    if tag in {'mi', 'mn', 'mo', 'mtext'}:
        text = MM._normalize_text(''.join(element.itertext()))
        var = element.get('mathvariant')
        if tag == 'mtext':
            return StyledText(text, TEXT)
        if var in ('bold', 'bold-italic', 'bold-upright'):
            return StyledText(text, VECTOR)
        if tag == 'mi' and (var == 'normal' or len(text) > 1):
            return StyledText(text, FUNCTION)
    return _orig_parse_element(element)


MM._parse_element = _parse_element


def _bigop_parts(el):
    tag = _ln(el)
    if tag == 'mo' and _txt(el) in BIG:
        return _txt(el), None, None, False
    if tag in {'msub', 'munder', 'msubsup', 'munderover', 'msup', 'mover'} and len(el) >= 2 and _txt(el[0]) in BIG:
        op = _txt(el[0])
        lower = upper = None
        if tag in {'msub', 'munder'}:
            lower = el[1]
        elif tag in {'msup', 'mover'}:
            upper = el[1]
        else:
            lower, upper = el[1], el[2]
        return op, lower, upper, BIG[op][2] and tag in {'munder', 'munderover', 'mover'}
    return None


def _is_open(el, children, i):
    t = _txt(el)
    return el.get('form') == 'prefix' or t in {'⟨', '⟦'} or (t in {'[', '('} and i + 1 < len(children) and _ln(children[i + 1]) == 'mtable')


def _is_close(el, children, j):
    t = _txt(el)
    return el.get('form') == 'postfix' or t in {'⟩', '⟧'} or (t in {']', ')'} and j > 0 and _ln(children[j - 1]) == 'mtable')


def _parse_list(children):
    """Siblings -> objects, turning fence pairs and big operators (with their integrands) into templates."""
    out, i = [], 0
    while i < len(children):
        el = children[i]
        b = _bigop_parts(el)
        if b:
            j = i + 1
            while j < len(children) and not (_ln(children[j]) == 'mo' and _txt(children[j]) in STOP):
                j += 1
            op, lo, up, ss = b
            body = _parse_list(children[i + 1:j]) if j > i + 1 else MT.Text('')
            out.append(BigOp(op, body, MM._parse_element(lo) if lo is not None else None,
                             MM._parse_element(up) if up is not None else None, ss))
            i = j
            continue
        if _ln(el) == 'mo' and _is_open(el, children, i):
            left, depth, close = _txt(el), 0, None
            for j in range(i, len(children)):
                c = children[j]
                if _ln(c) != 'mo':
                    continue
                if c is el or _is_open(c, children, j):
                    depth += 1
                elif _is_close(c, children, j):
                    depth -= 1
                    if depth == 0:
                        close = j
                        break
            right = _txt(children[close]) if close is not None else None
            if close is not None and ((left, right) in FENCE or (not left and right in CLOSE)
                                      or (not right and left in OPEN)):
                out.append(FenceT(_parse_list(children[i + 1:close]), left, right))
                i = close + 1
                continue
            if left == '{' and close is None:                   # cases: left brace only
                out.append(FenceT(_parse_list(children[i + 1:]), '{', ''))
                break
        out.append(MM._parse_element(el))
        i += 1
    return MT.seq(out) if out else MT.Text('')


MM._parse_children = lambda element: _parse_list([c for c in element if _ln(c) != 'annotation'])

if os.environ.get('MTEF_PROFILE_JSON'):
    configure(profile_json=os.environ['MTEF_PROFILE_JSON'])
