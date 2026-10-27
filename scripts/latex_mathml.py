"""LaTeX -> MathML for the MathType backend (latex2mathml + fixes).

- macros latex2mathml does not know (\\tfrac, \\dfrac, \\big..., \\llbracket) are normalised; \\, and \; become
  spaces (latex2mathml turns \; into a line break);
- Mathematical Alphanumeric Symbols (outside the BMP, rejected by the MTEF writer) become BMP letters with mathvariant;
- double brackets become fence operators; options: capital Greek upright (ISO), bold symbols upright.

    from latex_mathml import to_mathml; to_mathml(r'\\int_\\Omega f\\,d\\Omega', display=True)
"""
import re

from latex2mathml.converter import convert

SUBS = [(r'\\tfrac', r'\\frac'), (r'\\dfrac', r'\\frac'),
        (r'\\big(?=[(\[\]\)|.])', ''), (r'\\Big(?=[(\[\]\)|.])', ''), (r'\\bigg(?=[(\[\]\)|.])', ''),
        (r'\\big\\', '\\\\'), (r'\\Big\\', '\\\\'), (r'\\!', ''), (r'\;', r'\\ '), (r'\\,', r'\\ ')]


def normalise(tex):
    tex = tex.strip().replace(r'[\![', '⟦').replace(r']\!]', '⟧')
    tex = tex.replace(r'\llbracket', '⟦').replace(r'\rrbracket', '⟧')
    for a, b in SUBS:
        tex = re.sub(a, b, tex)
    return re.sub(r'\\mathsf\{?T\}?', 'T', tex)            # transpose: plain T


def _alnum_map():
    m = {}
    for base, start, var in (('A', 0x1D400, 'bold'), ('a', 0x1D41A, 'bold'), ('A', 0x1D434, 'italic'), ('a', 0x1D44E, 'italic'),
                             ('A', 0x1D468, 'bold-italic'), ('a', 0x1D482, 'bold-italic')):
        for k in range(26):
            m[chr(start + k)] = (chr(ord(base) + k), var)
    m['\u210E'] = ('h', 'italic')
    for k in range(10):
        m[chr(0x1D7CE + k)] = (chr(ord('0') + k), 'bold')
    greek = 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡϴΣΤΥΦΧΨΩ∇αβγδεζηθικλμνξοπρςστυφχψω∂ϵϑϰϕϱϖ'
    for k, g in enumerate(greek):
        m[chr(0x1D6A8 + k)], m[chr(0x1D6E2 + k)], m[chr(0x1D71C + k)] = (g, 'bold'), (g, 'italic'), (g, 'bold-italic')
    return m


AL = _alnum_map()


def fix_mathml(mml):
    def tok(mo):
        tag, attrs, text = mo.groups()
        if not any(ord(c) > 0xFFFF for c in text):
            return mo.group(0)
        out, var = [], None
        for c in text:
            if c in AL:
                ch, var = AL[c]
                out.append(ch)
            elif ord(c) > 0xFFFF:
                raise ValueError(f'unmapped non-BMP character {c!r} (U+{ord(c):X})')
            else:
                out.append(c)
        attrs = re.sub(r'\s*mathvariant="[^"]*"', '', attrs)
        return f'<{tag}{attrs} mathvariant="{var}">{"".join(out)}</{tag}>'
    return re.sub(r'<(mi|mn|mo|mtext)([^>]*)>([^<]*)</\1>', tok, mml)


def to_mathml(tex, display=False, upright_capital_greek=True, bold_upright=True):
    mml = convert(normalise(tex), display='block' if display else 'inline')
    mml = re.sub(r'&#x0*(1[0-9A-Fa-f]{4});', lambda m: chr(int(m.group(1), 16)), mml)   # non-BMP entities -> chars
    mml = fix_mathml(mml)
    if upright_capital_greek:
        def upright(m):
            c = m.group(1)
            ch = chr(int(c[3:-1], 16)) if c.startswith('&#x') else c
            return f'<mi mathvariant="normal">{c}</mi>' if 0x391 <= ord(ch) <= 0x3A9 else m.group(0)
        mml = re.sub(r'<mi>(&#x[0-9A-Fa-f]+;|[^<&])</mi>', upright, mml)
    if bold_upright:
        mml = mml.replace('mathvariant="bold-italic"', 'mathvariant="bold"')
    mml = re.sub(r'<mi>(\u27e6|&#x0*27E6;)</mi>', r'<mo fence="true" form="prefix">\1</mo>', mml)
    mml = re.sub(r'<mi>(\u27e7|&#x0*27E7;)</mi>', r'<mo fence="true" form="postfix">\1</mo>', mml)
    bad = [c for c in mml if ord(c) > 0xFFFF]
    if bad:
        raise ValueError(f'non-BMP characters left: {bad}')
    return mml
