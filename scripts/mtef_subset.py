#!/usr/bin/env python3
"""Conservative self-owned MTEF v5 encoder for a documented LaTeX subset.

The output uses a *style preamble* from an actual user-provided MathType v5
object; only the mathematical structure is synthesized, not copied from the
original formula. This is a compatibility experiment, not a full MTEF encoder.

Supported: Latin/Greek variables, digits, mathematical punctuation, simple
scripts/fractions/square-roots, and MathType semantic text/function/vector roles.
Unknown or ambiguous LaTeX commands fail, rather than silently losing styles.
"""
from __future__ import annotations
from dataclasses import dataclass
import re
import struct

GREEK={
 'alpha': '\u03b1', 'beta':'\u03b2','gamma':'\u03b3','delta':'\u03b4',
 'varepsilon':'\u03b5','zeta':'\u03b6','eta':'\u03b7',
 'theta':'\u03b8','lambda':'\u03bb','mu':'\u03bc','nu':'\u03bd',
 'xi':'\u03be','pi':'\u03c0','rho':'\u03c1','sigma':'\u03c3',
 'tau':'\u03c4','phi':'\u03c6','chi':'\u03c7',
 'psi':'\u03c8','omega':'\u03c9','Gamma':'\u0393','Delta':'\u0394',
 'Theta':'\u0398','Lambda':'\u039b','Pi':'\u03a0','Sigma':'\u03a3',
 'Phi':'\u03a6','Omega':'\u03a9',
}
# 8-bit positions within the Symbol encoding, where known. For a few glyphs
# no 8-bit encoding is emitted; MTCode remains the unambiguous Unicode value.
SYM_POS={
 '\u03b1':'a','\u03b2':'b','\u03b3':'g','\u03b4':'d','\u03b5':'e',
 '\u03b8':'q','\u03bb':'l','\u03bc':'m','\u03bd':'n','\u03be':'x',
 '\u03c0':'p','\u03c1':'r','\u03c3':'s','\u03c4':'t','\u03c6':'f',
 '\u03c7':'c','\u03c8':'y','\u03c9':'w',
 '\u0393':'G','\u0394':'D','\u0398':'Q','\u039b':'L',
 '\u03a0':'P','\u03a3':'S','\u03a6':'F','\u03a9':'W',
}
SYMBOLS={
 '=': (0x003D,0x3D), '+':(0x002B,0x2B),'-':(0x2212,0x2D),
 '*':(0x00D7,0xB4), '/':(0x002F,0x2F),'(':(0x0028,0x28),
 ')':(0x0029,0x29), ',':(0x002C,0x2C), '.':(0x002E,0x2E),
 ':':(0x003A,0x3A),
}
TOKENS=re.compile(r'\\[a-zA-Z]+|[A-Za-z0-9_^{}/=+*(),.\-:]|\s+')
FUNCTIONS={'sin','cos','tan','exp','log','ln','lim'}
STYLED_COMMANDS={'mathrm':'text','text':'text','operatorname':'function',
                 'mathbf':'vector','mathit':'variable'}
STYLE_TYPEFACE={'text':0x81,'function':0x82,'variable':0x83,
                'lcgreek':0x84,'ucgreek':0x85,'symbol':0x86,
                'vector':0x87,'number':0x88}

@dataclass(frozen=True)
class Glyph:
    ch: str
    role: str = "auto"
    function_start: bool = False
@dataclass(frozen=True)
class Script:
    sub: tuple | None
    sup: tuple | None
@dataclass(frozen=True)
class Fraction:
    num: tuple
    den: tuple
@dataclass(frozen=True)
class Root:
    rad: tuple

class LatexSubset:
    def __init__(self, latex: str):
        self.original=latex
        if re.search(r'\\(?:mathrm|text|mathbf|mathit|operatorname)\{[^{}]*\s',latex):
            raise ValueError('Whitespace in styled text is not supported; use a proper text encoder')
        self.tokens=[]
        pos=0
        for m in TOKENS.finditer(latex):
            if m.start()!=pos:
                raise ValueError(f'Unsupported LaTeX at char offset {pos}')
            pos=m.end()
            if not m.group().isspace():self.tokens.append(m.group())
        if pos!=len(latex):raise ValueError(f'Unsupported LaTeX at char offset {pos}')
        self.i=0
    def peek(self):
        return self.tokens[self.i] if self.i<len(self.tokens) else None
    def consume(self):
        t=self.peek()
        if t is None:raise ValueError('Unexpected end of LaTeX input')
        self.i+=1
        return t
    def group(self):
        if self.peek()=='{':
            self.consume(); content=self.parse(stop=True)
            if self.consume()!='}':raise ValueError('Missing close brace')
            return content
        return self.parse_atom()
    def parse_atom(self):
        t=self.consume()
        if t=='{':
            chunk=self.parse(stop=True)
            if self.consume()!='}':raise ValueError('Missing close brace')
            return chunk
        if t.startswith('\\'):
            cmd=t[1:]
            if cmd=='frac':
                return (Fraction(self.group(),self.group()),)
            if cmd=='sqrt':
                return (Root(self.group()),)
            if cmd in FUNCTIONS:
                return tuple(Glyph(c,'function',i==0) for i,c in enumerate(cmd))
            if cmd in STYLED_COMMANDS:
                if self.peek()!='{':
                    raise ValueError('Style command requires braced group: '+t)
                return _apply_style(self.group(),STYLED_COMMANDS[cmd])
            if cmd not in GREEK:
                raise ValueError('LaTeX command not supported: '+t)
            return (Glyph(GREEK[cmd]),)
        if t in ('^','_','}'):
            raise ValueError('Stray LaTeX script/group token: '+t)
        return (Glyph(t),)
    def parse(self,stop=False):
        result=[]
        while self.peek() is not None and self.peek()!='}':
            if self.peek() in ('^','_'):
                raise ValueError('Script marker without base symbol')
            result.extend(self.parse_atom())
            sub=sup=None
            while self.peek() in ('^','_'):
                typ=self.consume()
                val=self.group()
                if typ=='_':
                    if sub is not None:raise ValueError('Duplicate subscript')
                    sub=val
                else:
                    if sup is not None:raise ValueError('Duplicate superscript')
                    sup=val
            if sub is not None or sup is not None:
                result.append(Script(sub,sup))
        if not stop and self.peek() is not None:
            raise ValueError('Unexpected brace')
        return tuple(result)
    def read(self):
        result=self.parse()
        if self.i!=len(self.tokens):raise ValueError('Unconsumed LaTeX')
        if not result:raise ValueError('Empty formula')
        return result


def _apply_style(nodes: tuple, role: str) -> tuple:
    """Apply semantic MathType style without collapsing the mathematical AST."""
    out=[]
    for idx,node in enumerate(nodes):
        if isinstance(node,Glyph):
            if role=='vector' and (not node.ch.isascii() or not node.ch.isalpha()):
                raise ValueError('Bold style currently supports Latin variables only')
            if role in ('function','text','variable') and (not node.ch.isascii() or not node.ch.isalpha()):
                raise ValueError('Styled text currently supports Latin letters only')
            out.append(Glyph(node.ch,role,node.function_start or (role=='function' and idx==0)))
        elif isinstance(node,Script):
            out.append(Script(_apply_style(node.sub,role) if node.sub else None,
                              _apply_style(node.sup,role) if node.sup else None))
        else:
            raise ValueError('Font override over template is not yet supported')
    return tuple(out)


def semantic_role(node: Glyph) -> str:
    if node.role!='auto':return node.role
    ch=node.ch
    if ch in SYM_POS:
        return 'ucgreek' if ch.isupper() else 'lcgreek'
    if ch.isascii() and ch.isalpha():return 'variable'
    if ch in '0123456789':return 'number'
    if ch in SYMBOLS:return 'symbol'
    raise ValueError('Unsupported mathematical glyph '+repr(ch))


def glyph_record(glyph: Glyph | str) -> bytes:
    """Emit a CHAR record whose typeface selects the correct MathType style."""
    if isinstance(glyph,str):glyph=Glyph(glyph)
    ch=glyph.ch
    n=ord(ch)
    if n>0xFFFF:raise ValueError('Unicode above BMP unsupported')
    role=semantic_role(glyph)
    tp=STYLE_TYPEFACE[role]
    func=2 if glyph.function_start else 0
    if ch in SYM_POS:
        if role not in ('lcgreek','ucgreek'):
            raise ValueError('Greek font override not yet supported')
        return b'\x02'+bytes([4|func,tp])+struct.pack('<H',n)+SYM_POS[ch].encode('ascii')
    if ch.isascii() and ch.isalpha():
        if role not in ('variable','function','text','vector'):
            raise ValueError('Unsupported Latin role: '+role)
        return b'\x02'+bytes([func,tp])+struct.pack('<H',n)
    if ch in '0123456789':
        if role!='number':raise ValueError('Digit override not supported')
        return b'\x02'+bytes([func,tp])+struct.pack('<H',n)
    if ch in SYMBOLS:
        if role!='symbol':raise ValueError('Symbol override not supported')
        code,font=SYMBOLS[ch]
        return b'\x02'+bytes([4|func,tp])+struct.pack('<H',code)+bytes([font])
    raise ValueError(f'No MTEF glyph encoding for {ch!r}')


def _slot(content: tuple | None) -> bytes:
    if content is None:
        return b'\x01\x01' # null line, no END
    return b'\x01\x00'+encode_nodes(content)+b'\x00'


def encode_nodes(nodes: tuple) -> bytes:
    out=bytearray()
    for n in nodes:
        if isinstance(n,Glyph):
            out+=glyph_record(n)
        elif isinstance(n,Script):
            # 27 subscript, 28 superscript, 29 sub+sup. MTEF template
            # contains sub then super slots, with explicit size transitions.
            selector=29 if n.sub is not None and n.sup is not None else (27 if n.sub is not None else 28)
            out+=bytes([3,0,selector,0,0,11])
            out+=_slot(n.sub)
            out+=_slot(n.sup)
            out+=b'\x00\x0a' # END template, restore full size
        elif isinstance(n,Fraction):
            out+=b'\x03\x00\x0b\x00\x00'
            out+=_slot(n.num)+_slot(n.den)+b'\x00'
        elif isinstance(n,Root):
            out+=b'\x03\x00\x0a\x00\x00'
            out+=_slot(n.rad)
            out+=b'\x0b'+_slot(None)+b'\x00\x0a'
        else:raise TypeError(type(n))
    return bytes(out)


def find_style_prefix(original_mtef: bytes) -> bytes:
    if len(original_mtef)<200 or original_mtef[0]!=5:
        raise ValueError('Only real MTEF v5 MathType donor with style definitions')
    # Sampled real MTEF v5 styles have a FULL size followed by a LINE(0)
    # immediately before formula contents. Unlike a blind split, ensure
    # expected key and selected preamble structure are present.
    if b'DSMT' not in original_mtef[:30]:
        raise ValueError('Expected MathType style donor (DSMT)')
    candidates=[i for i in range(170,min(350,len(original_mtef)-2))
                if original_mtef[i:i+3]==b'\x0a\x01\x00']
    if len(candidates)!=1:
        raise ValueError(f'Ambiguous/unsupported style preamble: {candidates}')
    result=original_mtef[:candidates[0]]
    if b'\x12\x00' not in result[80:150]:
        raise ValueError('No equation preferences in style donor')
    return result


def make_mtef(latex: str, style_seed: bytes, inline=False) -> bytes:
    ast=LatexSubset(latex).read()
    preamble=bytearray(find_style_prefix(style_seed))
    # MTEF header: 5-byte version metadata, NUL-terminated app key,
    # followed by 1 byte equation options (0=display / 1=inline).
    app_end=preamble.index(b'\x00',5)
    eqopt=app_end+1
    if eqopt>=len(preamble):raise ValueError('Missing MTEF equation options')
    preamble[eqopt]=1 if inline else 0
    return bytes(preamble)+b'\x0a\x01\x00'+encode_nodes(ast)+b'\x00\x00'


def validate_subset_mtef(mtef: bytes) -> dict:
    # Validate nested object lists and char records for *our generated subset*.
    start=mtef.find(b'\x0a\x01\x00',170)
    if start<0:raise ValueError('Missing initial MTEF SIZE_FULL/LINE')
    i=start+3
    count=0
    templates=0
    style_counts={}
    fn_starts=0
    chars=[]
    def parse_list(end_required=True):
        nonlocal i,count,templates,fn_starts
        while i<len(mtef):
            r=mtef[i];i+=1
            if r==0:return
            if r==2:
                if i+2>len(mtef):raise ValueError('Truncated CHAR')
                opts,tp=mtef[i:i+2];i+=2
                if opts not in (0,2,4,6):raise ValueError('Unsupported CHAR flags in generated MTEF')
                if tp not in STYLE_TYPEFACE.values():
                    raise ValueError('Unknown typeface in generated MTEF: '+str(tp))
                if i+2>len(mtef):raise ValueError('Missing MTCode')
                code=int.from_bytes(mtef[i:i+2],'little'); i+=2
                fontpos=None
                if opts&4:
                    if i>=len(mtef):raise ValueError('Missing 8-bit font position')
                    fontpos=mtef[i];i+=1
                role=next(k for k,v in STYLE_TYPEFACE.items() if v==tp)
                ch=chr(code)
                if role in ('text','function','variable','vector'):
                    if not ch.isascii() or not ch.isalpha() or fontpos is not None:
                        raise ValueError('Invalid Latin char/style encoding')
                elif role=='number':
                    if ch not in '0123456789' or fontpos is not None:
                        raise ValueError('Invalid number encoding')
                elif role in ('lcgreek','ucgreek'):
                    if ch not in SYM_POS or fontpos!=ord(SYM_POS[ch]) or ch.isupper() != (role=='ucgreek'):
                        raise ValueError('Invalid Greek font position/typeface')
                elif role=='symbol':
                    if fontpos is None:
                        raise ValueError('Missing Symbol font position')
                if opts&2:
                    if role!='function':raise ValueError('Function-start flag used outside Function style')
                    fn_starts+=1
                style_counts[role]=style_counts.get(role,0)+1
                chars.append({'char':ch,'role':role,'function_start':bool(opts&2), 'font_pos':fontpos})
                count+=1
            elif r==3:
                if i+4>len(mtef):raise ValueError('Truncated TMPL')
                opt,sel,var,topt=mtef[i:i+4];i+=4
                if opt or var or topt or sel not in (10,11,27,28,29):
                    raise ValueError('Unsupported template in generated MTEF')
                templates+=1
                if sel in (27,28,29):
                    if mtef[i]!=11: raise ValueError('Expected sub size')
                    i+=1
                parse_slot()
                if sel==10:
                    if mtef[i]!=11:raise ValueError('Missing radical size SUB')
                    i+=1
                parse_slot()
                if mtef[i]!=0:raise ValueError('Missing template end')
                i+=1
                if sel in (10,27,28,29):
                    if mtef[i]!=10:raise ValueError('Missing size reset')
                    i+=1
            elif r==10 or r==11:
                # Explicit SIZE_FULL/SUB controls are legal in a nested line
                continue
            else:raise ValueError(f'Unexpected MTEF record {r} at byte {i-1}')
        raise ValueError('Missing END record')
    def parse_slot():
        nonlocal i
        if mtef[i]!=1:raise ValueError('Expected LINE slot')
        i+=1
        flag=mtef[i];i+=1
        if flag==1:return
        if flag!=0:raise ValueError('Bad LINE slot options')
        parse_list()
    parse_list()
    if i>=len(mtef) or mtef[i]!=0:
        raise ValueError('Missing final MTEF END')
    i+=1
    if i!=len(mtef):raise ValueError('Trailing data in generated MTEF')
    return {'glyphs':count,'templates':templates,'length':len(mtef),
            'style_counts':style_counts,'function_starts':fn_starts,'characters':chars}
