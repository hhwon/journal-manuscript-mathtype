#!/usr/bin/env python3
"""Experimental MTEF 5 explicit-font CHAR backend (separate from semantic styles).

Explicit FONT_STYLE_DEF references force direct Times New Roman (upright,
italic, bold) or Symbol (upright, italic) instead of relying on MathType's
Variable/Function/Greek style mapping. Do not use this as a production default
until a real MathType 7/Word test confirms font fidelity and edit/save behavior.
"""
from __future__ import annotations
from dataclasses import dataclass
from mtef_subset import (LatexSubset,Glyph,Script,Fraction,Root,semantic_role,
                         SYM_POS,SYMBOLS,find_style_prefix)
import struct

# Use existing FONT_DEF records in the original MathType style preamble:
# font #1 = Times New Roman; #2 = Symbol. Explicit font styles are assigned
# sequential indexes starting at 1 and selected by negative typeface values.
FONT_STYLES=(
    ('times_italic',1,2),
    ('times_upright',1,0),
    ('times_bold',1,1),
    ('symbol_italic',2,2),
    ('symbol_upright',2,0),
)
ROLE_TO_FONT={
    'variable':1,'function':2,'text':2,'vector':3,'number':2,
    'lcgreek':4,'ucgreek':5,'symbol':5,
}


def font_style_records():
    return b''.join(bytes((8,fid,style)) for _,fid,style in FONT_STYLES)


def explicit_char(g):
    ch=g.ch;role=semantic_role(g)
    num=ROLE_TO_FONT[role]
    tp=128-num  # MTEF signed typeface value, biased by 128
    flag=2 if g.function_start else 0
    if ch in SYM_POS:
        return b'\x02'+bytes((4|flag,tp))+struct.pack('<H',ord(ch))+SYM_POS[ch].encode('ascii')
    if ch in SYMBOLS:
        mt,enc=SYMBOLS[ch]
        return b'\x02'+bytes((4|flag,tp))+struct.pack('<H',mt)+bytes((enc,))
    if ch.isascii() and ch.isalnum():
        return b'\x02'+bytes((flag,tp))+struct.pack('<H',ord(ch))
    raise ValueError('Unsupported explicit font glyph: '+repr(ch))


def _slot(content):
    if content is None:return b'\x01\x01'
    return b'\x01\x00'+_nodes(content)+b'\x00'


def _nodes(nodes):
    out=bytearray()
    for n in nodes:
        if isinstance(n,Glyph):out+=explicit_char(n)
        elif isinstance(n,Script):
            sel=29 if n.sub is not None and n.sup is not None else (27 if n.sub is not None else 28)
            out+=bytes((3,0,sel,0,0,11))
            out+=_slot(n.sub)+_slot(n.sup)+b'\x00\x0a'
        elif isinstance(n,Fraction):
            out+=b'\x03\x00\x0b\x00\x00'+_slot(n.num)+_slot(n.den)+b'\x00'
        elif isinstance(n,Root):
            out+=b'\x03\x00\x0a\x00\x00'+_slot(n.rad)+b'\x0b'+_slot(None)+b'\x00\x0a'
        else:raise TypeError(type(n))
    return bytes(out)


def make_mtef_explicit(latex: str, seed: bytes,inline:bool=False):
    tree=LatexSubset(latex).read()
    preamble=bytearray(find_style_prefix(seed))
    if b'Times New Roman\x00' not in preamble or b'Symbol\x00' not in preamble:
        raise ValueError('A real TNR/Symbol MathType donor is required')
    # Reject unconventional font order rather than assigning the wrong font.
    ix_t=preamble.index(b'Times New Roman\x00');ix_s=preamble.index(b'Symbol\x00')
    if not (ix_t<ix_s):raise ValueError('Unexpected source font definition order')
    key_end=preamble.index(b'\x00',5)
    preamble[key_end+1]=1 if inline else 0
    return bytes(preamble)+font_style_records()+b'\x0a\x01\x00'+_nodes(tree)+b'\x00\x00'


def validate_explicit(mtef):
    d=font_style_records()
    begin=mtef.find(d)
    if begin<0:raise ValueError('Expected 5 explicit font style records')
    start=mtef.find(b'\x0a\x01\x00',begin+len(d))
    if start!=begin+len(d):raise ValueError('FONT_STYLE_DEF records not immediately before formula body')
    chars={}
    for idx in range(start+3,len(mtef)-2):
        if mtef[idx]==2 and idx+2<len(mtef) and mtef[idx+2] in range(123,128):
            v=mtef[idx+2]
            chars[v]=chars.get(v,0)+1
    if not chars:raise AssertionError('No explicit-style characters found')
    return {'style_defs':len(FONT_STYLES),'explicit_typefaces':chars,'length':len(mtef)}
