#!/usr/bin/env python3
"""Audit real MathType MTEF v5 font definitions, style flags and base size.

Read-only and fail-closed. This does not claim to validate MathType rendering.
Uses the documented FONT_DEF and EQN_PREFS record structure, not heuristic
string extraction from the OLE binary blob.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from zipfile import ZipFile
from ole_editor import FormulaDocument
from ole_native import CompoundFile

STYLE_NAMES=['Text','Function','Variable','L.C. Greek','U.C. Greek',
             'Symbol','Vector','Number','User 1','User 2','MT Extra','Text FE']
UNIT_NAMES={0:'inches',1:'cm',2:'pt',3:'picas',4:'%'}


def _uint(data, pos):
    if pos>=len(data): raise ValueError('Truncated unsigned integer')
    if data[pos]!=255: return data[pos],pos+1
    if pos+3>len(data): raise ValueError('Truncated extended unsigned integer')
    return int.from_bytes(data[pos+1:pos+3],'little'),pos+3


def _dimensions(data, pos):
    if pos>=len(data): raise ValueError('Truncated dimension array')
    count=data[pos];pos+=1
    nibble_idx=0
    def nibble():
        nonlocal nibble_idx
        i=pos+nibble_idx//2
        if i>=len(data): raise ValueError('Truncated dimension value')
        b=data[i];out=b>>4 if nibble_idx%2==0 else b&15
        nibble_idx+=1
        return out
    values=[]
    for _ in range(count):
        unit=nibble();digits=[]
        for unused in range(512):
            token=nibble()
            if token==15: break
            if token<=9:digits.append(str(token))
            elif token==10:digits.append('.')
            elif token==11:digits.append('-')
            else: raise ValueError('Unsupported dimension nibble')
        else:raise ValueError('Dimension text too long')
        values.append({'value':''.join(digits),'unit':UNIT_NAMES.get(unit,str(unit))})
    return values, pos+(nibble_idx+1)//2


def read_font_profile(mtef: bytes):
    if len(mtef)<20 or mtef[0]!=5: raise ValueError('MTEF v5 required')
    end=mtef.index(b'\x00',5)
    pos=end+2  # application key and equation options
    fonts=[]; encodings=[]; prefs=None; explicit_font_styles=[]
    while pos<len(mtef):
        rec=mtef[pos]; pos+=1
        if rec==10:break # SIZE_FULL is the start of equation body
        if rec==19:
            last=mtef.index(0,pos)
            encodings.append(mtef[pos:last].decode('latin1'))
            pos=last+1
        elif rec==17:
            encoding,pos=_uint(mtef,pos)
            last=mtef.index(0,pos)
            fonts.append({'encoding_index':encoding,
                          'name':mtef[pos:last].decode('latin1')})
            pos=last+1
        elif rec==8:
            font_id,pos=_uint(mtef,pos)
            if pos>=len(mtef):raise ValueError('Truncated FONT_STYLE_DEF')
            flags=mtef[pos];pos+=1
            explicit_font_styles.append({'index':len(explicit_font_styles)+1,
                              'font_id':font_id,
                              'font':fonts[font_id-1]['name'] if 0<font_id<=len(fonts) else None,
                              'bold':bool(flags&1),'italic':bool(flags&2)})
        elif rec==18:
            if prefs is not None: raise ValueError('Multiple EQN_PREFS records')
            options=mtef[pos];pos+=1
            sizes,pos=_dimensions(mtef,pos)
            spacing,pos=_dimensions(mtef,pos)
            style_cnt=mtef[pos];pos+=1
            styles=[]
            for i in range(style_cnt):
                font_id,pos=_uint(mtef,pos)
                char_style=None
                if font_id:
                    char_style=mtef[pos];pos+=1
                fname=fonts[font_id-1]['name'] if 0<font_id<=len(fonts) else None
                styles.append({ 'name':STYLE_NAMES[i] if i<len(STYLE_NAMES) else f'Extra {i}',
                               'font_id':font_id,'font':fname,
                               'bold':bool(char_style&1) if char_style is not None else False,
                               'italic':bool(char_style&2) if char_style is not None else False})
            prefs={'options':options,'sizes':sizes,'spacing':spacing,'styles':styles}
        else:
            raise ValueError(f'Unsupported MTEF definition record {rec} at {pos-1}')
    if prefs is None: raise ValueError('No EQN_PREFS record')
    return {'mtef_version':5, 'application':mtef[5:end].decode('ascii', errors='replace'),
            'fonts':fonts,'encodings':encodings,'explicit_font_styles':explicit_font_styles,
            'full_size':prefs['sizes'][0] if prefs['sizes'] else None,
            'styles':prefs['styles'],'sizes':prefs['sizes']}


def audit_docx(path, indexes=(1,)):
    d=FormulaDocument(path)
    with ZipFile(path) as z:
        result=[]
        for index in indexes:
            f=d.by_index(index)
            native=CompoundFile(z.read(f.object_part)).read('Equation Native')
            result.append({'formula_index':index, 'object_part':f.object_part,
                           'profile':read_font_profile(native[28:])})
    return result


def main():
    p=argparse.ArgumentParser()
    p.add_argument('docx',type=Path)
    p.add_argument('--indexes',type=int,nargs='*',default=[1])
    args=p.parse_args()
    print(json.dumps(audit_docx(args.docx,args.indexes),ensure_ascii=False,indent=2))

if __name__=='__main__':main()
