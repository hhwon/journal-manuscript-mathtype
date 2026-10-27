#!/usr/bin/env python3
"""Stable manuscript IDs, Word SEQ/REF fields and real OOXML bookmarks.

Caption numbers use SEQ fields and equations use the MathType backend's SEQ
fields. Internal references are Word REF fields with cached visible values.
"""
from __future__ import annotations
from dataclasses import dataclass
from hashlib import sha1
from collections.abc import Mapping
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


@dataclass(frozen=True)
class Anchor:
    identifier: str
    kind: str
    number: int
    bookmark: str
    bookmark_id: int


def _bookmark_name(identifier: str, kind: str) -> str:
    prefix = {'equation': 'EQ', 'figure': 'FIG', 'table': 'TAB'}[kind]
    return 'MS_' + prefix + '_' + sha1(identifier.encode('utf-8')).hexdigest()[:24]


def make_anchors(spec: Mapping[str, Any]) -> tuple[dict[str, Anchor], dict[int, Anchor]]:
    by_id: dict[str, Anchor] = {}
    by_block: dict[int, Anchor] = {}
    counters = {'equation': 0, 'figure': 0, 'table': 0}
    bookmark_id = 100
    for section in spec.get('sections', []):
        for item in section.get('blocks', []):
            kind = item.get('type')
            if kind not in counters:
                continue
            counters[kind] += 1
            number = counters[kind]
            if 'number' in item and item['number'] != number:
                raise ValueError(
                    f'Non-sequential {kind} number {item["number"]!r}: expected {number}. '
                    'Use automatically generated numbering to keep SEQ and REF fields consistent.'
                )
            identifier = item.get('id', f'{kind}:{number}')
            if not isinstance(identifier, str) or not identifier.strip():
                raise ValueError(f'{kind} identifier must be a nonempty string')
            if identifier in by_id:
                raise ValueError(f'Duplicate manuscript object identifier: {identifier!r}')
            bookmark_id += 1
            ref = Anchor(identifier, kind, number, _bookmark_name(identifier, kind), bookmark_id)
            by_id[identifier] = ref
            by_block[id(item)] = ref
    return by_id, by_block


def validate_refs(spec, by_id):
    """Check references before writing the file so forward refs fail clearly."""
    def scan(thing, path):
        if isinstance(thing, dict):
            if 'ref' in thing:
                name = thing['ref']
                if name not in by_id:
                    raise ValueError(f'{path}: unknown reference {name!r}')
            for key, value in thing.items():
                scan(value, f'{path}.{key}')
        elif isinstance(thing, list):
            for i, value in enumerate(thing):
                scan(value, f'{path}[{i}]')
    scan(spec, 'manuscript')


def _run_text(p, value):
    return p.add_run(str(value))


def _field(p, instruction: str, cached: str):
    """Append complex field with cached result (renderable without an update)."""
    r = p.add_run()
    f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'begin')
    r._r.append(f)
    r = p.add_run()
    t = OxmlElement('w:instrText'); t.set(qn('xml:space'), 'preserve')
    t.text = ' ' + instruction + ' '
    r._r.append(t)
    r = p.add_run()
    f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'separate')
    r._r.append(f)
    r = p.add_run(cached)
    r = p.add_run()
    f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), 'end')
    r._r.append(f)


def add_bookmark_start(p, anchor: Anchor):
    element = OxmlElement('w:bookmarkStart')
    element.set(qn('w:id'), str(anchor.bookmark_id))
    element.set(qn('w:name'), anchor.bookmark)
    p._p.append(element)


def add_bookmark_end(p, anchor: Anchor):
    element = OxmlElement('w:bookmarkEnd')
    element.set(qn('w:id'), str(anchor.bookmark_id))
    p._p.append(element)


def add_caption_prefix(p, anchor: Anchor):
    if anchor.kind not in {'figure', 'table'}:
        raise ValueError('Caption prefix is only for figures and tables')
    text = 'Figure' if anchor.kind == 'figure' else 'Table'
    _run_text(p, text + ' ')
    add_bookmark_start(p, anchor)
    first = anchor.number == 1
    restart = ' \\r 1' if first else ''
    _field(p, f'SEQ {text}{restart} \\* ARABIC', str(anchor.number))
    add_bookmark_end(p, anchor)
    _run_text(p, '. ')


def add_reference(p, anchor: Anchor, prefix: str | None = None):
    text_prefix = prefix
    if text_prefix is None:
        text_prefix = {'equation': 'Eq. ', 'figure': 'Figure ', 'table': 'Table '}[anchor.kind]
    _run_text(p, text_prefix)
    cached = f'({anchor.number})' if anchor.kind == 'equation' else str(anchor.number)
    _field(p, f'REF {anchor.bookmark} \\h', cached)
