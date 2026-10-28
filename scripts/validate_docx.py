#!/usr/bin/env python3
"""Fail-closed DOCX/MathType structural audit; does not certify Windows editing."""
from __future__ import annotations
import argparse
import json
import posixpath
import re
import struct
import sys
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from docx import Document
from ole_native import equation_native_summary

CFB = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
O = 'urn:schemas-microsoft-com:office:office'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W, 'o': O, 'r': R, 'm': M, 'pr': REL}


def cfb_stream_directory(blob: bytes) -> dict[str, int]:
    """Read the CFB directory via FAT, verifying its actual stream entries.

    This is a small *read-only* parser, not a replacement for olefile on Windows.
    It handles normal 512/4096-byte sectors and the common <=109 FAT sectors.
    Files with an extended DIFAT fail closed instead of being claimed valid.
    """
    if len(blob) < 512 or blob[:8] != CFB:
        raise ValueError('Invalid CFB signature/size')
    sec_shift = struct.unpack_from('<H', blob, 30)[0]
    if sec_shift not in (9, 12):
        raise ValueError('Unexpected CFB sector size')
    size = 1 << sec_shift
    n_fat = struct.unpack_from('<I', blob, 44)[0]
    dir_start = struct.unpack_from('<I', blob, 48)[0]
    if n_fat > 109:
        raise ValueError('Extended DIFAT needs an additional CFB reader')
    fat_sectors = struct.unpack_from('<109I', blob, 76)[:n_fat]
    if any(x >= 0xfffffffa for x in fat_sectors):
        raise ValueError('Invalid FAT sector number')

    def sector(sid):
        offset = 512 + sid * size
        if offset + size > len(blob):
            raise ValueError('CFB sector extends past file')
        return blob[offset:offset + size]

    fat = []
    for sid in fat_sectors:
        fat.extend(struct.unpack('<' + 'I' * (size // 4), sector(sid)))
    if not fat:
        raise ValueError('No CFB FAT sectors')
    ids = []
    sid = dir_start
    while sid not in (0xfffffffe, 0xffffffff):
        if sid >= len(fat) or sid in ids or len(ids) > len(fat):
            raise ValueError('Invalid/cyclic CFB directory FAT chain')
        ids.append(sid)
        sid = fat[sid]
    data = b''.join(sector(s) for s in ids)
    streams = {}
    for i in range(0, len(data), 128):
        rec = data[i:i + 128]
        if len(rec) != 128:
            continue
        length = struct.unpack_from('<H', rec, 64)[0]
        kind = rec[66]
        if kind not in (1, 2, 5) or length < 2 or length > 64 or length % 2:
            continue
        name = rec[:length - 2].decode('utf-16le', errors='replace')
        stream_size = struct.unpack_from('<Q', rec, 120)[0]
        if kind == 2:
            streams[name] = stream_size
    return streams


def inspect(path, expected=0, require_mathtype=False, allow_legacy_equation3=False):
    path = Path(path)
    warnings = []
    if not path.is_file():
        raise FileNotFoundError(path)
    with ZipFile(path) as z:
        bad_crc = z.testzip()
        if bad_crc:
            warnings.append(f'ZIP CRC check failed at {bad_crc}')
        names = set(z.namelist())
        for required in ('word/document.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml'):
            if required not in names:
                raise ValueError(f'Missing mandatory DOCX part: {required}')
        root = etree.fromstring(z.read('word/document.xml'))
        rels = etree.fromstring(z.read('word/_rels/document.xml.rels'))
        embeds = sorted(n for n in names if n.startswith('word/embeddings/') and n.endswith('.bin'))
        ole_objects = root.xpath('.//o:OLEObject', namespaces=NS)
        rel_by_id = {r.get('Id'): r for r in rels}
        native = 0
        compound = 0
        mtef_versions = {}
        for name in embeds:
            data = z.read(name)
            if data.startswith(CFB):
                compound += 1
                try:
                    directory = cfb_stream_directory(data)
                    if directory.get('Equation Native', 0) > 0:
                        native += 1
                        try:
                            mtef_versions[name] = equation_native_summary(data)['mtef_version']
                            if mtef_versions[name] not in (3, 5):
                                warnings.append(f'Unsupported MTEF version: {name}')
                        except ValueError as err:
                            warnings.append(f'Cannot read MTEF native stream {name}: {err}')
                    else:
                        warnings.append(f'Compound OLE has no nonempty Equation Native stream: {name}')
                except ValueError as exc:
                    warnings.append(f'Malformed CFB directory in {name}: {exc}')
            else:
                warnings.append(f'Invalid CFB header: {name}')
        prog_count = 0
        legacy_count = 0
        target_list = []
        for idx, node in enumerate(ole_objects, 1):
            prog = node.get('ProgID')
            # A few legacy Microsoft Equation Editor formulas are stored as
            # Equation.3 / MTEF v3. Only accept those with explicit opt-in and
            # a native v3 payload, never relabel them as MathType DSMT4.
            standard_prog = prog in {'Equation.DSMT4', 'Equation.DSMT6'}
            rid = node.get(f'{{{R}}}id')
            relation = rel_by_id.get(rid)
            if relation is None:
                warnings.append(f'OLE object {idx}: missing relationship {rid}')
                continue
            if not relation.get('Type', '').endswith('/oleObject') or relation.get('TargetMode') == 'External':
                warnings.append(f'OLE object {idx}: invalid relation type or external target')
                continue
            target = posixpath.normpath(posixpath.join('word', relation.get('Target', '')))
            target_list.append(target)
            if standard_prog and mtef_versions.get(target) in (3, 5):
                prog_count += 1
            elif prog == 'Equation.3' and allow_legacy_equation3 and mtef_versions.get(target) == 3:
                prog_count += 1
                legacy_count += 1
            else:
                warnings.append(f'OLE object {idx} unsupported ProgID/MTEF pair: {prog!r}, v{mtef_versions.get(target)}')
            if target not in names:
                warnings.append(f'OLE object {idx}: missing binary target {target}')
        if len(target_list) != len(set(target_list)):
            warnings.append('Multiple OLE references point to the same binary part')
        visible = ''.join(root.xpath('.//w:t/text()', namespaces=NS))
        if 'MERGEFORMAT' in visible or 'Equation Chapter' in visible:
            warnings.append('Legacy field code leaked into visible document text')
        if re.search(r'\[(?:FORMULA|MATH).*?PREVIEW ONLY\]', visible):
            warnings.append('NON-PRODUCTION formula preview marker detected')
        if re.search(r'\{\{DOCX_EQ_\d+\}\}', visible):
            warnings.append('Unreplaced equation placeholder visible')
        bookmarks = [x.get(f'{{{W}}}name') for x in root.xpath('.//w:bookmarkStart', namespaces=NS)]
        if len(bookmarks) != len(set(bookmarks)):
            warnings.append('Duplicate Word bookmark names')
        instructions = root.xpath('.//w:instrText/text()', namespaces=NS)
        referenced = []
        for instruction in instructions:
            match = re.search(r'\bREF\s+(MS_(?:EQ|FIG|TAB)_[0-9a-f]+|ZEqnNum\d+)', instruction)
            if match:
                referenced.append(match.group(1))
                if match.group(1) not in bookmarks:
                    warnings.append(f'Broken cross-reference: {match.group(1)}')
        settings = z.read('word/settings.xml') if 'word/settings.xml' in names else b''
        if referenced and not settings:
            warnings.append('Reference fields without Word settings.xml')
        document = Document(path)
        section = document.sections[0]
        report = {
            'docx': str(path),
            'paragraphs': len(document.paragraphs),
            'tables': len(document.tables),
            'inline_images': len(document.inline_shapes),
            'page_width_cm': round(section.page_width.cm, 3),
            'page_height_cm': round(section.page_height.cm, 3),
            'margin_left_cm': round(section.left_margin.cm, 3),
            'ole_bins': len(embeds),
            'valid_compound_ole_headers': compound,
            'ole_with_equation_native_stream_name': native,
            'ole_objects_in_document': len(ole_objects),
            'mathtype_progid_objects': prog_count,
            'legacy_equation3_objects': legacy_count,
            'mtef_v3_objects': sum(v==3 for v in mtef_versions.values()),
            'mtef_v5_objects': sum(v==5 for v in mtef_versions.values()),
            'word_ref_fields': len(referenced),
            'matched_ole_relationships': len(target_list),
            'omml_equations': len(root.xpath('.//m:oMath', namespaces=NS)),
            'expected_min_equations': int(expected),
            'warnings': warnings,
            'windows_mathtype_editability_verified': False,
        }
        if require_mathtype:
            if expected < 1:
                raise ValueError('--require-mathtype needs --expected > 0')
            for key in ('ole_bins', 'valid_compound_ole_headers',
                        'ole_with_equation_native_stream_name', 'ole_objects_in_document',
                        'mathtype_progid_objects', 'matched_ole_relationships'):
                if report[key] < expected:
                    warnings.append(f'{key} is below expected {expected}: {report[key]}')
        if not require_mathtype and len(ole_objects) == 0 and expected > 0:
            warnings.append('Expected equations but found no OLE objects')
        report['pass_structural_checks'] = not warnings
        return report


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('docx')
    ap.add_argument('--expected', type=int, default=0, help='Expected inline plus display equation count')
    ap.add_argument('--require-mathtype', action='store_true')
    ap.add_argument('--allow-legacy-equation3', action='store_true',
                    help='Count Equation.3/MTEF v3 objects as valid legacy editable OLE')
    ap.add_argument('--json', dest='json_output')
    args = ap.parse_args()
    result = inspect(args.docx, args.expected, args.require_mathtype, args.allow_legacy_equation3)
    output = json.dumps(result, indent=2, ensure_ascii=False)
    print(output)
    if args.json_output:
        Path(args.json_output).write_text(output + '\n', encoding='utf-8')
    return 0 if result['pass_structural_checks'] else 2


if __name__ == '__main__':
    sys.exit(main())
