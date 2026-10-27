#!/usr/bin/env python3
"""Minimal **read-only** Compound File Binary (CFB) stream reader.

Purpose: inspect original MathType Equation Native streams without Word, MathType
or the olefile Python package. It is intentionally strict: malformed/unsupported
CFB containers raise errors instead of being edited blindly. A separate self-owned CFB writer now exists in ole_cfb_writer.py.
"""
from __future__ import annotations
import dataclasses
import struct

MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC


@dataclasses.dataclass(frozen=True)
class StreamEntry:
    name: str
    type: int
    start_sid: int
    size: int


class CompoundFile:
    def __init__(self, blob: bytes):
        if len(blob) < 512 or blob[:8] != MAGIC:
            raise ValueError('Not a CFB/OLE2 file')
        self.blob = blob
        self.major = struct.unpack_from('<H', blob, 26)[0]
        self.sector_size = 1 << struct.unpack_from('<H', blob, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from('<H', blob, 32)[0]
        if self.major not in (3, 4) or self.sector_size not in (512, 4096):
            raise ValueError('Unsupported CFB major version / sector size')
        if self.mini_sector_size != 64:
            raise ValueError('Unsupported mini-sector size')
        self.cutoff = struct.unpack_from('<I', blob, 56)[0]
        count = struct.unpack_from('<I', blob, 44)[0]
        difat_start, difat_count = struct.unpack_from('<II', blob, 68)
        if count < 1 or count > 109 or difat_count or difat_start != ENDOFCHAIN:
            raise ValueError('Extended DIFAT is not supported (fail closed)')
        fat_sectors = struct.unpack_from('<109I', blob, 76)[:count]
        if any(x >= DIFSECT for x in fat_sectors):
            raise ValueError('Invalid FAT sector index')
        self.fat = []
        for sid in fat_sectors:
            self.fat.extend(struct.unpack('<%dI' % (self.sector_size // 4), self.sector(sid)))
        directory_start = struct.unpack_from('<I', blob, 48)[0]
        directory = self._regular_chain(directory_start)
        self.entries = {}
        self.root = None
        for i in range(0, len(directory), 128):
            rec = directory[i:i+128]
            if len(rec) < 128:
                raise ValueError('Truncated directory entry')
            length, kind = struct.unpack_from('<HB', rec, 64)
            if kind not in (2, 5):
                continue
            if length < 2 or length > 64 or length % 2:
                raise ValueError('Invalid directory name length')
            name = rec[:length - 2].decode('utf-16le', errors='strict')
            start = struct.unpack_from('<I', rec, 116)[0]
            size = struct.unpack_from('<Q', rec, 120)[0]
            if self.major == 3:
                size &= 0xFFFFFFFF
            item = StreamEntry(name, kind, start, size)
            if kind == 5:
                self.root = item
            else:
                if name in self.entries:
                    raise ValueError('Duplicate stream names are not supported')
                self.entries[name] = item
        if not self.root:
            raise ValueError('Missing CFB Root Entry')
        self.mini_fat = []
        mini_fat_start, mini_fat_count = struct.unpack_from('<II', blob, 60)
        if mini_fat_count:
            raw = self._regular_chain(mini_fat_start)
            if len(raw) != mini_fat_count * self.sector_size:
                raise ValueError('MiniFAT length does not match CFB header')
            self.mini_fat = list(struct.unpack('<%dI' % (len(raw) // 4), raw))
        self.mini_data = self._regular_chain(self.root.start_sid)[:self.root.size] if self.root.size else b''

    def sector(self, sid: int) -> bytes:
        if sid >= DIFSECT:
            raise ValueError('Special SID encountered in regular sector read')
        offset = self.sector_size + sid * self.sector_size
        if offset + self.sector_size > len(self.blob):
            raise ValueError('Sector exceeds CFB payload')
        return self.blob[offset:offset+self.sector_size]

    def _chain(self, start: int, fat: list[int], read_sector, max_steps: int):
        if start in (ENDOFCHAIN, FREESECT):
            return b''
        out = bytearray()
        sid = start
        seen = set()
        while sid != ENDOFCHAIN:
            if sid in seen or sid >= len(fat) or len(seen) >= max_steps:
                raise ValueError('Invalid/cyclic CFB allocation chain')
            seen.add(sid)
            out.extend(read_sector(sid))
            sid = fat[sid]
        return bytes(out)

    def _regular_chain(self, start: int) -> bytes:
        return self._chain(start, self.fat, self.sector, len(self.fat) + 1)

    def _mini_chain(self, start: int) -> bytes:
        def mini_sector(sid):
            offset = sid * self.mini_sector_size
            if offset + self.mini_sector_size > len(self.mini_data):
                raise ValueError('Mini sector exceeds mini stream')
            return self.mini_data[offset:offset+self.mini_sector_size]
        return self._chain(start, self.mini_fat, mini_sector, len(self.mini_fat)+1)

    def read(self, name: str) -> bytes:
        entry = self.entries.get(name)
        if entry is None:
            raise KeyError(name)
        if not entry.size:
            return b''
        if entry.size < self.cutoff:
            raw = self._mini_chain(entry.start_sid)
        else:
            raw = self._regular_chain(entry.start_sid)
        if len(raw) < entry.size:
            raise ValueError('CFB stream allocation shorter than declared stream size')
        return raw[:entry.size]

    def summary(self) -> dict:
        return {name: item.size for name, item in self.entries.items()}


def equation_native_summary(blob: bytes) -> dict:
    cf = CompoundFile(blob)
    native = cf.read('Equation Native')
    # A WIRIS OLE MathType native stream starts with EQNOLEFILEHDR (28 bytes),
    # followed by MTEF. Check obvious invariants, but do not pretend to fully
    # validate every structure and MathType-version-specific record.
    if len(native) < 34:
        raise ValueError('Equation Native stream too short for OLE header + MTEF')
    head = struct.unpack_from('<7I', native, 0)
    header_size = head[0]
    # Various legacy containers have a 28-byte native header; fail closed when
    # trying to locate MTEF if format differs.
    if header_size != 28:
        return {'stream_length': len(native), 'header_length': header_size,
                'mtef_version': None, 'header_supported': False, 'streams': cf.summary()}
    mtef_start = header_size
    mtef_version = native[mtef_start]
    return {'stream_length': len(native), 'header_length': header_size,
            'mtef_version': mtef_version, 'header_supported': True,
            'streams': cf.summary(), 'mtef_prefix_hex': native[mtef_start:mtef_start+16].hex()}
