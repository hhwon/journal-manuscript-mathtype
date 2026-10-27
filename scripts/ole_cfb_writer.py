#!/usr/bin/env python3
"""Self-owned CFB v3 stream rebuilder preserving donor storage metadata.

Reads existing directory tree, stream values, and CLSIDs, then rebuilds all
FAT/MiniFAT allocations for a replacement Equation Native stream. Does not
change other logical streams. No third-party OLE writer is used.

Only CFB v3 (512-byte sectors), <= 109 FAT sectors, flat storages are
supported; fail closed on any unsupported allocation.
"""
from __future__ import annotations
import math
import struct
from ole_native import CompoundFile, ENDOFCHAIN, FREESECT, FATSECT

S = 512
M = 64


def _padded(raw, length):
    return raw + b'\x00' * (length - len(raw))


def repack_stream(template: bytes, stream_name: str, replacement: bytes) -> bytes:
    # NOTE: keeps the template's directory tree as it is. Safe with a genuine MathType container; a container written by
    # docx-equation has an invalid tree, which Windows rejects -> run ole_fix.fix_dir_tree afterwards (cfb_treecheck.py).
    cf = CompoundFile(template)
    if cf.major != 3 or cf.sector_size != S or cf.cutoff != 4096:
        raise ValueError('Only 512-byte CFB v3 with cutoff 4096 is supported')
    if stream_name not in cf.entries:
        raise KeyError(stream_name)
    # We keep the original CFB directory record tree (including stream GUIDs,
    # children/sibling indices) and replace only stream allocation/size fields.
    raw_directory = bytearray(cf._regular_chain(struct.unpack_from('<I', template, 48)[0]))
    if not raw_directory or len(raw_directory) % 128:
        raise ValueError('Invalid CFB directory')
    entries = []
    names = set()
    for n in range(len(raw_directory)//128):
        pos = n*128
        entry = raw_directory[pos:pos+128]
        name_len, kind = struct.unpack_from('<HB', entry, 64)
        if kind not in (0, 2, 5):
            raise ValueError('Complex storage or unhandled directory entry')
        if kind == 0:
            continue
        name = entry[:name_len-2].decode('utf-16le')
        if name in names:
            raise ValueError('Repeated directory entry')
        names.add(name)
        if kind == 5:
            if name != 'Root Entry':
                raise ValueError('Unexpected root storage')
            entries.append((n,name,kind,None))
        else:
            value = replacement if name == stream_name else cf.read(name)
            entries.append((n,name,kind,value))
    # MiniFAT allocations are contiguous by stream for reproducibility.
    small = []
    large = []
    mini_alloc = {}
    mini_data = bytearray()
    mini_fat = []
    for n,name,kind,value in entries:
        if kind != 2:
            continue
        if len(value) < cf.cutoff:
            small.append((n,name,value))
            count = math.ceil(len(value)/M)
            start = len(mini_fat) if count else ENDOFCHAIN
            for i in range(count):
                mini_fat.append(start+i+1 if i+1<count else ENDOFCHAIN)
            mini_data.extend(_padded(value,count*M))
            mini_alloc[n]=(start,len(value))
        else:
            large.append((n,name,value))
    # All regular sectors follow deterministic order:
    # big streams, mini-stream, mini-FAT, directory, FAT sectors.
    sector_blobs = []
    fat_links = []
    allocations = {}
    def alloc_regular(data: bytes):
        if not data:
            return ENDOFCHAIN,0
        count=math.ceil(len(data)/S)
        first=len(sector_blobs)
        for i in range(count):
            sector_blobs.append(_padded(data[i*S:(i+1)*S],S))
            fat_links.append(first+i+1 if i+1<count else ENDOFCHAIN)
        return first,count
    for n,name,value in large:
        sid,_=alloc_regular(value)
        allocations[n]=(sid,len(value))
    root_sid,root_count=alloc_regular(bytes(mini_data))
    mini_bytes=struct.pack('<%dI'%len(mini_fat),*mini_fat) if mini_fat else b''
    mini_sid,mini_count=alloc_regular(mini_bytes)
    # Patch directory entries with new first sectors and lengths.
    for n,name,kind,value in entries:
        pos=n*128
        if kind == 5:
            first,size=root_sid,len(mini_data)
        elif n in mini_alloc:
            first,size=mini_alloc[n]
        else:
            first,size=allocations[n]
        struct.pack_into('<I',raw_directory,pos+116,first)
        struct.pack_into('<Q',raw_directory,pos+120,size)
    directory_sid,directory_count=alloc_regular(bytes(raw_directory))
    num_base=len(sector_blobs)
    fat_count=1
    while True:
        expected=math.ceil((num_base+fat_count)/(S//4))
        if fat_count==expected:break
        fat_count=expected
    if fat_count>109:
        raise ValueError('Large compound file would require extended DIFAT')
    fat_start=len(sector_blobs)
    fat_links.extend([FATSECT]*fat_count)
    all_fat=_padded(struct.pack('<%dI'%len(fat_links),*fat_links),fat_count*S)
    sector_blobs.extend([all_fat[i*S:(i+1)*S] for i in range(fat_count)])
    header=bytearray(template[:S])
    struct.pack_into('<I',header,40,0) # v3 directory sector count
    struct.pack_into('<I',header,44,fat_count)
    struct.pack_into('<I',header,48,directory_sid)
    struct.pack_into('<I',header,60,mini_sid)
    struct.pack_into('<I',header,64,mini_count)
    struct.pack_into('<I',header,68,ENDOFCHAIN)
    struct.pack_into('<I',header,72,0)
    difat=[fat_start+i for i in range(fat_count)]+[FREESECT]*(109-fat_count)
    struct.pack_into('<109I',header,76,*difat)
    output=bytes(header)+b''.join(sector_blobs)
    check=CompoundFile(output)
    for n,name,kind,value in entries:
        if kind==2 and check.read(name)!=value:
            raise AssertionError(f'CFB stream roundtrip mismatch: {name!r}')
    return output
