"""Make the MathType OLE objects of a generated DOCX acceptable to Windows Word + MathType.

- root storage CLSID = MathType {0002CE03-0000-0000-C000-000000000046} (without it Word shows nothing);
- the compound-file directory as a valid red-black tree sorted by (name length, upper-case name) (MS-CFB 2.6.4).
  docx-equation 0.3.3 writes the streams as an unsorted all-red chain; Windows searches the tree, does not find
  '\\x01Ole' or 'Equation Native', and MathType cannot open the object, although every byte of the streams is right;
- one v:shapetype _x0000_t75 per document (as Word writes it);
- no w:updateFields (Word would redraw every object on opening).

    python ole_fix.py manuscript.docx [out.docx]
"""
import re
import struct
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

MATHTYPE_CLSID = bytes.fromhex('03ce020000000000c000000000000046')
NONE = 0xFFFFFFFF


def set_root_clsid(blob):
    b = bytearray(blob)
    assert b[:8] == bytes.fromhex('d0cf11e0a1b11ae1'), 'not a compound file'
    ssz = 1 << struct.unpack_from('<H', b, 0x1E)[0]
    root = (struct.unpack_from('<I', b, 0x30)[0] + 1) * ssz
    assert b[root:root + 20].decode('utf-16-le').startswith('Root Entry'), 'root entry not found'
    b[root + 0x50:root + 0x60] = MATHTYPE_CLSID
    return bytes(b)


def cfb_key(name):
    """MS-CFB sibling order: shorter name first, then upper-case code units."""
    return (len(name), name.upper())


def fix_dir_tree(blob):
    """Relink every storage's children as a balanced binary search tree; nodes on the incomplete last level red (0),
    all others black (1): a valid red-black tree."""
    b = bytearray(blob)
    ssz = 1 << struct.unpack_from('<H', b, 0x1E)[0]
    sect = lambda s: (s + 1) * ssz
    fat = []
    for s in struct.unpack_from('<109I', b, 0x4C):
        if s >= 0xFFFFFFFA:
            break
        fat += struct.unpack_from(f'<{ssz // 4}I', b, sect(s))
    chain, s = [], struct.unpack_from('<I', b, 0x30)[0]
    while s < 0xFFFFFFFA:
        chain.append(s)
        s = fat[s]
    off = [sect(c) + k * 128 for c in chain for k in range(ssz // 128)]
    name = lambda i: b[off[i]:off[i] + max(struct.unpack_from('<H', b, off[i] + 0x40)[0] - 2, 0)].decode('utf-16-le')
    L, R, C = 0x44, 0x48, 0x4C
    get = lambda i, f: struct.unpack_from('<I', b, off[i] + f)[0]

    def put(i, f, v):
        struct.pack_into('<I', b, off[i] + f, v)

    def siblings(i):
        return [] if i == NONE else siblings(get(i, L)) + [i] + siblings(get(i, R))

    def relink(storage):
        ids = sorted(siblings(get(storage, C)), key=lambda i: cfb_key(name(i)))
        full_levels = (len(ids) + 1).bit_length() - 1

        def mk(lo, hi, depth):
            if lo >= hi:
                return NONE
            m = (lo + hi) // 2
            i = ids[m]
            put(i, L, mk(lo, m, depth + 1))
            put(i, R, mk(m + 1, hi, depth + 1))
            b[off[i] + 0x43] = 1 if depth < full_levels else 0
            return i
        put(storage, C, mk(0, len(ids), 0))
        for i in ids:
            if b[off[i] + 0x42] in (1, 5):
                relink(i)
    b[off[0] + 0x43] = 1
    relink(0)
    return bytes(b)


def fix_docx(path, out=None):
    path = Path(path)
    out = Path(out) if out else path
    with ZipFile(path) as z:
        infos = z.infolist()
        parts = {i.filename: z.read(i.filename) for i in infos}
    for n in parts:
        if n.startswith('word/embeddings/') and n.endswith('.bin') and parts[n][:8] == bytes.fromhex('d0cf11e0a1b11ae1'):
            parts[n] = fix_dir_tree(set_root_clsid(parts[n]))
    doc = parts['word/document.xml'].decode('utf-8')
    first = doc.find('<v:shapetype id="_x0000_t75"')
    if first >= 0:
        end = doc.find('</v:shapetype>', first) + len('</v:shapetype>')
        doc = doc[:end] + re.sub(r'<v:shapetype id="_x0000_t75".*?</v:shapetype>', '', doc[end:], flags=re.S)
    parts['word/document.xml'] = doc.encode('utf-8')
    if 'word/settings.xml' in parts:
        parts['word/settings.xml'] = re.sub(rb'<w:updateFields[^>]*/>', b'', parts['word/settings.xml'])
    tmp = out.with_suffix('.tmp')
    with ZipFile(tmp, 'w', ZIP_DEFLATED) as zo:
        for i in infos:
            zo.writestr(i.filename, parts[i.filename])
    tmp.replace(out)
    return out


if __name__ == '__main__':
    import sys
    print(fix_docx(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None))
