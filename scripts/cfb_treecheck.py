"""Check MS-CFB directory trees: BST order (length, upper-case) and red-black rules (1 = black, 0 = red)."""
import sys, io, zipfile, olefile
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from ole_fix import cfb_key as _cfb_key
NONE = 0xFFFFFFFF
def check(blob):
    o = olefile.OleFileIO(io.BytesIO(blob)); D = o.direntries
    errs = []
    def walk(i, lo, hi, parent_red):
        if i == NONE: return 1
        e = D[i]; k = _cfb_key(e.name)
        if lo is not None and not k > lo: errs.append(f'order {e.name!r}')
        if hi is not None and not k < hi: errs.append(f'order {e.name!r}')
        red = e.color == 0
        if red and parent_red: errs.append(f'red-red {e.name!r}')
        a = walk(e.sid_left, lo, k, red); b = walk(e.sid_right, k, hi, red)
        if a != b: errs.append(f'black height {e.name!r}')
        if e.entry_type in (1, 5): walk(e.sid_child, None, None, False)
        return a + (0 if red else 1)
    walk(D[0].sid_child, None, None, False)
    return errs
if __name__ == '__main__':
    z = zipfile.ZipFile(sys.argv[1]); bad = n = 0
    for name in z.namelist():
        if name.startswith('word/embeddings/') and name.endswith('.bin'):
            n += 1; e = check(z.read(name))
            if e:
                bad += 1
                if bad <= 3: print(name, e)
    print(f'{n} objects, {bad} with an invalid directory tree')
