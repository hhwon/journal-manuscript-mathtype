"""Strict reader for MTEF 5 (MathType 'Equation Native' data): parses every record, checks nesting and that the whole
stream is consumed. Layouts from the `mathtype` gem (lib/records5). Used to compare our equations with genuine MathType
ones (a reference document with 467 genuine MathType objects) and to reject malformed MTEF before delivery.

    python mtef_check.py file.docx [--dump N]
"""
import io
import struct
import sys
import zipfile

NUDGE, CHAR_EMBELL, CHAR_ENC8, CHAR_ENC16, CHAR_NOMT = 0x08, 0x01, 0x04, 0x10, 0x20
LINE_NULL, LINE_LSPACE, LP_RULER = 0x01, 0x04, 0x02


class Bad(Exception):
    pass


class Reader:
    def __init__(self, b):
        self.b, self.p, self.out = b, 0, []

    def u8(self):
        if self.p >= len(self.b):
            raise Bad(f'truncated at {self.p}')
        v = self.b[self.p]
        self.p += 1
        return v

    def u16(self):
        v = struct.unpack_from('<H', self.b, self.p)[0]
        self.p += 2
        return v

    def uint(self):
        v = self.u8()
        return self.u16() if v == 255 else v

    def zstr(self):
        e = self.b.index(0, self.p)
        s = self.b[self.p:e]
        self.p = e + 1
        return s.decode('latin1')

    def nudge(self, opt):
        if opt & NUDGE:
            dx, dy = self.u8(), self.u8()
            if dx == 128 and dy == 128:
                self.p += 4

    def ruler(self):
        tag = self.u8()
        if tag != 7:
            raise Bad(f'expected RULER at {self.p - 1}, got {tag}')
        n = self.u8()
        self.p += 3 * n

    def dims(self):
        n = self.u8()
        nib = 0
        for _ in range(n):
            while True:
                byte = self.b[self.p + nib // 2]
                v = byte >> 4 if nib % 2 == 0 else byte & 15
                nib += 1
                if v == 15 and nib > 1:
                    break
        self.p += (nib + 1) // 2

    def header(self):
        ver = self.u8()
        if ver != 5:
            raise Bad(f'MTEF version {ver}')
        self.p += 4
        app = self.zstr()
        self.u8()                                   # equation options
        return app

    def defs(self):
        """definition records up to the first equation-body record"""
        while True:
            tag = self.b[self.p]
            if tag == 17:                           # FONT_DEF
                self.p += 1; self.uint(); self.zstr()
            elif tag == 19:                         # ENCODING_DEF
                self.p += 1; self.zstr()
            elif tag == 18:                         # EQN_PREFS
                self.p += 1; self.u8(); self.dims(); self.dims()
                for _ in range(self.u8()):
                    if self.uint():
                        self.u8()
            elif tag == 8:                          # FONT_STYLE_DEF
                self.p += 1; self.uint(); self.u8()
            elif tag == 16:                         # COLOR_DEF
                self.p += 1
                opt = self.u8()
                self.p += 8 if opt & 1 else 6
                if opt & 4:
                    self.zstr()
            elif tag >= 100:
                self.p += 1; self.p += self.u16()
            else:
                return

    def obj_list(self, depth, what):
        n = 0
        while True:
            tag = self.u8()
            if tag == 0:
                return n
            self.record(tag, depth)
            n += 1
            if self.p > len(self.b):
                raise Bad(f'{what}: ran past end')

    def record(self, tag, depth):
        ind = '  ' * depth
        if tag == 1:                                # LINE
            opt = self.u8(); self.nudge(opt)
            if opt & LINE_LSPACE:
                self.u16()
            if opt & LP_RULER:
                self.ruler()
            if opt & LINE_NULL:
                self.out.append(f'{ind}LINE(null)')
                return
            self.out.append(f'{ind}LINE')
            self.obj_list(depth + 1, 'LINE')
        elif tag == 2:                              # CHAR
            opt = self.u8(); self.nudge(opt)
            tf = self.u8()
            mt = None if opt & CHAR_NOMT else self.u16()
            if opt & CHAR_ENC8:
                self.u8()
            elif opt & CHAR_ENC16:
                self.u16()
            self.out.append(f'{ind}CHAR {chr(mt) if mt else "?"} tf={tf:#x} opt={opt:#x}')
            if opt & CHAR_EMBELL:
                self.obj_list(depth + 1, 'EMBELL list')
        elif tag == 3:                              # TMPL
            opt = self.u8(); self.nudge(opt)
            sel = self.u8()
            v = self.u8()
            if v & 0x80:
                v = (v & 0x7F) | (self.u8() << 8)
            to = self.u8()
            self.out.append(f'{ind}TMPL sel={sel} var={v:#x} opt={to}')
            self.obj_list(depth + 1, 'TMPL')
        elif tag == 4:                              # PILE
            opt = self.u8(); self.nudge(opt)
            ha, va = self.u8(), self.u8()
            if opt & LP_RULER:
                self.ruler()
            self.out.append(f'{ind}PILE h={ha} v={va}')
            self.obj_list(depth + 1, 'PILE')
        elif tag == 5:                              # MATRIX
            opt = self.u8(); self.nudge(opt)
            va, hj, vj, rows, cols = self.u8(), self.u8(), self.u8(), self.u8(), self.u8()
            self.p += (rows + 4) // 4 + (cols + 4) // 4
            self.out.append(f'{ind}MATRIX {rows}x{cols} valign={va} hjust={hj} vjust={vj}')
            start = len(self.out)
            self.obj_list(depth + 1, 'MATRIX')
            n = sum(1 for s in self.out[start:] if s.startswith(ind + '  LINE'))
            if n != rows * cols:
                raise Bad(f'MATRIX {rows}x{cols} has {n} cells')
        elif tag == 6:                              # EMBELL
            opt = self.u8(); self.nudge(opt)
            self.out.append(f'{ind}EMBELL {self.u8()}')
        elif tag == 7:
            self.p -= 1; self.ruler()
        elif tag == 8:
            self.uint(); self.u8()
        elif tag == 9:                              # SIZE
            s = self.u8()
            if s == 101:
                self.u16()
            elif s == 100:
                self.u8(); self.u16()
            else:
                self.u8()
        elif 10 <= tag <= 14:                       # FULL, SUB, SUB2, SYM, SUBSYM
            self.out.append(f'{ind}SIZE{tag}')
        elif tag == 15:                             # COLOR
            self.uint()
        elif tag == 16:                             # COLOR_DEF
            opt = self.u8()
            self.p += 8 if opt & 1 else 6
            if opt & 4:
                self.zstr()
        elif tag >= 100:
            self.p += self.u16()
        else:
            raise Bad(f'unknown record {tag} at {self.p - 1}')


def check(mtef):
    r = Reader(mtef)
    app = r.header()
    r.defs()
    while r.p < len(r.b):
        tag = r.u8()
        if tag == 0:
            break
        r.record(tag, 0)
    rest = len(r.b) - r.p
    if rest > 0 and any(r.b[r.p:]):
        raise Bad(f'{rest} trailing bytes after the equation')
    return app, r.out


def equations(docx):
    import olefile
    z = zipfile.ZipFile(docx)
    for n in sorted(x for x in z.namelist() if x.startswith('word/embeddings/') and x.endswith('.bin')):
        o = olefile.OleFileIO(io.BytesIO(z.read(n)))
        if not o.exists('Equation Native'):
            continue
        en = o.openstream('Equation Native').read()
        hdr = struct.unpack_from('<H', en, 0)[0]
        size = struct.unpack_from('<I', en, 8)[0]
        yield n, en[hdr:hdr + size], (len(en) - hdr, size)


def main():
    docx = sys.argv[1]
    dump = int(sys.argv[sys.argv.index('--dump') + 1]) if '--dump' in sys.argv else -1
    ok = bad = 0
    for k, (n, m, sizes) in enumerate(equations(docx)):
        try:
            app, out = check(m)
            ok += 1
            if k == dump:
                print(n, app, sizes); print('\n'.join(out))
        except Exception as e:
            bad += 1
            if bad <= 8:
                print('BAD', n, sizes, type(e).__name__, e)
            if k == dump:
                print(n, sizes)
    print(f'{ok} parsed, {bad} malformed')


if __name__ == '__main__':
    main()
