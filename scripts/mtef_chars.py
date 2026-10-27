"""Character inventory of MTEF equations: (typeface, MTCode) -> set of 8-bit font positions and counts.
    python mtef_chars.py file.docx"""
import sys, collections
import mtef_check as M

class R(M.Reader):
    def __init__(self, b, acc):
        super().__init__(b); self.acc = acc
    def record(self, tag, depth):
        if tag != 2:
            return super().record(tag, depth)
        opt = self.u8(); self.nudge(opt)
        tf = self.u8()
        mt = None if opt & M.CHAR_NOMT else self.u16()
        pos = None
        if opt & M.CHAR_ENC8:
            pos = self.u8()
        elif opt & M.CHAR_ENC16:
            pos = self.u16()
        self.acc[(tf, mt)][pos] += 1
        if opt & M.CHAR_EMBELL:
            self.obj_list(depth + 1, 'EMBELL list')

def inventory(docx):
    acc = collections.defaultdict(collections.Counter)
    for n, m, s in M.equations(docx):
        r = R(m, acc); r.header(); r.defs()
        while r.p < len(r.b):
            t = r.u8()
            if t == 0: break
            r.record(t, 0)
    return acc

if __name__ == '__main__':
    acc = inventory(sys.argv[1])
    for (tf, mt), c in sorted(acc.items(), key=lambda x: (x[0][0], x[0][1] or 0)):
        if tf in (0x86, 0x96, 0x8a, 0x8b, 0x84, 0x85) or (tf & 0x80 == 0):
            print(f'tf={tf:#04x} U+{mt or 0:04X} {chr(mt) if mt else "?"}  pos={ {(hex(k) if k is not None else None): v for k, v in c.items()} }')
