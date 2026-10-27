# MathType objects in Word: verified facts

What a generated MathType object must look like for Windows Word + MathType to accept, edit and redraw it like a
genuine one, and what MathType's own layout looks like. Collected in October 2026 while bringing docx-equation's output
to genuine quality. Each fact names its evidence:

- **[genuine]** a reference Word document with 467 MathType 6/7 equations written by MathType itself (not
  distributed): MTEF read with `scripts/mtef_check.py` / `scripts/mtef_chars.py`, WMF previews parsed glyph by glyph;
- **[spec]** WIRIS, [MTEF v.5 (MathType 4.0 and later)](https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later),
  *How MTEF is stored in files and objects*, and the MTCode encoding tables (fences EC00–ECFF, non-marking EF00–EFFF);
- **[EqLab]** [EquationLab](https://github.com/Ne1ther/EquationLab) (MIT; `src/native.mjs`, `src/native-layout.mjs`,
  `src/new-formula.mjs`), which reconstructs MathType 7.8's layout rules — read, not run, not bundled;
- **[Windows]** tests in Word + MathType on Windows during development.

The modules that implement these facts: `mtef_profile.py` (MTEF), `preview_math.py` (preview), `equation_layout.py`
(placement), `ole_fix.py` (container); `mathtype_gate.py` checks the result.

## 1. OLE container

| fact | evidence |
|---|---|
| Root storage CLSID must be MathType's {0002CE03-0000-0000-C000-000000000046}; without it Word shows no equations. | Windows |
| The compound-file directory must be a valid red-black tree sorted by (name length, upper-case name) (MS-CFB 2.6.4). docx-equation 0.3.3 writes the streams as an unsorted all-red chain; Windows searches the tree, cannot find `\x01Ole` / `Equation Native`, and **no equation can be edited** although every stream byte is right. `ole_fix.fix_dir_tree` relinks the tree; `cfb_treecheck.py` checks it. `ole_cfb_writer.repack_stream` keeps the source's tree, so it is only safe on a genuine container. | Windows (bisection, see windows-acceptance.md: genuine container + own stream edits, own container + byte-identical genuine stream does not; after the repair all edit) |
| `Equation Native` = 28-byte EQNOLEFILEHDR (cbHdr 28, version 0x00020000, clipboard format, cbObject = MTEF length, 4 reserved DWORDs) + MTEF. | spec, genuine |
| No `w:updateFields` in settings: Word then redraws every object on opening, and objects with faulty MTEF go blank. | Windows |

## 2. MTEF records

| fact | evidence |
|---|---|
| An empty slot is a null LINE (`01 01`) **without** END record; an extra END closes the enclosing template and shifts everything after it (e.g. the superscript of C^T). | genuine, `mathtype` gem records5 |
| Scripts: TMPL 27/28/29, SUB size, slots, END, FULL size. | genuine |
| Big operator: TMPL, LINE body, SUB size, lower, upper, SYM size (0x0d), CHAR operator (Symbol style), END; variation tvBO_LOWER 0x10, tvBO_UPPER 0x20, tvBO_SUM 0x40, integrals tvINT_1..3. | genuine, spec, EqLab |
| Fence: TMPL selector, variation (left 0x01, right 0x02), LINE body, FULL, fence characters with typeface 22 (fnEXPAND, no font position), END. Selectors 0 angle, 1 parenthesis, 2 brace, 3 bracket, 4 bar, 5 double bar, 6 floor, 7 ceiling, 8 white bracket. | genuine, spec |
| Fence MTCodes: angle 2329/232A, bar EC07/EC08, double bar EC09/EC0A, floor F8F0/F8FB, ceiling F8EE/F8F9, white bracket 301A/301B. | spec, genuine, EqLab |
| Accent on one character: CHAR with mtefOPT_CHAR_EMBELL + EMBELL record (bar 17, tilde 8, hat 9, right arrow 11). | genuine, spec |
| CHAR typeface = 128 + style: 1 Text, 2 Function, 3 Variable, 4 LC Greek, 5 UC Greek, 6 Symbol, 7 Vector, 8 Number, 9 User 1, 10 User 2, 11 MT Extra, 22 fnEXPAND, 24 fnSPACE. | spec |
| Characters in a Symbol-font style carry their 8-bit **Symbol font position** (option 0x04): = 3D, + 2B, − 2D (MTCode 2212), ∗ 2A (2217), ∂ B6, ∑ E5, ∫ F2, ∏ D5, → AE, ∈ CE, ⊂ CC, ∇ D1, ≈ BB, ≠ B9, ≤ A3, ≥ B3, × B4, ∞ A5, 〈 E1, 〉 F1 (Adobe Symbol encoding); Greek in Symbol: α 61 … ω 77, ϕ (U+03D5) 66, φ (U+03C6) 6A. Times New Roman characters carry none. | genuine (every Symbol-style character), spec |
| Solidus, parentheses, comma, colon, period are Function style (Times). | genuine |
| Spaces are fnSPACE characters: EF01 zero, EF02 thin (1/6 em), EF03 medium, EF04 thick (1/3 em), EF05 em, EF06 2-em, EF08 1-point. | spec, genuine |
| MATRIX `h_just`: **0 left, 1 centre, 2 right**. The spec's alignment list says 1 = left, but genuine matrices carry 1 and are centred (a 3×1 column whose wider cell is shifted by half its extra width), and EqLab reads MathType 7.8 the same way. | genuine, EqLab; left alignment in Windows not yet confirmed |

## 3. Fonts, styles, sizes, spacing (EQN_PREFS)

| fact | evidence |
|---|---|
| Factory profile ('Times+Symbol 12.eqp'): fonts Times New Roman, Symbol, Courier New, MT Extra; Text and Function Times upright, Variable Times italic, LC Greek Symbol italic, UC Greek Symbol upright, **Symbol style in the Symbol font**, Vector Times bold, Number Times, User 1 Courier New, User 2 Times, MT Extra in MT Extra. | EqLab, genuine (all 467 objects map the Symbol style to Symbol) |
| MathType draws tall parentheses, brackets and braces from Symbol pieces (E6–E8/F6–F8, E9–EB/F9–FB, EC–EF/FC–FE) or scales the Symbol glyph, and integrals as Symbol F2 at the symbol size. Times New Roman has no such pieces, so a Symbol *style* mapped to Times breaks tall brackets and integrals when MathType redraws the equation. | genuine WMF, font check, Windows |
| Greek may be put in Times New Roman (an option, `--greek times`): Times has the Greek block; the Symbol style must stay Symbol. | Windows (all editable with Times Greek) |
| Sizes: 12 pt, 58 %, 42 %, 150 %, 100 % (+ 75 %, 150 %, 1 pt). | EqLab |
| Text (FE), the twelfth style, holds East Asian text (e.g. 宋体/SimSun). MathType then adds a second ENCODING_DEF `WinAllCodePages` (index 6) and writes the font name in the ANSI code page (SimSun = GBK `cb ce cc e5`); the style array's twelfth entry points to that font. `.eqp` key `TextFE`. | genuine |
| Spacing order: 0 line, 1 matrix row, 2 matrix column, 3 superscript height, 4 subscript depth, 5 sub/superscript gap, 6 limit height, 7 limit depth, 8 limit line spacing, 9 numerator height, 10 denominator depth, 11 bar overhang, 12 bar thickness, 13 sub-fraction bar, 14 slash gap, 15 fence overhang, 16 operator spacing, 17 non-operator spacing, 18 character width, 19 minimum gap, 20–22 radical, 23 embellishment gap, 24 prime height, 25–28 strokes, 29 horizontal fence gap. | spec |
| Factory spacing values (%): 150, 150, 100, 45, 25, 8, 25, 100, 100, 35, 100, 8, 5, 2.5, 8, 8, 100, 100, 0, 8, 17, 8, 100, 12.5, 45, 5, 5, 5, 5, 10 — identical in all 467 genuine objects. docx-equation writes 100/100/200 for the first three, which changes matrices when MathType redraws. | genuine, EqLab |
| A MathType preference file (.eqp, "MathType INI v1": `[Styles]` key=Font[,B|,I], `[Sizes]` and `[Spacing]` key=value pt or %, keys in the order above) holds these settings. `mtef_profile.read_eqp` reads it (`--profile file.eqp`); the factory file in `examples/` reproduces the built-in profile byte for byte. | EqLab (`src/eqp.mjs`), spec |

## 4. Layout: what MathType draws

| fact | evidence |
|---|---|
| Fraction: numerator baseline = axis + t/2 + max(min gap, numerator height − numerator depth) + numerator depth; denominator baseline = axis − t/2 − max(min gap, denominator depth − denominator height) − denominator height; bar on the axis (0.25 em), 5 % thick, 8 % overhang. At 12 pt: numerator 7.56 pt up, denominator 9.38 pt down. Fractions keep full size inside tables and other fractions. | EqLab, genuine WMF |
| Superscript baseline 0.45 em up, subscript 0.25 em down, scripts 58 % (second level 42 %). On a tall base (fence, matrix) the superscript follows the top edge (baseline ≈ 0.35 script-em below the top) and a subscript the bottom edge. | EqLab, genuine WMF |
| Spacing beyond the glyph advance at full size: relations and the solidus 0.21–0.23 em each side, binary + and − 0.15–0.20 em, after a comma 0.14 em, none before it, none after an opening parenthesis; inside scripts none. | genuine WMF (glyph positions from ExtTextOut dx arrays) |
| Integral (150 %): Symbol sign centred on the axis (0.357 em below the baseline); lower limit baseline at the sign's bottom (0.517 em below), upper limit top at its top, both after the sign's advance + 8 %; integrand after the wider of them + 0.1 em. | EqLab, Windows rendering |
| Bar accent: 0.5 em long whatever the letter, 0.04–0.05 em thick, 0.12–0.15 em above it, starting 0.13 em (italic) / 0.09 em (upright) after the letter's origin. | genuine WMF, Windows rendering |
| About 0.1 em inside an opening tall bracket or brace; tall brackets are built from fixed-height pieces and come out taller than the contents (not reproduced by the Chromium preview). | Windows rendering |
| Matrix rows at least font ascent + descent; pitch from matrix row spacing; column gap from matrix column spacing. | EqLab |
| Word places an object with `w:position` = −(depth of the preview below the formula's baseline) in half-points; genuine shapes have no VML `vertical-align`. | genuine (every object) |
| Preview box ≈ ink + about 2 pt margin, whole points. | genuine WMF |
| LibreOffice ignores `w:position` on OLE objects (genuine documents also show raised inline formulas there) and draws only the preview, never the MTEF: never judge baselines or MTEF from a LibreOffice render. | genuine |

## 5. Preview implementation (`scripts/preview_math.py`)

Chromium through Playwright (pages loaded from `file://`, otherwise file fonts do not load). The maths font is a renamed
copy of TeX Gyre Termes Math whose MATH constants carry section 4 (ScriptPercentScaleDown 58/42, fraction shifts
625/775, gaps 80, script shifts 450/250, SuperscriptBaselineDropMax 210 and SubscriptBaselineDropMin 100 for tall
bases, limits 250/1000, OverbarVerticalGap 110). A MathML transform sets explicit operator spacing, rebuilds integrals
with `mpadded`, draws bars as `mspace` rules, and marks Symbol-style characters, which are drawn with a Unicode-mapped
CFF copy of URW Standard Symbols PS (metric clone of Adobe Symbol). Crop = ink + 2 pt; the depth below the baseline is
recorded for `equation_layout.py`. Reference formulas agree with genuine MathType boxes to about 1.5 pt; side-by-side
with MathType's own redraw only the taller piece-built brackets differ. The preview is replaced by MathType's own as
soon as an equation is edited.
