# MathType 7 native MTEF font semantics (self-owned backend)

## Purpose

Distinguish **native MathType font styles inside Equation Native/MTEF** from
Word's external fallback PNG. A Times-like screenshot is NOT proof of an
editable Times New Roman formula.

## Example: settings read from a reference manuscript

- MTEF version: 5; original application key DSMT7, editor target Equation.DSMT4.
- Full formula font size: 12 pt.
- `Text`, `Function`, `Number`: Times New Roman, plain.
- `Variable`: Times New Roman, italic.
- `L.C. Greek`: Symbol, italic.
- `U.C. Greek`, `Symbol`: Symbol, plain.
- `Vector`: Times New Roman, bold and upright.
- `MT Extra`: used for specialized symbols, not a general Latin font.

This is the manuscript's actual style configuration, not a universal MathType
policy for every document. Existing papers may have different styles and sizes.

## Native implementation

MTEF v5 `CHAR` records use a signed *typeface* integer biased by +128:

| Mathematical role | MTEF typeface | Stored byte | Test IDs |
|---|---:|---:|---|
| Text upright | 1 | 0x81 | A7 |
| Function upright | 2 | 0x82 | A5, A6, A12 |
| Variable italic | 3 | 0x83 | A1, A3 |
| Lowercase Greek | 4 | 0x84 | A2 |
| Uppercase Greek | 5 | 0x85 | A8 |
| Symbol | 6 | 0x86 | A1-A13 |
| Vector bold upright | 7 | 0x87 | A7, A13 |
| Number upright | 8 | 0x88 | A4, A10 |

Variable letters are not forced to explicit font definitions; they reference
MathType's semantic style via typeface. `FONT_DEF` and `EQN_PREFS` define those
styles in the original style seed. Named functions use the MTEF `FUNC_START`
option where applicable. The self-owned parser checks CHAR options, MTCode,
typeface, Greek Symbol font position and style profile consistency.

## Supported LaTeX subset for semantic font style

The core adds `\sin`, `\cos`, `\tan`, `\exp`, `\log`, `\ln`, `\lim`,
`\mathrm{LATIN}`, `\text{LATIN}`, `\operatorname{LATIN}`,
`\mathit{LATIN}`, `\mathbf{LATIN}` with a strict grammar. These are not
unrestricted text or font renderers. Whitespace inside styled spans is rejected;
`\epsilon` and `\varphi` are rejected because mapping them to the
same glyph as another command would be a mathematical notation error.
Bold Greek, advanced accents, matrices, integrals and complex alignment
are not yet supported.

## Use

```bash
python scripts/test_font_semantics.py --source real-mathtype.docx
python scripts/build_font_role_test.py --source real-mathtype.docx -o font-test.docx
python scripts/validate_docx.py font-test.docx --expected 15 --require-mathtype
```

For Linux previews: real Times New Roman is used if already installed; otherwise
Tinos (or STIX) is an approximation. Set `MATHTYPE_PREVIEW_FONT` only to an
installed font, and never bundle proprietary font files. Preview typography
and native MTEF typography remain separate.

Windows editing is the required final gate. Compare A5/A7/A8/A12 with C1,
then edit each and save/reopen in MathType 7. Record *inside MathType* styles
independently from the Word PNG previews. A font profile audit may pass on Linux
even when Windows substitutes a missing font or displays a template differently.

## References

- WIRIS, MTEF v5: https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later
- WIRIS, MathType 7 tutorials: https://docs.wiris.com/en_US/using-mathtype-7/main-tutorials


## Experimental v4 direct-font override

See [font-diagnostic-v4.md](font-diagnostic-v4.md). The standalone
`scripts/mtef_explicit.py` selects explicit `FONT_STYLE_DEF` records and negative
CHAR typeface values rather than semantic MathType style identifiers. This is
not the production default until a Windows comparison proves that the original
MathType fonts display correctly and that editing/saving does not change them.
Do not assume that success at reading FONT_DEF proves correct visual glyphs.
