# v4 native MathType font A/B test

## What was wrong

The earlier tests only established that `EQN_PREFS` selects Times New Roman/Symbol semantic roles, and that the Linux fallback *preview* was rendered in Tinos (a metrically compatible but non-identical font). These facts did not establish that Windows MathType's actual glyph rendering was identical to original MathType-generated equations. The user confirmed that the display still looked wrong. The cause remains **unresolved** without screenshots from the MathType editor and from Word before/after re-editing.

## New experimental encoding

MTEF 5 permits both positive semantic style typeface IDs (Variable, Function, Number, Greek) and *negative* explicit font-style IDs. The explicit-font stream is encoded by `scripts/mtef_explicit.py`:

| Explicit FONT_STYLE_DEF index | Reference font definition | CHAR style bits | MTEF typeface byte |
|---:|---|---|---|
| 1 | #1 Times New Roman | italic | 127 (-1 + 128) |
| 2 | #1 Times New Roman | plain | 126 (-2 + 128) |
| 3 | #1 Times New Roman | bold | 125 (-3 + 128) |
| 4 | #2 Symbol | italic | 124 (-4 + 128) |
| 5 | #2 Symbol | plain | 123 (-5 + 128) |

Existing `FONT_DEF` and `EQN_PREFS` are preserved from a real donor. The explicit style references are *additional* MTEF records ahead of the math body. No Windows fonts are bundled or downloaded. This is a specification-level experiment, not a certified MathType rendering fix.

The test DOCX built by `scripts/build_font_comparison_test.py` from your own MathType document (not shipped) includes four A/B formula pairs:

- A1 semantic `E=kT` vs A2 explicit Times italic.
- A3 semantic Greek tensor vs A4 explicit Times/Symbol tensor.
- A5 semantic `sin`/`cos` vs A6 explicit Times/Symbol function.
- A7 semantic bold vector vs A8 explicit Times bold vector.
- C1: unchanged MathType OLE from the user's source document.

The Word preview of each pair is identical because the PNG renderer remains Tinos on Linux. **Therefore compare the equations INSIDE the MathType editor** rather than judging the Word preview before activation. After editing and closing MathType, compare its regenerated preview separately. If the explicit variant changes the glyph appearance, report which A/B pair and whether the difference is correct.

## Validation limits

- **Passed:** OLE2/CFB stream consistency; DOCX relationships; MTEF v5 structure; 5 direct font style definitions on even-numbered test formulas; 9 MathType MTEF OLE objects; original control unmodified; MathType field codes imported from source.
- **Not passed or not available:** Windows MathType UI font fidelity, cross-reference creation, reformatting of fields by MathType 7.
- **Known:** LibreOffice does not render `MTPlaceRef` numbering fields correctly and may show field codes; Word desktop is the target.

## User feedback needed

For each A/B pair, indicate which version visually matches a formula manually typed in MathType 7. Ideally provide a screenshot from inside MathType with one wrong formula selected and **Style > Define** open. Also report whether `MTPlaceRef` numbers are recognized by **MathType > Insert Reference**.

## Spec

- [WIRIS MTEF 5](https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later): `FONT_STYLE_DEF`, `FONT_DEF`, typeface signed values and CHAR records.
