# Native MathType OLE on Linux: tested architecture and deployment notes

## Required outcome

A Linux server produces a `.docx` containing CFB2/OLE2 binary streams with `Equation Native` (MTEF) data, associated `Equation.DSMT4` objects, valid OOXML relationships and visible previews. It does **not** merely add Word OMML math or pictures. Windows Word with compatible MathType is the only reliable end-to-end editor/activation test.

## Required backend and version consistency

The current adapter targets the upstream `npofsi/docx-equation` repository commit `b70517cc22e85cb003d541dc592af2ee0e70ac7c`, version **0.3.3**. The available PyPI release is 0.3.0 and may not support this adapter. Install from the pinned Git commit via `requirements-generator.txt` in this skill, *not* by requesting a non-published `docx-equation==0.3.3` wheel.

`docx-equation` is AGPL-3.0-only at that commit. Its dependencies and the older PyPI release may have different licence metadata; check the specific artifact actually installed. An EquationLab (MIT) Linux-native adapter is a potential future path, but is not included here.

## Flow

```
LaTeX string or supplied MathML
    -> latex2mathml (when LaTeX is supplied)
    -> docx-equation MathML parser
    -> MTEF v5
    -> OLE Compound File (Equation Native)
    -> word/embeddings/*.bin + word/media/*preview*.png
    -> word/document.xml OLEObject ProgID="Equation.DSMT4"
    -> patch equation ID bookmarks and restore paragraph styles
    -> structural audit, page rendering and Windows user edit test
```

Use `MathTypeOptions(embed_mode='png-preview', mathtype_version='DSMT4')`. The selected mode fails if Chromium cannot render a preview rather than silently writing a generic placeholder.

## Renderer/font sizing

The upstream default `preview_pt_per_px=0.15` is likely small for the 12 pt body of the reference manuscript. The skill starts at **0.38 pt/px**, with `max_width_pt=365`, but this is an engineering starting point and must be calibrated on the user's Windows machine. Change with `--math-preview-scale`. Check full-size equations, fractions, matrix height, baseline alignment, and long-equation wrapping. Font availability between Linux and Windows can change appearance.

## Commands

```bash
python scripts/doctor.py --strict
python scripts/test_skill.py
python scripts/smoke_math.py -o math-regression.docx
python scripts/validate_docx.py math-regression.docx --expected 12 --require-mathtype
python scripts/build_paper.py examples/demo_paper.json -o paper.docx \
  --write-equation-manifest --math-preview-scale 0.38
python scripts/validate_docx.py paper.docx --expected 3 --require-mathtype
```

`--expected` counts **both inline and display MathType objects**, including any formulas embedded in captions or table cells.

## Evidence levels

- **Level 1, ZIP:** `.docx` is a valid package with intact parts and relationships.
- **Level 2, CFB:** each embedded OLE file has a valid CFB header and a nonempty `Equation Native` directory stream.
- **Level 3, display:** document renders properly on Linux and all formulas have legible preview images.
- **Level 4, Windows:** Word+MathType 7 can edit-save-reopen a broad formula corpus. Only this stage confirms actual end-user editability.

The structural validator uses a read-only Compound File directory walker and does not pretend to decode or semantically prove the MTEF stream contents. The Windows checklist remains mandatory.

## Unsupported constructs

Complex MathML tables, nested scripts, custom LaTeX macros, under/overscripts and multi-line alignments can expose limitations in `latex2mathml` or in MTEF serialization. Run `scripts/smoke_math.py`, preserve original equation source, check rendering, and fail clearly on any unsuccessful conversion rather than silently downgrading to image-only formulae.
