# Self-owned DOCX/OLE editor: architectural contract and evidence

## Original object-level editing (v0.1)

Scripts: `scripts/ole_native.py`, `scripts/ole_editor.py`, `scripts/test_ole_editor.py`. Written specifically for this Skill without importing the source of EquationLab or docx-equation. Depends only on `lxml` plus Python standard library. Word and MathType are not needed on Linux for object-level editing.

### Three data layers

1. **DOCX OPC/OOXML:** `word/document.xml` contains `w:object` with `v:shape`, `v:imagedata` and `o:OLEObject`; document relationships resolve preview media and OLE binaries. The selected object's existing `ShapeID`, `ObjectID`, position and paragraph are kept. Clone relationship/part when shared by multiple objects.
2. **CFB/OLE2:** `word/embeddings/oleObject*.bin` is a Compound File Binary container. The read-only stream parser supports FAT, MiniFAT and standard 512/4096-byte sectors; it extracts `Equation Native`. Extended DIFAT is intentionally unsupported, and such data fails closed.
3. **MathType MTEF:** The native stream contains a 28-byte `EQNOLEFILEHDR`, followed by an MTEF stream (v3 or v5 in the supplied manuscripts). The current module **identifies and extracts** MTEF, but does **not** decode every record or convert arbitrary LaTeX into new MTEF.

Source references for the design: WIRIS MathType MTEF v5 specification (`docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later`) and Microsoft Open XML `o:OLEObject` schema (`learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.office.oleobject`).

### Commands

```bash
python -m pip install -r requirements.txt
python scripts/ole_editor.py list original.docx --out equations.json
python scripts/ole_editor.py extract original.docx --index 12 --out-dir extracted-12
python scripts/ole_editor.py replace original.docx --index 12 \
    --donor math-donor.docx --donor-index 4 -o revised.docx
python scripts/validate_docx.py revised.docx --expected 249 \
    --require-mathtype --allow-legacy-equation3
```

The `--index` refers to the **OLE object's order in document.xml**, not its printed equation number; an individual inline symbol may be an OLE object. Locate the correct index from the inventory's `paragraph`, `context`, and `preview_part` fields. Before writing, verify the intended target visually.

`--fit donor` (default) copies donor preview dimensions and Word baseline offset while retaining target paragraph style/numbering/ID. `--fit original` preserves target geometry (may clip a larger replacement). The replacement updates both the genuine OLE object and its preview, or refuses the input when unsafe.

### Supported and unsupported

| Operation | Status |
|---|---|
| Scan hundreds of MTEF v3/v5 OLE objects in a DOCX | Implemented and verified |
| Export original OLE2, `Equation Native` MTEF, preview, JSON provenance | Implemented |
| Replace any single known formula with a donor *that is already a real MathType OLE* | Implemented and verified structurally |
| Preserve every unrelated OLE/preview byte, cloning shared references | Implemented and verified structurally |
| Preserve original document/other files without in-place overwrite | Implemented |
| Arbitrary LaTeX -> MTEF conversion without external dependencies | Not yet implemented |
| Semantic changes inside an existing MTEF tree, e.g. rename a tensor index | Not yet implemented |
| Repair all Word+MathType UI/browser compatibility automatically | Not yet implemented |
| Guaranteed Windows Word/MathType 7 activation | Not yet demonstrated |

Important: `Equation.3` and MTEF v3 are legacy Equation Editor objects. They are **not** automatically equivalent to modern MathType OLE and might activate a different handler on Windows. The module refuses cross-ProgID replacement. The compatibility policy is **preserve, not silently upgrade**.

## Tested against the supplied private research manuscripts

Without bundling either manuscript into the Skill:

- Private sample manuscript A (not shipped): 249 equation objects, 240 MTEF v5 + 9 legacy Equation.3/MTEF v3.
- Private sample manuscript B (not shipped): 219 equation objects, all MTEF v5.
- Ran extraction and same-document, cross-document and legacy-object donor replacement on copies.
- Verified that **248 unedited formula OLE parts and previews remained identical** after replacing one formula in the 249-formula document.
- Exercised shared OLE relationship and shared preview relationship cloning, plus WMF -> PNG preview-part re-binding.
- Verified incompatible ProgID or in-place overwrite fails without producing an output.
- No Windows Word+MathType edit/save/reopen test has been run. Structural validation alone cannot certify successful activation.

## Experimental self-owned MTEF-v5 writer and formula replacement (v0.2)

New scripts: `ole_cfb_writer.py`, `mtef_subset.py`, `native_preview.py`, `build_native_ole_test.py`, `test_native_writer.py`. New command: `ole_editor.py replace-latex`.

- Build a constrained mathematical AST from a documented LaTeX subset, and emit native MTEF records (CHAR, LINE, TMPL with simple SUB/SUP, fraction, radical, plus SIZE and END). The output is *not* OMML. Unknown commands, inconsistent group syntax and unsupported operators cause errors.
- Use a real `Equation.DSMT4` MTEF-v5 formula as a **style-preamble/CFB metadata seed**. The source formula content is **not copied** into new formulas; only its valid MTEF font/size/spacing declarations and the non-native OLE streams (`CompObj`, `Ole`, `ObjInfo`) are retained. This is **not yet seed-free**.
- Use the original CFB's directory metadata (CLSID and sibling structure) while **rebuilding** all stream allocations, the MiniFAT and FAT. Tested v3 CFB with native payload sizes from 50 to 33,000 bytes including below and above the 4096-byte MiniFAT threshold.
- Generate PNG preview separately using Matplotlib mathtext. Its display correctness is **not evidence** that MathType interprets generated MTEF correctly; Word activation/re-render remains a required external check.
- `replace-latex` rewrites one selected MTEF v5 formula in an existing `.docx`, clones any shared relationships and updates only relevant package parts, and audits all unedited formulas/previews for byte-identical preservation. A real 249-formula manuscript yielded 248 unchanged objects after one rewrite, including a WMF -> PNG preview conversion.
- `build_native_ole_test.py` creates a fresh 2-page `.docx` with 7 synthesised MTEF OLE objects, one direct sigma-to-tau native MTEF CHAR patch and one untouched real OLE control. It includes expected LaTeX and a Windows feedback matrix.
- **Unverified:** actual Windows MathType edit, save, close and reopen for the newly synthesized objects. **Unsupported:** generic integrals, matrix layouts, alignments, arbitrary higher-grade macros, Equation.3/MTEF v3 generation, any MTEF5 font style outside the explicit subset.

Run:

```bash
pip install -r requirements-native-test.txt
python scripts/build_native_ole_test.py --source old.docx -o test-native.docx
python scripts/ole_editor.py replace-latex old.docx --index 14 \
    --latex '\sigma_{ij}=C_{ijkl}\varepsilon_{kl}' -o changed.docx
python scripts/test_native_writer.py --source old.docx
```

## Next-phase engineering (independent, no false promises)

1. Write a lossless MTEF-v5 **parser to a typed AST**, preserving unknown records byte-for-byte (record boundaries, styles, fonts, size, nudges). Implement a separate legacy-v3 parser only when needed; never treat v3 as v5.
2. Write AST rewrite transforms for limited, rigorously supported operations (symbol replace, index reassignment, insertion/deletion of well-defined templates). Keep original stream as canonical until transformation is validated.
3. **Extend the implemented CFB v3 writer** with broader storage layouts, DIFAT for exceptionally large embedded objects, and independent conformance tests. Preserve correct EQNOLEFILEHDR length fields; do not patch raw MTEF bytes without reallocating stream sectors.
4. Expand the current limited AST/MTEF encoder to cover advanced matrix, integral, fence, alignment, accent and template records; define symbol and size policies from verified MathType output. Avoid copying licensed templates/code without review.
5. Provide deterministic Linux preview generation for each equation, matching MTEF layout as closely as possible; compare against Word/MathType reference renderings.
6. Test against tensor indices, matrices, derivatives, multi-line aligned equations, integrals, Greek glyphs, font roles, formula numbering, inline baselines, mixed CFB versions, cross-references and legacy documents.
7. Require Windows Word and MathType compatibility proof (open -> double-click -> edit -> save -> close -> reopen). Preserve copies/hashes and provide rollback.

The self-owned editor is the preferred future backend. `docx-equation` remains an **optional temporary generation adapter**, isolated in `requirements-generator.txt` until a native MTEF encoder passes identical acceptance tests.

## Font-style and fallback image separation

MathType MTEF `FONT_DEF` associates fonts with encodings. `EQN_PREFS` carries full/sub sizes and style mappings with bold/italic bits. The journal donor currently uses Times New Roman 12 pt for variables/numbers, Symbol for Greek/symbol, and applies italic to variable and lower-case Greek styles. Do not overwrite these definitions when regenerating an object's visible preview. The DOCX VML image fallback was previously rendered in DejaVu Sans and caused a discrepancy; the native renderer now uses Tinos/ STIX and matches the 12 pt baseline more closely. `mtef_font_audit.py` emits a machine-readable profile, allowing the selection of a consistent donor. Matplotlib Tinos is not a pixel-perfect replacement for the actual Times New Roman font on Windows.

## v3: semantic typeface assignment, not just preview substitution

The initial binary generator already reproduced the donor's Times New Roman /
Symbol `FONT_DEF` and `EQN_PREFS` records, including the 12 pt base size.
The v3 encoder also controls the `CHAR` typeface byte for **each character**:
Text=1, Function=2, Variable=3, L.C. Greek=4, U.C. Greek=5, Symbol=6,
Vector=7, Number=8 (values are stored biased by 128 in MTEF v5). It supports
named function macros and the `FUNC_START` bit, as well as Latin text/roman
and bold vector spans through `\mathrm`, `\operatorname`, and `\mathbf`.
It does not silently conflate unsupported LaTeX Greek variants or bold Greek.
`mtef_subset.validate_subset_mtef` independently parses the generated MTEF
structure and checks character/role/font-position consistency.

`build_font_role_test.py` creates a DOCX with 15 real OLE objects. Original
MTEF style preferences are compared to those in each generated object; one
original OLE is copied as a Windows control. Neither MathType rendering nor
font-selection effects after reopening in Windows can be proved offline.
Run `python scripts/test_font_semantics.py --source source.docx`, then perform
the Windows acceptance in `references/windows-acceptance.md`.
