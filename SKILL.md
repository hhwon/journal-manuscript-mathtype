---
name: journal-manuscript-mathtype
description: "Create, revise, and typeset scientific journal DOCX manuscripts with MathType OLE formulas on Linux. Use for manuscripts requiring consistent Word layout, a restrained technical English style, complex mathematical notation, formula/figure/table numbering, and preservation or replacement of preexisting MathType objects. Contains an independent OLE2/CFB writer and a conservative, native MTEF-v5 LaTeX subset encoder for generating or rewriting existing formulas, with donor-based replacement and an optional external backend for unsupported equations. Enforces strict validation and Windows editing checks."
---

# Journal manuscript with native MathType equations

Produce an evidence-grounded scientific manuscript with faithfully preserved technical content, controlled English style, high-quality Word layout, and native MathType OLE equations. The Linux server needs neither Word nor MathType installed; the Windows editing check requires them.

## Load the right reference

- **Writing/revision:** Read [language-profile.md](references/language-profile.md). Distil the rhetoric, not source typos or defective scientific claims. Default to restrained, method-oriented scientific English and minimal local edits.
- **DOCX formatting:** Read [format-profile.md](references/format-profile.md). Use `assets/manuscript-template.docx` (house style: see the indent, spacing, caption and table rules there).
- **Equations:** Read [mathtype-linux.md](references/mathtype-linux.md) and [api_reference.md](references/api_reference.md). Never quietly substitute Word OMML or pictures for MathType.
- **Document construction and references:** Read [input-schema.md](references/input-schema.md).
- **Existing DOCX with MathType OLE:** Read [custom-ole-editor.md](references/custom-ole-editor.md). Use `scripts/ole_editor.py` to inventory, extract and replace an existing formula using a genuine editable donor equation. No MathType/Word/external conversion library is needed for these operations. It indexes both MTEF v3 and v5; donor replacement covers both within compatible ProgIDs. A separate native MTEF-v5 subset encoder can now write or replace common formulas. Arbitrary MTEF AST edits and complete LaTeX support are not available. For safe prose-only edits use `scripts/safe_text_patch.py`, preserving all other DOCX package parts byte-for-byte.
- **Checks and supported claims:** Read [quality-gates.md](references/quality-gates.md) and [windows-acceptance.md](references/windows-acceptance.md).
- **Backend ownership:** Read [backend-design.md](references/backend-design.md). The preferred long-term design is a self-owned DOCX/OLE/MTEF core, not permanent reliance on EquationLab or docx-equation.

## Research and writing workflow

1. Identify the manuscript task: new manuscript, a section, minor revisions, reviewer response, or only layout. Preserve the given research domain, data, figures, symbols, boundary conditions and cited literature.
2. Do not invent mathematics, simulations, references, citations, error statistics, convergence behaviour, experimental evidence, authors, funding or journals. Make unsupported statements explicitly conditional or mark the evidence gap outside final submission content.
3. Maintain the source papers' scientific sequence: physical/numerical difficulty -> governing relations and constraints -> discretisation/computational implementation -> validation -> parameter analysis -> bounded conclusions.
4. Draft technical prose using the distilled model: informative topic sentences, precise defined terminology, one physical/computational result per claim, immediate symbol explanations after equations. Avoid generic AI transitions, inflated novelty and vague `excellent`, `significant` or `robust` without numerical evidence. Preserve the source text's author voice while correcting grammar.
5. For numerical examples, state geometry, material properties, boundary conditions, control parameters, numerical settings, benchmarks and observables. Discuss each figure/table by concrete values or a properly qualified trend.
6. On an existing paper, use minimal changes and a concise edit log. For exact single-run prose corrections, use the OLE-safe text patch; otherwise preserve the DOCX parts and manually review local XML changes. Never override a user's actual conclusion because a template sentence sounds more fluent.

## Surgically edit existing MathType equations (preferred self-owned path)

```bash
python -m pip install -r requirements.txt
python scripts/ole_editor.py list original.docx --out inventory.json
python scripts/ole_editor.py extract original.docx --index 12 --out-dir extracted-12
python scripts/ole_editor.py replace original.docx --index 12 \
    --donor math-donor.docx --donor-index 4 -o revised.docx
python scripts/validate_docx.py revised.docx --expected 249 \
    --require-mathtype --allow-legacy-equation3
```

An index is the location of the OLE object in `document.xml`, not necessarily a displayed equation number. Verify its `paragraph` and `context` before editing. `replace` updates the selected OLE CFB and corresponding preview; it copies the donor formula's width, height and baseline, preserves existing paragraph/numbering, isolates shared relationships, and verifies every unrelated formula and preview byte. Never replace an `Equation.3` (MTEF v3) using `Equation.DSMT4` without an explicit compatible conversion process.

**Status distinction:** `replace` uses an existing OLE donor. `replace-latex` below uses a newly synthesized MTEF-v5 record tree but covers only an explicit subset. Neither pathway proves MathType 7 activation without a Windows test.

## Self-owned MTEF v5 authoring and surgical LaTeX rewrite (experimental)

Use `scripts/mtef_subset.py`, `scripts/ole_cfb_writer.py` and `scripts/native_preview.py`. A **genuine MTEF v5 Equation.DSMT4 source DOCX is required as a *style and OLE-metadata seed***. The mathematical formula records are synthesized without MathType, Word, docx-equation or EquationLab; the style/CompObj wrapper is deliberately retained from the source. Do not claim a fully seed-free OLE generator or complete LaTeX parser.

Supported: Latin and common Greek variables, digits, basic operators, simple subscripts/superscripts, fractions, square roots, and native semantic font roles for `\sin`, `\cos`, `\tan`, `\exp`, `\log`, `\ln`, `\lim`, `\mathrm{...}`, `\mathbf{...}`, `\mathit{...}`, `\text{...}` and `\operatorname{...}` (Latin letters only for styled groups). Style variants not yet encoded correctly, such as `\epsilon`, `\varphi` and `\boldsymbol`, must be rejected rather than silently substituted. Refuse unsupported macros (e.g. integrals, matrices, alignment or advanced embellishments). Use the external adapter only with separate user approval and licensing review, **never silently**.

```bash
python -m pip install -r requirements-native-test.txt
# Create a reproducible Word test containing 9 real OLE parts:
python scripts/build_native_ole_test.py --source source-with-mathtype.docx \
    -o native-ole-test.docx
# Replace one existing MathType v5 formula from a supported LaTeX subset:
python scripts/ole_editor.py replace-latex original.docx --index 14 \
    --latex '\sigma_{ij}=C_{ijkl}\varepsilon_{kl}' -o revised.docx
python scripts/validate_docx.py revised.docx --expected 249 \
    --require-mathtype --allow-legacy-equation3
python scripts/test_native_writer.py --source source-with-mathtype.docx
```

For direct font-role regression (15 MathType OLE parts, including 13 self-generated expressions, one single-CHAR patch and one original control), run `python scripts/build_font_role_test.py --source source-with-mathtype.docx -o font-test.docx` and `python scripts/test_font_semantics.py --source source-with-mathtype.docx`. See [math-font-semantics.md](references/math-font-semantics.md).

**Font fidelity (required):** Editable MathType typography is controlled by MTEF `FONT_DEF` and `EQN_PREFS`, not Word's normal text style or the fallback PNG. Run `python scripts/mtef_font_audit.py source-with-mathtype.docx --indexes 1` to inspect the donor's variable, Greek, operator and number styles and full point size. Require genuine locally installed Times New Roman in `native_preview.py` by default; do not silently substitute Tinos or STIX. A missing TNR family or style face must raise an error. See [true-times-new-roman.md](references/true-times-new-roman.md) for font preflight and limitations. `MATHTYPE_PREVIEW_FONT` optionally selects an already-installed font; do not bundle font files. When matching a reference manuscript, prefer 12 pt Times New Roman variable/number with the appropriate italic flags and Symbol Greek/operator styles; require a real donor MTEF preamble with those settings. Test the unedited Word preview separately from the MathType editor view; never claim changing the preview fixes the editable formula's internal font. Native generator checks the original MTEF `EQN_PREFS` and the actual `CHAR` typeface roles separately. Linux verification is structural, not proof of identical Windows glyph metrics or spacing. See `references/windows-acceptance.md`.

`replace-latex` does not edit the original in place. It regenerates `Equation Native` with an updated MTEF record sequence and header length, rebuilds the compound file with new MiniFAT/FAT allocation, generates a PNG fallback preview, and updates only the target OLE and required Word relationships. The module audits *all* other OLE binaries/previews for byte-identical preservation. Source `Equation.3`/MTEF v3 is **not supported by the native writer**. A template script `build_native_ole_test.py` produces seven generated examples, one 1-symbol binary patch, one untouched original donor. Windows Word/MathType testing is still required; see [windows-acceptance.md](references/windows-acceptance.md).


## Experimental MathType-font and numbered-equation verification (v4)

When a user reports "MathType font is wrong" or "equation number is not native", do **not** treat a Times-like Word preview or ordinary `SEQ Equation` as success. Read [font-diagnostic-v4.md](references/font-diagnostic-v4.md) and [native-numbering.md](references/native-numbering.md). The existing Word generation backend's default `SEQ`/`REF` equation labels are **Word-only** and do not meet the user's MathType-native numbering requirement. Do not market these fallback labels as MathType 7 numbering.

Run `scripts/build_font_comparison_test.py` to create 4 matched semantic-vs-explicit font pairs. Each even-numbered formula carries 5 self-owned MTEF 5 `FONT_STYLE_DEF` records and signed negative typefaces. This route uses genuine donor Times New Roman and Symbol `FONT_DEF` entries, but whether it looks correct inside MathType **must be validated on Windows**.

For a manuscript built by this skill, run `scripts/mathtype_native_numbering.py out.docx` after `ole_fix` (or call `convert(path)` from the builder). It rewrites the `(SEQ Eq)` and fixed `(A.1)` labels as MathType's `MACROBUTTON MTPlaceRef`/`SEQ MTEqn` number fields (lettered labels through `SEQ MTChap \c \* Alphabetic`), the `REF` citations as `GOTOBUTTON`/`REF ZEqnNum…`, adds the hidden `MTEditEquationSection2` section breaks and the `MTDisplayEquation`/`MTEquationSection` styles, keeps every OLE part byte for byte, and checks the counts (`--check`). The field layout is the one MathType writes (see native-numbering.md).

Run `scripts/mathtype_numbering.py` with `--mode exact` to transplant the actual `MACROBUTTON MTPlaceRef` and nested `SEQ MTEqn` fields from a reference Word document whose numbers MathType inserted, and import `MTDisplayEquation` style. This is an experimental native-field transplant, not a proven alternative to MathType's Insert Right-numbered Equation ribbon action; use `test_font_numbering_v4.py` to verify ZIP and XML structure. In particular, using plain `(1)` text or generic `SEQ Equation` is not sufficient for genuine MathType numbering.

```bash
python scripts/build_font_comparison_test.py \
  --source manuscript-with-mathtype.docx -o font-numbering-v4.docx
python scripts/test_font_numbering_v4.py font-numbering-v4.docx \
  --reference manuscript-with-mathtype.docx
```

LibreOffice may show literal `MERGEFORMAT` artifacts for MathType macro fields. Do not claim layout QA passes for numbering based on LibreOffice. Validate the numbers, actual font, native Add-in Insert Reference workflow, and save/reopen behaviour in Windows Word with MathType 7. If a user requests a submission-ready document and this acceptance fails, do not deliver the experimental numbering as finished; use MathType's real Windows add-in to regenerate numbering or explicitly disclose this limitation.

## Generated equations must be genuine-equivalent (required for the generator backend)

docx-equation 0.3.3 alone writes MathType objects that Windows Word cannot edit and whose fonts, spacing and previews
differ from MathType's. `build_paper.py` therefore always applies the following modules; any other pipeline that calls
docx-equation must do the same, in this order. Evidence for every rule: [mathtype-facts.md](references/mathtype-facts.md).

1. `latex_mathml.to_mathml` — LaTeX -> MathML (non-BMP letters -> mathvariant, `\;` fix, double brackets as fences).
2. `mtef_profile.configure(greek=..., matrix_align=...)` before saving — patches the MTEF writer: MathType profile
   (Times New Roman letters/numbers, Symbol style in the **Symbol** font — never Times, or MathType cannot draw tall
   brackets and integrals when it redraws), factory spacing table, Symbol font positions, fnSPACE spaces, genuine
   template layouts (null LINE without END, scripts, big operators, fences, accents), matrix alignment.
3. `preview_math.install()` before saving — MathType-like PNG previews (scripts 58 %, MathType fractions, operator
   spacing, Symbol-shaped integrals and operators, bar accents) and measured baselines.
4. After saving: `repair_equation_layout.repair`, `patch_equation_bookmarks.patch`, `equation_layout.fix_layout`
   (`w:position` = −depth as MathType does, number tab inside the margin, display alignment), `ole_fix.fix_docx`
   (MathType CLSID, **valid compound-file directory tree** — without it no equation can be edited — single shapetype,
   no `w:updateFields`).
5. `mathtype_gate.py out.docx` must pass: MTEF of every object parses (`mtef_check.py`), every directory tree is valid
   (`cfb_treecheck.py`), no `updateFields`, no Symbol-style character without its Symbol position.

**House-style defaults** (configurable; the defaults of this skill): Times New Roman for every letter including Greek, bold always upright,
matrix contents left-aligned, display equations left-aligned at the 0.85 cm body indent with the number right-aligned
by a tab. MathType's own defaults instead: `--greek symbol --matrix-align center --display-align center`.

**MathType preference file:** `--profile settings.eqp` takes all styles (fonts, bold/italic), sizes and the 30 spacing
values from a MathType .eqp file ("MathType INI v1"; `examples/mathtype-factory.eqp` shows the format). They go into
the EQN_PREFS of every equation and drive the preview geometry; keys missing from the file keep factory values and are
reported. If the user supplies an .eqp, use it; otherwise use the defaults.

Requirements of the generator path (not optional): Playwright with its Chromium, genuine Times New Roman, TeX Gyre
Termes Math (`fonts-texgyre-math` or TeX Live) and URW Standard Symbols PS (`fonts-urw-base35`); `doctor.py` and
`python scripts/preview_math.py` check them.

Never judge baselines or MTEF from LibreOffice: it ignores `w:position` on OLE objects and draws only the preview.
To compare a preview with MathType, ask the user to re-save one equation in MathType and send a screenshot of both.

## Native DOCX construction (optional generator backend)

Use a structured JSON manuscript (`examples/demo_paper.json`) with paragraphs, inline/display MathType equations, images, tables, caption and equation IDs, and `REF` cross-references. Run the standard scripts from the unpacked skill root:

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements-generator.txt
python -m playwright install chromium        # REQUIRED: equation previews are rendered by preview_math.py
python scripts/doctor.py --strict
python scripts/preview_math.py               # fonts (Times New Roman, TeX Gyre Termes Math, URW Standard Symbols PS) + browser
python scripts/lint_manuscript.py examples/demo_paper.json
python scripts/build_paper.py examples/demo_paper.json -o manuscript.docx --write-equation-manifest
python scripts/validate_docx.py manuscript.docx --expected 3 --require-mathtype
python scripts/mathtype_gate.py manuscript.docx
python scripts/audit_format.py manuscript.docx
```

**Important:** Only optional `requirements-generator.txt` installs `docx-equation` directly from a specific GitHub commit with API `0.3.3`, not from the older PyPI wheel `0.3.0`. This backend is externally licensed under AGPL-3.0-only; review license implications before commercial distribution.

- Make all displayed formula IDs stable, e.g. `eq:weak-form`. Use paragraph `{"ref":"eq:weak-form"}` segments to insert **Word REF fields** with cached text `Eq. (n)`. Figure and table captions use `SEQ` fields. Field update: Word Ctrl+A, F9.
- Display equations: left-aligned at the 0.85 cm indent (default) or centred, with right-aligned `(n)` labels on a tab inside the margin. The backend replaces display paragraphs, so the layout repair, bookmark patch and `equation_layout` run after embedding.
- Center figures without first-line indent; plain `Figure n.` caption (not bold, justified, 12 pt) under them; plain `Table n.` caption above each table; three-line tables (rules at top, under the header, at the bottom; `"style": "grid"` keeps all borders). Add legible captions and meaningful alt text. Tables have repeatable header rows and no fixed row heights.
- Keep original figure assets unaltered except safe scaling and necessary SVG -> PNG conversion; never generate invented scientific curves.
- Normal text: Letter, 2.54 cm margins, TNR 12 pt, 1.5 lines, justified, 0.85 cm first line, **0 pt before and after** (no blank space between paragraphs). Section headings use Word's built-in **Heading 1 / Heading 2** styles (outline levels 1/2, so they appear in the Navigation Pane and a table of contents), restyled to the manuscript format. Headings, keywords label, figures, captions, equations and table cells have **no first-line indent**; a sentence continued after a display equation ("where ...") starts at the margin. Authors carry superscript affiliation marks (`["Name", "1*"]`), references are numbered `[n]` with a hanging indent, the footer shows the page number only. See [format-profile.md](references/format-profile.md).

## Preflight and verification

1. Run `python scripts/test_skill.py` for offline logic tests; on supplied real manuscripts run `python scripts/test_ole_editor.py --source old.docx --donor other.docx` and `python scripts/test_native_writer.py --source old.docx`. Use `python scripts/doctor.py --strict` only when generating NEW equations through the optional external backend.
2. Run `python scripts/smoke_math.py -o math-regression.docx` on a Linux host with dependencies; review its tensor, fraction, matrix, integration, Greek and index cases. Run the OLE audit on that output. The regression corpus is **not scientific data**.
3. Require all MathType conversions **and surgical edits** to succeed and satisfy the OLE structure checker. Check OLE CFB streams named `Equation Native`, MathType ProgID, `word/embeddings/*.bin`, working relationships, SEQ/REF bookmark targets and no leftover placeholders. Otherwise fail and remove incomplete output.
4. Render the DOCX to PNGs with a reliable DOCX renderer, inspect every page at full size, and iterate to fix overlaps, broken breaks, strange fonts, clipped formulas, plots, tables or footers. `--layout-preview` is a non-production, placeholder-only mode used when native MathType dependencies are unavailable.
5. **Do not certify Windows editability from Linux alone.** Follow [windows-acceptance.md](references/windows-acceptance.md): double-click diverse formulas in Word+MathType 7, edit, save, close, reopen, and update Word fields. Report Word compatibility as `unverified` until genuinely tested.
6. Confirm the intended journal's author guidelines, even if they conflict with the reference draft layout.

## Limits and change control

- Avoid broad rewriting, placeholder claims, simulated results, broken citation links, invented references and publishing synthetic examples.
- Keep formula LaTeX/MathML source with IDs for reproducibility. If an unsupported macro or equation construct fails, report it explicitly rather than substituting a picture.
- The independent **DOCX/OLE object-level edit module is implemented and integration-tested**, including legacy Equation.3/MTEF v3 and MTEF v5. Subset MTEF-v5 authoring and local replacement are implemented and structurally tested; a full spec-complete MTEF AST and cross-version edits remain unimplemented. EquationLab is used only as a documented reference for MathType's layout rules (read, not run, not bundled).
- Windows status of the generator path (2026-10): objects open and edit after the directory-tree repair (Times-Greek profile); MathType's own redraw of integrals and brackets with the Symbol style, left-aligned matrices (h_just 0) and the Symbol-Greek factory profile are not yet confirmed in Windows. Do not claim native arbitrary LaTeX formula modifications or Windows MathType activation are complete until independently tested.
- Use the final `quality-gates.md` checklist before declaring a paper submission-ready. A successful ZIP/OOXML inspection is not a substitute for scientific review or Microsoft Word/MathType verification.

## Install real Times New Roman on the Linux host

**Before production preview rendering**, install the licensed original font family
in the current Linux host (not merely in the DOCX metadata). Run
`python -m pip install -r requirements-native-test.txt` and choose one of:

```bash
# Debian/Ubuntu with the appropriate repository enabled and Internet access:
bash scripts/install_tnr.sh --apt
# Or install four licensed local font faces on a server without Internet:
bash scripts/install_tnr.sh --from-dir /path/to/licensed-fonts
```

The APT route uses the Microsoft Core Fonts installer and leaves any license
agreement to the user; it does not auto-accept. On Debian 13 it requires the
**contrib** component; Ubuntu requires **multiverse**. The local route verifies
the font's *internal family name and style flags* before installing. Then run:

```bash
python scripts/check_tnr.py --strict --summary
python scripts/native_preview.py --check
python scripts/test_tnr_font_policy.py
```

A report of Tinos from `fc-match 'Times New Roman'` is an alias, **not a real
font installation**. The Skill's checks explicitly reject such a substitution.
Never copy proprietary font files into the Skill ZIP or deliver them as assets.
For complete requirements, sources and network troubleshooting, see
[true-times-new-roman.md](references/true-times-new-roman.md).

## Strict preview rendering (native equations)

For Linux authoring of a Word `.docx` with real MathType OLE formulas, run
`python scripts/native_preview.py --check` before generating preview images.
If Times New Roman Regular/Italic/Bold/Bold Italic is unavailable, **fail** and
explain installation. Do not claim Tinos is Times New Roman. The user may
explicitly set `MATHTYPE_PREVIEW_FONT=Tinos` for visual-only regression tests,
never for publication-ready fidelity. STIX may still supply symbols missing
from genuine TNR; exact MathType equation layout and MathType-native numbering
require separate validation. See [true-times-new-roman.md](references/true-times-new-roman.md).
