# Journal manuscript production release gates

## 0. Scientific integrity (required for every publication draft)

- Preserve source text, equations, parameter definitions, datasets, figure paths and reference provenance. Do not invent citations, numerical results, experimental validation, model derivations or performance claims.
- Run `scripts/lint_manuscript.py input.json --submission`. Resolve all errors and inspect warnings. A linter cannot independently fact-check the science.
- For user-provided drafts, prefer surgical edits. Keep an editorial change summary separate from the manuscript.
- Label *synthetic tests* as such. The sample JSON and illustration must never be used as scientific evidence.

## 1. Environment

Run `python scripts/doctor.py --strict`. Match `docx-equation 0.3.3` and its pinned API; check Chromium and mathematical fonts. Keep backend source revision and machine environment reproducible.

## 2. Structural output

Run the builder without `--layout-preview`. It must report all expected equations converted, write `.docx`, and pass `scripts/validate_docx.py --expected N --require-mathtype` where N includes inline equations, display equations and any formulas in captions/tables.

- Genuine OLE CFB + non-empty Equation Native directory stream
- MathType ProgID `Equation.DSMT4`
- Valid OOXML OLE part relationships
- Correct Word `SEQ` / `REF` field and bookmark targets
- No stray placeholder or `MERGEFORMAT` visible text

### MathType gate (generated equations; release blocker)

`python scripts/mathtype_gate.py out.docx` must pass (`build_paper.py` runs it and refuses the output otherwise):
every `Equation Native` parses as MTEF 5, every compound-file directory tree is valid, no `w:updateFields`, no
Symbol-style character without its Symbol font position. These checks are structural; they found every defect that
made Windows Word/MathType reject generated objects, but they do not replace the Windows test.

## 3. Word layout

Run `scripts/audit_format.py file.docx`. Render *all pages* with the packaged DOCX renderer where available. Review each page image for text clipping, equation preview sizing, figure captions, table headings, page overflow, poor page breaks, missing Unicode glyphs and invalid references. Correct and re-render until clean.

Headless LibreOffice may show MathType OLE differently from Windows Word; it is a **visual check**, not an interoperability certificate. It ignores `w:position` on OLE objects (genuine MathType documents also show raised inline formulas there) and draws only the preview PNG, never the MTEF: never judge equation baselines or MathType content from it.

## 4. Math corpus

On Linux with dependencies: `python scripts/smoke_math.py -o math-regression.docx` followed by structure audit. The corpus includes tensor derivatives, weak forms, integrals, matrices, cases, non-dimensional variables and Greek symbols. Failures are a release blocker unless explicitly resolved or documented as unsupported symbols that never appear in the manuscript.

## 5. Windows compatibility

Perform the edit-save-close-reopen checklist in [windows-acceptance.md](windows-acceptance.md) on Word + MathType 7. Do not mark this `pass` without actually running it.

## 6. Submission preflight

Verify targeted journal author instructions, bibliography format, total words/figures, official disclosures, order of authors and affiliations, physical units, figures resolution, numbering, citations and plagiarism policy. User-provided sample layout is not automatically the journal's final submission style.

## Existing OLE manuscript edit gate (self-owned module)

When changing an original DOCX formula, execute `scripts/ole_editor.py list`, identify the exact object by paragraph and preview context, then `replace` with a compatible genuine donor. Verify the output contains the same number of OLE objects; compare all unedited OLE and preview hashes. Require `--allow-legacy-equation3` only for genuinely legacy `Equation.3`/MTEF v3 objects. Do not silently upgrade v3 objects or treat them as newer DSMT4 objects. Run `scripts/test_ole_editor.py --source ... --donor ...` on a trusted local corpus before relying on batch edits. Finish with Windows Word/MathType edit/save/reopen; Linux inspections do not imply editor activation.

The minimal custom editor does not require the optional `docx-equation` package. `doctor.py --strict` is for **generating new MathType formulas**, not extracting or replacing existing formulas from real donor DOCX files.


### Native MTEF experimental gates

Use `python scripts/test_native_writer.py --source original-with-mathtype.docx`. Ensure new output has real `Equation Native` CFB streams, correct MTEF v5 header and ProgID, payload length consistency, independent previews and unique OPC relationships. Before accepting a rewrite, verify all other OLE and preview hashes remain unchanged. Record the Windows application check as `unverified` until the actual editor is tested. The native encoder uses only a limited LaTeX grammar and must throw on unsupported constructs.

## Font-specific regression gate

- Verify MTEF real `FONT_DEF` and `EQN_PREFS` styles with `scripts/mtef_font_audit.py`, independently of PNG fallback.
- For a Times New Roman reference, ensure Latin variables and digits have the intended face and bold/italic flags; Greek and operator symbols may intentionally use the Symbol font.
- Generator previews come from `preview_math.py` (Times New Roman, MathType metrics); preview scale 0.158 pt/px gives 12 pt maths. The native-subset previews use Matplotlib with a serif maths font; never use default DejaVu Sans or an arbitrary scaling factor.
- Compare Word before editing, MathType editor, and Word after editing/reopening as **three distinct font-validation phases**.
- For a preview-only update, assert that every pre-existing OLE `.bin` payload is unchanged via SHA-256.

## Native font-regression gate (v3)

- Verify `EQN_PREFS`: Text/Function/Number = Times New Roman regular;
  Variable = Times New Roman italic; Vector = Times New Roman bold;
  lowercase Greek = Symbol italic; uppercase Greek/Symbol = Symbol upright.
  Full size follows the intended reference (12 pt for manuscript 2).
- Verify every synthesized `CHAR` uses the expected MTEF typeface (biased by
  128), valid MTCode and valid 8-bit Symbol font position where required.
- Verify `FUNC_START` is used for named functions; reject unknown variants
  such as `\epsilon` and `\varphi` until their MTCode/font positions are verified.
- Compare the original and generated font preference profiles, check each OLE
  and relationship structurally, then inspect PNG page renders.
- Do **not** declare MathType font parity until the Windows editor displays the
  expected italics, upright functions, Greek and bold vectors and an edited
  equation survives save/close/reopen. Preview images are a separate test.


### Required genuine Times New Roman preflight

- Run `scripts/check_tnr.py --strict` before Linux Word equation preview
  production. It checks **font files' internal family/style metadata**, not just
  `fc-match` substitution. Verify all 4 faces.
- Run `scripts/native_preview.py --check` to ensure Matplotlib resolves exactly
  those font files for upright, italic, bold and bold italic.
- Do not publish outputs produced with `MATHTYPE_PREVIEW_FONT=Tinos`; that override
  is limited to diagnostics.
- Do not conflate successfully rendering a Times New Roman-based preview with
  correct MathType MTEF styles or MathType-native numbering.
- If the network prevents Microsoft font installation, explicitly label the
  environment as unprepared and require a licensed local installation instead.
