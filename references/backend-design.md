# Backend architecture: self-owned edit first, generation swappable

The self-owned **DOCX/OLE editing core** (`scripts/ole_editor.py` and `scripts/ole_native.py`) is now implemented and tested on real manuscript copies. Read [custom-ole-editor.md](custom-ole-editor.md) for its exact API, supported editing scope, failure handling and verification record.

## Division of responsibilities

1. **Narrative + Word layout:** `build_paper.py`, Word styles and SEQ/REF fields, independent of equation backend.
2. **Existing-equation editing:** Our own `ole_editor.py` reads the DOCX OLE and preview relationships, extracts CFB `Equation Native`, and can atomically replace a specific OLE object with a genuine preexisting formula donor. No Word/MathType/external converter needed on Linux.
3. **New-equation generation (temporary):** Existing `docx-equation` 0.3.3 adapter, installed separately via `requirements-generator.txt`. It is not required for object-level OLE editing and is externally licensed under AGPL-3.0-only.
4. **Experimental self-owned MTEF authoring:** a limited typed MTEF-v5 AST, independent CFB v3 writer, and Linux PNG preview generator are implemented and structurally tested. A complete v5/v3 parser/encoder, arbitrary LaTeX, exact MathType preview metrics, and native semantic edits of arbitrary existing MTEF trees are *not* implemented. The Windows edit/save/reopen test is still pending.

## Standard new-formula backend contract

The source LaTeX/MathML and equation ID are independent of the format backend. Implement `parse`, `encode`, `render_preview`, `insert`, `verify`. Keep MTEF/visual previews consistent, perform the same strict tests on all backends, and require Windows compatibility checks for final sign-off.

## Generator output made genuine-equivalent (2026-10)

The docx-equation path is wrapped by self-owned modules: `latex_mathml.py` (input), `mtef_profile.py` (MTEF profile and
record layouts), `preview_math.py` (MathType-like previews and baselines), `equation_layout.py` (placement),
`ole_fix.py` (container: CLSID, directory tree) and `mathtype_gate.py` (release check). The rules come from 467 genuine
MathType objects, the WIRIS MTEF 5 specification and EquationLab's layout code; see
[mathtype-facts.md](mathtype-facts.md). They would apply unchanged to a future self-owned encoder.

## What to borrow conceptually from EquationLab

EquationLab's approach to editable MTEF records and preservation of native formatting offers design ideas, but no EquationLab source is bundled. Its reconstruction of MathType's layout rules (fractions, scripts, integrals, spacing table indices) is cited, with attribution, in mathtype-facts.md and THIRD_PARTY_NOTICES.md. Review third-party licenses for any concrete reuse. Its macOS CoreText font-metrics layer must not leak into a Linux-only core.

Select the future backend on reproducible results and licensing: high-order tensor indices, fractions, matrices, multiline equations, Greek/unicode, font styles, inline baseline, exact MTEF/CFB validity and Word+MathType editability. Do not infer conversion accuracy from README claims.


## 2026-10-26 native subset implementation

The architecture now has a new native MTEF subset pipeline: `mtef_subset.py` -> `ole_cfb_writer.py` -> `ole_editor.py replace-latex` or `build_native_ole_test.py` -> PNG preview -> DOCX OPC repair. The MTEF style and OLE metadata remain seeded from an actual `Equation.DSMT4` object; the mathematical record sequence is independently encoded. This avoids dependency on the optional external generator for simple supported formulas. The full MTEF AST and Windows activation acceptance remain incomplete.
