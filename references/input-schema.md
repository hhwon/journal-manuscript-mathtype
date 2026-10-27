# Structured manuscript JSON

The builder accepts UTF-8 JSON. See `examples/demo_paper.json` for a complete model. Keep mathematical source information separate from formatted output; use `--write-equation-manifest` when a formula audit trail is required.

## Root fields

- `title`: required string; do not invent.
- `authors`: optional ordered list; each entry `"Name"` or `["Name", "1*"]` with superscript affiliation/corresponding marks; preserve actual order.
- `affiliations`: optional ordered list; each `"text"` or `["1", "text"]` (superscript mark, text).
- `corresponding_author`: optional known contact.
- `abstract`: supported technical summary, string.
- `keywords`: list of terms.
- `sections`: array of objects, each `{"heading":"1. Introduction", "blocks":[...]}`.
- `references`: list of bibliographic strings, based on verified source metadata; printed as `[n]` with a hanging indent in the given order (cite them in order of first citation). Entries that already start with `[` are kept; `"reference_style": "plain"` disables numbering.
- `acknowledgements`: optional sourced statement.

## Supported block types

- Paragraph: `{"type":"paragraph","text":"An arbitrary domain is considered."}`. A paragraph right after a display equation that starts with a lower-case letter ("where ...") is set without first-line indent; force with `"continues": true/false`.
- Mixed paragraph: `{"type":"paragraph","segments":[{"text":"Let "},{"latex":"u(x)"},{"text":" be ..."}]}`. Each inline formula becomes its own MathType OLE object.
- Subsection: `{"type":"subsection","heading":"2.1 Boundary conditions"}`.
- Heading: `{"type":"heading","level":1,"heading":"3. Numerical examples"}`. Level 2 is supported.
- Formula: `{"type":"equation","id":"eq:balance","latex":"\\nabla\\cdot\\mathbf{q}=0"}`. Alternative `mathml` field may contain full `<math ...>` XML.
- Figure: `{"type":"figure","id":"fig:mesh","path":"mesh.png","width_inches":4.9,"caption":"Mesh ...","alt_text":"..."}`. PNG/JPEG/SVG; SVG uses CairoSVG.
- Table: `{"type":"table","id":"tab:convergence","caption":"Convergence...","headers":["L","Error"],"rows":[["10","0.01"]],"note":"..."}`. Cells and captions can be segment objects with inline formulas. Three-line table by default; `"style": "grid"` keeps all borders.
- Note: `{"type":"note","text":"..."}`.

## Word cross-references

Every formula/figure/table may have a unique `id`. Omit `id` for automatically assigned IDs like `equation:1`, `figure:1`, `table:1`.

Insert a reference segment anywhere that accepts segments:

```json
{
  "type": "paragraph",
  "segments": [
    {"text": "The governing relation in "},
    {"ref": "eq:balance"},
    {"text": " is evaluated using "},
    {"ref": "fig:mesh"},
    {"text": "."}
  ]
}
```

This produces `Eq. (1)` and `Figure 1` as **Word REF fields**, not literal numbers. Captions use Word `SEQ` fields; formula numbering is controlled by the equation backend. Before production, duplicate IDs, missing reference targets and non-consecutive manually assigned numbers cause an error. Update fields in Word using Ctrl+A, F9 after rearrangements.

## Important boundaries

- The current builder uses sequential `(1)`, `(2)` display numbers. Appendix formats such as `(A.1)` or per-chapter numbering require a tested extension; do not fake them by overwriting XML text.
- No unsupported formulas are silently replaced by pictures or OMML.
- `--layout-preview` produces labelled placeholders instead of equations; it is unsuitable for publication.
- Figure paths are relative to the input JSON file, not the output DOCX.
- `--math-preview-scale 0.158` (default) gives 12 pt maths: Chromium renders at device scale 2 with 38 px per em. Equation options: `--greek times|symbol`, `--matrix-align left|center|right`, `--display-align left|center`, `--indent-twips 482`, `--profile file.json` (see SKILL.md).
- The optional `.equations.json` file exposes source formula text; treat it as confidential research material.

## Non-production sample guard

Set `"sample_data": true` in illustrative examples and test manuscripts. `scripts/lint_manuscript.py manuscript.json --submission` will refuse the synthetic sample; remove the flag only after replacing all simulated claims, authors, affiliations, references, plots and tables with verified content. The linter cannot itself verify whether the replacement research is genuine.
