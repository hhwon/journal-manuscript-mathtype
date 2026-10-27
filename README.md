# journal-manuscript-mathtype

A [Claude Code](https://claude.com/claude-code) skill for writing and typesetting scientific journal manuscripts as Word
documents with **editable MathType equations**, on Linux, without Word or MathType installed.

- Builds a DOCX manuscript from structured JSON: title block, sections, inline and display equations, figures,
  tables, captions, cross-references, numbered references.
- Equations are genuine `Equation.DSMT4` MathType objects: Windows Word + MathType can open, edit, re-save and
  redraw them like objects created in MathType.
- Edits equations in existing MathType manuscripts (replace one object, keep everything else byte for byte).
- Writing guidance for restrained technical English (computational mechanics style).

## Install

```bash
git clone https://github.com/hhwon/journal-manuscript-mathtype ~/.claude/skills/journal-manuscript-mathtype
cd ~/.claude/skills/journal-manuscript-mathtype
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements-generator.txt       # includes the optional docx-equation backend (AGPL-3.0)
python -m playwright install chromium
python scripts/doctor.py --strict                # checks packages, fonts and browser
```

Fonts needed on the host (not bundled): Times New Roman (see `references/true-times-new-roman.md`),
TeX Gyre Termes Math (`fonts-texgyre-math`) and URW Standard Symbols PS (`fonts-urw-base35`).

Only editing existing equations? `pip install -r requirements.txt` is enough.

## Use

In Claude Code, ask for a manuscript, a section, a revision or an equation edit; the skill is picked up from
`SKILL.md`. From the shell:

```bash
# new manuscript from JSON (see examples/demo_paper.json and references/input-schema.md)
python scripts/build_paper.py examples/demo_paper.json -o paper.docx
python scripts/mathtype_gate.py paper.docx       # structural check of every MathType object

# replace one equation in an existing manuscript with an equation from another MathType document
python scripts/ole_editor.py list old.docx --out inventory.json
python scripts/ole_editor.py replace old.docx --index 3 --donor donor.docx --donor-index 1 -o new.docx
```

Default layout: Times New Roman 12 pt, 1.5 lines, display equations left-aligned at the 0.85 cm indent with the
number on a right tab, matrices left-aligned, Greek in Times New Roman, three-line tables, plain `Figure n.` /
`Table n.` captions, numbered references. MathType's own equation defaults: `--greek symbol --matrix-align center
--display-align center`. Details: `references/format-profile.md`.

## Equation preferences (.eqp)

Give a MathType preference file to use its fonts, sizes and spacing:

```bash
python scripts/build_paper.py paper.json -o paper.docx --profile my-settings.eqp
```

Save your settings in MathType (Preferences → Equation Preferences → Save to File) or start from
`examples/mathtype-factory.eqp`. The settings are written into every equation, so MathType keeps them when an equation
is edited, and the previews are laid out with them. `--greek` and `--matrix-align` still apply on top. Without
`--profile`, the default layout above is used.

## How it stays editable

The generator's raw output is not accepted by Windows MathType. The skill fixes the OLE container (directory tree,
CLSID), writes MathType's own record layouts, font profile and Symbol font positions, renders previews with
MathType's spacing and measures their baselines, and refuses to deliver a file that fails `mathtype_gate.py`.
Every rule and its evidence: [`references/mathtype-facts.md`](references/mathtype-facts.md).

## Limits

- Linux checks are structural. Confirm editing in Windows Word + MathType before submission
  (`references/windows-acceptance.md`).
- The preview shown before an equation is edited approximates MathType's layout; MathType replaces it on the first edit.
- The native MTEF encoder covers a LaTeX subset; other constructs go through the optional docx-equation backend.
- Numbering uses Word fields; MathType's own numbering fields are experimental (`references/native-numbering.md`).

## License

MIT (see [LICENSE](LICENSE)). The optional docx-equation backend is AGPL-3.0-only and installed separately; a program
that combines it with this skill is subject to the AGPL. No fonts or third-party code are included; see
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
