# Experimental MathType 7 equation numbering in DOCX

## Why normal Word numbers are insufficient

MathType 7's Word add-in uses an extended Microsoft Word field structure rather than a literal `(1)` or generic `{ SEQ Equation }` counter. A reference paper whose numbers were inserted by MathType provides authentic examples containing:

- `MACROBUTTON MTPlaceRef` (MathType reference-placement field)
- `SEQ MTEqn \\h` (hidden sequence increment)
- `SEQ MTEqn \\c \\* Arabic` (current equation number)
- Word paragraph/character styles `MTDisplayEquation`, `MTDisplayEquationChar`, and `MTEquationSection`

Native MathType 7 numbering is required, not Word-only visual imitations. Import the field structure and style from a genuine MathType paper. Preserve the MTEqn sequence name, macro name, control field and referenced display style exactly.

## Experiment in v4

`scripts/mathtype_numbering.py` upgrades the generated diagnostic document from static numbered tables into real MathType field paragraphs:

```bash
python scripts/mathtype_numbering.py unnumbered.docx \
  --reference reference-with-mathtype-numbers.docx \
  --mode exact -o native-numbering-test.docx
```

`--mode exact` copies the authentic `MTPlaceRef` compound field from the paper and changes only its cached numeral. A separate `canonical` reconstruction mode, building OOXML field separators and caches, is **diagnostic only** and not an endorsed substitute. The comparison document uses `exact` by default because the reference source was produced by MathType itself.

Centered and right-justified tab positions are updated to the actual document text width. Existing mathematical OLE binaries, previews and relationships are byte-identical before/after numbering, and all paragraphs retain genuine MathType OLE objects.

**Do not claim Word/MathType recognizes these fields until tested on Windows.** Linux's LibreOffice does not interpret the `MACROBUTTON MTPlaceRef` field as MathType does. It may display corrupted field-code fragments such as `MERGEFORMAT` even though the OOXML fields are syntactically present. This must not be called a successful visual rendering. Only Word + MathType 7 can settle whether numbering references work. Structural checks are necessary but insufficient.

## Windows acceptance test

1. Open a disposable copy in Windows desktop Word with MathType 7 add-in enabled.
2. Check whether numbers display `(1)`, `(2)` and stay at the right margin. If code is displayed, toggle field codes with `Alt+F9`, then update fields (`Ctrl+A`, `F9`) and observe.
3. Use MathType's **Insert Reference** command, select a target number, and verify that the reference can be inserted and updated, including after inserting/deleting an equation.
4. Save, close and reopen; check all references and numbering again.
5. Report whether the MathType right-numbered equation format dialog can recognize or reformat the existing numbers.

If any of the above fail, this field synthesis **is not fully MathType-native functional**, and the appropriate production path is to let the actual Windows MathType add-in create the numbering structure or use a verified original-document template without reconstructing its number fields.

## Sources

- WIRIS: [Can't change the font of a MathType equation number](https://docs.wiris.com/editing-formatting-equations/cant-change-the-font-of-a-mathtype-equation-number)
- WIRIS: [Display equations and equation numbers are out of position](https://docs.wiris.com/en_US/display-and-rendering/display-equations-and-equation-numbers-are-out-of-position-after-re-configuring-document)
- A user-supplied reference DOCX with MathType-inserted numbers (not bundled)
