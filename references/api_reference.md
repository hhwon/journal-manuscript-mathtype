# Reproducible backend interface (version-specific)

## Exact upstream source

- Repository: https://github.com/npofsi/docx-equation
- Repository commit: `b70517cc22e85cb003d541dc592af2ee0e70ac7c`
- Code version at that commit: `0.3.3`, AGPL-3.0-only.
- The last verified PyPI distribution is `0.3.0` (MIT, June 12, 2026). It does **not** provide the same `EquationRegistry` interface. Do not use `pip install docx-equation==0.3.3` without checking release availability. Run `pip install -r requirements-generator.txt` for the Git-pinned adapter.

The verified API is defined in `src/docx_equation/document.py` and `src/docx_equation/shared/models.py` of the pinned GitHub repository, not inferred from the PyPI 0.3.0 wheel.

## Imported symbols

```python
from docx_equation import (
    EquationRegistry, EquationStyle, ExportOptions,
    MathTypeOptions, NumberingOptions,
)
from latex2mathml.converter import convert
```

## Minimum flow

```python
registry = EquationRegistry()
p = doc.add_paragraph('Let ')
registry.add_inline(p, convert(r'u(x)'))
registry.add_display(doc, convert(r'\sigma_{ij,j}=0'), number=1)
result = registry.save(doc, 'paper.docx', ExportOptions(
    target='mathtype',
    style=EquationStyle(font_family='Times New Roman', font_size_pt=12,
                        number_font_size_pt=12),
    numbering=NumberingOptions(number_format='(1)', use_seq_field=True),
    mathtype=MathTypeOptions(embed_mode='png-preview', mathtype_version='DSMT4',
                             preview_pt_per_px=0.38),
))
assert result.converted == len(registry)
```

`EquationRegistry` saves hidden formula anchors first and replaces them with OLE objects. Its display formatter **replaces the whole equation paragraph**, so this Skill runs `repair_equation_layout.py` and `patch_equation_bookmarks.py` *after* registry save.

## Renderer calibration

The 0.3.3 `MathTypeOptions` default `preview_pt_per_px=0.15` can make typical equations very small, because its preview sizing multiplies the source PNG dimensions by points-per-pixel. This Skill defaults to 0.38, an *engineering starting value*, not yet cross-checked on Windows. Use `--math-preview-scale` and test with the local MathType 7 installation. Preserve the aspect ratio and cap wide equations to the available text width. Do not claim perfect fidelity from this numerical setting alone.

The `png-preview` mode **raises** if its Chromium renderer fails. The alternate-content mode may replace failed previews with placeholder images and is therefore not recommended for strict MathType-first manuscripts.

## Failure contract

- Missing Python import, API mismatch or non-0.3.3 backend: stop.
- LaTeX-to-MathML failure: stop and report the original LaTeX.
- ConversionSummary errors or converted count mismatch: delete incomplete document and stop.
- Missing MTEF OLE stream or relationship: delete invalid document and stop.
- Never silently generate OMML or image-only equations when the user asked for MathType objects.
- Windows Word + MathType activation remains a separate acceptance test.
