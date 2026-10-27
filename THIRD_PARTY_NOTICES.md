# Third-party components and sources

This repository is MIT-licensed (see LICENSE). No third-party code or font file is bundled in this Skill. Components are installed by the user or found on the host;
modified font copies are generated at run time in `~/.cache/journal-manuscript-mathtype/`.

| component | how it is used | licence / note |
|---|---|---|
| docx-equation (npofsi/docx-equation, pinned commit, API 0.3.3) | optional generator backend, installed by `requirements-generator.txt`; patched at run time by `mtef_profile.py` and `preview_math.py` (no source copied) | AGPL-3.0-only — review before commercial redistribution |
| latex2mathml, python-docx, lxml, olefile, Pillow, numpy, fontTools | Python dependencies | MIT / BSD / PIL licences |
| Playwright + Chromium | renders equation previews | Apache-2.0 / BSD |
| EquationLab (github.com/Ne1ther/EquationLab) | **reference only**: its reconstruction of MathType 7.8's layout rules (`src/native-layout.mjs`: fraction, script, integral geometry; `src/new-formula.mjs`: factory 'Times+Symbol 12' profile; `src/native.mjs`: fence MTCodes, matrix alignment codes) informed `preview_math.py`, `mtef_profile.py` and references/mathtype-facts.md. No code copied, nothing run. | MIT, © its authors |
| WIRIS / Design Science, MTEF v.5 specification and MTCode tables | reference for record layouts, spacing-table order, space and fence codes | linked, not copied |
| TeX Gyre Termes Math (GUST e-foundry) | found on the host; a copy with modified MATH constants is written under the new name "MTPreview Math", as the GUST Font License requires for modified versions | GUST Font License |
| URW Standard Symbols PS (URW base35) | found on the host; its glyph outlines are converted to a Unicode-mapped CFF font "MTPreview Symbol" that stays in the user's cache (never distributed) and only rasterises preview images | AGPL-3 with font exception (© 2015 URW Software, 2013–2014 (URW)++; Debian package fonts-urw-base35) |
| Times New Roman | must be installed by the user (see references/true-times-new-roman.md); never bundled | Monotype / Microsoft licence |
| `mathtype` Ruby gem (jure/mathtype) and mathtype_to_mathml | reference for MTEF 5 record decoding (`mtef_check.py` follows its layouts) | ISC / MIT |
