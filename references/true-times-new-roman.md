# Installing and enforcing **real Times New Roman** on Linux

This Skill **does not distribute Microsoft fonts**. The four requested faces
(Regular, Italic, Bold, Bold Italic) are commercial/proprietary font software.
Obtain them from a licensed Microsoft source or Microsoft's Core Fonts for the
Web installer, under the applicable license. Tinos, Liberation Serif and
other metric-compatible alternatives do **not** satisfy the production check.

## A. Debian 13 / Ubuntu with Internet (recommended)

Microsoft's legacy Core Fonts may be installed through
`ttf-mscorefonts-installer`. It is in Debian's **contrib** repository or
Ubuntu's **multiverse** component; if `apt-cache policy ttf-mscorefonts-installer`
has no Candidate, enable the correct component in your package manager.

For Debian 13 using a `*.sources` file, check the active stanza under
`/etc/apt/sources.list.d/`; its `Components:` line normally needs
`main contrib` rather than only `main`. On Ubuntu, enable `multiverse` in the
APT software sources. Run `sudo apt-get update` after making the change.

Install from the Skill root:

```bash
sudo apt-get update
bash scripts/install_tnr.sh --apt
python scripts/check_tnr.py --strict --summary
python scripts/native_preview.py --check
```

The installer **may require explicit agreement to Microsoft's license**. Read
and accept only if appropriate; the Skill never pre-accepts an EULA. The
installer may download font archives separately from Debian, so network access
is required. Debian reference: https://www.debian.org/doc/manuals/debian-reference/debian-reference.en.pdf

## B. Using licensed Windows font files without Linux network access

If the Linux server cannot reach font repositories, obtain **licensed**
Times New Roman files from a source whose license permits their installation
on that Linux computer. Common Windows filename conventions are `times.ttf`,
`timesi.ttf`, `timesbd.ttf`, and `timesbi.ttf` (naming varies).

Copy the four files **directly to the server**, not into the Skill ZIP, and run:

```bash
bash scripts/install_tnr.sh --from-dir /path/to/licensed-font-files
python scripts/check_tnr.py --strict --summary
python scripts/native_preview.py --check
```

The script validates the **internal font family and real style flags** via
`fontTools` *before copying* and installs the faces for the current Linux user
under `~/.local/share/fonts/TimesNewRoman/`. The original fonts are never
bundled in, exported from, or delivered with the Skill. Install `fontconfig`
and the Skill's `requirements-native-test.txt` Python dependencies first.

## Validation requirements

`fc-match 'Times New Roman'` alone is **not valid evidence**: fontconfig may
return an alias (e.g. Tinos). Use the strict check, which reads the actual
font's metadata for the four styles:

```bash
fc-match 'Times New Roman'
python scripts/check_tnr.py --strict --summary
python scripts/native_preview.py --check
```

If Matplotlib still has a stale cache after installation, restart the Python
process. If necessary remove only Matplotlib's `fontlist-v*.json` cache from
its configured cache directory and retry. Do not set
`MATHTYPE_PREVIEW_FONT=Tinos` for a production run; this env var is only an
**explicit temporary diagnostic override**.

## What this fixes and what it does not

- Fixes Linux **preview PNG** Latin glyph families/metrics and establishes real
  Times New Roman availability for DOCX rendering.
- **Does not** independently change MathType's editable MTEF `EQN_PREFS` or
  `CHAR` typefaces; inspect these using `scripts/mtef_font_audit.py`.
- Matplotlib MathText is not the MathType 7 typesetter. Roots, fractions,
  baselines, stretchy delimiters and specialized Greek/math glyphs may still
  differ; missing glyphs can fall back to STIX. MathType's numbering add-in
  has its own Word fields and must be tested on Windows.
- Production acceptance still requires Word + MathType 7 double-click, edit,
  save, close/reopen and numbering/reference tests.

## Environment note

If DNS/network access is disabled, `apt-get update` and the font download will
fail even when the installation commands are correct. **Do not claim the font
is installed** until all four faces pass the strict check. Running `--check`
in an offline sandbox tests current availability, not the user's remote server.
