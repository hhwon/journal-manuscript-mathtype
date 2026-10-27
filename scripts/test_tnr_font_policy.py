#!/usr/bin/env python3
"""Tests: an alias is not genuine Times New Roman; diagnostic fallback is explicit."""
from __future__ import annotations

import os
from pathlib import Path
import shutil
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
from check_tnr import inventory, metadata
from native_preview import preview_font_settings, preview_png
from matplotlib.font_manager import FontProperties, findfont

# A metric-compatible font renamed at filesystem level is still not TNR.
with tempfile.TemporaryDirectory() as tmp:
    tinos = Path(findfont(FontProperties(family='Tinos'), fallback_to_default=False))
    imposter = Path(tmp) / 'Times_New_Roman.ttf'
    shutil.copyfile(tinos, imposter)
    assert metadata(imposter) is None, 'Tinos must not pass Times New Roman metadata check'
    assert not inventory(tmp)['ready'], 'Filesystem name must not bypass internal family check'

original = os.environ.get('MATHTYPE_PREVIEW_FONT')
try:
    os.environ.pop('MATHTYPE_PREVIEW_FONT', None)
    preview_font_settings.cache_clear()
    actual = inventory()
    if not actual['ready']:
        assert actual['missing_faces'], 'No genuine TNR but no missing faces reported'
        try:
            preview_font_settings()
        except RuntimeError as exc:
            assert 'True Times New Roman' in str(exc)
        else:
            raise AssertionError('TNR absent; silent substitution must not be possible')
    else:
        assert preview_font_settings()[0] == 'Times New Roman'

    os.environ['MATHTYPE_PREVIEW_FONT'] = 'Tinos' # diagnostic override only
    preview_font_settings.cache_clear()
    family, settings = preview_font_settings()
    assert family == 'Tinos' and settings['mathtext.fontset'] == 'custom'
    png, width, height = preview_png(r'E=kT')
    assert png.startswith(b'\x89PNG\r\n\x1a\n') and width > 0 and height > 0
    print('PASS: strict TNR identity, four-face gate, and explicit diagnostic override')
finally:
    if original is None:
        os.environ.pop('MATHTYPE_PREVIEW_FONT', None)
    else:
        os.environ['MATHTYPE_PREVIEW_FONT'] = original
    preview_font_settings.cache_clear()
