#!/usr/bin/env python3
"""Render editable MathType object previews with a strict installed-font policy.

For scientific paper production, require a licensed, locally installed genuine
Times New Roman family. Do not silently substitute Tinos/STIX for it. Matplotlib
may still use STIX for mathematical symbols missing from TNR; this is not a
pixel-perfect MathType layout engine. No font binaries are packaged.

Only the fallback PNG changes here. Editable fonts come from MTEF records.
"""
from __future__ import annotations
from io import BytesIO
import re
from functools import lru_cache
import os
import argparse
import json
import matplotlib as mpl
mpl.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties, findfont
from PIL import Image
from check_tnr import inventory as inspect_tnr
from pathlib import Path


@lru_cache(maxsize=1)
def preview_font_settings() -> tuple[str, dict[str, str]]:
    """Resolve a genuine installed TNR family; fail instead of silently falling back.

    MATHTYPE_PREVIEW_FONT may be explicitly set to a substitute for *tests only*.
    The production default is Times New Roman. Require all four real faces to
    avoid Matplotlib synthesizing fake italic/bold variants.
    """
    font_name = os.environ.get('MATHTYPE_PREVIEW_FONT', 'Times New Roman').strip()
    if not font_name:
        font_name = 'Times New Roman'
    faces = {}
    genuine_fonts = None
    if font_name.casefold() == 'times new roman':
        genuine_fonts = inspect_tnr()
        if not genuine_fonts['ready']:
            raise RuntimeError(
                'True Times New Roman (Regular, Italic, Bold, Bold Italic) is not installed. '
                'Fontconfig may substitute Tinos; that is not accepted. '
                'Run scripts/install_tnr.sh --apt or '
                'scripts/install_tnr.sh --from-dir /licensed/font/files; then rerun --check. '
                f"Missing faces: {', '.join(genuine_fonts['missing_faces'])}."
            )
    try:
        for role, style, weight in (
            ('regular', 'normal', 'normal'),
            ('italic', 'italic', 'normal'),
            ('bold', 'normal', 'bold'),
            ('bold_italic', 'italic', 'bold'),
        ):
            faces[role] = findfont(
                FontProperties(family=font_name, style=style, weight=weight),
                fallback_to_default=False,
            )
    except ValueError as exc:
        raise RuntimeError(
            f'Font family {font_name!r} not installed or not registered with Matplotlib. '
            'For production install genuine Times New Roman with a valid license; '
            'run fc-cache -f and refresh the Matplotlib font cache. '
            'To deliberately use a substitute in diagnostic tests only, set '
            'MATHTYPE_PREVIEW_FONT=Tinos.'
        ) from exc
    if genuine_fonts is not None:
        for role, path in faces.items():
            if str(Path(path).resolve()) not in genuine_fonts['faces'][role]:
                raise RuntimeError(
                    f'Matplotlib selected {path!r} for {role}, not the genuine Times New Roman face. '
                    'Refresh the Matplotlib font cache after installing the original fonts.'
                )
    if len(set(faces.values())) < 4:
        raise RuntimeError(
            f'Font family {font_name!r} lacks distinct Regular, Italic, Bold and '
            'Bold Italic font files; install all four actual faces.'
        )
    return font_name, {
        'font.family': font_name,
        'mathtext.fontset': 'custom',
        'mathtext.rm': font_name,
        'mathtext.it': font_name + ':italic',
        'mathtext.bf': font_name + ':bold',
        'mathtext.bfit': font_name + ':italic:bold',
        'mathtext.sf': font_name,
        'mathtext.tt': font_name,
        'mathtext.fallback': 'stix',
    }


def preview_png(latex: str, font_size_pt: float = 12.0):
    """Return (PNG bytes, width pt, height pt) for an editable OLE's preview.

    Preview PNG is *not* the math object. It is re-rendered independently of
    the CFB/MTEF object, using genuine Times New Roman by default. Mathematical
    symbols absent from Times New Roman may still fall back to STIX.
    """
    if not (7.0 <= font_size_pt <= 36.0):
        raise ValueError('Preview font_size_pt must be between 7 and 36 points')
    _, config = preview_font_settings()
    dpi = 240
    # MathText lacks MathType's operatorname macro: visual-only normalization.
    display_latex = re.sub(r'\\operatorname\{([A-Za-z]+)\}',r'\\mathrm{\1}',latex)
    display_latex = re.sub(r'\\text\{([A-Za-z]+)\}',r'\\mathrm{\1}',display_latex)
    with mpl.rc_context(config):
        fig = plt.figure(figsize=(7.0, 1.1), dpi=dpi)
        fig.patch.set_alpha(0)
        fig.text(0.01, 0.5, '$'+display_latex+'$', fontsize=font_size_pt,
                 horizontalalignment='left', verticalalignment='center')
        buf = BytesIO()
        try:
            fig.savefig(buf, format='png', dpi=dpi, transparent=True,
                        bbox_inches='tight', pad_inches=0)
        finally:
            plt.close(fig)
    im = Image.open(BytesIO(buf.getvalue())).convert('RGBA')
    bbox = im.getbbox()
    if not bbox:
        raise ValueError('Formula preview rendered as empty')
    # Tight crop; two transparent pixels ensure no glyph clipping.
    x0,y0,x1,y1 = bbox
    im = im.crop((max(0,x0-2),max(0,y0-2),min(im.width,x1+2),min(im.height,y1+2)))
    out = BytesIO()
    im.save(out, format='PNG')
    return out.getvalue(), round(im.width*72/dpi,2), round(im.height*72/dpi,2)


def main():
    ap = argparse.ArgumentParser(description='Check genuine Times New Roman availability')
    ap.add_argument('--check', action='store_true', help='Validate all four font faces')
    args = ap.parse_args()
    try:
        name, _ = preview_font_settings()
    except RuntimeError as exc:
        print(json.dumps({'status': 'missing_or_incomplete', 'error': str(exc)}, indent=2))
        return 2
    print(json.dumps({'status': 'ok', 'selected_font': name,
                      'production_ready': name == 'Times New Roman',
                      'note': 'Math symbols can use STIX fallback; native MTEF style is independent.',
                      'font_faces_verified': name == 'Times New Roman'},
                     indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
