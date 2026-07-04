"""
marathi_text_render.py
Renders Devanagari (Marathi) text into a small raster image with correct
glyph shaping (conjuncts like क्ल्प, matras, etc). This is needed because
reportlab's built-in text drawing only does a 1:1 Unicode-codepoint-to-glyph
mapping with no OpenType shaping, which breaks up Devanagari - e.g. "कल्पना"
renders as separate unjoined marks instead of properly formed conjuncts.

Uses uharfbuzz (shaping) + freetype (rasterizing each shaped glyph). Both
ship prebuilt wheels for Windows/macOS/Linux, so no system font-shaping
library (e.g. libraqm) needs to be installed on the target machine.
"""

import numpy as np
import uharfbuzz as hb
import freetype
from PIL import Image

_hb_font_cache = {}
_ft_face_cache = {}


def _get_hb_font(font_path):
    font = _hb_font_cache.get(font_path)
    if font is None:
        blob = hb.Blob.from_file_path(font_path)
        face = hb.Face(blob)
        font = hb.Font(face)
        _hb_font_cache[font_path] = font
    return font


def _get_ft_face(font_path):
    face = _ft_face_cache.get(font_path)
    if face is None:
        face = freetype.Face(font_path)
        _ft_face_cache[font_path] = face
    return face


def _bitmap_to_array(bitmap):
    """Converts a FreeType bitmap to a numpy array, respecting row pitch
    (FreeType often pads each row, so width != pitch)."""
    rows, width, pitch = bitmap.rows, bitmap.width, bitmap.pitch
    buf = bitmap.buffer
    arr = np.zeros((rows, width), dtype=np.uint8)
    for r in range(rows):
        start = r * pitch
        arr[r, :width] = list(buf[start:start + width])
    return arr


def render_text_line(text, font_path, font_size_pt):
    """Shapes and rasterizes one line of text.
    Returns (PIL.Image in 'L' mode with black-on-transparent glyphs,
    ascent_px) where ascent_px is the distance from the image's top edge
    down to the text baseline (needed by callers to align it with other
    text). Returns None if text is blank.
    """
    text = (text or "").strip()
    if not text:
        return None

    hbfont = _get_hb_font(font_path)
    scale = int(font_size_pt * 64)
    hbfont.scale = (scale, scale)

    buf = hb.Buffer()
    buf.add_str(text)
    buf.guess_segment_properties()
    hb.shape(hbfont, buf)

    infos = buf.glyph_infos
    positions = buf.glyph_positions

    ft_face = _get_ft_face(font_path)
    ft_face.set_char_size(int(font_size_pt * 64))

    pen_x = pen_y = 0.0
    glyphs = []
    min_x = min_y = 0.0
    max_x = max_y = 0.0
    for info, pos in zip(infos, positions):
        ft_face.load_glyph(info.codepoint, freetype.FT_LOAD_RENDER)
        bitmap = ft_face.glyph.bitmap
        left = ft_face.glyph.bitmap_left
        top = ft_face.glyph.bitmap_top
        x = pen_x + pos.x_offset / 64
        y = pen_y + pos.y_offset / 64
        arr = _bitmap_to_array(bitmap) if bitmap.rows and bitmap.width else None
        glyphs.append((arr, bitmap.width, bitmap.rows, left, top, x, y))
        if arr is not None:
            min_x = min(min_x, x + left)
            max_x = max(max_x, x + left + bitmap.width)
            max_y = max(max_y, y + top)
            min_y = min(min_y, y + top - bitmap.rows)
        pen_x += pos.x_advance / 64
        pen_y += pos.y_advance / 64
    max_x = max(max_x, pen_x)

    if max_x <= min_x:
        return None

    pad = 2
    width = int(round(max_x - min_x)) + pad * 2
    height = int(round(max_y - min_y)) + pad * 2
    canvas = np.zeros((height, width), dtype=np.uint8)

    for arr, gw, gh, left, top, x, y in glyphs:
        if arr is None:
            continue
        px = int(round(x + left - min_x)) + pad
        py = int(round(max_y - (y + top))) + pad
        for r in range(gh):
            row = py + r
            if row < 0 or row >= height:
                continue
            row_slice = canvas[row, max(0, px):min(width, px + gw)]
            src_start = max(0, -px)
            src = arr[r, src_start:src_start + row_slice.shape[0]]
            np.maximum(row_slice, src, out=row_slice)

    img = Image.fromarray(canvas, mode="L")
    ascent_px = int(round(max_y)) + pad
    return img, ascent_px
