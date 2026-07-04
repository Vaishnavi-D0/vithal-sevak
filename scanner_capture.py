"""
scanner_capture.py
Captures a photo from a WIA-compatible scanner (Windows only).
Works with most flatbed / passport-photo scanners.

The scanner bed is A4-sized, but the actual photo placed on it is always a
small passport-size photo somewhere on that bed - so instead of assuming
the photo fills the whole scanned frame, we auto-detect the photo's actual
bounds (the non-background region) and crop to just that.
"""

import os
from datetime import datetime

PHOTOS_DIR = "photos"
os.makedirs(PHOTOS_DIR, exist_ok=True)

SCAN_DPI = 1200
FINAL_LONG_EDGE_PX = 1200
# How far (out of 255, grayscale) a pixel must differ from the detected
# background shade to be considered part of the actual photo.
BACKGROUND_DIFF_THRESHOLD = 25
# Small margin (px, at SCAN_DPI) kept around the detected photo edges so
# auto-crop doesn't clip into the picture.
CROP_PADDING = 15
# Width (px) of the border strip sampled to figure out what the scanner
# bed/lid background actually looks like - it isn't always white; many
# scanners (e.g. most Epson flatbeds) have a black or dark grey lid.
BORDER_SAMPLE_PX = 40


def _detect_background_shade(gray):
    """Estimates the scanner bed/lid background's grayscale shade by
    sampling a strip along all 4 edges of the scan (the photo is assumed
    to not touch the very edges of the bed) and taking the median."""
    w, h = gray.size
    strip = min(BORDER_SAMPLE_PX, w // 4 or 1, h // 4 or 1)
    if strip <= 0:
        return 255

    samples = []
    top_strip = gray.crop((0, 0, w, strip))
    bottom_strip = gray.crop((0, h - strip, w, h))
    left_strip = gray.crop((0, 0, strip, h))
    right_strip = gray.crop((w - strip, 0, w, h))
    for region in (top_strip, bottom_strip, left_strip, right_strip):
        samples.extend(region.getdata())

    if not samples:
        return 255
    samples.sort()
    return samples[len(samples) // 2]  # median


def _autocrop_to_photo(img):
    """Finds the bounding box of the actual photo content on the scanned
    A4 page (i.e. everything that differs from the scanner bed/lid
    background, whether that background is light or dark) and crops to
    it, with a small padding margin. Falls back to the full image if no
    distinct content region is found."""
    gray = img.convert("L")
    background_shade = _detect_background_shade(gray)
    binary = gray.point(lambda p: 255 if abs(p - background_shade) > BACKGROUND_DIFF_THRESHOLD else 0)
    bbox = binary.getbbox()
    if not bbox:
        return img

    left, top, right, bottom = bbox
    left = max(0, left - CROP_PADDING)
    top = max(0, top - CROP_PADDING)
    right = min(img.width, right + CROP_PADDING)
    bottom = min(img.height, bottom + CROP_PADDING)
    return img.crop((left, top, right, bottom))


def _resize_long_edge(img, target_px):
    w, h = img.size
    if w <= 0 or h <= 0:
        return img
    if w >= h:
        new_w = target_px
        new_h = max(1, round(h * (target_px / w)))
    else:
        new_h = target_px
        new_w = max(1, round(w * (target_px / h)))
    from PIL import Image
    resample = getattr(Image, "LANCZOS", getattr(Image, "ANTIALIAS", None))
    return img.resize((new_w, new_h), resample)


def scan_photo():
    """
    Opens the scanner, scans the A4 bed, auto-crops to just the passport
    photo actually placed on it, resizes so its longer edge is
    FINAL_LONG_EDGE_PX, and returns the saved file path.
    Only works on Windows with a WIA-compatible scanner connected.
    """
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32 not installed or not on Windows. Run: pip install pywin32"
        )

    wia = win32com.client.Dispatch("WIA.CommonDialog")
    device = wia.ShowSelectDevice()  # lets user pick scanner if multiple connected

    if device is None:
        raise RuntimeError("No scanner selected.")

    # WIA_ITEM_ID for flatbed scanners; item(1) is usually the scan bed
    item = device.Items[1]

    # Set scan properties: color, high DPI for a precise auto-crop and a
    # sharp final image once cropped down to just the passport photo.
    try:
        item.Properties("6146").Value = 1              # 1 = Color intent
        item.Properties("6147").Value = SCAN_DPI        # Horizontal resolution (DPI)
        item.Properties("6148").Value = SCAN_DPI        # Vertical resolution (DPI)
    except Exception:
        pass  # some scanners don't expose all properties; safe to skip

    image = item.Transfer()  # triggers the actual scan

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    raw_path = os.path.join(PHOTOS_DIR, f"scan_raw_{timestamp}.bmp")
    image.SaveFile(raw_path)

    from PIL import Image
    img = Image.open(raw_path)
    img = _autocrop_to_photo(img)
    img = _resize_long_edge(img, FINAL_LONG_EDGE_PX)

    final_path = os.path.join(PHOTOS_DIR, f"photo_{timestamp}.jpg")
    img.convert("RGB").save(final_path, "JPEG", quality=95)

    os.remove(raw_path)  # cleanup raw bmp
    return final_path
