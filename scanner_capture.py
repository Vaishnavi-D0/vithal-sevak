"""
scanner_capture.py
Captures a photo from a WIA-compatible scanner (Windows only).
Works with most flatbed / passport-photo scanners.
"""

import os
from datetime import datetime

PHOTOS_DIR = "photos"
os.makedirs(PHOTOS_DIR, exist_ok=True)


def scan_photo():
    """
    Opens the scanner, scans an image, saves it cropped to
    passport-size proportions, returns the saved file path.
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

    # Set scan properties: color, 300 DPI is good for passport photo printing
    try:
        item.Properties("6146").Value = 1      # 1 = Color intent
        item.Properties("6147").Value = 300     # Horizontal resolution (DPI)
        item.Properties("6148").Value = 300     # Vertical resolution (DPI)
    except Exception:
        pass  # some scanners don't expose all properties; safe to skip

    image = item.Transfer()  # triggers the actual scan

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    raw_path = os.path.join(PHOTOS_DIR, f"scan_raw_{timestamp}.bmp")
    image.SaveFile(raw_path)

    # Crop to passport-size ratio and save as JPG using Pillow
    from PIL import Image
    img = Image.open(raw_path)
    w, h = img.size
    target_ratio = 3.5 / 4.5

    if w / h > target_ratio:
        new_w = int(h * target_ratio)
        left = (w - new_w) // 2
        img = img.crop((left, 0, left + new_w, h))
    else:
        new_h = int(w / target_ratio)
        top = (h - new_h) // 2
        img = img.crop((0, top, w, top + new_h))

    img = img.resize((413, 531))  # ~3.5x4.5cm @ 300dpi
    final_path = os.path.join(PHOTOS_DIR, f"photo_{timestamp}.jpg")
    img.convert("RGB").save(final_path, "JPEG", quality=95)

    os.remove(raw_path)  # cleanup raw bmp
    return final_path
