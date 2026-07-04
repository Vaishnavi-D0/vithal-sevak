"""
translator.py
English -> Marathi text translation (uses Google Translate via deep-translator).
Requires an internet connection.
"""

from deep_translator import GoogleTranslator


def translate_to_marathi(text: str) -> str:
    """Translates English text to Marathi. Returns '' for blank input."""
    text = (text or "").strip()
    if not text:
        return ""
    try:
        return GoogleTranslator(source="en", target="mr").translate(text)
    except Exception as e:
        raise RuntimeError(f"Translation failed: {e}")
