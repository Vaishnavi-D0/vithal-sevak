"""
translator.py
English -> Marathi text translation (uses Google Translate via deep-translator).
Requires an internet connection.
"""

import socket

from deep_translator import GoogleTranslator

# A stalled/blocked connection (common on older or restricted networks)
# would otherwise hang indefinitely, making the whole app look like it
# has frozen or "gone away". Cap it so a failure surfaces in seconds.
NETWORK_TIMEOUT_SECONDS = 12


def translate_to_marathi(text: str) -> str:
    """Translates English text to Marathi. Returns '' for blank input."""
    text = (text or "").strip()
    if not text:
        return ""
    previous_timeout = socket.getdefaulttimeout()
    socket.setdefaulttimeout(NETWORK_TIMEOUT_SECONDS)
    try:
        return GoogleTranslator(source="en", target="mr").translate(text)
    except Exception as e:
        raise RuntimeError(f"Translation failed (check your internet connection): {e}")
    finally:
        socket.setdefaulttimeout(previous_timeout)
