"""
legacy_font_converter.py
Best-effort converter from legacy glyph-substitution Marathi fonts (the
AkrutiDynamicMar_BYogini family and similar Akruti-style fonts) to Unicode
Devanagari.

IMPORTANT - accuracy caveat: these legacy fonts work by remapping ordinary
ASCII/Latin-1 byte codes to Devanagari-look glyphs at the font-rendering
level; the underlying stored characters are NOT Unicode Devanagari. There is
no single official "Akruti" mapping - it varies by font file, and this
module's table has NOT been verified byte-for-byte against the specific
AkrutiDynamicMar_BYogini font. Treat every converted string as a draft:
review it (and correct it with the app's on-screen Marathi keyboard) before
relying on it.

The conversion works by:
1. Matching the longest possible multi-character glyph sequences first
   (legacy fonts often use 2-3 ASCII characters to form one Devanagari
   glyph/conjunct), falling back to single characters.
2. Leaving any unmapped character untouched, so unsupported input degrades
   to "partially converted" rather than silently dropping data.
"""

# Ordered longest-match-first: (source_sequence, devanagari_unicode)
# Multi-character sequences MUST come before their single-character prefixes.
_MAPPING_TABLE = [
    # Common conjuncts / matra combinations (best-effort, longest first)
    ("Ø«", "ज्ञ"), ("Ù", "क्ष"), ("i{", "त्र"), ("Sar", "श्र"),
    # Independent vowels
    ("v", "अ"), ("vk", "आ"), ("b", "इ"), ("bZ", "ई"), ("m", "उ"), ("mZ", "ऊ"),
    ("_", "ऋ"), (",", "ए"), (",s", "ऐ"), ("vks", "ओ"), ("vkS", "औ"),
    # Anusvara / visarga / chandrabindu
    ("a", "ं"), ("%", "ः"), ("¡", "ँ"),
    # Consonants
    ("d", "क"), ("[k", "ख"), ("x", "ग"), ("?k", "घ"), ("³", "ङ"),
    ("p", "च"), ("N", "छ"), ("t", "ज"), (">", "झ"), ("¥", "ञ"),
    ("V", "ट"), ("B", "ठ"), ("M", "ड"), ("<", "ढ"), (".k", "ण"),
    ("r", "त"), ("Fk", "थ"), ("n", "द"), ("/k", "ध"), ("u", "न"),
    ("i", "प"), ("Q", "फ"), ("c", "ब"), ("Hk", "भ"), ("e", "म"),
    (";", "य"), ("j", "र"), ("y", "ल"), ("G", "ळ"), ("o", "व"),
    ("'k", "श"), ("\"k", "श"), ("\"", "ष"), ("l", "स"), ("g", "ह"),
    # Halant (virama) - joins consonants into conjuncts
    ("~", "्"),
    # Dependent vowel signs (matras)
    ("k", "ा"), ("f", "ि"), ("h", "ी"), ("q", "ु"), ("w", "ू"),
    ("s", "े"), ("S", "ै"), ("ks", "ो"), ("kS", "ौ"),
    # Digits (Devanagari)
    ("0", "०"), ("1", "१"), ("2", "२"), ("3", "३"), ("4", "४"),
    ("5", "५"), ("6", "६"), ("7", "७"), ("8", "८"), ("9", "९"),
]
# Sort by descending source-sequence length so multi-char sequences are
# tried before any single-character prefix they contain.
_MAPPING_TABLE.sort(key=lambda pair: -len(pair[0]))


def convert_legacy_marathi(text):
    """Best-effort conversion of legacy-font-encoded text to Unicode
    Devanagari. Unmapped characters are left as-is. NOT verified for
    AkrutiDynamicMar_BYogini specifically - always review the output."""
    if not text:
        return text

    result = []
    i = 0
    n = len(text)
    while i < n:
        matched = False
        for source, devanagari in _MAPPING_TABLE:
            if text.startswith(source, i):
                result.append(devanagari)
                i += len(source)
                matched = True
                break
        if not matched:
            result.append(text[i])
            i += 1
    return "".join(result)
