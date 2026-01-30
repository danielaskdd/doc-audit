#!/usr/bin/env python3
"""
ABOUTME: XML utility functions for document processing
ABOUTME: Provides sanitization for XML-incompatible characters
"""


def sanitize_xml_string(text: str) -> str:
    """
    Remove control characters that are illegal in XML 1.0.

    XML 1.0 allows: #x9 (tab), #xA (LF), #xD (CR), and #x20-#xD7FF, #xE000-#xFFFD, #x10000-#x10FFFF
    This function removes all other control characters (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F).

    Args:
        text: Text that may contain control characters

    Returns:
        Sanitized text safe for XML. Returns input unchanged if not a non-empty string.
    """
    if not text or not isinstance(text, str):
        return text
    # Build a translation table to remove illegal control characters
    # Keep: \t (0x09), \n (0x0A), \r (0x0D)
    # Remove: 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F
    illegal_chars = ''.join(
        chr(c) for c in range(0x20)
        if c not in (0x09, 0x0A, 0x0D)
    )
    return text.translate(str.maketrans('', '', illegal_chars))
