"""
md2docx - Markdown to Word (DOCX) Converter

A professional tool for converting Markdown documents to Word format with advanced
formatting capabilities.
"""

from .converter import (
    MarkdownToDocxConverter,
    DocumentConfig,
    PaperSize,
    DocumentStyle
)

__version__ = "1.1.3"
__all__ = [
    "MarkdownToDocxConverter",
    "DocumentConfig",
    "PaperSize",
    "DocumentStyle"
]
