from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ConversionResult:
    """Result of a file conversion operation"""

    input: Path
    output: Path
    engine: str  # 'pdf2docx' | 'docx2pdf' | 'soffice' | 'pdfplumber' | 'mammoth' | 'python-docx'


class ConversionError(Exception):
    """Raised when a conversion operation fails"""

    pass


class UnsupportedFormatError(ConversionError):
    """Raised when the input file format is not supported"""

    pass


class EngineNotFoundError(ConversionError):
    """Raised when required conversion engine is not available"""

    pass
