"""
docx2pdfmake
============
Python library for converting .docx files to pdfmake Document Definition
Objects (DDO).
"""

from .converter import DocxConverter
from .models import ConversionOptions

__all__ = ["DocxConverter", "ConversionOptions"]
__version__ = "0.1.0"
