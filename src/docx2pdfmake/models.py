"""
Configuration model for DocxConverter.
"""
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ConversionOptions:
    """Controls the converter behaviour."""

    # ── Page layout ─────────────────────────────────────────────────────────
    page_size: str = "A4"                   # "A4" | "LETTER" | "A3" …
    page_orientation: str = "portrait"      # "portrait" | "landscape"

    # Margins in pt (pdfmake default: 40 pt ≈ 14 mm)
    margin_top: float = 40.0
    margin_right: float = 40.0
    margin_bottom: float = 60.0
    margin_left: float = 40.0

    # ── Styles ──────────────────────────────────────────────────────────────
    # Carry named styles from the DOCX into the pdfmake styles block
    emit_named_styles: bool = True
    # Inline formatting (bold, italic, underline …) applied directly to the 
    # text object
    emit_inline_overrides: bool = True

    # ── Images ──────────────────────────────────────────────────────────────
    # Embed images as base64 data URLs (True) or keep as external refs (False)
    embed_images: bool = True
    # Maximum image width in pt (None = no limit)
    max_image_width: Optional[float] = 500.0
    # Maximum image height in pt (None = no limit)
    max_image_height: Optional[float] = 700.0

    # ── Tables ──────────────────────────────────────────────────────────────
    # Default table width in pt (None = automatic via columnWidths)
    default_table_width: Optional[float] = None

    # ── Typography ──────────────────────────────────────────────────────────
    default_font: str = "Roboto"
    default_font_size: float = 11.0
    default_line_height: float = 1.2

    # ── Headings (Heading 1–6) ──────────────────────────────────────────────
    heading_font_sizes: list = field(
        default_factory=lambda: [22, 18, 16, 14, 13, 12]
    )
    heading_bold: list = field(
        default_factory=lambda: [True, True, True, True, True, True]
    )
    heading_color: list = field(
        default_factory=lambda: ["#000000"] * 6
    )
    heading_margin_before: list = field(
        default_factory=lambda: [16, 14, 12, 10, 8, 6]
    )
    heading_margin_after: list = field(
        default_factory=lambda: [8, 6, 6, 4, 4, 4]
    )

    # ── Header / Footer ─────────────────────────────────────────────────────
    include_header: bool = True
    include_footer: bool = True
