"""
converter.py
------------
Main entry point for DOCX → pdfmake DDO conversion.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, BinaryIO, Union

from docx import Document

from .content_builder import ContentBuilder
from .header_footer import HeaderFooterExtractor
from .image_handler import ImageHandler
from .models import ConversionOptions
from .style_resolver import StyleResolver


class DocxConverter:
    """
    Converts DOCX documents to pdfmake Document Definition Objects (DDO).

    Example::

        from docx2pdfmake import DocxConverter, ConversionOptions

        opts = ConversionOptions(
            page_size="A4",
            embed_images=True,
            default_font="Roboto",
        )
        conv = DocxConverter(opts)
        ddo = conv.convert("report.docx")

        import json
        with open("report_ddo.json", "w", encoding="utf-8") as f:
            json.dump(ddo, f, indent=2, ensure_ascii=False)
    """

    def __init__(self, options: ConversionOptions | None = None):
        self._opts = options or ConversionOptions()

    # ── Public API ────────────────────────────────────────────────────────────

    def convert(
        self,
        source: Union[str, Path, BinaryIO],
        *,
        options: ConversionOptions | None = None,
    ) -> dict[str, Any]:
        """
        Converts a DOCX document.

        Parameters
        ----------
        source:
            File path (str/Path) or an open Binary-IO object.
        options:
            Optional settings; override constructor options.

        Returns
        -------
        dict
            Complete pdfmake DDO (JSON-serializable).
        """
        opts = options or self._opts
        doc = Document(source)

        # Initialize sub-components
        sr = StyleResolver(doc, opts)
        ih = ImageHandler(doc, opts)
        cb = ContentBuilder(doc, sr, ih, opts)
        hfe = HeaderFooterExtractor(doc, opts)

        # Build content
        content = cb.build()

        # Styles block
        styles = sr.pdfmake_styles_block() if opts.emit_named_styles else {}

        # Page layout
        page_margins = [
            opts.margin_left,
            opts.margin_top,
            opts.margin_right,
            opts.margin_bottom,
        ]

        # Assemble DDO
        ddo: dict[str, Any] = {
            "content": content,
            "pageSize": opts.page_size.upper(),
            "pageOrientation": opts.page_orientation,
            "pageMargins": page_margins,
            "defaultStyle": {
                "font": opts.default_font,
                "fontSize": opts.default_font_size,
                "lineHeight": opts.default_line_height,
            },
        }

        if styles:
            ddo["styles"] = styles

        # Header / Footer
        header_info = hfe.get_header()
        footer_info = hfe.get_footer()

        if header_info:
            ddo["header"] = header_info["static"]
            ddo["_header_fn_body"] = header_info["fn_body"]  # usable from JS

        if footer_info:
            ddo["footer"] = footer_info["static"]
            ddo["_footer_fn_body"] = footer_info["fn_body"]

        return ddo

    def convert_to_json(
        self,
        source: Union[str, Path, BinaryIO],
        *,
        options: ConversionOptions | None = None,
        indent: int = 2,
        ensure_ascii: bool = False,
    ) -> str:
        """Like :meth:`convert`, but returns a JSON string directly."""
        ddo = self.convert(source, options=options)
        return json.dumps(ddo, indent=indent, ensure_ascii=ensure_ascii)

    def convert_to_file(
        self,
        source: Union[str, Path, BinaryIO],
        output_path: Union[str, Path],
        *,
        options: ConversionOptions | None = None,
        indent: int = 2,
        ensure_ascii: bool = False,
    ) -> Path:
        """
        Converts and writes the result directly to a JSON file.

        Returns
        -------
        Path
            Path to the written JSON file.
        """
        output = Path(output_path)
        json_str = self.convert_to_json(
            source,
            options=options,
            indent=indent,
            ensure_ascii=ensure_ascii,
        )
        output.write_text(json_str, encoding="utf-8")
        return output
