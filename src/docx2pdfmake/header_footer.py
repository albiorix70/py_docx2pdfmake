"""
header_footer.py
----------------
Extracts header and footer content from the DOCX document and converts them
to pdfmake-compatible header/footer function stubs.

Since pdfmake expects header/footer as JavaScript functions, two formats
are supported:
  1. ``static``  – a simple pdfmake content object (ready to use)
  2. ``fn_body`` – a string containing the body of a JS arrow function
                   (for direct insertion into pdfmake configuration)
"""

from __future__ import annotations

import logging
from typing import Any, Optional

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .models import ConversionOptions

logger = logging.getLogger(__name__)


class HeaderFooterExtractor:
    """Reads the header and footer parts of the DOCX and converts them."""

    def __init__(self, doc: Document, opts: ConversionOptions):
        self._doc = doc
        self._opts = opts

    # ── Public API ───────────────────────────────────────────────────────

    def get_header(self) -> Optional[dict]:
        """
        Returns a pdfmake header object (static) or None.
        Format: { "static": <pdfmake-node>, "fn_body": <js-string> }
        """
        if not self._opts.include_header:
            return None
        return self._extract_section("header")

    def get_footer(self) -> Optional[dict]:
        """Returns a pdfmake footer object or None."""
        if not self._opts.include_footer:
            return None
        return self._extract_section("footer")

    # ── Internal logic ───────────────────────────────────────────────────

    def _extract_section(self, section_type: str) -> Optional[dict]:
        """Generic extractor for header/footer."""
        try:
            section = self._doc.sections[0]
            if section_type == "header":
                part = section.header
            else:
                part = section.footer

            if part is None or part.is_linked_to_previous:
                return None

            paragraphs = part.paragraphs
            if not paragraphs:
                return None

            nodes = []
            for para in paragraphs:
                node = self._para_to_node(para)
                if node is not None:
                    nodes.append(node)

            if not nodes:
                return None

            # Static pdfmake object
            if len(nodes) == 1:
                static_node = nodes[0]
            else:
                static_node = {"stack": nodes}

            # JS function body (with page number placeholders)
            fn_body = self._make_fn_body(nodes, section_type)

            return {
                "static": static_node,
                "fn_body": fn_body,
            }
        except Exception:
            logger.warning("Failed to extract %s", section_type, exc_info=True)
            return None

    def _para_to_node(self, para: Paragraph) -> Optional[Any]:
        """Converts a paragraph to a pdfmake text node."""
        texts = []
        for child in para._p:
            tag = _local(child.tag)
            if tag == "r":
                for t in child.findall(qn("w:t")):
                    if t.text:
                        texts.append(t.text)
                # Page number field
                for fld in child.findall(qn("w:fldChar")):
                    pass
            elif tag == "fldSimple":
                instr = child.get(qn("w:instr"), "")
                if "PAGE" in instr:
                    texts.append("__PAGE__")
                elif "NUMPAGES" in instr:
                    texts.append("__NUMPAGES__")

        # Also instrText (complex fields)
        for instr in para._p.iter(qn("w:instrText")):
            if instr.text:
                if "PAGE" in instr.text and "__PAGE__" not in texts:
                    texts.append("__PAGE__")
                if "NUMPAGES" in instr.text and "__NUMPAGES__" not in texts:
                    texts.append("__NUMPAGES__")

        text = "".join(texts).strip()
        if not text:
            return None

        # Alignment
        alignment = self._para_alignment(para)
        node: dict[str, Any] = {
            "text": text,
            "fontSize": 9,
            "color": "#888888",
        }
        if alignment:
            node["alignment"] = alignment
        node["margin"] = [40, 10, 40, 0]
        return node

    def _make_fn_body(self, nodes: list, section_type: str) -> str:
        """
        Builds the string body of a pdfmake header/footer function.
        Page number placeholders are replaced with pdfmake currentPage
        expressions.

        Usage in pdfmake::

            const dd = {
              header: new Function('currentPage', 'pageCount', <fn_body>),
              ...
            }
        """
        import json

        serialized = json.dumps(nodes, ensure_ascii=False)
        # Replace placeholders with JS expressions
        serialized = serialized.replace(
            '"__PAGE__"', 'currentPage.toString()'
        ).replace(
            '"__NUMPAGES__"', 'pageCount.toString()'
        )
        return f"return {serialized};"

    def _para_alignment(self, para: Paragraph) -> Optional[str]:
        ppr = para._p.find(qn("w:pPr"))
        if ppr is None:
            return None
        jc = ppr.find(qn("w:jc"))
        if jc is None:
            return None
        _MAP = {
            "center": "center",
            "right": "right",
            "both": "justify",
            "left": "left",
        }
        return _MAP.get(jc.get(qn("w:val"), ""), None)


def _local(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag
