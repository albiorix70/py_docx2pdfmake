"""
style_resolver.py
-----------------
Reads all styles from the DOCX document and produces:
  1. A pdfmake-compatible ``styles`` block (dict)
  2. An internal lookup cache for inline overrides
"""

from __future__ import annotations

import logging
import re
from typing import Any, Optional

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from .models import ConversionOptions

logger = logging.getLogger(__name__)


# ── Helper functions ──────────────────────────────────────────────────

def _emu_to_pt(emu: int) -> float:
    """English Metric Units → Points (1 pt = 12700 EMU)."""
    return emu / 12700.0


def _twip_to_pt(twip: int) -> float:
    """Twips → Points (1 pt = 20 twip)."""
    return twip / 20.0


def _half_pt_to_pt(half: int) -> float:
    """Half-points → Points."""
    return half / 2.0


def _rgb_to_hex(r: int, g: int, b: int) -> str:
    return f"#{r:02X}{g:02X}{b:02X}"


def _parse_color(color_str: Optional[str]) -> Optional[str]:
    """Converts 'RRGGBB' (DOCX format) to '#RRGGBB'."""
    if not color_str or color_str.upper() in ("AUTO", "NONE", ""):
        return None
    color_str = color_str.strip("#")
    if re.fullmatch(r"[0-9A-Fa-f]{6}", color_str):
        return f"#{color_str.upper()}"
    return None


def _is_truthy_flag(el) -> bool:
    """True unless a w:b/w:i/w:strike element explicitly disables itself."""
    return el.get(qn("w:val"), "true").lower() not in ("0", "false")


# ── Main class ────────────────────────────────────────────────────────

class StyleResolver:
    """
    Reads paragraph and character styles from the DOCX document and
    produces pdfmake style definitions.
    """

    # DOCX heading style names → pdfmake style name
    HEADING_MAP = {
        "Heading 1": "heading1",
        "Heading 2": "heading2",
        "Heading 3": "heading3",
        "Heading 4": "heading4",
        "Heading 5": "heading5",
        "Heading 6": "heading6",
        # German locale variants
        "berschrift 1": "heading1",
        "berschrift 2": "heading2",
        "berschrift 3": "heading3",
        "berschrift 4": "heading4",
        "berschrift 5": "heading5",
        "berschrift 6": "heading6",
    }

    def __init__(self, doc: Document, opts: ConversionOptions):
        self._doc = doc
        self._opts = opts
        # style_id → pdfmake style name
        self._id_to_name: dict[str, str] = {}
        # style_id → pdfmake properties (dict)
        self._id_to_props: dict[str, dict] = {}
        self._build()

    # ── Public API ───────────────────────────────────────────────────────

    def pdfmake_styles_block(self) -> dict[str, dict]:
        """Returns the complete pdfmake ``styles`` block."""
        result: dict[str, dict] = {}
        for sid, name in self._id_to_name.items():
            props = self._id_to_props.get(sid, {})
            if props:
                result[name] = props
        # Always include heading defaults from ConversionOptions
        result.update(self._heading_defaults())
        return result

    def resolve_paragraph_style(self, para) -> Optional[str]:
        """Returns the pdfmake style name for a paragraph (or None)."""
        style = para.style
        if style is None:
            return None
        return self._id_to_name.get(style.style_id)

    def resolve_run_props(self, run) -> dict[str, Any]:
        """
        Returns a dict of pdfmake inline properties for a run.
        Only properties that are explicitly set (overrides).
        """
        if not self._opts.emit_inline_overrides:
            return {}

        props: dict[str, Any] = {}
        rpr = run._r.find(qn("w:rPr"))
        if rpr is None:
            return props

        # bold
        b = rpr.find(qn("w:b"))
        if b is not None and _is_truthy_flag(b):
            props["bold"] = True

        # italic
        i = rpr.find(qn("w:i"))
        if i is not None and _is_truthy_flag(i):
            props["italics"] = True

        # underline
        u = rpr.find(qn("w:u"))
        if u is not None:
            val = u.get(qn("w:val"), "none")
            if val not in ("none", "0"):
                props["decoration"] = "underline"

        # strikethrough
        strike = rpr.find(qn("w:strike"))
        if strike is not None and _is_truthy_flag(strike):
            props["decoration"] = "lineThrough"

        # font size (half-points)
        sz = rpr.find(qn("w:sz"))
        if sz is not None:
            try:
                raw_sz = int(sz.get(qn("w:val"), "0"))
                props["fontSize"] = _half_pt_to_pt(raw_sz)
            except ValueError:
                pass

        # color
        color_el = rpr.find(qn("w:color"))
        if color_el is not None:
            hex_color = _parse_color(color_el.get(qn("w:val")))
            if hex_color:
                props["color"] = hex_color

        # highlight
        hl = rpr.find(qn("w:highlight"))
        if hl is not None:
            color_name = hl.get(qn("w:val"), "")
            pdf_color = _highlight_to_hex(color_name)
            if pdf_color:
                props["background"] = pdf_color

        return props

    def is_heading(self, para) -> bool:
        """True if the paragraph uses a heading style."""
        style = para.style
        if style is None:
            return False
        return self._id_to_name.get(style.style_id, "").startswith("heading")

    def heading_level(self, para) -> int:
        """Returns 1–6 (or 0 if not a heading)."""
        style = para.style
        if style is None:
            return 0
        name = self._id_to_name.get(style.style_id, "")
        if name.startswith("heading") and len(name) == 8:
            try:
                return int(name[-1])
            except ValueError:
                pass
        return 0

    # ── Internal build logic ──────────────────────────────────────────────

    def _build(self):
        """Iterates over all document styles and populates the cache."""
        for style in self._doc.styles:
            sid = style.style_id
            # Determine pdfmake name
            pdf_name = self._pdf_name_for(style)
            self._id_to_name[sid] = pdf_name

            if not self._opts.emit_named_styles:
                continue

            props = self._extract_para_props(style)
            props.update(self._extract_run_props_from_style(style))
            if props:
                self._id_to_props[sid] = props
        logger.debug("Resolved %d DOCX style(s)", len(self._id_to_name))

    def _pdf_name_for(self, style) -> str:
        """Computes a clean pdfmake style name."""
        raw = style.name or style.style_id or "unknown"
        # Heading special cases
        for key, val in self.HEADING_MAP.items():
            if key in raw:
                return val
        # General: CamelCase → camelCase, strip spaces and special characters
        clean = re.sub(r"[^a-zA-Z0-9_ ]", "", raw)
        parts = clean.split()
        if not parts:
            return "style_" + style.style_id
        head = parts[0][0].lower() + parts[0][1:]
        tail = "".join(p.capitalize() for p in parts[1:])
        return head + tail

    def _extract_para_props(self, style) -> dict[str, Any]:
        """Extracts paragraph properties (pPr) from a style."""
        props: dict[str, Any] = {}
        try:
            ppr = style.element.find(qn("w:pPr"))
        except Exception:
            return props
        if ppr is None:
            return props

        # Alignment
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            alignment_map = {
                "center": "center",
                "right": "right",
                "both": "justify",
                "left": "left",
            }
            val = jc.get(qn("w:val"), "")
            if val in alignment_map:
                props["alignment"] = alignment_map[val]

        # Spacing
        spacing = ppr.find(qn("w:spacing"))
        if spacing is not None:
            before = spacing.get(qn("w:before"))
            after = spacing.get(qn("w:after"))
            line = spacing.get(qn("w:line"))
            if before:
                try:
                    props["marginTop"] = _twip_to_pt(int(before))
                except ValueError:
                    pass
            if after:
                try:
                    props["marginBottom"] = _twip_to_pt(int(after))
                except ValueError:
                    pass
            if line:
                try:
                    # lineRule="auto" → value is in 240ths
                    props["lineHeight"] = round(int(line) / 240.0, 2)
                except ValueError:
                    pass

        # Indent
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            left = ind.get(qn("w:left"))
            if left:
                try:
                    props["marginLeft"] = _twip_to_pt(int(left))
                except ValueError:
                    pass

        return props

    def _extract_run_props_from_style(self, style) -> dict[str, Any]:
        """Extracts character run properties (rPr) from a style."""
        props: dict[str, Any] = {}
        try:
            rpr = style.element.find(qn("w:rPr"))
        except Exception:
            return props
        if rpr is None:
            return props

        b = rpr.find(qn("w:b"))
        if b is not None and _is_truthy_flag(b):
            props["bold"] = True

        i = rpr.find(qn("w:i"))
        if i is not None and _is_truthy_flag(i):
            props["italics"] = True

        sz = rpr.find(qn("w:sz"))
        if sz is not None:
            try:
                raw_sz = int(sz.get(qn("w:val"), "0"))
                props["fontSize"] = _half_pt_to_pt(raw_sz)
            except ValueError:
                pass

        color_el = rpr.find(qn("w:color"))
        if color_el is not None:
            hex_color = _parse_color(color_el.get(qn("w:val")))
            if hex_color:
                props["color"] = hex_color

        return props

    def _heading_defaults(self) -> dict[str, dict]:
        """Builds heading style definitions from ConversionOptions."""
        opts = self._opts
        result = {}
        for i in range(6):
            lvl = i + 1
            style_def: dict[str, Any] = {
                "fontSize": opts.heading_font_sizes[i],
                "bold": opts.heading_bold[i],
            }
            color = opts.heading_color[i]
            if color and color.upper() != "#000000":
                style_def["color"] = color
            mb = opts.heading_margin_before[i]
            ma = opts.heading_margin_after[i]
            if mb or ma:
                style_def["margin"] = [0, mb, 0, ma]
            result[f"heading{lvl}"] = style_def
        return result


# ── Highlight colours ─────────────────────────────────────────────────

_HIGHLIGHT_COLORS = {
    "yellow": "#FFFF00",
    "green": "#00FF00",
    "cyan": "#00FFFF",
    "magenta": "#FF00FF",
    "blue": "#0000FF",
    "red": "#FF0000",
    "darkBlue": "#000080",
    "darkCyan": "#008080",
    "darkGreen": "#008000",
    "darkMagenta": "#800080",
    "darkRed": "#800000",
    "darkYellow": "#808000",
    "darkGray": "#808080",
    "lightGray": "#C0C0C0",
    "black": "#000000",
    "white": "#FFFFFF",
}


def _highlight_to_hex(name: str) -> Optional[str]:
    return _HIGHLIGHT_COLORS.get(name)
