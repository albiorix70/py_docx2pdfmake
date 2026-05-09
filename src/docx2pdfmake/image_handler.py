"""
image_handler.py
----------------
Extracts images from the DOCX and returns them as base64 data URLs
or path references.
"""

from __future__ import annotations

import base64
import mimetypes
from typing import Optional

from docx import Document
from docx.oxml.ns import qn

from .models import ConversionOptions


class ImageHandler:
    """Manages all document images and provides pdfmake image objects."""

    def __init__(self, doc: Document, opts: ConversionOptions):
        self._doc = doc
        self._opts = opts
        # rId → base64 data URL (or None on error)
        self._cache: dict[str, Optional[str]] = {}
        self._build_cache()

    # ── Public API ────────────────────────────────────────────────────────────

    def get_image_node(self, inline_or_anchor) -> Optional[dict]:
        """
        Expects a ``<wp:inline>`` or ``<wp:anchor>`` element and returns
        a pdfmake image object (or None on error).
        """
        try:
            # Determine relationship ID
            blip = inline_or_anchor.find(".//" + qn("a:blip"))
            if blip is None:
                return None
            r_embed = blip.get(qn("r:embed"))
            if not r_embed:
                return None

            # Dimensions (EMU → pt)
            extent = inline_or_anchor.find(qn("wp:extent"))
            width_pt = height_pt = None
            if extent is not None:
                cx = extent.get("cx")
                cy = extent.get("cy")
                if cx:
                    width_pt = int(cx) / 12700.0
                if cy:
                    height_pt = int(cy) / 12700.0

            # Clamp size
            width_pt, height_pt = self._clamp_size(width_pt, height_pt)

            data_url = self._cache.get(r_embed)
            if not data_url:
                return None

            node: dict = {"image": data_url}
            if width_pt:
                node["width"] = round(width_pt, 2)
            if height_pt:
                node["height"] = round(height_pt, 2)
            # Centered by default
            node["alignment"] = "center"
            node["margin"] = [0, 4, 0, 4]
            return node
        except Exception:
            return None

    # ── Internal logic ────────────────────────────────────────────────────────

    def _build_cache(self):
        """Reads all image relationships and populates the cache."""
        try:
            rels = self._doc.part.rels
            for r_id, rel in rels.items():
                if "image" not in rel.reltype:
                    continue
                try:
                    img_part = rel.target_part
                    mime = img_part.content_type or "image/png"
                    data = img_part.blob
                    b64 = base64.b64encode(data).decode("ascii")
                    self._cache[r_id] = f"data:{mime};base64,{b64}"
                except Exception:
                    self._cache[r_id] = None
        except Exception:
            pass

    def _clamp_size(
        self,
        w: Optional[float],
        h: Optional[float],
    ) -> tuple[Optional[float], Optional[float]]:
        """Clamps image dimensions to the configured maximum values."""
        max_w = self._opts.max_image_width
        max_h = self._opts.max_image_height

        if w and max_w and w > max_w:
            if h:
                h = h * (max_w / w)
            w = max_w

        if h and max_h and h > max_h:
            if w:
                w = w * (max_h / h)
            h = max_h

        return w, h
