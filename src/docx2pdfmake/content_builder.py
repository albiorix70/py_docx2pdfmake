"""
content_builder.py
------------------
Converts DOCX body elements (paragraphs, tables) to pdfmake content nodes.
Supports:
  - Body text with inline formatting (bold, italic, underline, color, …)
  - Headings (Heading 1–6)
  - Unordered and ordered lists (including nesting)
  - Tables (including merged cells)
  - Inline images
"""

from __future__ import annotations

import logging
from typing import Any, Optional

from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from .image_handler import ImageHandler
from .models import ConversionOptions
from .style_resolver import StyleResolver

logger = logging.getLogger(__name__)


# ── Constants ───────────────────────────────────────────────────────────────

_ALIGNMENT_MAP = {
    "center": "center",
    "right": "right",
    "both": "justify",
    "left": "left",
}


class ContentBuilder:
    """
    Main content builder: iterates over all body elements and builds the
    pdfmake ``content`` array.
    """

    def __init__(
        self,
        doc: DocxDocument,
        style_resolver: StyleResolver,
        image_handler: ImageHandler,
        opts: ConversionOptions,
    ):
        self._doc = doc
        self._sr = style_resolver
        self._ih = image_handler
        self._opts = opts

    # ── Public API ─────────────────────────────────────────────────────────

    def build(self) -> list:
        """Returns the complete pdfmake content array."""
        content: list = []
        self._process_body(self._doc.element.body, content)
        logger.debug("Content build complete: %d node(s)", len(content))
        return content

    # ── Iteration ──────────────────────────────────────────────────────────

    def _process_body(self, body, content: list):
        """Iterates over the direct children of the body element."""
        i = 0
        children = list(body)
        while i < len(children):
            child = children[i]
            tag = _local(child.tag)

            if tag == "p":
                para = Paragraph(child, self._doc)
                # Lists are collected and processed as a group
                if self._is_list_item(para):
                    nodes, i = self._collect_list(children, i)
                    content.extend(nodes)
                    i += 1
                    continue
                node = self._process_paragraph(para)
                if node is not None:
                    content.append(node)

            elif tag == "tbl":
                node = self._process_table(Table(child, self._doc))
                if node is not None:
                    content.append(node)

            i += 1

    # ── Paragraphs ─────────────────────────────────────────────────────────

    def _process_paragraph(self, para: Paragraph) -> Optional[Any]:
        """Processes a paragraph → pdfmake node."""
        # Check for images first
        images = self._extract_images(para)
        if images:
            if len(images) == 1:
                return images[0]
            return {"stack": images, "margin": [0, 2, 0, 2]}

        # Text
        inline_nodes = self._build_inline_nodes(para)
        if not inline_nodes:
            # Empty paragraph → small spacing
            return {"text": "", "margin": [0, 2, 0, 2]}

        node = self._wrap_inline(inline_nodes, para)
        return node

    def _build_inline_nodes(self, para: Paragraph) -> list:
        """Builds a list of inline text objects for all runs."""
        nodes = []
        for child in para._p:
            tag = _local(child.tag)
            if tag == "r":
                # Regular run
                run_nodes = self._process_run(child, para)
                nodes.extend(run_nodes)
            elif tag == "hyperlink":
                # Hyperlink (contains runs)
                link_nodes = self._process_hyperlink(child, para)
                if link_nodes:
                    nodes.extend(link_nodes)
        return nodes

    def _process_run(self, r_el, para: Paragraph) -> list:
        """Processes a w:r element → list of pdfmake text nodes."""
        result = []

        # Inline images in the run
        for inline in r_el.iter(qn("wp:inline")):
            img = self._ih.get_image_node(inline)
            if img:
                result.append(img)
        for anchor in r_el.iter(qn("wp:anchor")):
            img = self._ih.get_image_node(anchor)
            if img:
                result.append(img)

        # Text
        text_parts = []
        for t in r_el.findall(qn("w:t")):
            if t.text:
                text_parts.append(t.text)
        for br in r_el.findall(qn("w:br")):
            br_type = br.get(qn("w:type"), "")
            if br_type == "page":
                result.append({"text": "", "pageBreak": "before"})
            else:
                text_parts.append("\n")
        for tab in r_el.findall(qn("w:tab")):
            text_parts.append("\t")

        text = "".join(text_parts)
        if not text:
            return result

        # Fetch inline overrides from the StyleResolver
        # Simulate a run proxy
        class _RunProxy:
            def __init__(self, r):
                self._r = r

        run_proxy = _RunProxy(r_el)
        props = self._sr.resolve_run_props(run_proxy)

        if props:
            node = {"text": text, **props}
        else:
            node = text  # a plain string is valid in pdfmake

        result.append(node)
        return result

    def _process_hyperlink(self, hl_el, para: Paragraph) -> list:
        """Processes a w:hyperlink node."""
        nodes = []
        for r_el in hl_el.findall(qn("w:r")):
            nodes.extend(self._process_run(r_el, para))

        # Retrieve URL from relationship
        r_id = hl_el.get(qn("r:id"))
        url = None
        if r_id:
            try:
                url = para.part.rels[r_id].target_ref
            except (KeyError, AttributeError):
                pass

        if not nodes:
            return []

        if url:
            text_content = _flatten_text(nodes)
            link_node = {
                "text": text_content,
                "link": url,
                "color": "#0563C1",
                "decoration": "underline",
            }
            return [link_node]
        return nodes

    def _wrap_inline(self, inline_nodes: list, para: Paragraph) -> dict:
        """Wraps inline nodes in the appropriate pdfmake paragraph node."""
        node: dict[str, Any] = {}

        # Style
        pdf_style = self._sr.resolve_paragraph_style(para)
        if pdf_style:
            node["style"] = pdf_style

        # Heading handling
        lvl = self._sr.heading_level(para)
        if lvl > 0:
            node["style"] = f"heading{lvl}"

        # Alignment (from paragraph XML, overrides style)
        alignment = self._para_alignment(para)
        if alignment:
            node["alignment"] = alignment

        # Text
        if len(inline_nodes) == 1 and isinstance(inline_nodes[0], str):
            node["text"] = inline_nodes[0]
        else:
            node["text"] = inline_nodes

        # Margin from paragraph spacing
        margin = self._para_margin(para)
        if margin:
            node["margin"] = margin

        return node

    # ── Lists ──────────────────────────────────────────────────────────────

    def _is_list_item(self, para: Paragraph) -> bool:
        """Returns True if the paragraph is a list item."""
        ppr = para._p.find(qn("w:pPr"))
        if ppr is None:
            return False
        return ppr.find(qn("w:numPr")) is not None

    def _get_list_props(self, para: Paragraph) -> tuple[int, int]:
        """Returns (ilvl, numId)."""
        ppr = para._p.find(qn("w:pPr"))
        if ppr is None:
            return 0, 0
        num_pr = ppr.find(qn("w:numPr"))
        if num_pr is None:
            return 0, 0
        ilvl_el = num_pr.find(qn("w:ilvl"))
        num_id_el = num_pr.find(qn("w:numId"))
        ilvl = int(ilvl_el.get(qn("w:val"), "0")) if ilvl_el is not None else 0
        num_id = (
            int(num_id_el.get(qn("w:val"), "0"))
            if num_id_el is not None
            else 0
        )
        return ilvl, num_id

    def _is_ordered(self, num_id: int, ilvl: int) -> bool:
        """Returns True when the list is ordered (ol)."""
        try:
            numbering = self._doc.part.numbering_part
            if numbering is None:
                return False
            # Find abstractNumId
            num_el = numbering._element.find(
                f'.//{qn("w:num")}[@{qn("w:numId")}="{num_id}"]'
            )
            if num_el is None:
                # Fallback: search by w:numId value
                for n in numbering._element.findall(qn("w:num")):
                    nid = n.find(qn("w:numId"))
                    if nid is not None and nid.get(qn("w:val")) == str(num_id):
                        num_el = n
                        break
            if num_el is None:
                return False

            abs_ref = num_el.find(qn("w:abstractNumId"))
            if abs_ref is None:
                return False
            abs_id = abs_ref.get(qn("w:val"))

            abs_xpath = (
                f'.//{qn("w:abstractNum")}'
                f'[@{qn("w:abstractNumId")}="{abs_id}"]'
            )
            abs_num = numbering._element.find(abs_xpath)
            if abs_num is None:
                for an in numbering._element.findall(qn("w:abstractNum")):
                    aid_el = an.find(qn("w:abstractNumId"))
                    if (
                        aid_el is not None
                        and aid_el.get(qn("w:val")) == abs_id
                    ):
                        abs_num = an
                        break
            if abs_num is None:
                return False

            lvl_el = abs_num.find(
                f'.//{qn("w:lvl")}[@{qn("w:ilvl")}="{ilvl}"]'
            )
            if lvl_el is None:
                return False

            num_fmt = lvl_el.find(qn("w:numFmt"))
            if num_fmt is None:
                return False
            val = num_fmt.get(qn("w:val"), "bullet")
            return val not in ("bullet", "none")
        except Exception:
            return False

    def _collect_list(self, children: list, start: int) -> tuple[list, int]:
        """
        Collects consecutive list items starting at index ``start`` and builds
        a nested pdfmake list tree.
        Returns (nodes, next_index).
        """
        i = start
        # Collect all consecutive list paragraphs
        items: list[tuple[int, int, Paragraph]] = []  # (ilvl, numId, para)
        while i < len(children):
            child = children[i]
            if _local(child.tag) != "p":
                break
            para = Paragraph(child, self._doc)
            if not self._is_list_item(para):
                break
            ilvl, num_id = self._get_list_props(para)
            items.append((ilvl, num_id, para))
            i += 1

        # Items → nested tree
        nodes = self._build_list_tree(items)
        return nodes, i - 1  # -1 because the caller increments i

    def _build_list_tree(
        self, items: list[tuple[int, int, Paragraph]]
    ) -> list:
        """Builds a pdfmake-compatible list node (possibly nested)."""
        if not items:
            return []

        # All entries at level 0 → a single ul/ol node
        # Nesting is written as sub-lists inside the items array
        root_nodes = []
        i = 0
        while i < len(items):
            ilvl, num_id, para = items[i]
            if ilvl > 0:
                # Skip — consumed by parent item
                i += 1
                continue

            # Entry text
            inline_nodes = self._build_inline_nodes(para)
            if len(inline_nodes) == 1 and isinstance(inline_nodes[0], str):
                text = _flatten_text(inline_nodes)
            else:
                text = inline_nodes

            # Collect children (deeper levels)
            children_items = []
            j = i + 1
            while j < len(items) and items[j][0] > ilvl:
                children_items.append(items[j])
                j += 1

            if children_items:
                # Children as sub-list
                sub_nodes = self._build_list_tree_recursive(
                    children_items, ilvl + 1
                )
                item_content = [{"text": text}]
                item_content.extend(sub_nodes)
                root_nodes.append(item_content)
            else:
                root_nodes.append({"text": text})
            i = j if children_items else i + 1

        if not root_nodes:
            return []

        # Type: ordered or unordered (from the first item)
        first_num_id = items[0][1]
        is_ord_root = self._is_ordered(first_num_id, 0)
        key = "ol" if is_ord_root else "ul"
        return [{key: root_nodes, "margin": [0, 2, 0, 2]}]

    def _build_list_tree_recursive(
        self, items: list, expected_ilvl: int
    ) -> list:
        """Recursive builder for nested lists."""
        if not items:
            return []

        root_nodes = []
        i = 0
        while i < len(items):
            ilvl, num_id, para = items[i]
            if ilvl != expected_ilvl:
                i += 1
                continue

            inline_nodes = self._build_inline_nodes(para)
            if len(inline_nodes) != 1:
                text = inline_nodes
            elif isinstance(inline_nodes[0], str):
                text = inline_nodes[0]
            else:
                text = inline_nodes

            # Collect children
            children_items = []
            j = i + 1
            while j < len(items) and items[j][0] > ilvl:
                children_items.append(items[j])
                j += 1

            if children_items:
                sub = self._build_list_tree_recursive(children_items, ilvl + 1)
                root_nodes.append(
                    {"text": text, "stack": sub} if sub else {"text": text}
                )
            else:
                root_nodes.append({"text": text})
            i = j if children_items else i + 1

        if not root_nodes:
            return []
        first_num_id = items[0][1]
        is_ord_root = self._is_ordered(first_num_id, expected_ilvl)
        key = "ol" if is_ord_root else "ul"
        return [{key: root_nodes}]

    # ── Tables ─────────────────────────────────────────────────────────────

    def _process_table(self, table: Table) -> Optional[dict]:
        """Processes a table → pdfmake table node."""
        try:
            body = []
            for row in table.rows:
                pdf_row = []
                for cell in row.cells:
                    pdf_cell = self._process_cell(cell)
                    pdf_row.append(pdf_cell)
                body.append(pdf_row)

            if not body:
                return None

            col_count = max(len(r) for r in body)
            # Equal column widths ('*' lets pdfmake auto-calculate)
            col_widths = ["*"] * col_count

            return {
                "table": {
                    "headerRows": 1,
                    "widths": col_widths,
                    "body": body,
                },
                "layout": "lightHorizontalLines",
                "margin": [0, 6, 0, 6],
            }
        except Exception:
            logger.warning("Failed to process table", exc_info=True)
            return None

    def _process_cell(self, cell: _Cell) -> dict:
        """Processes a table cell → pdfmake cell object."""
        if cell._tc.find(qn("w:tbl")) is not None:
            logger.warning("Nested table found in cell — flattening (unsupported)")

        stack = []
        for para in cell.paragraphs:
            if self._is_list_item(para):
                # Simplify list rendering inside cells
                inline = self._build_inline_nodes(para)
                text = (
                    _flatten_text(inline)
                    if all(isinstance(n, str) for n in inline)
                    else inline
                )
                stack.append({"text": text, "margin": [2, 1, 2, 1]})
            else:
                node = self._process_paragraph(para)
                if node:
                    stack.append(node)

        if not stack:
            return {"text": ""}
        if len(stack) == 1:
            return stack[0]
        return {"stack": stack}

    # ── Images ─────────────────────────────────────────────────────────────

    def _extract_images(self, para: Paragraph) -> list:
        """Extracts all images from a paragraph."""
        images = []
        for inline in para._p.iter(qn("wp:inline")):
            img = self._ih.get_image_node(inline)
            if img:
                images.append(img)
        for anchor in para._p.iter(qn("wp:anchor")):
            img = self._ih.get_image_node(anchor)
            if img:
                images.append(img)
        return images

    # ── Helpers ────────────────────────────────────────────────────────────

    def _para_alignment(self, para: Paragraph) -> Optional[str]:
        ppr = para._p.find(qn("w:pPr"))
        if ppr is None:
            return None
        jc = ppr.find(qn("w:jc"))
        if jc is None:
            return None
        return _ALIGNMENT_MAP.get(jc.get(qn("w:val"), ""), None)

    def _para_margin(self, para: Paragraph) -> Optional[list]:
        ppr = para._p.find(qn("w:pPr"))
        if ppr is None:
            return None
        spacing = ppr.find(qn("w:spacing"))
        if spacing is None:
            return None
        before = spacing.get(qn("w:before"))
        after = spacing.get(qn("w:after"))
        if not before and not after:
            return None
        try:
            mb = round(int(before) / 20.0, 1) if before else 0
            ma = round(int(after) / 20.0, 1) if after else 0
            return [0, mb, 0, ma]
        except ValueError:
            return None


# ── Helper functions ───────────────────────────────────────────────────────

def _local(tag: str) -> str:
    """Returns the local part of a Clark-notation tag name."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _flatten_text(nodes: list) -> str:
    """Flattens a list of strings / dicts into a single string."""
    parts = []
    for n in nodes:
        if isinstance(n, str):
            parts.append(n)
        elif isinstance(n, dict) and "text" in n:
            t = n["text"]
            parts.append(_flatten_text(t) if isinstance(t, list) else str(t))
    return "".join(parts)
