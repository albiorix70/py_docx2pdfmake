"""Integration tests for DocxConverter (require python-docx)."""
import io
import json
import pytest

from docx2pdfmake import DocxConverter, ConversionOptions


def _make_minimal_docx() -> io.BytesIO:
    """Creates an in-memory DOCX with a single paragraph."""
    from docx import Document as DocxDocument

    doc = DocxDocument()
    doc.add_paragraph("Hello, pdfmake!")
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def test_convert_returns_dict():
    conv = DocxConverter()
    ddo = conv.convert(_make_minimal_docx())
    assert isinstance(ddo, dict)
    assert "content" in ddo
    assert "pageSize" in ddo
    assert "pageMargins" in ddo
    assert "defaultStyle" in ddo


def test_convert_page_size_uppercase():
    opts = ConversionOptions(page_size="a4")
    conv = DocxConverter(opts)
    ddo = conv.convert(_make_minimal_docx())
    assert ddo["pageSize"] == "A4"


def test_convert_to_json_is_valid():
    conv = DocxConverter()
    json_str = conv.convert_to_json(_make_minimal_docx())
    parsed = json.loads(json_str)
    assert isinstance(parsed, dict)


def test_convert_to_file(tmp_path):
    conv = DocxConverter()
    out = tmp_path / "output.json"
    result = conv.convert_to_file(_make_minimal_docx(), out)
    assert result == out
    assert out.exists()
    parsed = json.loads(out.read_text(encoding="utf-8"))
    assert "content" in parsed


def test_no_styles_block_when_disabled():
    opts = ConversionOptions(emit_named_styles=False)
    conv = DocxConverter(opts)
    ddo = conv.convert(_make_minimal_docx())
    assert "styles" not in ddo


def test_content_is_list():
    conv = DocxConverter()
    ddo = conv.convert(_make_minimal_docx())
    assert isinstance(ddo["content"], list)
