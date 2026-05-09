"""Tests for ConversionOptions defaults and validation."""
import pytest
from docx2pdfmake import ConversionOptions


def test_defaults():
    opts = ConversionOptions()
    assert opts.page_size == "A4"
    assert opts.page_orientation == "portrait"
    assert opts.embed_images is True
    assert opts.emit_named_styles is True
    assert opts.emit_inline_overrides is True
    assert opts.default_font == "Roboto"
    assert opts.default_font_size == 11.0
    assert opts.default_line_height == 1.2
    assert opts.include_header is True
    assert opts.include_footer is True


def test_heading_sizes_length():
    opts = ConversionOptions()
    assert len(opts.heading_font_sizes) == 6
    assert len(opts.heading_bold) == 6
    assert len(opts.heading_color) == 6
    assert len(opts.heading_margin_before) == 6
    assert len(opts.heading_margin_after) == 6


def test_custom_options():
    opts = ConversionOptions(
        page_size="LETTER",
        page_orientation="landscape",
        embed_images=False,
        default_font="Arial",
    )
    assert opts.page_size == "LETTER"
    assert opts.page_orientation == "landscape"
    assert opts.embed_images is False
    assert opts.default_font == "Arial"
