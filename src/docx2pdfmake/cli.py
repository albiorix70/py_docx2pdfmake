"""
cli.py
------
Command-line interface for docx2pdfmake.

Usage::

    python -m docx2pdfmake input.docx                  # JSON to stdout
    python -m docx2pdfmake input.docx -o output.json   # write to file
    python -m docx2pdfmake input.docx --page-size LETTER --no-embed-images
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .converter import DocxConverter
from .models import ConversionOptions


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="docx2pdfmake",
        description="Converts DOCX files to pdfmake DDO (JSON).",
    )
    parser.add_argument("input", help="Path to the DOCX file")
    parser.add_argument(
        "-o", "--output",
        help="Output file (default: stdout)",
        default=None,
    )
    parser.add_argument("--page-size", default="A4", help="Page size (default: A4)")
    parser.add_argument(
        "--page-orientation",
        choices=["portrait", "landscape"],
        default="portrait",
    )
    parser.add_argument(
        "--no-embed-images",
        action="store_true",
        help="Do not embed images as base64",
    )
    parser.add_argument(
        "--no-named-styles",
        action="store_true",
        help="Do not emit a styles block",
    )
    parser.add_argument(
        "--no-inline-overrides",
        action="store_true",
        help="Do not apply inline formatting",
    )
    parser.add_argument(
        "--default-font",
        default="Roboto",
        help="Default font (default: Roboto)",
    )
    parser.add_argument(
        "--default-font-size",
        type=float,
        default=11.0,
        help="Default font size in pt (default: 11)",
    )
    parser.add_argument(
        "--indent",
        type=int,
        default=2,
        help="JSON indentation (default: 2)",
    )

    args = parser.parse_args(argv)

    opts = ConversionOptions(
        page_size=args.page_size,
        page_orientation=args.page_orientation,
        embed_images=not args.no_embed_images,
        emit_named_styles=not args.no_named_styles,
        emit_inline_overrides=not args.no_inline_overrides,
        default_font=args.default_font,
        default_font_size=args.default_font_size,
    )

    conv = DocxConverter(opts)
    try:
        json_str = conv.convert_to_json(
            args.input,
            indent=args.indent,
        )
    except FileNotFoundError:
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error during conversion: {e}", file=sys.stderr)
        sys.exit(1)

    if args.output:
        Path(args.output).write_text(json_str, encoding="utf-8")
        print(f"Written: {args.output}", file=sys.stderr)
    else:
        print(json_str)


if __name__ == "__main__":
    main()
