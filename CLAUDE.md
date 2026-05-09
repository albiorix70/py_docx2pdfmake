# docx2pdfmake — Claude Code Context

## What this project does

Converts `.docx` files to **pdfmake Document Definition Objects (DDO)** — a JSON structure that the pdfmake JavaScript library uses to generate PDFs. The output is JSON that can be used server-side (Node.js) or client-side in a browser.

## Project layout

```
py_docx2pdfmake/
├── src/
│   └── docx2pdfmake/       # installable package
│       ├── __init__.py     # public API: DocxConverter, ConversionOptions
│       ├── __main__.py     # python -m docx2pdfmake entry point
│       ├── cli.py          # argparse CLI
│       ├── converter.py    # DocxConverter — main entry point
│       ├── models.py       # ConversionOptions dataclass
│       ├── style_resolver.py  # DOCX styles → pdfmake styles block + inline lookup
│       ├── image_handler.py   # image extraction + base64 encoding
│       ├── content_builder.py # paragraphs, lists, tables → pdfmake nodes
│       └── header_footer.py   # header/footer extraction
├── tests/
│   ├── test_models.py
│   └── test_converter.py
├── pyproject.toml          # hatchling build, pytest config
├── .gitignore
├── README.md
└── LICENSE
```

## Architecture

The conversion pipeline is:

```
DocxConverter.convert(source)
  └── StyleResolver(doc, opts)      — builds style name/props cache
  └── ImageHandler(doc, opts)       — pre-loads all images as base64
  └── ContentBuilder(doc, sr, ih, opts)
        ├── _process_body()          — iterates body children
        ├── _process_paragraph()     — text, headings, images
        ├── _collect_list()          — groups consecutive list items
        ├── _build_list_tree()       — builds ul/ol pdfmake nodes
        └── _process_table()        — table → pdfmake table node
  └── HeaderFooterExtractor(doc, opts)
        └── returns { static, fn_body } for JS usage
```

The final DDO dict contains: `content`, `pageSize`, `pageOrientation`, `pageMargins`, `defaultStyle`, optionally `styles`, `header`, `footer`, `_header_fn_body`, `_footer_fn_body`.

## Key design decisions

- **`src/` layout** — the package lives under `src/docx2pdfmake/` to prevent accidental imports of the source tree.
- **No external dependencies beyond `python-docx`** — keeps installation trivial.
- **`ConversionOptions` is a dataclass** — simple, typed, no validation overhead.
- **`_header_fn_body` / `_footer_fn_body`** — raw JS strings for page-number functions; the caller embeds them with `new Function(...)` in pdfmake.
- **`embed_images=True` by default** — base64 data URLs make the DDO self-contained.

## Development setup

```bash
# Create and activate virtual environment
python -m venv .env
.env\Scripts\activate          # Windows
# source .env/bin/activate     # Linux/macOS

# Install in editable mode + test dependencies
pip install -e ".[dev]"
# or, without extras:
pip install -e . && pip install pytest

# Run tests
pytest

# Run CLI
python -m docx2pdfmake input.docx
python -m docx2pdfmake input.docx -o output.json
```

## Known limitations

- Nested tables (table inside a table cell) are flattened
- Text boxes / Shapes (DrawingML) are not converted
- Equations (OMML) are skipped
- Custom numbering with complex overrides may fall back to `ul`
- Bidi text (Arabic, Hebrew) is passed through but not explicitly mirrored

## Coding conventions

- Python 3.9+ (`from __future__ import annotations` used throughout)
- No comments except where the *why* is non-obvious
- Type hints on all public methods
- Private helpers prefixed with `_`
- Module-level helper functions (not methods) at the bottom of each file
