# docx2pdfmake

Python library that converts `.docx` files into **pdfmake Document Definition Objects (DDO)** — the JSON structure consumed by the [pdfmake](https://pdfmake.github.io/docs/) JavaScript library to generate PDFs.

## Features

| Feature | Status |
|---|---|
| Body text (bold, italic, underline, color, highlight) | ✅ |
| Headings 1–6 | ✅ |
| Unordered & ordered lists (ul/ol, nested) | ✅ |
| Tables (including header row, merged cells) | ✅ |
| Inline images (base64 embedded) | ✅ |
| Page layout (size, orientation, margins) | ✅ |
| Header / footer (including page-number placeholder) | ✅ |
| Named styles from DOCX → pdfmake `styles` block | ✅ |
| Inline overrides (formatting from run properties) | ✅ |
| Hyperlinks | ✅ |
| Page breaks | ✅ |

## Installation

The package is not yet on PyPI. Install locally from source:

```bash
pip install -e /path/to/py_docx2pdfmake
```

The only runtime dependency is `python-docx>=1.1.0`, which is pulled in automatically.

After installation a `docx2pdfmake` command is registered on your `PATH`.

## Quick start

```python
from docx2pdfmake import DocxConverter

conv = DocxConverter()
ddo = conv.convert("document.docx")

import json
print(json.dumps(ddo, indent=2, ensure_ascii=False))
```

## Options

```python
from docx2pdfmake import DocxConverter, ConversionOptions

opts = ConversionOptions(
    # Page layout
    page_size="A4",                   # "A4" | "LETTER" | "A3" …
    page_orientation="portrait",      # "portrait" | "landscape"
    margin_top=40.0,                  # in pt
    margin_right=40.0,
    margin_bottom=60.0,
    margin_left=40.0,

    # Styles
    emit_named_styles=True,           # DOCX styles → pdfmake styles block
    emit_inline_overrides=True,       # embed run formatting inline

    # Images
    embed_images=True,                # True = base64 data URL, False = external ref
    max_image_width=500.0,            # max image width in pt (None = no limit)
    max_image_height=700.0,           # max image height in pt (None = no limit)

    # Tables
    default_table_width=None,         # total table width in pt (None = auto)

    # Typography
    default_font="Roboto",
    default_font_size=11.0,
    default_line_height=1.2,

    # Headings (one value per heading level 1–6)
    heading_font_sizes=[22, 18, 16, 14, 13, 12],
    heading_bold=[True] * 6,
    heading_color=["#000000"] * 6,
    heading_margin_before=[16, 14, 12, 10, 8, 6],  # pt above heading
    heading_margin_after=[8, 6, 6, 4, 4, 4],        # pt below heading

    # Header / footer
    include_header=True,
    include_footer=True,
)

conv = DocxConverter(opts)
ddo = conv.convert("document.docx")
```

## Convenience methods

```python
# Return JSON string directly
json_str = conv.convert_to_json("input.docx", indent=2)

# Write directly to a file; returns the Path of the written file
output_path = conv.convert_to_file("input.docx", "output.json")

# Pass a binary IO object (e.g. from a Django file upload)
from io import BytesIO
with open("input.docx", "rb") as f:
    ddo = conv.convert(BytesIO(f.read()))
```

Per-call options can override the constructor options:

```python
ddo = conv.convert("input.docx", options=ConversionOptions(embed_images=False))
```

## Header / footer with page numbers (pdfmake JS)

The converter returns two representations for header and footer:

```python
ddo = conv.convert("input.docx")

# Static object — ready for server-side pdfmake (Node.js)
print(ddo.get("header"))          # {"text": "…", "fontSize": 9, …}

# JS function body — for client-side pdfmake
print(ddo.get("_header_fn_body")) # 'return [{"text": currentPage.toString(), …}];'
```

In JavaScript:

```javascript
const docDefinition = {
  ...ddo,
  header: new Function('currentPage', 'pageCount', ddo._header_fn_body),
  footer: new Function('currentPage', 'pageCount', ddo._footer_fn_body),
};
pdfMake.createPdf(docDefinition).download('output.pdf');
```

## CLI

After installing the package a `docx2pdfmake` command is available:

```bash
# JSON to stdout
docx2pdfmake input.docx

# Write to file
docx2pdfmake input.docx -o output.json

# Common options
docx2pdfmake input.docx \
    --page-size LETTER \
    --page-orientation landscape \
    --no-embed-images \
    --no-named-styles \
    --no-inline-overrides \
    --default-font "Times New Roman" \
    --default-font-size 10 \
    --indent 4
```

Alternatively, without installation:

```bash
python -m docx2pdfmake input.docx
```

## Django / REST example

```python
import json
from io import BytesIO
from django.http import JsonResponse
from docx2pdfmake import DocxConverter, ConversionOptions

def docx_to_pdfmake_view(request):
    docx_file = request.FILES["file"]
    opts = ConversionOptions(embed_images=True)
    conv = DocxConverter(opts)
    ddo = conv.convert(BytesIO(docx_file.read()))
    return JsonResponse(ddo)
```

## Architecture

```
DocxConverter.convert(source)
  └── StyleResolver(doc, opts)      — builds style name/props cache
  └── ImageHandler(doc, opts)       — pre-loads all images as base64
  └── ContentBuilder(doc, sr, ih, opts)
        ├── _process_body()          — iterates body children
        ├── _process_paragraph()     — text, headings, images
        ├── _collect_list()          — groups consecutive list items
        ├── _build_list_tree()       — builds ul/ol pdfmake nodes
        └── _process_table()         — table → pdfmake table node
  └── HeaderFooterExtractor(doc, opts)
        └── returns { static, fn_body } for JS usage
```

Module layout:

```
src/docx2pdfmake/
├── __init__.py          # public API: DocxConverter, ConversionOptions
├── __main__.py          # python -m docx2pdfmake entry point
├── converter.py         # DocxConverter — main entry point
├── models.py            # ConversionOptions dataclass
├── style_resolver.py    # DOCX styles → pdfmake styles block + inline lookup
├── image_handler.py     # image extraction + base64 encoding
├── content_builder.py   # paragraphs, lists, tables → pdfmake nodes
├── header_footer.py     # header/footer extraction
└── cli.py               # argparse CLI
```

## Known limitations

- **Nested tables** (table inside a table cell) are flattened
- **Text boxes / Shapes** (DrawingML) are not converted
- **Equations** (OMML) are skipped
- **Custom numbering** with complex overrides may fall back to `ul`
- **Bidi text** (Arabic, Hebrew) is passed through but not explicitly mirrored
