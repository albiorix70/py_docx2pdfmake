# docx2pdfmake

Python-Bibliothek zum Konvertieren von `.docx`-Dateien in **pdfmake Document Definition Objects (DDO)**.

## Features

| Feature | Status |
|---|---|
| Fließtext (bold, italic, underline, Farbe, Highlight) | ✅ |
| Überschriften Heading 1–6 | ✅ |
| Ungeordnete & geordnete Listen (ul/ol, verschachtelt) | ✅ |
| Tabellen (inkl. Header-Zeile, merged cells) | ✅ |
| Inline-Bilder (base64 eingebettet) | ✅ |
| Seitenlayout (Größe, Orientierung, Margins) | ✅ |
| Header / Footer (inkl. Seitenzahl-Platzhalter) | ✅ |
| Named Styles aus DOCX → pdfmake `styles`-Block | ✅ |
| Inline-Overrides (Style aus Run-Formatting) | ✅ |
| Hyperlinks | ✅ |
| Seitenumbrüche | ✅ |

## Installation

```bash
pip install python-docx       # einzige Abhängigkeit
# (Paket noch nicht auf PyPI – lokale Installation:)
pip install -e /pfad/zu/docx2pdfmake
```

## Schnellstart

```python
from docx2pdfmake import DocxConverter

conv = DocxConverter()
ddo = conv.convert("mein_dokument.docx")

import json
print(json.dumps(ddo, indent=2, ensure_ascii=False))
```

## Optionen

```python
from docx2pdfmake import DocxConverter, ConversionOptions

opts = ConversionOptions(
    # Seitenlayout
    page_size="A4",                  # "A4" | "LETTER" | "A3" …
    page_orientation="portrait",     # "portrait" | "landscape"
    margin_top=40.0,                 # in pt
    margin_right=40.0,
    margin_bottom=60.0,
    margin_left=40.0,

    # Styles
    emit_named_styles=True,          # DOCX-Styles → pdfmake styles-Block
    emit_inline_overrides=True,      # Run-Formatierung inline einbetten

    # Bilder
    embed_images=True,               # base64 einbetten (False = externe Ref)
    max_image_width=500.0,           # Maximale Bildbreite in pt
    max_image_height=700.0,

    # Typografie
    default_font="Roboto",
    default_font_size=11.0,
    default_line_height=1.2,

    # Überschriften
    heading_font_sizes=[22, 18, 16, 14, 13, 12],
    heading_bold=[True]*6,
    heading_color=["#000000"]*6,

    # Header/Footer aus DOCX übernehmen
    include_header=True,
    include_footer=True,
)

conv = DocxConverter(opts)
ddo = conv.convert("dokument.docx")
```

## Convenience-Methoden

```python
# Direkt als JSON-String
json_str = conv.convert_to_json("input.docx", indent=2)

# Direkt in Datei schreiben
output_path = conv.convert_to_file("input.docx", "output.json")

# Binary-IO-Objekt (z. B. aus Django-Upload)
with open("input.docx", "rb") as f:
    ddo = conv.convert(f)
```

## Header / Footer mit Seitenzahlen (pdfmake JS)

Der Converter liefert für Header/Footer zwei Varianten:

```python
ddo = conv.convert("input.docx")

# Statisches Objekt (für server-seitiges pdfmake via node)
print(ddo.get("header"))   # { "text": "...", "fontSize": 9, ... }

# JS-Function-Body (für client-seitiges pdfmake)
print(ddo.get("_header_fn_body"))
# → 'return [{"text": currentPage.toString(), ...}];'
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

```bash
# JSON auf stdout
python -m docx2pdfmake input.docx

# In Datei schreiben
python -m docx2pdfmake input.docx -o output.json

# Optionen
python -m docx2pdfmake input.docx \
    --page-size LETTER \
    --page-orientation landscape \
    --no-embed-images \
    --default-font "Times New Roman" \
    --default-font-size 10
```

## Verwendung im Django-Backend (APEX/REST)

```python
import json
from io import BytesIO
from docx2pdfmake import DocxConverter, ConversionOptions

def docx_to_pdfmake_view(request):
    docx_file = request.FILES["file"]
    opts = ConversionOptions(embed_images=True)
    conv = DocxConverter(opts)
    
    buffer = BytesIO(docx_file.read())
    ddo = conv.convert(buffer)
    
    return JsonResponse(ddo)
```

## Architektur

```
docx2pdfmake/
├── __init__.py          # Öffentliche API
├── converter.py         # DocxConverter – Einstiegspunkt
├── models.py            # ConversionOptions (Dataclass)
├── style_resolver.py    # DOCX-Styles → pdfmake-styles-Block + Inline-Lookup
├── image_handler.py     # Bild-Extraktion + base64-Enkodierung
├── content_builder.py   # Paragraphen, Listen, Tabellen → pdfmake-Nodes
├── header_footer.py     # Header/Footer-Extraktion
└── cli.py               # CLI-Interface (python -m docx2pdfmake)
```

## Bekannte Einschränkungen

- **Verschachtelte Tabellen** (Tabelle in Tabellenzelle) werden vereinfacht
- **Textfelder / Shapes** (DrawingML) werden nicht konvertiert
- **Formeln** (OMML) werden übersprungen
- **Benutzerdefinierte Nummerierungen** mit komplexen Overrides können auf `ul` fallen
- **Bidi-Text** (Arabisch, Hebräisch) wird übergeben, aber nicht explizit gespiegelt
