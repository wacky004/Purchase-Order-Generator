# Purchase Order Generator (Takara Int'l Export Corp.)

A Windows desktop app that fills a digital Purchase Order template from form
inputs and exports print-ready **PDF** + **PNG** files.

## Features

- Order form: P.O. No., customer name, order/delivery dates, dynamic item rows
  (Item No., Description, Qty., Rate, U/M) with automatic Amount and Total
- Signatories: Approved by / Verified by / Received by, with dates
- Calendar date pickers on all date fields
- Clean digital PO template matching the company paper form (multi-page,
  compact rows, repeated table header)
- Saved PO library with search, reopen-to-edit, duplicate/Keep-Both handling
- Generate, Preview, and Print (default Windows printer)
- Self-contained rendering: Edge/Chrome when available, built-in fallback
  renderer otherwise

## Run from source

Requires Python 3.11+ on Windows (uses `tkinter`, `Pillow`):

```bat
python app.py
```

Saved POs and generated files go to `Documents\Takara Purchase Orders\`
(`orders\` records, `documents\` PDFs/PNGs).

## Build the exe

Requires [PyInstaller](https://pyinstaller.org/):

```bat
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "Takara PO Generator" --icon "assets\app.ico" ^
  --add-data "template.html;." --add-data "assets;assets" app.py
```

Or rebuild from the spec file:

```bat
python -m PyInstaller "Takara PO Generator.spec"
```

`python app.py --selftest` validates document generation without opening the UI.
