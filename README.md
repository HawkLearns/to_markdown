# Document to Markdown Converter

A small desktop app for converting PDF, Excel, and Visio files to Markdown.

## Features

- Select and convert multiple files in one batch.
- Save each Markdown file beside its source file with the same name and a `.md` extension.
- Convert all sheets in an Excel workbook in workbook order.
- Extract PDF content with PyMuPDF4LLM.
- Convert `.vsdx` diagrams to Mermaid Markdown.

## Requirements

- Python 3
- Tkinter (usually included with Python)

Install the required packages:

```bash
pip install -U pymupdf4llm pymupdf pandas openpyxl
```

For legacy `.xls` files, install `xlrd` as well:

```bash
pip install xlrd
```

## Run

From the project directory:

```bash
python md_converter_rev3.py
```

Use **Browse** to select one or more supported files, adjust the available PDF options if needed, then click **Convert**. If an output Markdown file already exists, the app asks before overwriting it.

## Supported formats

- PDF: `.pdf`
- Excel: `.xlsx`, `.xls` (legacy `.xls` requires `xlrd`)
- Visio: `.vsdx`

Legacy binary Visio `.vsd` files can be selected in the file picker but are not supported for conversion.
