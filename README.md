# Ticket Report Generator

Generate PDF ticket reports from an Excel template.

## Requirements

- Python 3
- LibreOffice (`libreoffice` or `soffice`)
- Python dependencies from `requirements.txt`

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python3 report.py
```

The script prompts for ticket information, issue summary rows, and creates a PDF in `temp/`.

The generated PDF uses the ticket ID as the filename:

```text
temp/ABC-123.pdf
```

The Excel template remains unchanged. Temporary XLSX files are deleted after PDF generation.

## Template

The source template is:

```text
report-template.xlsx
```

Do not edit generated files in `temp/`; update the template instead.

## License

MIT License.

Copyright (c) 2026 Matheus Tecchio
