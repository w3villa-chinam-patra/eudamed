# Swissamed Scripts

## Python Setup

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install openpyxl lxml
```

## Run Commands

Create Excel template:

```bash
python create_eudamed_excel_template.py --output sample_excel/eudamed_upload_template.xlsx
```

Generate XML from Excel:

```bash
python excel_to_eudamed_xml.py --excel sample_excel/eudamed_upload_template.xlsx --output output/eudamed_upload.xml
```
