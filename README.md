# docx2md GUI Tool

A local Python GUI utility that supports:

- DOCX to Markdown (extracts images to `xxx_images/` and references them in Markdown)
- DOCX to PDF (via `docx2pdf`; Microsoft Word is usually required on Windows)
- One-click convert both formats
- Warning classification: style hint / content risk / general warning
- Strict mode: popup reminder when content-risk warnings are detected

## Install

```powershell
pip install -r requirements.txt
```

## Run

```powershell
python app.py
```

## Usage

1. Choose a `.docx` file
2. Choose output folder
3. Optionally enable strict mode
4. Click one of:
   - Convert to Markdown
   - Convert to PDF
   - Convert Both

## Outputs

For input `report.docx`, the output folder will contain:

- `report.md`
- `report_images/`
- `report.pdf`
