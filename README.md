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

## Build EXE (Windows)

### Option 1: one-click

```powershell
.\build.bat
```

### Option 2: PowerShell

```powershell
.\build_exe.ps1
```

### Clean build

```powershell
.\build_exe.ps1 -Clean
```

After build, executable will be at:

- `dist\docx2md-gui.exe`

## Notes for PDF conversion

- `docx -> pdf` depends on Microsoft Word via COM automation.
- If target machine has no Word installed, PDF conversion may fail.
