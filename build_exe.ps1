param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

if ($Clean) {
    if (Test-Path ".\\build") { Remove-Item -Recurse -Force ".\\build" }
    if (Test-Path ".\\dist") { Remove-Item -Recurse -Force ".\\dist" }
    if (Test-Path ".\\docx2md-gui.spec") { Remove-Item -Force ".\\docx2md-gui.spec" }
}

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install -r requirements-build.txt

pyinstaller --noconfirm --clean --windowed --onefile `
  --name docx2md-gui `
  --hidden-import=win32com `
  --hidden-import=win32com.client `
  --collect-submodules=docx2pdf `
  app.py

Write-Host "Build complete: .\\dist\\docx2md-gui.exe"
