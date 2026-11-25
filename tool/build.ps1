# 打包（在 pads-logic-eco 根目录运行:  powershell -ExecutionPolicy Bypass -File tool\build.ps1）
# 需要: pip install pyinstaller openpyxl pywin32
# 产物: dist\pads-eco\pads-eco.exe（窗口版）+ library\ + 使用说明.txt
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

pyinstaller --noconfirm --clean --onefile --windowed `
  --name pads-eco `
  --paths scripts `
  --add-data "$root\library\Electronic_components_library.xlsx;library" `
  --add-data "$root\tool\pads-eco.ico;tool" `
  --icon "$root\tool\pads-eco.ico" `
  --hidden-import ecn_to_pads_eco --hidden-import verify_eco --hidden-import bomscan --hidden-import ecn_reverse `
  --hidden-import xlsx_open --hidden-import textio --hidden-import win32com.client --hidden-import pythoncom --hidden-import win32timezone `
  --distpath dist\pads-eco `
  --workpath build\pyinstaller `
  --specpath build `
  tool\pads_eco_app.py

python -X utf8 -c "import shutil; from pathlib import Path; d=Path('dist/pads-eco'); (d/'library').mkdir(parents=True, exist_ok=True); shutil.copy('library/Electronic_components_library.xlsx', d/'library'); [shutil.copy(t, d) for t in Path('tool').glob('*.txt')]; (d/'pads-eco-cli.exe').unlink(missing_ok=True); shutil.rmtree('build', ignore_errors=True)"
Write-Host "OK -> dist\pads-eco\"
