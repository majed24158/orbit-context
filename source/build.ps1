$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --distpath ../ --workpath ./build OrbitContext.spec
