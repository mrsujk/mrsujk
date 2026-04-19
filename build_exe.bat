@echo off
cd /d "%~dp0"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name ExcelParaTxt ^
  --collect-all tkinterdnd2 ^
  aula.py

echo.
echo Build concluido. O .exe estara em dist\ExcelParaTxt.exe
pause
