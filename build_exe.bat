@echo off
setlocal
cd /d "%~dp0"

python -m pip install --upgrade pip
if errorlevel 1 goto :error

python -m pip install -r requirements.txt
if errorlevel 1 goto :error

pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name ExcelParaTxt ^
  --collect-all tkinterdnd2 ^
  Excel_TotxT.py

if errorlevel 1 goto :error

for /f "usebackq delims=" %%I in (`powershell -NoProfile -STA -Command "Add-Type -AssemblyName System.Windows.Forms; $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = 'Escolha onde salvar o arquivo .exe'; $dialog.ShowNewFolderButton = $true; if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $dialog.SelectedPath }"`) do set "DESTINO=%%I"

if not defined DESTINO (
  echo.
  echo Nenhuma pasta foi selecionada. O .exe continuara em dist\ExcelParaTxt.exe
  pause
  exit /b 0
)

if not exist "%DESTINO%" (
  echo.
  echo Pasta invalida: %DESTINO%
  pause
  exit /b 1
)

copy /Y "dist\ExcelParaTxt.exe" "%DESTINO%\ExcelParaTxt.exe" >nul
if errorlevel 1 goto :error

echo.
echo Build concluido. O .exe foi salvo em:
echo %DESTINO%\ExcelParaTxt.exe
pause
exit /b 0

:error
echo.
echo Ocorreu um erro durante a geracao ou copia do executavel.
pause
exit /b 1
