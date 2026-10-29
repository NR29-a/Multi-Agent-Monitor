@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
  echo FEHLER: Python wurde nicht gefunden.
  echo Python muss installiert sein. Danach: py -m pip install openpyxl
  pause
  exit /b 1
)
py "%~dp0excel_zu_json.py" %*
echo.
pause
