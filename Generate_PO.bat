@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python 3 not found. Please install Python 3 and try again.
  pause
  exit /b 1
)
python generate_po_one_click.py
if errorlevel 1 (
  echo.
  echo PO generation failed.
  pause
  exit /b 1
)
start "" "%CD%\OUTPUT"
