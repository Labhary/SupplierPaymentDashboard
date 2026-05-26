@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 exit /b 1
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 exit /b 1

python -m pip install --upgrade pip
if errorlevel 1 exit /b 1

python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

python -m pip install pyinstaller
if errorlevel 1 exit /b 1

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onedir ^
  --name SupplierPaymentDashboard ^
  --collect-all streamlit ^
  --collect-all altair ^
  --collect-all pydeck ^
  --collect-all plotly ^
  --collect-all reportlab ^
  --collect-all matplotlib ^
  --add-data "app.py;." ^
  --add-data "src;src" ^
  --add-data "requirements.txt;." ^
  launcher.py

if errorlevel 1 exit /b 1

echo.
echo Build complete.
echo Run: dist\SupplierPaymentDashboard\SupplierPaymentDashboard.exe

endlocal
