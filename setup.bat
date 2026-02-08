@echo off
REM Quick start script for Automatisk vakansberäkning (Windows)

echo ╔══════════════════════════════════════════════════════════╗
echo ║      Automatisk vakansberäkning - Quick Start             ║
echo ╚══════════════════════════════════════════════════════════╝
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed. Please install Python 3.8 or higher.
    pause
    exit /b 1
)

echo ✓ Python found
python --version

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo.
    echo 📦 Creating virtual environment...
    python -m venv venv
    echo ✓ Virtual environment created
)

REM Activate virtual environment
echo.
echo 🔌 Activating virtual environment...
call venv\Scripts\activate.bat

REM Install dependencies
echo.
echo 📥 Installing dependencies...
pip install -q -r requirements.txt
echo ✓ Dependencies installed

REM Create input/output directories
echo.
echo 📁 Creating directories...
if not exist "input" mkdir input
if not exist "output" mkdir output
echo ✓ Directories created

REM Run tests
echo.
echo 🧪 Running tests...
python test_examples.py

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║                  Installation Complete!                  ║
echo ╚══════════════════════════════════════════════════════════╝
echo.
echo 🚀 Start the web app with:
echo    start_web.bat
echo.
echo Or use CLI:
echo    python vakant_karens_app.py --sick_pdf input\sjuklista.pdf --payslips input\*.pdf --out output\rapport.xlsx
echo.
echo 📚 See README.md for full documentation
echo.

pause
