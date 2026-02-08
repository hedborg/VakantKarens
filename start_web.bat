@echo off
REM Start the Streamlit web app (Windows)

echo 🚀 Starting Automatisk vakansberäkning...
echo.

REM Activate virtual environment if it exists
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

REM Start Streamlit
streamlit run vakant_karens_streamlit.py ^
    --server.port=8502 ^
    --server.address=localhost ^
    --browser.gatherUsageStats=false

REM Note: Press Ctrl+C to stop the server
