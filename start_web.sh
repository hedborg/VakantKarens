#!/bin/bash
# Start the Streamlit web app

echo "🚀 Starting Automatisk vakansberäkning..."
echo ""

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    source venv/bin/activate
fi

# Start Streamlit
streamlit run vakant_karens_streamlit.py \
    --server.port=8501 \
    --server.address=localhost \
    --browser.gatherUsageStats=false

# Note: Press Ctrl+C to stop the server
