#!/bin/bash
# Quick start script for Automatisk vakansberäkning

set -e

echo "╔══════════════════════════════════════════════════════════╗"
echo "║      Automatisk vakansberäkning - Quick Start             ║"
echo "╚══════════════════════════════════════════════════════════╝"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3.8 or higher."
    exit 1
fi

echo "✓ Python 3 found: $(python3 --version)"

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip3 is not installed. Please install pip."
    exit 1
fi

echo "✓ pip3 found"

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo ""
    echo "📦 Creating virtual environment..."
    python3 -m venv venv
    echo "✓ Virtual environment created"
fi

# Activate virtual environment
echo ""
echo "🔌 Activating virtual environment..."
source venv/bin/activate

# Install dependencies
echo ""
echo "📥 Installing dependencies..."
pip install -q -r requirements.txt
echo "✓ Dependencies installed"

# Create input/output directories
echo ""
echo "📁 Creating directories..."
mkdir -p input output
echo "✓ Directories created"

# Run tests
echo ""
echo "🧪 Running tests..."
python test_examples.py

echo ""
echo "╔══════════════════════════════════════════════════════════╗"
echo "║                  Installation Complete!                  ║"
echo "╚══════════════════════════════════════════════════════════╝"
echo ""
echo "🚀 Start the web app with:"
echo "   ./start_web.sh"
echo ""
echo "Or use CLI:"
echo "   python vakant_karens_app.py --sick_pdf input/sjuklista.pdf --payslips input/*.pdf --out output/rapport.xlsx"
echo ""
echo "📚 See README.md for full documentation"
echo ""
