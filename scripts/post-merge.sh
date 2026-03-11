#!/bin/bash
set -e

echo "Post-merge setup: installing Python dependencies..."
pip install -q streamlit pdfplumber python-pptx openpyxl XlsxWriter 2>/dev/null || true

echo "Post-merge setup: verifying app imports..."
python3 -c "import app" 2>/dev/null || echo "Warning: app import check failed"

echo "Post-merge setup complete."
