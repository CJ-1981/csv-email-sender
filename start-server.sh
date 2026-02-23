#!/bin/bash

echo "============================================"
echo "CSV Email Sender - Local Development Server"
echo "============================================"
echo ""
echo "Starting local web server on http://localhost:8000"
echo ""
echo "Press Ctrl+C to stop the server"
echo ""
echo "============================================"
echo ""

# Try Python 3 first
if command -v python3 &> /dev/null; then
    echo "Using Python 3 to start server..."
    python3 -m http.server 8000
    exit 0
fi

# Try Python 2
if command -v python2 &> /dev/null; then
    echo "Using Python 2 to start server..."
    python2 -m SimpleHTTPServer 8000
    exit 0
fi

# Try python command (might be Python 2 or 3)
if command -v python &> /dev/null; then
    echo "Using Python to start server..."
    python -m http.server 8000 2>/dev/null || python -m SimpleHTTPServer 8000
    exit 0
fi

# If no Python found
echo "============================================"
echo "ERROR: Python not found!"
echo "============================================"
echo ""
echo "Please install Python or use another method:"
echo ""
echo "Option 1: Install Python"
echo "  - Mac: brew install python3"
echo "  - Ubuntu/Debian: sudo apt-get install python3"
echo "  - Or download from https://python.org"
echo ""
echo "Option 2: Use Node.js:"
echo "  npx serve"
echo ""
echo "Option 3: Use PHP:"
echo "  php -S localhost:8000"
echo ""
echo "============================================"
exit 1
