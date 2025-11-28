#!/bin/bash

# IB Question Entry System - Mac Launcher Script
# Double-click this file to start the application

# Change to the script's directory
cd "$(dirname "$0")"

echo "========================================="
echo "IB Question Entry System"
echo "========================================="
echo ""

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ ERROR: Python 3 is not installed on your Mac."
    echo ""
    echo "Please install Python 3:"
    echo "  1. Visit https://www.python.org/downloads/"
    echo "  2. Download Python 3 for macOS"
    echo "  3. Run the installer"
    echo "  4. Come back and try again!"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "✅ Python 3 found: $(python3 --version)"
echo ""

# Check if required packages are installed
echo "Checking for required packages..."
python3 -c "import PyQt6" 2>/dev/null
if [ $? -ne 0 ]; then
    echo ""
    echo "⚠️  Required packages are not installed."
    echo "Installing required packages (this may take a minute)..."
    echo ""
    
    # Try to install pip if it doesn't exist
    if ! command -v pip3 &> /dev/null; then
        echo "Installing pip..."
        python3 -m ensurepip --upgrade
    fi
    
    # Install requirements
    pip3 install -r requirements.txt
    
    if [ $? -ne 0 ]; then
        echo ""
        echo "❌ ERROR: Failed to install required packages."
        echo ""
        echo "Please try installing manually:"
        echo "  1. Open Terminal"
        echo "  2. Navigate to this folder"
        echo "  3. Run: pip3 install -r requirements.txt"
        echo ""
        read -p "Press Enter to exit..."
        exit 1
    fi
    
    echo ""
    echo "✅ Packages installed successfully!"
    echo ""
fi

echo "✅ All requirements are ready!"
echo ""
echo "Starting application..."
echo "========================================="
echo ""

# Run the application
python3 app.py

# Keep window open if there's an error
if [ $? -ne 0 ]; then
    echo ""
    echo "❌ The application encountered an error."
    echo ""
    read -p "Press Enter to exit..."
fi

