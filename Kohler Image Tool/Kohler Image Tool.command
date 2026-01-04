#!/bin/bash
# Kohler Image Tool - Application Launcher

cd "$(dirname "$0")/.app_files"

# Check Python installation
if ! command -v python3 &> /dev/null; then
    osascript -e 'tell app "System Events" to display dialog "Please install Python first.\n\n1. Go to python.org\n2. Download Python 3.11\n3. Install it\n4. Run this tool again" buttons {"OK"} default button 1 with title "Setup Required"'
    open "https://www.python.org/downloads/"
    exit 1
fi

# Check if packages are installed, if not, install them automatically
python3 -c "import pdfplumber, PIL, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    osascript -e 'tell app "System Events" to display dialog "Installing required components...\nThis will take about 1 minute.\n\nClick OK to continue." buttons {"OK"} default button 1 with title "First Time Setup"'
    
    # Install in background with progress
    python3 -m pip install --quiet --user pdfplumber Pillow openpyxl
    
    if [ $? -eq 0 ]; then
        osascript -e 'tell app "System Events" to display dialog "Setup complete! Starting the tool now..." buttons {"OK"} default button 1 with title "Ready"'
    else
        osascript -e 'tell app "System Events" to display dialog "Installation failed. Please check your internet connection and try again." buttons {"OK"} default button 1 with icon stop'
        exit 1
    fi
fi

# Run the application
python3 run_gui.py
