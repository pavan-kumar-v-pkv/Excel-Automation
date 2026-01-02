#!/usr/bin/env python3
"""
Kohler Image Automation - GUI Launcher
Double-click this file to start the application.
"""

# Add python package to path
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

# Import and run GUI
from python.gui_app import main

if __name__ == "__main__":
    main()