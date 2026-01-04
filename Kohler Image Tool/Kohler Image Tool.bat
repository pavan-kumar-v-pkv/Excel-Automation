@echo off
REM Kohler Image Tool - Fully Automatic Windows Launcher
TITLE Kohler Image Tool

cd /d "%~dp0\.app_files"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ========================================
    echo   KOHLER IMAGE TOOL - FIRST TIME SETUP
    echo ========================================
    echo.
    echo Python is not installed. Installing automatically...
    echo This will take 2-3 minutes. Please wait...
    echo.
    
    REM Download Python installer
    echo [1/3] Downloading Python installer...
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe' -OutFile '%TEMP%\python_installer.exe'}"
    
    if errorlevel 1 (
        echo.
        echo Failed to download Python. Please check your internet connection.
        echo.
        echo Alternative: Install Python manually from python.org
        echo Make sure to check "Add Python to PATH" during installation.
        pause
        exit /b 1
    )
    
    echo [2/3] Installing Python (this may take 1-2 minutes)...
    echo Please wait, do not close this window...
    REM Install Python silently with PATH enabled
    "%TEMP%\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
    
    REM Wait for installation to complete
    timeout /t 5 /nobreak >nul
    
    REM Refresh environment variables
    call refreshenv >nul 2>&1
    
    REM Clean up installer
    del "%TEMP%\python_installer.exe" >nul 2>&1
    
    echo [3/3] Python installed successfully!
    echo.
    echo Restarting tool to complete setup...
    timeout /t 2 /nobreak >nul
    
    REM Restart this script to continue with package installation
    start "" "%~f0"
    exit /b 0
)

REM Check if packages are installed
python -c "import pdfplumber, PIL, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo ========================================
    echo   Installing required components...
    echo ========================================
    echo.
    echo This will take about 1 minute.
    echo Please keep this window open...
    echo.
    
    python -m pip install --upgrade pip --quiet --user >nul 2>&1
    python -m pip install --quiet --user pdfplumber Pillow openpyxl
    
    if errorlevel 1 (
        echo.
        echo Installation failed. Please check your internet connection.
        timeout /t 3 /nobreak >nul
        exit /b 1
    )
    
    echo.
    echo Setup complete! Starting Kohler Image Tool...
    echo.
    timeout /t 2 /nobreak >nul
)

REM Run the application
start "Kohler Image Tool" python run_gui.py

REM Close the command window after launching GUI
exit
