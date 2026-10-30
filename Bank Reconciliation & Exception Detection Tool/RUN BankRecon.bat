@echo off
title Bank Reconciliation Tool
cd /d "%~dp0"

:: Check that python-embed exists
if not exist "python-embed\python.exe" (
    echo.
    echo  ERROR: python-embed\python.exe not found.
    echo  Please place this .bat file in the same folder as python-embed\
    echo.
    pause
    exit /b 1
)

:: Check that BankRecon.py exists
if not exist "BankRecon.py" (
    echo.
    echo  ERROR: BankRecon.py not found.
    echo  Please place BankRecon.py in the same folder as this .bat file.
    echo.
    pause
    exit /b 1
)

:: Set paths
set PYTHONPATH=%~dp0python-embed\Lib\site-packages
set PATH=%~dp0python-embed;%PATH%

:: Check flask is installed
python-embed\python.exe -c "import flask, openpyxl" 2>nul
if errorlevel 1 (
    echo.
    echo  Installing required packages ^(flask, openpyxl^)...
    echo  This only happens once.
    echo.
    python-embed\python.exe -m pip install flask openpyxl --target="%~dp0python-embed\Lib\site-packages" --quiet
    if errorlevel 1 (
        echo.
        echo  ERROR: Failed to install packages. Check your internet connection.
        echo.
        pause
        exit /b 1
    )
    echo  Done. Packages installed successfully.
    echo.
)

:: Run the app
echo.
echo  Starting Bank Reconciliation Tool...
echo  Your browser will open automatically.
echo  Keep this window open while using the app.
echo  Close this window when you are done.
echo.
python-embed\python.exe BankRecon.py
pause
