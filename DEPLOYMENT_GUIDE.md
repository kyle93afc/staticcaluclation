@echo off
echo ===============================================
echo Steel Chimney Calculator - Installation Guide
echo ===============================================
echo.

echo Step 1: Checking Python installation...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python not found!
    echo Please download Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b
)

echo.
echo Step 2: Installing required packages...
pip install jupyter pandas numpy matplotlib seaborn ipywidgets
if %errorlevel% neq 0 (
    echo ERROR: Package installation failed!
    pause
    exit /b
)

echo.
echo Step 3: Starting Jupyter Notebook...
echo The calculator will open in your web browser
echo Look for "chimney_calculation.ipynb" and click to open
echo.
pause
jupyter notebook

echo.
echo Installation completed successfully!
echo The chimney calculator is now ready to use.
pause