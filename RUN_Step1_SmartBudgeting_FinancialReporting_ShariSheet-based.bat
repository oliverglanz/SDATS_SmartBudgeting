@echo off
setlocal enabledelayedexpansion
title Smart Budgeting - Step 1
echo =========================================
echo Starting Smart Budgeting Step 1...
echo =========================================
echo.

REM === Navigate to the directory where this batch file is located ===
cd /d "%~dp0"

REM === Use flexible Anaconda Python path (works for any user) ===
set "PYTHON_PATH=%LOCALAPPDATA%\anaconda3\python.exe"

if not exist "%PYTHON_PATH%" (
    set "PYTHON_PATH=%LOCALAPPDATA%\Programs\anaconda3\python.exe"
)

if not exist "%PYTHON_PATH%" (
    set "PYTHON_PATH=%USERPROFILE%\anaconda3\python.exe"
)

if not exist "%PYTHON_PATH%" (
    set "PYTHON_PATH=%USERPROFILE%\Anaconda3\python.exe"
)

if not exist "%PYTHON_PATH%" (
    set "PYTHON_PATH=%ProgramData%\anaconda3\python.exe"
)

if not exist "%PYTHON_PATH%" (
    echo ERROR: Could not find Anaconda Python installation.
    echo Please install Anaconda for this program to run.
    echo Expected in one of:
    echo   %%LOCALAPPDATA%%\anaconda3\
    echo   %%LOCALAPPDATA%%\Programs\anaconda3\
    echo   %%USERPROFILE%%\anaconda3\
    echo   %%USERPROFILE%%\Anaconda3\
    echo   %%ProgramData%%\anaconda3\
    echo.
    pause
    exit /b 1
)

echo Using Python at: %PYTHON_PATH%
echo.

REM === Activate base Conda environment if possible ===
if exist "%LOCALAPPDATA%\anaconda3\Scripts\activate.bat" (
    echo Activating base Conda environment...
    call "%LOCALAPPDATA%\anaconda3\Scripts\activate.bat"
) else if exist "%USERPROFILE%\anaconda3\Scripts\activate.bat" (
    echo Activating base Conda environment...
    call "%USERPROFILE%\anaconda3\Scripts\activate.bat"
) else if exist "%ProgramData%\anaconda3\Scripts\activate.bat" (
    echo Activating base Conda environment...
    call "%ProgramData%\anaconda3\Scripts\activate.bat"
) else (
    echo Proceeding without explicit environment activation.
)

echo.
echo Checking required Python packages...
echo.


REM === Check and install required packages ===

REM --- openpyxl (Excel reading/writing + formatting) ---
"%PYTHON_PATH%" -c "import openpyxl" 2>nul
if errorlevel 1 (
    echo Installing openpyxl...
    "%PYTHON_PATH%" -m pip install openpyxl
)

"%PYTHON_PATH%" -c "import xlrd" 2>nul
if errorlevel 1 (
    echo Installing xlrd...
    "%PYTHON_PATH%" -m pip install xlrd
)

REM --- xlsxwriter (Excel export with formatting) ---
"%PYTHON_PATH%" -c "import xlsxwriter" 2>nul
if errorlevel 1 (
    echo Installing xlsxwriter...
    "%PYTHON_PATH%" -m pip install XlsxWriter
)

REM --- pandas (core data processing) ---
"%PYTHON_PATH%" -c "import pandas" 2>nul
if errorlevel 1 (
    echo Installing pandas...
    "%PYTHON_PATH%" -m pip install pandas
)

REM --- numpy (numeric operations) ---
"%PYTHON_PATH%" -c "import numpy" 2>nul
if errorlevel 1 (
    echo Installing numpy...
    "%PYTHON_PATH%" -m pip install numpy
)

REM --- matplotlib (backend only, but required) ---
"%PYTHON_PATH%" -c "import matplotlib" 2>nul
if errorlevel 1 (
    echo Installing matplotlib...
    "%PYTHON_PATH%" -m pip install matplotlib
)


echo.
echo All required packages are installed.
echo.

REM Run the Python script
"%PYTHON_PATH%" "%~dp0source_Step1_SmartBudgeting_FinancialReporting_ShariSheet-based_v20260509.py"

echo.
echo =========================================
echo Script completed!
echo =========================================
pause