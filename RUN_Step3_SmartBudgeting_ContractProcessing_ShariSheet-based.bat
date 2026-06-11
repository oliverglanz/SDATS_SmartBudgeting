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


REM === Check and install required packages (for SmartContract PDF filling + DocuSign notebook) ===

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

REM --- openpyxl (needed by pandas for .xlsx read/write) ---
"%PYTHON_PATH%" -c "import openpyxl" 2>nul
if errorlevel 1 (
    echo Installing openpyxl...
    "%PYTHON_PATH%" -m pip install openpyxl
)

REM --- xlrd (needed by pandas for legacy .xls files) ---
"%PYTHON_PATH%" -c "import xlrd" 2>nul
if errorlevel 1 (
    echo Installing xlrd...
    "%PYTHON_PATH%" -m pip install xlrd
)

REM --- PyPDF2 (PDF reading/writing used in notebook) ---
"%PYTHON_PATH%" -c "import PyPDF2" 2>nul
if errorlevel 1 (
    echo Installing PyPDF2...
    "%PYTHON_PATH%" -m pip install PyPDF2
)

REM --- pypdf (PDF reading/writing used in notebook) ---
"%PYTHON_PATH%" -c "import pypdf" 2>nul
if errorlevel 1 (
    echo Installing pypdf...
    "%PYTHON_PATH%" -m pip install pypdf
)

REM --- reportlab (PDF reading/writing used in notebook) ---
"%PYTHON_PATH%" -c "import reportlab" 2>nul
if errorlevel 1 (
    echo Installing reportlab...
    "%PYTHON_PATH%" -m pip install reportlab
)

REM --- docusign-esign (DocuSign eSignature SDK; imports as docusign_esign) ---
"%PYTHON_PATH%" -c "import docusign_esign" 2>nul
if errorlevel 1 (
    echo Installing docusign-esign...
    "%PYTHON_PATH%" -m pip install docusign-esign
)


echo.
echo All required packages are installed.
echo.

REM Run the Python script
"%PYTHON_PATH%" "%~dp0source_Step3_SmartContract_PDF_filling_v20260416_woDocuSign_wTemplate_OneCoursePerContract.py"

echo.
echo =========================================
echo Script completed!
echo =========================================
pause