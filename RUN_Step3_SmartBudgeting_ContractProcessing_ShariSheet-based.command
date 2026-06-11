#!/bin/zsh

echo "========================================="
echo "Starting Smart Budgeting Step 1..."
echo "========================================="
echo ""

# Navigate to the directory where this script is located
cd "$(dirname "$0")"

# Try to find Python in common Anaconda locations
PYTHON_PATH=""

if [ -f "$HOME/anaconda3/bin/python" ]; then
    PYTHON_PATH="$HOME/anaconda3/bin/python"
elif [ -f "/opt/anaconda3/bin/python" ]; then
    PYTHON_PATH="/opt/anaconda3/bin/python"
elif [ -f "$HOME/opt/anaconda3/bin/python" ]; then
    PYTHON_PATH="$HOME/opt/anaconda3/bin/python"
elif [ -f "/usr/local/anaconda3/bin/python" ]; then
    PYTHON_PATH="/usr/local/anaconda3/bin/python"
fi

if [ -z "$PYTHON_PATH" ]; then
    echo "ERROR: Could not find Anaconda Python installation"
    echo "Please check your Anaconda installation"
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Using Python at: $PYTHON_PATH"
echo ""


# Check and install required packages
echo "Checking required packages..."
echo ""

# --- pandas (core data processing) ---
"$PYTHON_PATH" -c "import pandas" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing pandas..."
    "$PYTHON_PATH" -m pip install pandas
fi

# --- numpy (numeric operations) ---
"$PYTHON_PATH" -c "import numpy" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing numpy..."
    "$PYTHON_PATH" -m pip install numpy
fi

# --- openpyxl (Excel .xlsx support) ---
"$PYTHON_PATH" -c "import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing openpyxl..."
    "$PYTHON_PATH" -m pip install openpyxl
fi

# --- xlrd (Excel .xls support) ---
"$PYTHON_PATH" -c "import xlrd" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing xlrd..."
    "$PYTHON_PATH" -m pip install xlrd
fi

# --- pypdf (PDF manipulation used in notebook) ---
"$PYTHON_PATH" -c "import pypdf" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing pypdf..."
    "$PYTHON_PATH" -m pip install pypdf
fi

# --- PyPDF2 (legacy compatibility used in notebook) ---
"$PYTHON_PATH" -c "import PyPDF2" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing PyPDF2..."
    "$PYTHON_PATH" -m pip install PyPDF2
fi


# --- reportlab (legacy compatibility used in notebook) ---
"$PYTHON_PATH" -c "import reportlab" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing reportlab..."
    "$PYTHON_PATH" -m pip install reportlab
fi

# --- DocuSign eSignature SDK ---
"$PYTHON_PATH" -c "import docusign_esign" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing docusign-esign..."
    "$PYTHON_PATH" -m pip install docusign-esign
fi

echo ""
echo "All required packages checked."

echo ""
echo "All required packages are installed."
echo ""

# Run the Python script
"$PYTHON_PATH" source_Step3_SmartContract_PDF_filling_v20260416_woDocuSign_wTemplate_OneCoursePerContract.py

# Capture exit status
EXIT_CODE=$?

echo ""
echo "========================================="
if [ $EXIT_CODE -eq 0 ]; then
    echo "Script completed successfully!"
else
    echo "Script failed with error code: $EXIT_CODE"
fi
echo "========================================="
echo ""
read -p "Press Enter to exit..."