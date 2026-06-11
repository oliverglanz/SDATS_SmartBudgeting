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

# --- openpyxl (Excel reading/writing + formatting) ---
"$PYTHON_PATH" -c "import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing openpyxl..."
    "$PYTHON_PATH" -m pip install openpyxl
fi

"$PYTHON_PATH" -c "import xlrd" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing xlrd..."
    "$PYTHON_PATH" -m pip install xlrd
fi

# --- xlsxwriter (Excel export with formatting) ---
"$PYTHON_PATH" -c "import xlsxwriter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing xlsxwriter..."
    "$PYTHON_PATH" -m pip install XlsxWriter
fi

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

# --- matplotlib (backend only, but required) ---
"$PYTHON_PATH" -c "import matplotlib" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing matplotlib..."
    "$PYTHON_PATH" -m pip install matplotlib
fi


echo ""
echo "All required packages are installed."
echo ""

# Run the Python script
"$PYTHON_PATH" source_Step1_SmartBudgeting_FinancialReporting_ShariSheet-based_v20260509.py

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