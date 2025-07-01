#!/bin/bash

# Setup script for AIG Class List Processor
# This script sets up the Python virtual environment and installs dependencies

echo "Setting up AIG Class List Processor..."

# Check if venv directory exists
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Install dependencies
echo "Installing Python packages..."
pip install --upgrade pip
pip install -r requirements.txt

echo "Setup complete!"
echo ""
echo "To run the application:"
echo "1. Activate the virtual environment: source venv/bin/activate"
echo "2. Run the processor: python aig_processor.py"
echo ""
echo "Make sure your PDF and Excel files are in the current directory:"
echo "- SalemAIGRoster6.24.25.pdf"
echo "- HEINZE of  25-26 Class Lists.xlsx"
