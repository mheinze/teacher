# AIG Class List Processor

This Python application processes class lists from PDF files and Word documents, then combines them with Excel spreadsheets to generate updated class lists showing AIG (Academically/Intellectually Gifted) status.

## Features

- Extracts student data from PDF files with columns: Name, Student Id, Grade, Reading, Math
- Extracts student data from Word documents with table format: Name, Reading, Math  
- Updates the original Excel spreadsheet by adding AIG columns to each sheet
- **Generates a second Excel file with only AIG students (removes rows where AIG status is None)**
- **Generates a third Excel file listing students from PDF/Word who are not found in Excel spreadsheet**
- Handles grade, track, and classroom information from Excel worksheet first rows
- Uses only the first two columns from Excel sheets for student names
- **Preserves original Excel sheet colors/formatting**
- Color codes student rows based on AIG status:
  - **Light Blue (ADD8E6)**: AIG Math only
  - **Orange**: AIG Reading only  
  - **Yellow**: Both AIG Math and Reading
- **Provides detailed statistics on standard output:**
  - Total number of students with AIG status
  - Number of students who are only TD (not AG, IG, or AIG)
  - Breakdown by Math only, Reading only, and Both

## Key Requirements

- **PDF Structure**: Expects columns "Name Student Id Grade Reading Math"
- **Word Structure**: Table with three columns - Name, Reading, Math (TD means AIG)
- **Excel Structure**: First row contains grade/track/classroom info, uses only first two columns
- **Name Handling**: 
  - PDF: "Last, First" format
  - Word: "Last, First" or "First Last" format (automatically detected)
  - Excel: Separate LASTNAME/FIRSTNAME columns
- **Data Source**: Uses PDF and Word data for AIG status (not Excel data)
- **AIG Recognition**: Treats "TD", "AG", "IG", and "AIG" as AIG for both subjects
- **Output**: Single updated Excel file with AIG columns added to each sheet
- **Cleanup**: Removes all temporary files, keeping only the final Excel spreadsheet

## Setup

### Prerequisites
- Python 3.7 or higher
- The following files in the `input/` directory:
  - `input/SalemAIGRoster6.24.25.pdf` (AIG roster PDF)
  - `input/HEINZE of  25-26 Class Lists.xlsx` (Class lists Excel file)

### Installation

1. **Clone or download the project files**

2. **Run the setup script:**
   ```bash
   chmod +x setup.sh
   ./setup.sh
   ```

   Or manually set up:
   ```bash
   # Create virtual environment
   python3 -m venv venv
   
   # Activate virtual environment
   source venv/bin/activate
   
   # Install dependencies
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage

1. **Activate the virtual environment:**
   ```bash
   source venv/bin/activate
   ```

2. **Place your input files in the `input/` directory:**
   - `SalemAIGRoster6.24.25.pdf` (PDF with AIG roster)
   - `HEINZE of  25-26 Class Lists.xlsx` (Excel with class lists)
   - `TD from Finch WCPSS file.docx` (Word document with additional AIG students - optional)

3. **Run the main processor:**
   ```bash
   python aig_processor.py
   ```

4. **Check the output in the `output/` directory:**
   - Updated Excel file: `output/updated_class_lists.xlsx` 
   - AIG-only Excel file: `output/updated_class_lists_AIG_Only.xlsx`
   - Missing students file: `output/students_not_in_excel.xlsx`
   - Contains all original sheets with added AIG columns
   - Students are color-coded based on AIG status
   - Sheet colors from original Excel file are preserved
   - All temporary files are automatically removed
   - Statistics are displayed on standard output including missing students analysis

### Alternative Usage Options

**Test Mode (creates separate test output):**
```bash
python test_processor.py
```

**Custom Configuration:**
```bash
python example_usage.py
```

**Batch Processing (process multiple file combinations):**
```bash
python batch_processor.py
```

## File Structure

```
teacher/
├── venv/                          # Virtual environment
├── output/                        # Generated reports (created automatically)
├── aig_processor.py              # Main application
├── test_processor.py             # Test script with detailed output
├── example_usage.py              # Custom configuration example
├── batch_processor.py            # Batch processing for multiple files
├── requirements.txt              # Python dependencies
├── setup.sh                     # Setup script
├── README.md                    # This file
├── SalemAIGRoster6.24.25.pdf   # AIG roster (your file)
└── HEINZE of  25-26 Class Lists.xlsx  # Class lists (your file)
```

## How It Works

1. **PDF Processing**: The application extracts student names from the AIG roster PDF, identifying which students are in AIG Math and/or AIG Reading programs.

2. **Excel Processing**: It reads all worksheets from the class lists Excel file, treating each worksheet as a separate classroom.

3. **Data Matching**: For each student in each classroom, it checks if they appear in the AIG lists.

4. **Report Generation**: Creates individual Excel files for each classroom AND a combined Excel file with:
   - Individual files: One Excel file per classroom with original student data plus AIG status
   - Combined file: Single Excel workbook with separate sheets for each teacher/classroom
   - AIG Math status (Yes/No)
   - AIG Reading status (Yes/No)
   - AIG Status summary
   - Color coding for easy visual identification

## Customization

### Modifying File Paths
If your files have different names, edit the `main()` function in `aig_processor.py`:

```python
def main():
    # Update these paths to match your files
    pdf_file = "your_aig_roster.pdf"
    excel_file = "your_class_lists.xlsx"
```

### Adjusting Colors
To change the color scheme, modify the `colors` dictionary in the `__init__` method:

```python
self.colors = {
    'math': PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid'),     # Light Blue
    'reading': PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid'),  # Orange  
    'both': PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')     # Yellow
}
```

## Troubleshooting

### Common Issues

1. **Import Errors**: Make sure you've activated the virtual environment and installed all dependencies.

2. **File Not Found**: Ensure your PDF and Excel files are in the correct location and have the exact names specified in the code.

3. **Name Matching Issues**: The application uses fuzzy name matching. If students aren't being matched correctly, check that the names in both files are formatted consistently.

### Logging

The application includes detailed logging. Check the console output for information about:
- How many AIG students were found
- How many classrooms were processed
- Where output files were saved

## Dependencies

- `pandas`: Excel file processing and data manipulation
- `openpyxl`: Excel file reading/writing with formatting
- `PyPDF2`: PDF text extraction
- `xlsxwriter`: Additional Excel formatting capabilities

## License

This project is for educational use.
