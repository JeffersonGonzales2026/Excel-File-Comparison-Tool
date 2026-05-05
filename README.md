# Excel File Comparison Tool

A Streamlit-based automation tool to compare two Excel files and highlight the differences.

## Features

- 📊 Upload and compare two Excel files
- 🔍 Automatic detection of structural changes (added/removed columns, row count changes)
- 🎯 Detailed cell-level comparison highlighting data changes
- 📋 Generates a comparison report with all changes highlighted
- 💾 Export results to an Excel file with color-coded differences
- 📈 Visual metrics showing differences between files

## Setup Instructions

### 1. Install Python (if not already installed)
Download Python from [python.org](https://www.python.org/downloads/) and ensure you select "Add Python to PATH" during installation.

### 2. Install Required Dependencies

Open PowerShell or Command Prompt and navigate to the folder:

```bash
cd "c:\Users\SPM\Documents\Jefferson Gonzales\Ms. Ronalyn\TPAP\Automation\ComRevAuto"
```

Then install the requirements:

```bash
pip install -r requirements.txt
```

### 3. Run the Application

From the same folder, run:

```bash
streamlit run ComRevAuto.py
```

This will open a browser window with the Streamlit app.

## How to Use

1. **Upload Files:**
   - In the left sidebar, upload your Original Excel file (data from Sheet 2 will be used)
   - Upload your Revised Excel file

2. **View Comparison:**
   - The tool automatically displays both files side-by-side
   - Metrics show row count differences
   - View the original and revised data

3. **Analyze Changes:**
   - The app identifies structural changes (column additions/removals)
   - Shows detailed data changes with row and column references
   - Displays which values changed from original to revised

4. **Generate Report:**
   - Click "Generate Highlighted Excel Report" button
   - Download the comparison_report.xlsx file
   - The report contains:
     - Original sheet with all original data
     - Revised sheet with all revised data
     - Changes sheet highlighting all detected differences

## File Structure

```
ComRevAuto/
├── ComRevAuto.py          # Main Streamlit application
├── requirements.txt       # Python package dependencies
└── README.md             # This file
```

## Requirements

- Python 3.8 or higher
- streamlit
- pandas
- openpyxl
- numpy

## Troubleshooting

**Issue:** "No module named 'streamlit'"
- **Solution:** Make sure you ran `pip install -r requirements.txt`

**Issue:** File upload not working
- **Solution:** Ensure files are valid Excel files (.xlsx or .xls format)

**Issue:** "Sheet 2 not found"
- **Solution:** Your original file must have at least 2 sheets. The tool uses the second sheet (index 1) for the original data.

## Notes

- The tool reads the **second sheet (Sheet 2)** from the original file
- The **first sheet** from the revised file is used for comparison
- Maximum 10 changes are displayed in the summary; full details are in the "Detailed Changes Report" section
- The comparison is cell-by-cell and case-sensitive
