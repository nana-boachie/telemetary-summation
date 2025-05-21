# Telemetry Data Processing Suite

## Overview
A collection of cross-platform tools for processing and analyzing telemetry data from Excel files. The suite includes tools for both individual file processing and automated report generation across multiple files.

## Tools

### 1. Telemetry Summation Tool
Process individual Excel files by grouping sheets with common name prefixes and combining their data.

**Features:**
- **Generic Column Selection**: Process any columns of interest, not just limited to "Raw"
- **Flexible Grouping**: Group Excel sheets by configurable name prefixes
- **Timestamp-Based Organization**: Organize data using any timestamp column
- **Sum or Preserve**: Option to sum values or keep individual data points
- **Cross-Platform**: Works on both macOS and Windows
- **Visual Preview**: Preview which sheets will be processed
- **Sheet Tracking**: Track source sheets for each data point

### 2. Annual Report Generator
Automatically organize and process monthly telemetry data files into a structured annual report.

**Features:**
- **Automatic File Organization**: Automatically sorts files by year and month
- **Monthly Data Processing**: Processes multiple Excel files for a given month
- **Annual Report Generation**: Combines monthly data into comprehensive annual reports
- **Intelligent Date Detection**: Automatically detects date information from filenames
- **Duplicate Prevention**: Prevents processing the same file multiple times
- **User-Friendly Interface**: Simple PyQt6-based GUI for easy operation

## Requirements
- Python 3.9+
- Required packages (install via requirements.txt):
  - pandas
  - openpyxl
  - PyQt6
  - numpy

## Installation

```bash
# Clone the repository
git clone https://github.com/nana-boachie/telemetary-summation.git
cd telemetary-summation

# Create and activate virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Telemetry Summation Tool
```bash
python sum_telemetry_generic.py
```

1. Launch the application
2. Browse for your Excel file
3. Click "Analyze" to extract available columns
4. Select columns to process and a timestamp column (optional)
5. Preview results or process the file

### Annual Report Generator
```bash
python annual_report_generator.py
```

1. Launch the application
2. Select the source directory containing your monthly Excel files
3. Choose a destination directory for the organized files and reports
4. Click "Generate Report" to process the files

## Output

### Telemetry Summation Tool
Generates a summary Excel file with aggregated data organized by groups and timestamps. The output filename will be the same as the input with "_summed" appended.

### Annual Report Generator
Creates a structured directory layout:
```
output_directory/
├── organized/
│   ├── 2024/
│   │   ├── 01_January/
│   │   ├── 02_February/
│   │   └── ...
│   └── 2025/
└── reports/
    ├── 2024_Annual_Report.xlsx
    └── ...
```

## Building from Source

Pre-built executables are available in the GitHub releases section. To build from source:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Build the applications:
   ```bash
   # Build Telemetry Summation Tool
   pyinstaller --name="Telemetry Summation Tool" --windowed --onedir --add-data="requirements.txt:." sum_telemetry_generic.py
   
   # Build Annual Report Generator
   pyinstaller --name="Annual Report Generator" --windowed --onedir --add-data="requirements.txt:." --add-data="data_organizer.py:." annual_report_generator.py
   ```

## License
This project is licensed under the MIT License - see the LICENSE file for details.
