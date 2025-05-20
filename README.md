# Telemetry Summation Tool

## Overview
A cross-platform tool for processing Excel files containing telemetry data. The tool groups Excel sheets that share common name prefixes and combines their data. It features a modern PyQt6 interface and works seamlessly on both macOS and Windows.

## Features
- **Generic Column Selection**: Select any columns of interest to process, not just limited to "Raw"
- **Flexible Grouping**: Group Excel sheets by a configurable number of prefix characters in their names
- **Timestamp-Based Organization**: Organize data based on any selected timestamp column
- **Sum or Preserve**: Option to sum values or keep individual data points
- **Cross-Platform**: Works identically on macOS and Windows
- **Visual Preview**: See which sheets will be processed before running
- **Sheet Tracking**: Keeps track of which source sheets contributed to each data point

## Versions

### Generic Version (Recommended)
The most flexible version that allows selection of any columns:
```bash
python sum_telemetry_generic.py
```

### PyQt Basic Version
A version that works like the original but with a better UI:
```bash
python sum_telemetry_pyqt.py
```

## Requirements
- Python 3.6+
- Required packages (automatically installed via requirements.txt):
  - pandas
  - openpyxl
  - PyQt6

## Installation

```bash
# Clone the repository
git clone https://github.com/YourUsername/telemetry-summation.git
cd telemetry-summation

# Create and activate virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

1. Launch the application:
   ```bash
   python sum_telemetry_generic.py
   ```

2. Browse for your Excel file
3. Click "Analyze" to extract available columns
4. Select columns to process and a timestamp column (optional)
5. Preview results or process the file

## Output
The tool generates a summary Excel file with the aggregated data organized by groups and timestamps. The output filename will be the same as the input with "_summed" appended.
