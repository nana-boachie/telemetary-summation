# Telemetry Sum

## Overview
This tool processes Excel files containing telemetry data, specifically summing values from the "Raw" column across Excel sheets that share the first 6 characters in their names (e.g., all "VM ACC" sheets grouped together). The summation is organized according to timestamps.

## Features
- Groups Excel sheets by the first 6 characters of their filenames
- Extracts and sums values from the "Raw" column
- Organizes results by timestamps
- Supports sheets with names like "VM ACCRA_CENTRAL-+S" and "VM Afienya"

## Requirements
- Python 3.6+
- Required packages:
  - pandas
  - openpyxl
  - xlrd (for older .xls files)

## Usage
```bash
python sum_telemetry.py [directory_path]
```

Where `directory_path` is the folder containing your Excel files. If no path is provided, the script will use the current directory.

## Output
The script generates a summary Excel file with the aggregated data organized by groups and timestamps.
