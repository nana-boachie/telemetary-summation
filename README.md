# Telemetry Analysis Suite

## Overview
A simplified, comprehensive cross-platform application for processing, analyzing, and organizing telemetry data from Excel files. The suite combines multiple tools into a single, user-friendly interface with the following capabilities:

## Features

### 1. Telemetry Summation Tool
Process individual Excel files by grouping sheets with common name prefixes and combining their data.

**Key Features:**
- **Column Selection**: Process any columns of interest from your data
- **Intelligent Grouping**: Automatically group Excel sheets by configurable name prefixes
- **Timestamp Handling**: Organize data using any timestamp column
- **Flexible Aggregation**: Choose to sum values or keep individual data points
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **Visual Preview**: Preview which sheets will be processed before execution
- **Source Tracking**: Maintain traceability with source sheet information for each data point

### 2. Generic Telemetry Tool
A versatile tool for processing telemetry data with customizable column selection and processing options.

**Key Features:**
- **Custom Column Selection**: Choose exactly which columns to process
- **Flexible Grouping**: Group data by any number of leading characters in sheet names
- **Value Processing**: Sum values or keep them separate as needed
- **Timestamp Support**: Use any datetime column for time-based organization
- **Preview Functionality**: See how your data will be processed before final execution

### 3. Data Organizer
Automatically organize and process telemetry data files into a structured directory system.

**Key Features:**
- **Automatic Organization**: Sort files by year and month into a logical folder structure
- **Batch Processing**: Process multiple files at once with consistent settings
- **Date Detection**: Automatically detect dates from filenames or use manual override
- **Duplicate Prevention**: Skip already processed files to avoid duplicates
- **Detailed Logging**: Comprehensive logs of all operations for auditing

### 4. Report Generator
Create comprehensive reports from your organized telemetry data.

**Key Features:**
- **Annual Reports**: Generate yearly summaries from monthly data
- **Customizable Outputs**: Choose which metrics and visualizations to include
- **Flexible Time Ranges**: Generate reports for any date range
- **Multiple Formats**: Export reports in various formats (CSV, Excel, PDF)
- **Template Support**: Use custom templates for consistent report formatting

## Requirements
- Python 3.9+
- Required packages (install via requirements.txt):
  - pandas
  - openpyxl
  - PyQt6
  - numpy
  - xlrd (for .xls file support)
  - xlsxwriter

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/nana-boachie/telemetary-summation.git
   cd telemetary-summation
   ```

2. **Set up a virtual environment**
   ```bash
   # On Windows
   python -m venv venv
   .\venv\Scripts\activate
   
   # On macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## Quick Start

1. **Launch the application**:
   ```bash
   python telemetry_analysis_suite.py
   ```

2. **The main window will open with the following tabs**:
   - **Data Manager**: For organizing files and generating reports (main tab)
   - **Telemetry Sum**: For processing individual Excel files
   - **Generic Telemetry**: For advanced processing with custom column selection

## Detailed Usage

### 1. Telemetry Sum Tool

1. Click on the "Telemetry Sum" tab
2. Click "Browse" to select an Excel file
3. Click "Preview" to see how sheets will be grouped
4. Adjust the "Group sheets by first N characters" if needed
5. Click "Process Files" to generate the output
6. The processed file will be saved in the same directory as the input file

### 2. Generic Telemetry Tool

1. Click on the "Generic Telemetry" tab
2. Click "Browse" to select an Excel file
3. Click "Analyze" to load the available columns
4. Select columns to process from the "Available Columns" list
5. Use the arrow buttons to move them to "Selected Value Columns"
6. Choose a timestamp column from the dropdown
7. Configure the grouping prefix length
8. Choose whether to sum values or keep them separate
9. Click "Preview" to verify the processing
10. Click "Process Files" to generate the output

### 3. Data Manager

1. The "Data Manager" tab provides two sub-tabs:
   - **Organize Files**: For organizing telemetry data files
   - **Generate Reports**: For creating annual reports

#### Organizing Files:
1. Set the "Data Storage Location" where organized files will be saved
2. Click "Browse..." next to "Source Folder" to select files to organize
3. (Optional) Specify the year and month for the files
4. Click "Organize Files" to begin processing
5. View the progress in the log area
6. Organized files will be placed in `[Data Storage Location]/[Year]/[Month]/`

#### Report Generation:
1. In the "Generate Reports" sub-tab, select the year for the report
2. (Optional) Specify an output file location
3. Click "Check Available Data" to see which months have data available
4. Click "Generate Annual Report" to create the report
5. When complete, you'll be prompted with the report location

## Command Line Usage

You can also use the tools from the command line:

```bash
# Process a single file with default settings
python sum_telemetry.py input_file.xlsx

# Process with custom settings
python sum_telemetry.py input_file.xlsx --prefix 8 --output output_file.xlsx

# For help with command line options
python sum_telemetry.py --help
```

## Configuration

You can customize the application by creating a `config.ini` file in the application directory. Example:

```ini
[directories]
data_dir = ./data
reports_dir = ./reports

[processing]
default_prefix_length = 6
auto_detect_dates = true
```
```

## Command Line Usage

While the main application provides a complete GUI interface, you can also use individual components from the command line if needed:

```bash
# Process a single file with the Telemetry Summation Tool
python sum_telemetry.py input_file.xlsx

# Get help on command line options
python sum_telemetry.py --help
```

However, using the main GUI application is recommended for most users:

```bash
python telemetry_analysis_suite.py
```

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

2. Build the application:
   ```bash
   pyinstaller --name="Telemetry Analysis Suite" --windowed --onedir --add-data="requirements.txt;." --add-data="data_organizer.py;." --add-data="assets;assets" telemetry_analysis_suite.py
   ```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with Python and PyQt6
- Uses pandas for data manipulation
- Icons by [Font Awesome](https://fontawesome.com/)

## Support

For support, please open an issue on the [GitHub repository](https://github.com/nana-boachie/telemetary-summation/issues).

## Additional Resources

- [Telemetry Summation Tool Documentation](https://nana-boachie.github.io/telemetary-summation/telemetry_sum.html)
- [Data Organization Tool Documentation](https://nana-boachie.github.io/telemetary-summation/data_organizer.html)
- [Report Generation Tool Documentation](https://nana-boachie.github.io/telemetary-summation/report_generator.html)
