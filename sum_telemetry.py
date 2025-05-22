import os
import sys
import pandas as pd
from collections import defaultdict
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QSpinBox, QTextEdit, 
    QFileDialog, QMessageBox, QStatusBar, QFrame, QScrollArea
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QAction

class TelemetrySumTool(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Default values
        self.prefix_length = 6
        
        # Set up the UI
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface"""
        # Set window properties
        self.setWindowTitle("Excel Telemetry Summation Tool")
        self.setMinimumSize(800, 600)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Title
        title_label = QLabel("Excel Telemetry Summation Tool")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # Input file selection
        input_frame = QFrame()
        input_layout = QVBoxLayout(input_frame)
        
        input_label = QLabel("Input Excel File:")
        input_layout.addWidget(input_label)
        
        input_file_layout = QHBoxLayout()
        self.input_file_edit = QLineEdit()
        self.input_file_edit.setReadOnly(True)
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self.browse_input_file)
        
        input_file_layout.addWidget(self.input_file_edit)
        input_file_layout.addWidget(browse_button)
        input_layout.addLayout(input_file_layout)
        
        main_layout.addWidget(input_frame)
        
        # Prefix length selection
        prefix_frame = QFrame()
        prefix_layout = QHBoxLayout(prefix_frame)
        
        prefix_label = QLabel("Group sheets by first N characters:")
        self.prefix_spinbox = QSpinBox()
        self.prefix_spinbox.setRange(1, 20)
        self.prefix_spinbox.setValue(self.prefix_length)
        
        prefix_layout.addWidget(prefix_label)
        prefix_layout.addWidget(self.prefix_spinbox)
        prefix_layout.addStretch()
        
        main_layout.addWidget(prefix_frame)
        
        # Help text
        help_text = QLabel(
            "This tool will sum 'Raw' column values from Excel sheets that share "
            "the first N characters in their names.\n"
            "Results will be organized according to timestamps."
        )
        help_text.setWordWrap(True)
        main_layout.addWidget(help_text)
        
        # Add stretch to push buttons to the bottom
        main_layout.addStretch()
        
        # Button panel
        button_layout = QHBoxLayout()
        
        self.preview_button = QPushButton("Preview")
        self.preview_button.setMinimumSize(QSize(120, 40))
        self.preview_button.clicked.connect(self.preview_file)
        
        self.process_button = QPushButton("Process Files")
        self.process_button.setMinimumSize(QSize(120, 40))
        self.process_button.clicked.connect(self.process_files)
        
        button_layout.addStretch()
        button_layout.addWidget(self.preview_button)
        button_layout.addWidget(self.process_button)
        button_layout.addStretch()
        
        main_layout.addLayout(button_layout)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status("Ready. Please select an Excel file.")
        
        # Show the window
        self.show()
    
    def browse_input_file(self):
        """Open file dialog to select input Excel file"""
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls);;All Files (*)")
        
        if file_dialog.exec():
            file_path = file_dialog.selectedFiles()[0]
            self.input_file_edit.setText(file_path)
            self.update_status(f"Selected: {os.path.basename(file_path)}")
    
    def auto_generate_output_path(self, input_path):
        """Generate output path based on input path"""
        if not input_path:
            return ""
        
        # Get directory and filename
        directory, filename = os.path.split(input_path)
        name, ext = os.path.splitext(filename)
        
        # Create output filename
        output_filename = f"{name}_summed{ext}"
        output_path = os.path.join(directory, output_filename)
        
        return output_path
    
    def analyze_excel_file(self, file_path, prefix_length=6):
        """Analyze the Excel file and return information about sheet grouping without processing"""
        # Determine file extension
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # Try to detect file type by content if extension is not reliable
        if not file_ext or file_ext not in ['.xlsx', '.xls']:
            with open(file_path, 'rb') as f:
                header = f.read(8)  # Read first 8 bytes to detect file type
                if header.startswith(b'\x50\x4B\x03\x04'):  # PK header (ZIP)
                    file_ext = '.xlsx'
                elif header.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'):  # OLE header
                    file_ext = '.xls'
        
        # Try with appropriate engine based on detected file type
        if file_ext == '.xlsx':
            try:
                xl = pd.ExcelFile(file_path, engine='openpyxl')
            except Exception as e:
                try:
                    xl = pd.ExcelFile(file_path, engine='xlsxwriter')
                except Exception as e2:
                    try:
                        xl = pd.ExcelFile(file_path)  # Let pandas decide
                    except Exception as e3:
                        raise ValueError(f"Could not read Excel file. Error: {str(e3)}")
        elif file_ext == '.xls':
            try:
                xl = pd.ExcelFile(file_path, engine='xlrd')
            except Exception as e:
                try:
                    xl = pd.ExcelFile(file_path)  # Let pandas decide
                except Exception as e2:
                    raise ValueError(f"Could not read .xls file. Error: {str(e2)}")
        else:
            # For unknown extensions, try all engines
            for engine in ['openpyxl', 'xlsxwriter', 'xlrd']:
                try:
                    xl = pd.ExcelFile(file_path, engine=engine)
                    break
                except Exception as e:
                    continue
            else:
                # If all engines fail, try without specifying an engine
                try:
                    xl = pd.ExcelFile(file_path)
                except Exception as e:
                    raise ValueError(f"Could not determine Excel file format. Error: {str(e)}")
        
        # Group sheet names by first N characters (default=6)
        sheet_groups = defaultdict(list)
        for sheet_name in xl.sheet_names:
            if len(sheet_name) >= prefix_length:
                prefix = sheet_name[:prefix_length]
                sheet_groups[prefix].append(sheet_name)
            else:
                # For sheet names shorter than prefix_length, use the entire name
                sheet_groups[sheet_name].append(sheet_name)
        
        # Analyze each sheet for Raw column and timestamp
        sheet_analysis = {}
        for sheet_name in xl.sheet_names:
            try:
                # Try with appropriate engine based on file extension
                if file_ext == '.xlsx':
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5, engine='openpyxl')
                    except Exception as e:
                        try:
                            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5, engine='xlsxwriter')
                        except Exception as e2:
                            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)
                elif file_ext == '.xls':
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5, engine='xlrd')
                    except Exception as e:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)
                else:
                    # For unknown extensions, try all engines
                    for engine in ['openpyxl', 'xlrd', 'xlsxwriter']:
                        try:
                            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5, engine=engine)
                            break
                        except Exception as e:
                            continue
                    else:
                        # If all engines fail, try without specifying an engine
                        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)
                
                has_raw = 'Raw' in df.columns
                
                # Check for timestamp column
                timestamp_col = None
                for col in df.columns:
                    if 'time' in str(col).lower() or 'date' in str(col).lower():
                        timestamp_col = col
                        break
                
                sheet_analysis[sheet_name] = {
                    'has_raw': has_raw,
                    'timestamp_col': timestamp_col,
                    'processable': has_raw and timestamp_col is not None
                }
            except Exception as e:
                sheet_analysis[sheet_name] = {
                    'has_raw': False,
                    'timestamp_col': None,
                    'processable': False,
                    'error': str(e)
                }
        
        return {
            'sheet_groups': sheet_groups,
            'sheet_analysis': sheet_analysis,
            'total_sheets': len(xl.sheet_names),
            'processable_groups': sum(1 for prefix, sheets in sheet_groups.items() 
                                    if any(sheet_analysis.get(sheet, {}).get('processable', False) for sheet in sheets))
        }
    
    def process_excel_file(self, file_path, output_path, prefix_length=6):
        """Process Excel file and create output file with summed telemetry data"""
        # Determine file extension
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # Try to detect file type by content if extension is not reliable
        if not file_ext or file_ext not in ['.xlsx', '.xls']:
            with open(file_path, 'rb') as f:
                header = f.read(8)  # Read first 8 bytes to detect file type
                if header.startswith(b'\x50\x4B\x03\x04'):  # PK header (ZIP)
                    file_ext = '.xlsx'
                elif header.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'):  # OLE header
                    file_ext = '.xls'
        
        # Try with appropriate engine based on detected file type
        if file_ext == '.xlsx':
            try:
                xl = pd.ExcelFile(file_path, engine='openpyxl')
            except Exception as e:
                try:
                    xl = pd.ExcelFile(file_path, engine='xlsxwriter')
                except Exception as e2:
                    try:
                        xl = pd.ExcelFile(file_path)  # Let pandas decide
                    except Exception as e3:
                        raise ValueError(f"Could not read Excel file. Error: {str(e3)}")
        elif file_ext == '.xls':
            try:
                xl = pd.ExcelFile(file_path, engine='xlrd')
            except Exception as e:
                try:
                    xl = pd.ExcelFile(file_path)  # Let pandas decide
                except Exception as e2:
                    raise ValueError(f"Could not read .xls file. Error: {str(e2)}")
        else:
            # For unknown extensions, try all engines
            for engine in ['openpyxl', 'xlsxwriter', 'xlrd']:
                try:
                    xl = pd.ExcelFile(file_path, engine=engine)
                    break
                except Exception as e:
                    continue
            else:
                # If all engines fail, try without specifying an engine
                try:
                    xl = pd.ExcelFile(file_path)
                except Exception as e:
                    raise ValueError(f"Could not determine Excel file format. Error: {str(e)}")
        
        # Group sheet names by first N characters (default=6)
        sheet_groups = defaultdict(list)
        for sheet_name in xl.sheet_names:
            if len(sheet_name) >= prefix_length:
                prefix = sheet_name[:prefix_length]
                sheet_groups[prefix].append(sheet_name)
            else:
                # For sheet names shorter than prefix_length, use the entire name
                sheet_groups[sheet_name].append(sheet_name)
        
        # Process each group of sheets
        results = {}
        
        for prefix, sheets in sheet_groups.items():
            if len(sheets) > 0:
                # Initialize a DataFrame to store combined data
                combined_data = pd.DataFrame()
                
                for sheet in sheets:
                    try:
                        # Read data from the sheet with appropriate engine
                        if file_ext == '.xlsx':
                            try:
                                df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
                            except Exception as e:
                                try:
                                    df = pd.read_excel(file_path, sheet_name=sheet, engine='xlsxwriter')
                                except Exception as e2:
                                    df = pd.read_excel(file_path, sheet_name=sheet)
                        elif file_ext == '.xls':
                            try:
                                df = pd.read_excel(file_path, sheet_name=sheet, engine='xlrd')
                            except Exception as e:
                                df = pd.read_excel(file_path, sheet_name=sheet)
                        else:
                            # For unknown extensions, try all engines
                            for engine in ['openpyxl', 'xlrd', 'xlsxwriter']:
                                try:
                                    df = pd.read_excel(file_path, sheet_name=sheet, engine=engine)
                                    break
                                except Exception as e:
                                    continue
                            else:
                                # If all engines fail, try without specifying an engine
                                df = pd.read_excel(file_path, sheet_name=sheet)
                        
                        # Check if 'Raw' column exists
                        if 'Raw' in df.columns:
                            # Ensure there's a timestamp column
                            timestamp_col = None
                            for col in df.columns:
                                if 'time' in col.lower() or 'date' in col.lower():
                                    timestamp_col = col
                                    break
                            
                            if timestamp_col:
                                # Rename for consistency
                                df = df.rename(columns={timestamp_col: 'Timestamp'})
                                
                                # Select only the timestamp and Raw columns
                                subset_df = df[['Timestamp', 'Raw']].copy()
                                
                                # Append to combined data
                                combined_data = pd.concat([combined_data, subset_df])
                    except Exception as e:
                        print(f"Error processing sheet {sheet}: {str(e)}")
                
                # Process combined data if we have any
                if not combined_data.empty:
                    # Sort by timestamp
                    combined_data = combined_data.sort_values('Timestamp')
                    
                    # Store in results
                    results[prefix] = combined_data
        
        # Now write each group to a separate sheet in the output Excel file
        try:
            # Try with openpyxl first (for .xlsx files)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Write each group's data to a separate sheet
                for prefix, data in results.items():
                    if not data.empty:
                        data.to_excel(writer, sheet_name=prefix[:31], index=False)  # Sheet name max 31 chars
        except Exception as e:
            # Fallback to xlsxwriter if openpyxl fails
            try:
                with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                    # Write each group's data to a separate sheet
                    for prefix, data in results.items():
                        if not data.empty:
                            data.to_excel(writer, sheet_name=prefix[:31], index=False)  # Sheet name max 31 chars
            except Exception as e2:
                # Final fallback to default engine
                with pd.ExcelWriter(output_path) as writer:
                    # Write each group's data to a separate sheet
                    for prefix, data in results.items():
                        if not data.empty:
                            data.to_excel(writer, sheet_name=prefix[:31], index=False)  # Sheet name max 31 chars
        
        return {
            'processed_groups': len(results),
            'total_groups': len(sheet_groups),
            'output_path': output_path
        }
    
    def preview_file(self):
        """Preview sheet grouping in the input file"""
        input_file = self.input_file_edit.text()
        
        if not input_file:
            self.update_status("No input file selected", True)
            QMessageBox.critical(self, "Error", "Please select an input file first")
            return
        
        if not os.path.exists(input_file):
            self.update_status(f"File not found: {input_file}", True)
            QMessageBox.critical(self, "Error", f"File not found: {input_file}")
            return
            
        try:
            # Try to read the file to check if it's a valid Excel file
            with open(input_file, 'rb') as f:
                header = f.read(8)
                if not (header.startswith(b'\x50\x4B\x03\x04') or  # ZIP (XLSX)
                       header.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1')):  # OLE (XLS)
                    raise ValueError("Not a valid Excel file (invalid file signature)")
        except Exception as e:
            self.update_status(f"Error reading file: {str(e)}", True)
            QMessageBox.critical(self, "Error", f"Could not read file: {str(e)}")
            return
        
        self.update_status(f"Analyzing file: {os.path.basename(input_file)}...")
        
        try:
            # Get the selected prefix length
            prefix_length = self.prefix_spinbox.value()
            analysis = self.analyze_excel_file(input_file, prefix_length)
            
            # Create a preview dialog
            preview_dialog = QMainWindow(self)
            preview_dialog.setWindowTitle(f"Preview: {os.path.basename(input_file)}")
            preview_dialog.setMinimumSize(700, 500)
            
            # Central widget and layout
            dialog_central = QWidget()
            preview_dialog.setCentralWidget(dialog_central)
            dialog_layout = QVBoxLayout(dialog_central)
            
            # Create a text edit for displaying information
            text_edit = QTextEdit()
            text_edit.setReadOnly(True)
            
            # Create a scroll area
            scroll_area = QScrollArea()
            scroll_area.setWidget(text_edit)
            scroll_area.setWidgetResizable(True)
            
            dialog_layout.addWidget(scroll_area)
            
            # Insert analysis results
            text_content = f"File: {input_file}\n"
            text_content += f"Total sheets: {analysis['total_sheets']}\n"
            text_content += f"Processable groups: {analysis['processable_groups']}\n\n"
            text_content += f"Sheet grouping (prefix length: {prefix_length}):\n"
            
            for prefix, sheets in analysis['sheet_groups'].items():
                text_content += f"\nGroup: '{prefix}'\n"
                text_content += f"Sheets in this group: {len(sheets)}\n"
                
                for sheet in sheets:
                    sheet_info = analysis['sheet_analysis'][sheet]
                    processable = sheet_info['processable']
                    status = "Will be processed" if processable else "Will be SKIPPED"
                    text_content += f"  - {sheet}: {status}\n"
                    
                    if not processable:
                        if not sheet_info['has_raw']:
                            text_content += f"    Missing 'Raw' column\n"
                        if not sheet_info['timestamp_col']:
                            text_content += f"    Missing timestamp column\n"
                        if 'error' in sheet_info:
                            text_content += f"    Error: {sheet_info['error']}\n"
                    else:
                        text_content += f"    Timestamp column: {sheet_info['timestamp_col']}\n"
            
            text_edit.setText(text_content)
            
            # Add close button
            close_button = QPushButton("Close")
            close_button.clicked.connect(preview_dialog.close)
            dialog_layout.addWidget(close_button)
            
            # Show the dialog
            preview_dialog.show()
            
            self.update_status(f"Preview generated for {os.path.basename(input_file)}")
            
        except Exception as e:
            error_msg = str(e)
            self.update_status(f"Error during preview: {error_msg}", True)
            QMessageBox.critical(self, "Error", f"An error occurred during preview: {error_msg}")
    
    def process_files(self):
        """Process the input file and generate output"""
        input_file = self.input_file_edit.text()
        
        if not input_file:
            self.update_status("No input file selected", True)
            QMessageBox.critical(self, "Error", "Please select an input file first")
            return
            
        if not os.path.exists(input_file):
            self.update_status(f"File not found: {input_file}", True)
            QMessageBox.critical(self, "Error", f"File not found: {input_file}")
            return
            
        try:
            # Try to read the file to check if it's a valid Excel file
            with open(input_file, 'rb') as f:
                header = f.read(8)
                if not (header.startswith(b'\x50\x4B\x03\x04') or  # ZIP (XLSX)
                       header.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1')):  # OLE (XLS)
                    raise ValueError("Not a valid Excel file (invalid file signature)")
        except Exception as e:
            self.update_status(f"Error reading file: {str(e)}", True)
            QMessageBox.critical(self, "Error", f"Could not read file: {str(e)}")
            return
        
        # Generate output path
        output_file = self.auto_generate_output_path(input_file)
        
        # Confirm if output file exists
        if os.path.exists(output_file):
            reply = QMessageBox.question(
                self, 
                "Warning", 
                f"Output file {os.path.basename(output_file)} already exists. Overwrite?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.No:
                return
        
        self.update_status(f"Processing {os.path.basename(input_file)}...")
        
        try:
            # Get prefix length
            prefix_length = self.prefix_spinbox.value()
            
            # Process the file
            result = self.process_excel_file(input_file, output_file, prefix_length)
            
            self.update_status(
                f"Processing complete. Processed {result['processed_groups']} group(s). "
                f"Output saved to {os.path.basename(output_file)}"
            )
            
            # Ask if user wants to open the output file
            reply = QMessageBox.question(
                self,
                "Success",
                f"Processing complete. Processed {result['processed_groups']} of {result['total_groups']} groups.\n\n"
                f"Output saved to: {output_file}\n\n"
                f"Open output file?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                # Open output file with system default application
                if sys.platform == 'darwin':  # macOS
                    os.system(f"open '{output_file}'")
                elif sys.platform == 'win32':  # Windows
                    os.startfile(output_file)
                else:  # Linux or other
                    os.system(f"xdg-open '{output_file}'")
                    
        except Exception as e:
            error_msg = str(e)
            self.update_status(f"Error during processing: {error_msg}", True)
            QMessageBox.critical(self, "Error", f"An error occurred during processing: {error_msg}")
    
    def update_status(self, message, is_error=False):
        """Update the status bar with message"""
        self.status_bar.showMessage(message)
        if is_error:
            print(f"ERROR: {message}")
        else:
            print(message)


def main():
    app = QApplication(sys.argv)
    ex = TelemetrySumTool()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
