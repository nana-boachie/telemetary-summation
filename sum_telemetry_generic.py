import os
import sys
import pandas as pd
from collections import defaultdict
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QSpinBox, QTextEdit, 
    QFileDialog, QMessageBox, QStatusBar, QFrame, QScrollArea,
    QListWidget, QListWidgetItem, QSplitter, QComboBox, QGroupBox,
    QCheckBox
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QAction

class GenericTelemetrySumTool(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Default values
        self.prefix_length = 6
        self.available_columns = []
        self.selected_value_columns = []
        self.timestamp_column = None
        
        # Set up the UI
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface"""
        # Set window properties
        self.setWindowTitle("Generic Telemetry Summation Tool")
        self.setMinimumSize(900, 700)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Title
        title_label = QLabel("Generic Telemetry Summation Tool")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # Description
        description = QLabel(
            "This tool helps you summarize telemetry data across multiple Excel sheets. "
            "Select an Excel file, choose which columns to process, and the tool will combine "
            "data from sheets with similar names."
        )
        description.setWordWrap(True)
        main_layout.addWidget(description)
        
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
        analyze_button = QPushButton("Analyze")
        analyze_button.clicked.connect(self.analyze_columns)
        
        input_file_layout.addWidget(self.input_file_edit)
        input_file_layout.addWidget(browse_button)
        input_file_layout.addWidget(analyze_button)
        input_layout.addLayout(input_file_layout)
        
        main_layout.addWidget(input_frame)
        
        # Create a splitter for column selection
        column_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left side: Available columns
        available_group = QGroupBox("Available Columns")
        available_layout = QVBoxLayout(available_group)
        
        self.available_list = QListWidget()
        self.available_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        available_layout.addWidget(self.available_list)
        
        add_button = QPushButton("Add →")
        add_button.clicked.connect(self.add_selected_columns)
        available_layout.addWidget(add_button)
        
        # Right side: Selected value columns
        selected_group = QGroupBox("Selected Value Columns")
        selected_layout = QVBoxLayout(selected_group)
        
        self.selected_list = QListWidget()
        self.selected_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        selected_layout.addWidget(self.selected_list)
        
        remove_button = QPushButton("← Remove")
        remove_button.clicked.connect(self.remove_selected_columns)
        selected_layout.addWidget(remove_button)
        
        # Add both panels to the splitter
        column_splitter.addWidget(available_group)
        column_splitter.addWidget(selected_group)
        column_splitter.setStretchFactor(0, 1)
        column_splitter.setStretchFactor(1, 1)
        
        # Timestamp column selection
        timestamp_frame = QFrame()
        timestamp_layout = QVBoxLayout(timestamp_frame)
        
        timestamp_label = QLabel("Timestamp Column:")
        timestamp_layout.addWidget(timestamp_label)
        
        self.timestamp_combo = QComboBox()
        timestamp_layout.addWidget(self.timestamp_combo)
        
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
        
        # Checkbox for sum vs individual values
        self.sum_checkbox = QCheckBox("Sum values (unchecked = keep individual values)")
        self.sum_checkbox.setChecked(True)
        
        # Container for column selection components
        columns_container = QWidget()
        columns_layout = QVBoxLayout(columns_container)
        columns_layout.addWidget(column_splitter)
        columns_layout.addWidget(timestamp_frame)
        columns_layout.addWidget(prefix_frame)
        columns_layout.addWidget(self.sum_checkbox)
        
        main_layout.addWidget(columns_container)
        
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
        self.update_status("Ready. Please select an Excel file and analyze it to get available columns.")
        
        # Disable buttons until file is analyzed
        self.preview_button.setEnabled(False)
        self.process_button.setEnabled(False)
        
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
            
            # Reset column selections
            self.available_list.clear()
            self.selected_list.clear()
            self.timestamp_combo.clear()
            self.available_columns = []
            self.selected_value_columns = []
            self.timestamp_column = None
            
            # Prompt user to analyze the file
            QMessageBox.information(
                self, 
                "Analyze File", 
                "Please click the 'Analyze' button to extract available columns from this file."
            )
    
    def analyze_columns(self):
        """Analyze the Excel file to extract available columns"""
        input_file = self.input_file_edit.text()
        
        if not input_file:
            self.update_status("No input file selected", True)
            QMessageBox.critical(self, "Error", "Please select an input file first")
            return
        
        if not os.path.exists(input_file):
            self.update_status(f"File not found: {input_file}", True)
            QMessageBox.critical(self, "Error", f"File not found: {input_file}")
            return
        
        self.update_status(f"Analyzing file: {os.path.basename(input_file)}...")
        
        try:
            # First, try to determine the file type and use the appropriate engine
            file_ext = os.path.splitext(input_file)[1].lower()
            df = None
            
            # Try different engines based on file extension
            if file_ext == '.xlsx':
                try:
                    df = pd.read_excel(input_file, nrows=5, engine='openpyxl')
                except Exception as e:
                    try:
                        df = pd.read_excel(input_file, nrows=5, engine='xlrd')
                    except:
                        df = pd.read_excel(input_file, nrows=5)
            elif file_ext == '.xls':
                try:
                    df = pd.read_excel(input_file, nrows=5, engine='xlrd')
                except Exception as e:
                    try:
                        df = pd.read_excel(input_file, nrows=5, engine='openpyxl')
                    except:
                        df = pd.read_excel(input_file, nrows=5)
            else:
                # For other extensions or no extension, try all engines
                try:
                    df = pd.read_excel(input_file, nrows=5, engine='openpyxl')
                except Exception as e1:
                    try:
                        df = pd.read_excel(input_file, nrows=5, engine='xlrd')
                    except Exception as e2:
                        try:
                            df = pd.read_excel(input_file, nrows=5)
                        except Exception as e3:
                            raise Exception(f"Could not read file. Tried openpyxl, xlrd, and default engines. Last error: {str(e3)}")
            
            if df is None or df.empty:
                raise Exception("The Excel file is empty or could not be read")
            
            # Clear previous data
            self.available_columns = df.columns.tolist()
            self.available_list.clear()
            self.selected_list.clear()
            self.timestamp_combo.clear()
            self.selected_value_columns = []
            
            # Add to available columns list
            for col in self.available_columns:
                self.available_list.addItem(str(col))
            
            # Add to timestamp dropdown (add a "None" option first)
            self.timestamp_combo.addItem("-- Select Timestamp Column --")
            for col in self.available_columns:
                self.timestamp_combo.addItem(str(col))
            
            # Try to auto-select timestamp column
            for i, col in enumerate(self.available_columns):
                col_str = str(col).lower()
                if 'time' in col_str or 'date' in col_str:
                    self.timestamp_combo.setCurrentIndex(i + 1)  # +1 because of the "None" option
                    break
            
            # Enable preview and process buttons
            self.preview_button.setEnabled(True)
            self.process_button.setEnabled(True)
            
            if not self.available_columns:
                self.update_status("No columns found in the Excel file", True)
                QMessageBox.critical(self, "Error", "No columns found in the Excel file")
            else:
                self.update_status(f"Found {len(self.available_columns)} columns in {os.path.basename(input_file)}")
                
        except Exception as e:
            error_msg = str(e)
            self.update_status(f"Error analyzing columns: {error_msg}", True)
            QMessageBox.critical(
                self, 
                "Error", 
                f"Could not read the Excel file. Please ensure it's a valid Excel file.\n\nError details: {error_msg}"
            )
    
    def add_selected_columns(self):
        """Add selected columns from available list to selected list"""
        selected_items = self.available_list.selectedItems()
        
        for item in selected_items:
            column_name = item.text()
            
            # Check if column is already in selected list
            already_selected = False
            for i in range(self.selected_list.count()):
                if self.selected_list.item(i).text() == column_name:
                    already_selected = True
                    break
            
            if not already_selected:
                self.selected_list.addItem(column_name)
                self.selected_value_columns.append(column_name)
        
        # Update status
        self.update_status(f"Selected {len(self.selected_value_columns)} value columns")
    
    def remove_selected_columns(self):
        """Remove selected columns from the selected list"""
        selected_items = self.selected_list.selectedItems()
        
        for item in selected_items:
            column_name = item.text()
            
            # Remove from list and from selected_value_columns
            for i in range(self.selected_list.count()):
                if self.selected_list.item(i).text() == column_name:
                    self.selected_list.takeItem(i)
                    if column_name in self.selected_value_columns:
                        self.selected_value_columns.remove(column_name)
                    break
        
        # Update status
        self.update_status(f"Selected {len(self.selected_value_columns)} value columns")
    
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
        # Get selected columns
        value_columns = []
        for i in range(self.selected_list.count()):
            value_columns.append(self.selected_list.item(i).text())
        
        if not value_columns:
            raise ValueError("Please select at least one value column to process")
        
        timestamp_column = self.timestamp_combo.currentText()
        if timestamp_column == "-- Select Timestamp Column --":
            timestamp_column = None
        
        # Load the Excel file
        xl = pd.ExcelFile(file_path)
        
        # Group sheet names by first N characters (default=6)
        sheet_groups = defaultdict(list)
        for sheet_name in xl.sheet_names:
            if len(sheet_name) >= prefix_length:
                prefix = sheet_name[:prefix_length]
                sheet_groups[prefix].append(sheet_name)
            else:
                # For sheet names shorter than prefix_length, use the entire name
                sheet_groups[sheet_name].append(sheet_name)
        
        # Analyze each sheet for selected columns
        sheet_analysis = {}
        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)  # Read just a few rows for analysis
                
                # Check if selected columns exist in this sheet
                present_value_columns = [col for col in value_columns if col in df.columns]
                has_timestamp = timestamp_column is None or timestamp_column in df.columns
                
                sheet_analysis[sheet_name] = {
                    'value_columns': present_value_columns,
                    'has_timestamp': has_timestamp,
                    'processable': len(present_value_columns) > 0 and has_timestamp
                }
            except Exception as e:
                sheet_analysis[sheet_name] = {
                    'value_columns': [],
                    'has_timestamp': False,
                    'processable': False,
                    'error': str(e)
                }
        
        return {
            'sheet_groups': sheet_groups,
            'sheet_analysis': sheet_analysis,
            'total_sheets': len(xl.sheet_names),
            'value_columns': value_columns,
            'timestamp_column': timestamp_column,
            'processable_groups': sum(1 for prefix, sheets in sheet_groups.items() 
                                    if any(sheet_analysis.get(sheet, {}).get('processable', False) for sheet in sheets))
        }
    
    def process_excel_file(self, file_path, output_path, prefix_length=6):
        """Process Excel file and create output file with summed telemetry data"""
        # Get selected columns
        value_columns = []
        for i in range(self.selected_list.count()):
            value_columns.append(self.selected_list.item(i).text())
        
        if not value_columns:
            raise ValueError("Please select at least one value column to process")
        
        timestamp_column = self.timestamp_combo.currentText()
        if timestamp_column == "-- Select Timestamp Column --":
            timestamp_column = None
        
        sum_values = self.sum_checkbox.isChecked()
        
        # Load the Excel file
        xl = pd.ExcelFile(file_path)
        
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
                        # Read data from the sheet
                        df = pd.read_excel(file_path, sheet_name=sheet)
                        
                        # Get the columns that exist in this sheet
                        present_cols = [col for col in value_columns if col in df.columns]
                        
                        if present_cols:
                            if timestamp_column and timestamp_column in df.columns:
                                # Use specified timestamp column
                                subset_cols = [timestamp_column] + present_cols
                                subset_df = df[subset_cols].copy()
                                subset_df.rename(columns={timestamp_column: 'Timestamp'}, inplace=True)
                                
                                # Add source sheet as a column for reference
                                subset_df['Source_Sheet'] = sheet
                                
                                # Append to combined data
                                combined_data = pd.concat([combined_data, subset_df])
                            elif timestamp_column is None:
                                # No timestamp column specified, just use value columns
                                subset_df = df[present_cols].copy()
                                
                                # Add source sheet as a column for reference
                                subset_df['Source_Sheet'] = sheet
                                
                                # Append to combined data
                                combined_data = pd.concat([combined_data, subset_df])
                    except Exception as e:
                        print(f"Error processing sheet {sheet}: {str(e)}")
                
                # Process combined data if we have any
                if not combined_data.empty:
                    # Sort by timestamp if we have one
                    if 'Timestamp' in combined_data.columns:
                        combined_data = combined_data.sort_values('Timestamp')
                    
                    # If summing values
                    if sum_values and 'Timestamp' in combined_data.columns:
                        # Group by timestamp and sum the value columns
                        agg_dict = {col: 'sum' for col in value_columns if col in combined_data.columns}
                        # Keep track of source sheets
                        agg_dict['Source_Sheet'] = lambda x: ', '.join(set(x))
                        
                        combined_data = combined_data.groupby('Timestamp').agg(agg_dict).reset_index()
                    
                    # Store in results
                    results[prefix] = combined_data
        
        # Now write each group to a separate sheet in the output Excel file
        with pd.ExcelWriter(output_path) as writer:
            for prefix, data in results.items():
                # Write to Excel
                data.to_excel(writer, sheet_name=prefix, index=False)
        
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
        
        if not input_file.lower().endswith(('.xls', '.xlsx')):
            self.update_status(f"Not a valid Excel file: {input_file}", True)
            QMessageBox.critical(self, "Error", f"Not a valid Excel file: {input_file}")
            return
        
        # Check if any columns are selected
        if self.selected_list.count() == 0:
            self.update_status("No value columns selected", True)
            QMessageBox.critical(self, "Error", "Please select at least one value column to process")
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
            
            text_content += f"Selected value columns: {', '.join(analysis['value_columns'])}\n"
            if analysis['timestamp_column']:
                text_content += f"Timestamp column: {analysis['timestamp_column']}\n"
            else:
                text_content += "No timestamp column selected\n"
            
            text_content += f"Sum values: {'Yes' if self.sum_checkbox.isChecked() else 'No (keeping individual values)'}\n"
            text_content += f"Sheet grouping (prefix length: {prefix_length}):\n"
            
            for prefix, sheets in analysis['sheet_groups'].items():
                text_content += f"\nGroup: '{prefix}'\n"
                text_content += f"Sheets in this group: {len(sheets)}\n"
                
                for sheet in sheets:
                    sheet_info = analysis['sheet_analysis'][sheet]
                    processable = sheet_info.get('processable', False)
                    status = "Will be processed" if processable else "Will be SKIPPED"
                    text_content += f"  - {sheet}: {status}\n"
                    
                    if not processable:
                        if not sheet_info.get('value_columns', []):
                            text_content += f"    Missing selected value columns\n"
                        if not sheet_info.get('has_timestamp', True) and analysis['timestamp_column']:
                            text_content += f"    Missing timestamp column: {analysis['timestamp_column']}\n"
                        if 'error' in sheet_info:
                            text_content += f"    Error: {sheet_info['error']}\n"
                    else:
                        present_cols = sheet_info.get('value_columns', [])
                        if present_cols:
                            text_content += f"    Found columns: {', '.join(present_cols)}\n"
            
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
        
        if not input_file.lower().endswith(('.xls', '.xlsx')):
            self.update_status(f"Not a valid Excel file: {input_file}", True)
            QMessageBox.critical(self, "Error", f"Not a valid Excel file: {input_file}")
            return
        
        # Check if any columns are selected
        if self.selected_list.count() == 0:
            self.update_status("No value columns selected", True)
            QMessageBox.critical(self, "Error", "Please select at least one value column to process")
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
    ex = GenericTelemetrySumTool()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
