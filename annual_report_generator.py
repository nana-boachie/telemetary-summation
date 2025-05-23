import os
import sys

# Add the current directory to the Python path to ensure modules can be found
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTabWidget, QFrame, QComboBox,
                             QCheckBox, QGroupBox, QTextEdit, QScrollArea, QMessageBox,
                             QFileDialog, QGridLayout)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import pandas as pd
from datetime import datetime
import calendar
from data_organizer import TelemetryDataOrganizer

class AnnualReportGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Telemetry Annual Report Generator")
        self.resize(800, 600)
        
        # Initialize the data organizer
        self.organizer = TelemetryDataOrganizer()
        
        # Variables
        self.data_dir = self.organizer.base_directory
        self.input_dir = ""
        self.selected_year = str(datetime.now().year)
        self.setup_ui()
    
    def setup_ui(self):
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Header label
        header_label = QLabel("Telemetry Annual Report Generator")
        header_label.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        main_layout.addWidget(header_label)
        
        # Tab widget
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # Create tabs
        organize_tab = QWidget()
        reports_tab = QWidget()
        
        self.tab_widget.addTab(organize_tab, "Organize Monthly Files")
        self.tab_widget.addTab(reports_tab, "Generate Annual Reports")
        
        # Setup content for each tab
        self.setup_organize_tab(organize_tab)
        self.setup_reports_tab(reports_tab)
        
        # Status bar
        self.statusBar().showMessage("Ready")
    
    def setup_organize_tab(self, tab):
        layout = QVBoxLayout(tab)
        
        # Data directory frame
        data_dir_group = QGroupBox("Data Storage Location")
        data_dir_layout = QHBoxLayout(data_dir_group)
        
        self.data_dir_entry = QLineEdit(self.data_dir)
        data_dir_layout.addWidget(self.data_dir_entry)
        
        browse_data_btn = QPushButton("Browse...")
        browse_data_btn.clicked.connect(self.browse_data_dir)
        data_dir_layout.addWidget(browse_data_btn)
        
        layout.addWidget(data_dir_group)
        
        # Input directory frame
        input_dir_group = QGroupBox("Monthly Files to Organize")
        input_dir_layout = QHBoxLayout(input_dir_group)
        
        self.input_dir_entry = QLineEdit(self.input_dir)
        input_dir_layout.addWidget(self.input_dir_entry)
        
        browse_input_btn = QPushButton("Browse...")
        browse_input_btn.clicked.connect(self.browse_input_dir)
        input_dir_layout.addWidget(browse_input_btn)
        
        layout.addWidget(input_dir_group)
        
        # Year and month selection
        date_group = QGroupBox("Date Assignment (Optional)")
        date_layout = QGridLayout(date_group)
        
        date_layout.addWidget(QLabel("Year:"), 0, 0)
        self.year_combo = QComboBox()
        years = [str(year) for year in range(datetime.now().year - 10, datetime.now().year + 2)]
        self.year_combo.addItems(years)
        self.year_combo.setCurrentText(str(datetime.now().year))
        date_layout.addWidget(self.year_combo, 0, 1)
        
        date_layout.addWidget(QLabel("Month:"), 0, 2)
        self.month_combo = QComboBox()
        months = [""] + [f"{i}: {calendar.month_name[i]}" for i in range(1, 13)]
        self.month_combo.addItems(months)
        date_layout.addWidget(self.month_combo, 0, 3)
        
        note_label = QLabel("NOTE: If not specified, date will be determined from filenames or contents")
        date_layout.addWidget(note_label, 1, 0, 1, 4)
        
        layout.addWidget(date_group)
        
        # Process button area
        process_frame = QWidget()
        process_layout = QHBoxLayout(process_frame)
        
        self.process_check = QCheckBox("Process files immediately after organizing")
        process_layout.addWidget(self.process_check)
        
        organize_btn = QPushButton("Organize Files")
        organize_btn.clicked.connect(self.organize_files)
        process_layout.addWidget(organize_btn)
        
        layout.addWidget(process_frame)
        
        # Add stretching space at the bottom
        layout.addStretch()
    
    def setup_reports_tab(self, tab):
        layout = QVBoxLayout(tab)
        
        # Year selection
        year_group = QGroupBox("Select Year for Annual Report")
        year_layout = QHBoxLayout(year_group)
        
        self.report_year_combo = QComboBox()
        years = [str(year) for year in range(datetime.now().year - 10, datetime.now().year + 2)]
        self.report_year_combo.addItems(years)
        self.report_year_combo.setCurrentText(str(datetime.now().year))
        year_layout.addWidget(self.report_year_combo)
        
        refresh_btn = QPushButton("Refresh Available Years")
        refresh_btn.clicked.connect(self.refresh_years)
        year_layout.addWidget(refresh_btn)
        
        layout.addWidget(year_group)
        
        # Month availability display
        months_group = QGroupBox("Available Months")
        months_layout = QVBoxLayout(months_group)
        
        self.months_display = QTextEdit()
        self.months_display.setReadOnly(True)
        months_layout.addWidget(self.months_display)
        
        check_btn = QPushButton("Check Available Data")
        check_btn.clicked.connect(self.check_available_months)
        months_layout.addWidget(check_btn)
        
        layout.addWidget(months_group)
        
        # Output location
        output_group = QGroupBox("Report Output Location")
        output_layout = QHBoxLayout(output_group)
        
        self.output_path_entry = QLineEdit()
        output_layout.addWidget(self.output_path_entry)
        
        browse_output_btn = QPushButton("Browse...")
        browse_output_btn.clicked.connect(self.browse_output_file)
        output_layout.addWidget(browse_output_btn)
        
        layout.addWidget(output_group)
        
        # Generate report button
        generate_btn = QPushButton("Generate Annual Report")
        generate_btn.clicked.connect(self.generate_report)
        layout.addWidget(generate_btn)
        
        # Add stretching space at the bottom
        layout.addStretch()
    
    def browse_data_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Data Storage Directory")
        if directory:
            self.data_dir_entry.setText(directory)
            self.organizer.base_directory = directory
            self.statusBar().showMessage(f"Data directory set to: {directory}")
    
    def browse_input_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Directory with Monthly Files")
        if directory:
            self.input_dir_entry.setText(directory)
            self.statusBar().showMessage(f"Input directory set to: {directory}")
    
    def browse_output_file(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Save Annual Report As", 
            "", 
            "Excel files (*.xlsx);;All files (*.*)"
        )
        if file_path:
            self.output_path_entry.setText(file_path)
            self.statusBar().showMessage(f"Output file set to: {file_path}")
    
    def organize_files(self):
        input_dir = self.input_dir_entry.text()
        if not input_dir or not os.path.exists(input_dir):
            QMessageBox.critical(self, "Error", "Please select a valid input directory")
            return
        
        # Get year and month
        year = self.year_combo.currentText() if self.year_combo.currentText() else None
        
        month = None
        if self.month_combo.currentText():
            try:
                month = int(self.month_combo.currentText().split(":")[0])
            except:
                pass
        
        try:
            self.statusBar().showMessage("Organizing files...")
            report = self.organizer.process_new_files(
                input_dir, 
                year=year, 
                month=month,
                process_immediately=self.process_check.isChecked()
            )
            
            # Show results
            result_message = f"Processed {report['total_files']} files:\n"
            result_message += f"✓ Successfully organized: {len(report['organized'])}\n"
            result_message += f"✗ Errors: {len(report['errors'])}"
            
            self.statusBar().showMessage(f"Organized {len(report['organized'])} files successfully")
            QMessageBox.information(self, "Organization Complete", result_message)
            
            # Show detailed report
            if report['organized'] or report['errors']:
                self.show_detailed_report(report)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
            self.statusBar().showMessage(f"Error: {str(e)}")
    
    def refresh_years(self):
        try:
            # Get list of years from the data directory
            if not os.path.exists(self.organizer.base_directory):
                QMessageBox.information(self, "Info", "Data directory doesn't exist yet")
                return
                
            years = [d for d in os.listdir(self.organizer.base_directory) 
                    if os.path.isdir(os.path.join(self.organizer.base_directory, d)) and d.isdigit()]
            
            if not years:
                QMessageBox.information(self, "Info", "No year directories found in data location")
                return
            
            # Update the combobox
            self.report_year_combo.clear()
            self.report_year_combo.addItems(years)
            self.report_year_combo.setCurrentText(max(years))  # Set to most recent year
            self.statusBar().showMessage(f"Found {len(years)} years with data")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
    
    def check_available_months(self):
        year = self.report_year_combo.currentText()
        if not year:
            QMessageBox.critical(self, "Error", "Please select a year")
            return
        
        try:
            year_files = self.organizer.list_files_for_year(year)
            
            # Clear display
            self.months_display.clear()
            
            if not year_files:
                self.months_display.setText(f"No data found for year {year}")
                return
            
            # Display months and files
            month_text = ""
            for month_num in sorted(year_files.keys()):
                month_name = calendar.month_name[month_num]
                files = year_files[month_num]
                
                if files:
                    file_count = len(files)
                    month_text += f"✓ {month_name}: {file_count} files\n"
                else:
                    month_text += f"✗ {month_name}: No files\n"
            
            self.months_display.setText(month_text)
            self.statusBar().showMessage(f"Checked availability for year {year}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
    
    def generate_report(self):
        year = self.report_year_combo.currentText()
        if not year:
            QMessageBox.critical(self, "Error", "Please select a year")
            return
        
        output_path = self.output_path_entry.text()
        if not output_path:
            # Create default path
            output_path = os.path.join(self.organizer.base_directory, f"Annual_Report_{year}.xlsx")
            self.output_path_entry.setText(output_path)
        
        try:
            self.statusBar().showMessage(f"Generating annual report for {year}...")
            
            # Define a custom process function to ensure all models are included
            def enhanced_process_func(file_path):
                try:
                    # Import the necessary module for processing
                    from sum_telemetry import process_excel_file
                    
                    # Create a temporary output path
                    temp_output = os.path.join(os.path.dirname(file_path), f"temp_{os.path.basename(file_path)}")
                    
                    try:
                        # Process the file
                        process_excel_file(file_path, temp_output)
                        
                        # Load the processed data if file exists
                        if os.path.exists(temp_output):
                            # Read all sheets to capture all models
                            result_data = pd.read_excel(temp_output, sheet_name=None)
                            
                            # Combine all sheets into one DataFrame
                            all_sheets_data = []
                            for sheet_name, sheet_data in result_data.items():
                                # Add sheet name as Model column if not already present
                                if 'Model' not in sheet_data.columns:
                                    sheet_data['Model'] = sheet_name
                                all_sheets_data.append(sheet_data)
                            
                            if all_sheets_data:
                                combined_data = pd.concat(all_sheets_data, ignore_index=True)
                                return combined_data
                            return None
                        return None
                    except Exception as e:
                        raise Exception(f"Error processing file {file_path}: {str(e)}")
                    finally:
                        # Clean up temporary file regardless of success or failure
                        if os.path.exists(temp_output):
                            try:
                                os.remove(temp_output)
                            except Exception as cleanup_error:
                                print(f"Warning: Could not remove temporary file {temp_output}: {cleanup_error}")
                except Exception as e:
                    print(f"Error in enhanced_process_func: {str(e)}")
                    return None
            
            # Generate the report with our enhanced processing function
            combined_data, report_path = self.organizer.generate_annual_report(year, output_path, enhanced_process_func)
            
            if combined_data.empty:
                QMessageBox.information(self, "Result", f"No data found for year {year}")
                self.statusBar().showMessage(f"No data found for year {year}")
                return
            
            if report_path:
                QMessageBox.information(self, "Success", f"Annual report saved to:\n{report_path}")
                self.statusBar().showMessage(f"Report saved to {os.path.basename(report_path)}")
            else:
                QMessageBox.critical(self, "Error", "Failed to save report")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
            self.statusBar().showMessage(f"Error: {str(e)}")
    
    def show_detailed_report(self, report):
        # Create a dialog for detailed report
        dialog = QDialog(self)
        dialog.setWindowTitle("Organization Report")
        dialog.resize(600, 400)
        
        layout = QVBoxLayout(dialog)
        
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        layout.addWidget(text_edit)
        
        # Insert report content
        report_text = "ORGANIZATION REPORT\n\n"
        report_text += f"Total files processed: {report['total_files']}\n\n"
        
        # Successful files
        report_text += f"SUCCESSFULLY ORGANIZED FILES ({len(report['organized'])}):\n"
        for i, item in enumerate(report['organized'], 1):
            report_text += f"{i}. {os.path.basename(item['original'])} → {os.path.basename(item['destination'])}\n"
            if 'processed' in item:
                report_text += f"   Processed: {os.path.basename(item['processed'])}\n"
            if 'processing_error' in item:
                report_text += f"   Processing error: {item['processing_error']}\n"
        
        # Errors
        if report['errors']:
            report_text += f"\nERRORS ({len(report['errors'])}):\n"
            for i, error in enumerate(report['errors'], 1):
                report_text += f"{i}. {os.path.basename(error['file'])}: {error['error']}\n"
        
        text_edit.setText(report_text)
        
        # Add close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)
        
        dialog.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AnnualReportGeneratorApp()
    window.show()
    sys.exit(app.exec())