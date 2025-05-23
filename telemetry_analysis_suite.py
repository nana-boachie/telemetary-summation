import os
import sys

# Add the current directory to the Python path to ensure modules can be found
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                           QLabel, QLineEdit, QPushButton, QTabWidget, QFrame, QComboBox,
                           QCheckBox, QGroupBox, QTextEdit, QScrollArea, QMessageBox,
                           QFileDialog, QGridLayout, QSpinBox, QStatusBar)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QAction, QIcon, QPixmap
import pandas as pd
from datetime import datetime
import calendar

# Import core functionality
from data_organizer import TelemetryDataOrganizer

# Import telemetry modules
import importlib.util

def load_module(module_name):
    """Helper function to load a module using absolute paths"""
    # Get the directory of the current script
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Try with absolute paths
    try:
        file_path = os.path.join(current_dir, f"{module_name}.py")
        
        if not os.path.exists(file_path):
            raise ImportError(f"File not found: {file_path}")
                
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        if spec is None:
            raise ImportError(f"Failed to create spec for {file_path}")
                
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        return module
    except Exception as e:
        print(f"Error loading module {module_name}: {e}")
        raise
    
# Load modules
try:
    sum_telemetry = load_module('sum_telemetry')
    TelemetrySumTool = sum_telemetry.TelemetrySumTool
except ImportError as e:
    print(f"Warning: Could not load sum_telemetry: {e}")
    TelemetrySumTool = None

try:
    sum_telemetry_generic = load_module('sum_telemetry_generic')
    GenericTelemetrySumTool = sum_telemetry_generic.GenericTelemetrySumTool
except (ImportError, AttributeError) as e:
    print(f"Warning: Could not load sum_telemetry_generic: {e}")
    sum_telemetry_generic = None
    GenericTelemetrySumTool = None

class TelemetryAnalysisSuite(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Telemetry Analysis Suite")
        self.resize(1000, 800)
        
        # Set application logo
        try:
            app_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "applogo.png")
            if os.path.exists(app_icon_path):
                self.setWindowIcon(QIcon(app_icon_path))
        except Exception as e:
            print(f"Error setting application logo: {str(e)}")
        
        # Initialize the data organizer
        self.organizer = TelemetryDataOrganizer()
        
        # Variables
        self.data_dir = self.organizer.base_directory
        
        # Initialize status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        # Check if required modules were loaded
        self.modules_loaded = {
            'sum_telemetry': TelemetrySumTool is not None,
            'sum_telemetry_generic': sum_telemetry_generic is not None
        }
        
        # Initialize UI
        self.setup_ui()
        
        if not all(self.modules_loaded.values()):
            missing = [name for name, loaded in self.modules_loaded.items() if not loaded]
            QMessageBox.warning(
                self,
                "Warning: Some features disabled",
                f"Could not load the following modules: {', '.join(missing)}. "
                "Some features will be disabled.\n\n"
                "If you're running from a built version, this is expected.\n"
                "If running from source, please ensure all Python files are present."
            )
        
        self.update_status("Ready")
        
    def update_status(self, message, is_error=False):
        """Update the status bar with a message"""
        if hasattr(self, 'status_bar'):
            self.status_bar.showMessage(message)
            if is_error:
                self.status_bar.setStyleSheet("background-color: #ffdddd; color: #cc0000;")
            else:
                self.status_bar.setStyleSheet("")
    
    def setup_ui(self):
        # Create main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        
        # Add tabs - simplified version with just core tabs
        self.setup_organizer_tab()
        
        # Only add the telemetry tabs if the modules were loaded
        if TelemetrySumTool is not None:
            self.setup_telemetry_sum_tab()
        
        if GenericTelemetrySumTool is not None:
            self.setup_generic_telemetry_tab()
        
        # Set up the main layout
        layout = QVBoxLayout()
        layout.addWidget(self.tab_widget)
        self.central_widget.setLayout(layout)
    
    def setup_organizer_tab(self):
        """Set up the Data Organizer tab"""
        tab = QWidget()
        
        # Use a vertical layout for the main tab
        layout = QVBoxLayout(tab)
        
        # Header label
        header_label = QLabel("Telemetry Data Manager")
        header_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        # Create tab widget for the organizer features
        organizer_tabs = QTabWidget()
        layout.addWidget(organizer_tabs)
        
        # Add the organize files tab
        organize_tab = QWidget()
        self.setup_organize_tab(organize_tab)
        organizer_tabs.addTab(organize_tab, "Organize Files")
        
        # Add the report generation tab
        reports_tab = QWidget()
        self.setup_reports_tab(reports_tab)
        organizer_tabs.addTab(reports_tab, "Generate Reports")
        
        # Add to main tabs - make this the first tab
        self.tab_widget.insertTab(0, tab, "Data Manager")
        
        # Set the Data Manager as the active tab
        self.tab_widget.setCurrentIndex(0)
    
    def setup_organize_tab(self, tab):
        """Set up the file organization tab"""
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
        input_dir_group = QGroupBox("Source Folder with Files to Organize")
        input_dir_layout = QHBoxLayout(input_dir_group)
        
        self.input_dir_entry = QLineEdit()
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
        months = [""] + [f"{i:02d}: {calendar.month_name[i]}" for i in range(1, 13)]
        self.month_combo.addItems(months)
        date_layout.addWidget(self.month_combo, 0, 3)
        
        # Add a note
        note_label = QLabel("NOTE: If not specified, date will be determined from filenames or contents")
        note_label.setWordWrap(True)
        date_layout.addWidget(note_label, 1, 0, 1, 4)
        
        layout.addWidget(date_group)
        
        # Process button
        process_btn = QPushButton("Organize Files")
        process_btn.clicked.connect(self.organize_files)
        layout.addWidget(process_btn)
        
        # Log area
        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        layout.addWidget(log_group, 1)
    
    def setup_reports_tab(self, tab):
        """Set up the reports generation tab"""
        layout = QVBoxLayout(tab)
        
        # Year selection
        year_group = QGroupBox("Select Year for Report")
        year_layout = QHBoxLayout(year_group)
        
        self.report_year_combo = QComboBox()
        current_year = datetime.now().year
        self.report_year_combo.addItems([str(y) for y in range(current_year - 5, current_year + 1)])
        self.report_year_combo.setCurrentText(str(current_year))
        year_layout.addWidget(QLabel("Year:"))
        year_layout.addWidget(self.report_year_combo)
        year_layout.addStretch()
        
        # Check available data button
        check_btn = QPushButton("Check Available Data")
        check_btn.clicked.connect(self.check_available_months)
        year_layout.addWidget(check_btn)
        
        layout.addWidget(year_group)
        
        # Output file selection
        output_group = QGroupBox("Output File (Optional)")
        output_layout = QHBoxLayout(output_group)
        
        self.output_path_entry = QLineEdit()
        output_layout.addWidget(self.output_path_entry)
        
        browse_output_btn = QPushButton("Browse...")
        browse_output_btn.clicked.connect(self.browse_output_file)
        output_layout.addWidget(browse_output_btn)
        
        layout.addWidget(output_group)
        
        # Available months display
        months_group = QGroupBox("Available Monthly Data")
        months_layout = QVBoxLayout(months_group)
        
        self.months_display = QTextEdit()
        self.months_display.setReadOnly(True)
        months_layout.addWidget(self.months_display)
        
        layout.addWidget(months_group)
        
        # Generate button
        generate_btn = QPushButton("Generate Annual Report")
        generate_btn.clicked.connect(self.generate_annual_report)
        layout.addWidget(generate_btn)
    
    def setup_telemetry_sum_tab(self):
        """Set up the Telemetry Sum Tool tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Header label
        header_label = QLabel("Telemetry Sum Tool")
        header_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        if TelemetrySumTool is None:
            error_label = QLabel("The Telemetry Sum Tool could not be loaded.")
            error_label.setWordWrap(True)
            error_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(error_label)
        else:
            # Create but don't show the TelemetrySumTool
            self.sum_tool = TelemetrySumTool()
            self.sum_tool.hide()  # Hide the main window
            
            # Instead of trying to extract widgets, we'll embed the entire central widget
            if hasattr(self.sum_tool, 'centralWidget') and self.sum_tool.centralWidget():
                # Remove it from its parent first
                central_widget = self.sum_tool.centralWidget()
                central_widget.setParent(None)
                # Add it to our tab
                layout.addWidget(central_widget)
                
                # Make sure to keep connections to buttons and other UI elements
                self.sum_tool.setParent(tab)  # Set parent to keep connections alive
        
        self.tab_widget.addTab(tab, "Telemetry Sum")
    
    def setup_generic_telemetry_tab(self):
        """Set up the Generic Telemetry Tool tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        if GenericTelemetrySumTool is None:
            error_label = QLabel("The Generic Telemetry Tool could not be loaded.")
            error_label.setWordWrap(True)
            error_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(error_label)
            self.tab_widget.addTab(tab, "Generic Telemetry Tool (Error)")
            return
            
        try:
            # Header label
            header_label = QLabel("Generic Telemetry Tool")
            header_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
            header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(header_label)
            
            # Create and add the generic telemetry tool - prevent it from showing its own window
            self.generic_tool = GenericTelemetrySumTool()
            self.generic_tool.hide()  # Hide the main window
            
            # Instead of trying to extract widgets, we'll embed the entire central widget
            if hasattr(self.generic_tool, 'centralWidget') and self.generic_tool.centralWidget():
                # Remove it from its parent first
                central_widget = self.generic_tool.centralWidget()
                central_widget.setParent(None)
                # Add it to our tab
                layout.addWidget(central_widget)
                
                # Make sure to keep connections to buttons and other UI elements
                self.generic_tool.setParent(tab)  # Set parent to keep connections alive
            
            # Add the tab
            self.tab_widget.addTab(tab, "Generic Telemetry Tool")
            
        except Exception as e:
            error_label = QLabel(f"Failed to initialize Generic Telemetry Tool: {str(e)}")
            error_label.setWordWrap(True)
            error_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(error_label)
            self.tab_widget.addTab(tab, "Generic Telemetry Tool (Error)")
    
    def browse_data_dir(self):
        """Open a dialog to select the data directory"""
        dir_path = QFileDialog.getExistingDirectory(self, "Select Data Directory", self.data_dir)
        if dir_path:
            self.data_dir = dir_path
            self.data_dir_entry.setText(dir_path)
            self.organizer.base_directory = dir_path
    
    def browse_input_dir(self):
        """Open a dialog to select input folder"""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Folder Containing Files to Organize",
            "",
            QFileDialog.Option.ShowDirsOnly
        )
        if folder_path:
            self.input_dir_entry.setText(folder_path)
    
    def browse_output_file(self):
        """Open a dialog to select output file location"""
        year = self.report_year_combo.currentText()
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Annual Report",
            os.path.join(self.data_dir, f"Annual_Report_{year}.xlsx"),
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.output_path_entry.setText(file_path)
    
    def organize_files(self):
        """Organize files from the selected folder into the data directory"""
        folder_path = self.input_dir_entry.text().strip()
        
        # Get the selected year and month from the UI
        year = self.year_combo.currentText()
        month_combo_text = self.month_combo.currentText()
        month = month_combo_text.split(":")[0] if month_combo_text else None
        
        if not folder_path or not os.path.isdir(folder_path):
            QMessageBox.warning(
                self, 
                "Invalid Folder", 
                "Please select a valid folder containing files to organize."
            )
            return
        
        # Get all files in the selected directory
        try:
            file_paths = [
                os.path.join(folder_path, f) 
                for f in os.listdir(folder_path) 
                if os.path.isfile(os.path.join(folder_path, f))
            ]
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Error", 
                f"Could not read folder contents: {str(e)}"
            )
            return
        
        if not file_paths:
            QMessageBox.information(
                self, 
                "No Files", 
                "The selected folder is empty or contains no accessible files."
            )
            return
        
        # Clear previous logs and show progress
        self.log_text.clear()
        self.log_text.append("Starting file organization...\n")
        self.status_bar.showMessage("Organizing files...")
        
        try:
            # Process files using the data organizer
            report = self.organizer.process_new_files(
                folder_path,
                year=year if year and year.isdigit() else None,
                month=month if month and month.isdigit() else None
            )
            
            # Log results
            self.log_text.append(f"Processed {report['total_files']} files:\n")
            
            # Log successful files
            for item in report['organized']:
                self.log_text.append(f"✅ {os.path.basename(item['original'])} → {os.path.basename(item['destination'])}")
                if 'processed' in item:
                    self.log_text.append(f"   → Also processed: {os.path.basename(item['processed'])}")
            
            # Log errors
            if report['errors']:
                self.log_text.append("\nErrors:")
                for error in report['errors']:
                    self.log_text.append(f"❌ {os.path.basename(error['file'])}: {error['error']}")
            
            # Final status
            success_count = len(report['organized'])
            error_count = len(report['errors'])
            
            status_msg = (
                f"Completed: {success_count} files organized successfully, "
                f"{error_count} files had errors."
            )
            self.status_bar.showMessage(status_msg)
            
            # Show completion dialog
            QMessageBox.information(
                self,
                "File Organization Complete",
                status_msg
            )
            
        except Exception as e:
            error_msg = f"An error occurred during organization: {str(e)}"
            self.log_text.append(f"ERROR: {error_msg}")
            self.status_bar.showMessage("Error during organization", True)
            QMessageBox.critical(self, "Error", error_msg)
    
    def check_available_months(self):
        """Check which months have data available for the selected year"""
        year = self.report_year_combo.currentText()
        
        try:
            # Get files organized by month for the selected year
            year_files = self.organizer.list_files_for_year(year)
            
            if not any(year_files.values()):
                self.months_display.setText(f"No data files found for year {year}")
                return
            
            # Display months and files
            month_text = ""
            for month_num in range(1, 13):
                month_name = calendar.month_name[month_num]
                files = year_files.get(month_num, [])
                
                if files:
                    file_count = len(files)
                    month_text += f"✅ {month_name}: {file_count} files\n"
                else:
                    month_text += f"❌ {month_name}: No files\n"
            
            self.months_display.setText(month_text)
            self.status_bar.showMessage(f"Checked availability for year {year}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
    
    def generate_annual_report(self):
        """Generate an annual report for the selected year"""
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
            self.status_bar.showMessage(f"Generating annual report for {year}...")
            
            # Generate the report
            combined_data, report_path = self.organizer.generate_annual_report(year, output_path)
            
            if combined_data.empty:
                QMessageBox.information(self, "Result", f"No data found for year {year}")
                self.status_bar.showMessage(f"No data found for year {year}")
                return
            
            if report_path:
                QMessageBox.information(self, "Success", f"Annual report saved to:\n{report_path}")
                self.status_bar.showMessage(f"Report saved to {os.path.basename(report_path)}")
            else:
                QMessageBox.critical(self, "Error", "Failed to save report")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
            self.status_bar.showMessage(f"Error: {str(e)}")


def main():
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle('Fusion')
    
    # Set application icon for all windows
    try:
        app_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "applogo.png")
        if os.path.exists(app_icon_path):
            app_icon = QIcon(app_icon_path)
            app.setWindowIcon(app_icon)
    except Exception as e:
        print(f"Error setting application logo: {str(e)}")
    
    window = TelemetryAnalysisSuite()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
