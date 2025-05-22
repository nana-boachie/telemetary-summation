import os
import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                           QLabel, QLineEdit, QPushButton, QTabWidget, QFrame, QComboBox,
                           QCheckBox, QGroupBox, QTextEdit, QScrollArea, QMessageBox,
                           QFileDialog, QGridLayout, QSpinBox, QStatusBar)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QAction
import pandas as pd
from datetime import datetime
import calendar

# Import core functionality
from data_organizer import TelemetryDataOrganizer

# Import sum_telemetry and prevent it from running __main__
import sys
import importlib.util

def load_module(module_name):
    """Helper function to load a module, trying both .py and .py.old extensions"""
    base_path = module_name
    
    # Try .py first, then .py.old
    for ext in ['', '.old']:
        try:
            file_path = f"{base_path}.py{ext}" if ext else f"{base_path}.py"
            spec = importlib.util.spec_from_file_location(module_name, file_path)
            if spec is None:
                continue
                
            module = importlib.util.module_from_spec(spec)
            sys.modules[module_name] = module
            spec.loader.exec_module(module)
            return module
        except FileNotFoundError:
            continue
    
    raise ImportError(f"Could not find {module_name}.py or {module_name}.py.old")

# Load sum_telemetry without executing its __main__ section
try:
    sum_telemetry = load_module('sum_telemetry')
    TelemetrySumTool = sum_telemetry.TelemetrySumTool
except ImportError as e:
    print(f"Warning: Could not load sum_telemetry: {e}")
    TelemetrySumTool = None

# Load sum_telemetry_generic
try:
    sum_telemetry_generic = load_module('sum_telemetry_generic')
except ImportError as e:
    print(f"Warning: Could not load sum_telemetry_generic: {e}")
    sum_telemetry_generic = None
GenericTelemetrySumTool = sum_telemetry_generic.GenericTelemetrySumTool

class TelemetryAnalysisSuite(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Telemetry Analysis Suite")
        self.resize(1000, 800)
        
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
        
        self.setup_ui()
        
    def update_status(self, message, is_error=False):
        """Update the status bar with a message"""
        self.status_bar.showMessage(message)
        if is_error:
            self.status_bar.setStyleSheet("color: red;")
        else:
            self.status_bar.setStyleSheet("")
    
    def setup_ui(self):
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Create tab widget
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Add tools as tabs
        self.setup_telemetry_sum_tab()
        self.setup_generic_telemetry_tab()
        self.setup_organizer_tab()
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status("Ready")
        
    def setup_telemetry_sum_tab(self):
        """Set up the Telemetry Sum Tool tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Header label
        header_label = QLabel("Telemetry Sum Tool")
        header_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        # Add the TelemetrySumTool widget
        self.sum_tool = TelemetrySumTool()
        
        # Remove the main window elements we don't need in the tab
        self.sum_tool.setParent(None)
        
        # Get the central widget's layout and add it to our tab
        if self.sum_tool.centralWidget():
            layout.addWidget(self.sum_tool.centralWidget())
        
        self.tabs.addTab(tab, "Telemetry Sum")
    
    def setup_generic_telemetry_tab(self):
        """Set up the Generic Telemetry Tool tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Header label
        header_label = QLabel("Generic Telemetry Tool")
        header_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        # Add the GenericTelemetrySumTool widget
        self.generic_tool = GenericTelemetrySumTool()
        
        # Remove the main window elements we don't need in the tab
        self.generic_tool.setParent(None)
        
        # Get the central widget's layout and add it to our tab
        if self.generic_tool.centralWidget():
            layout.addWidget(self.generic_tool.centralWidget())
        
        self.tabs.addTab(tab, "Generic Telemetry")
        
    def setup_organizer_tab(self):
        """Set up the Data Organizer tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Header label
        header_label = QLabel("Data Organizer")
        header_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        # Tab widget for organizer features
        tab_widget = QTabWidget()
        layout.addWidget(tab_widget)
        
        # Create tabs for organizer features
        organize_tab = QWidget()
        reports_tab = QWidget()
        
        tab_widget.addTab(organize_tab, "Organize Files")
        tab_widget.addTab(reports_tab, "Generate Reports")
        
        # Setup content for each organizer tab
        self.setup_organize_tab(organize_tab)
        self.setup_reports_tab(reports_tab)
        
        # Add to main tabs
        self.tabs.addTab(tab, "Data Organizer")
    
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
        
        self.year_combo = QComboBox()
        current_year = datetime.now().year
        self.year_combo.addItems([str(y) for y in range(current_year - 5, current_year + 1)])
        self.year_combo.setCurrentText(str(current_year))
        year_layout.addWidget(QLabel("Year:"))
        year_layout.addWidget(self.year_combo)
        year_layout.addStretch()
        
        layout.addWidget(year_group)
        
        # Generate button
        generate_btn = QPushButton("Generate Annual Report")
        generate_btn.clicked.connect(self.generate_annual_report)
        layout.addWidget(generate_btn)
        
        # Status text area
        self.report_status = QTextEdit()
        self.report_status.setReadOnly(True)
        layout.addWidget(QLabel("Status:"))
        layout.addWidget(self.report_status)
        
        # Report options
        options_group = QGroupBox("Report Options")
        options_layout = QVBoxLayout(options_group)
        
        self.include_summary = QCheckBox("Include monthly summary")
        self.include_summary.setChecked(True)
        options_layout.addWidget(self.include_summary)
        
        self.include_charts = QCheckBox("Include charts (if available)")
        self.include_charts.setChecked(True)
        options_layout.addWidget(self.include_charts)
        
        layout.addWidget(options_group)
        
        # Report preview area
        preview_group = QGroupBox("Report Preview")
        preview_layout = QVBoxLayout(preview_group)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)
        
        # Save report button
        save_btn = QPushButton("Save Report")
        save_btn.clicked.connect(self.save_report)
        preview_layout.addWidget(save_btn)
        
        layout.addWidget(preview_group, 1)
    
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
        
        success_count = 0
        
        # Process each file
        for i, file_path in enumerate(file_paths, 1):
            try:
                self.status_bar.showMessage(
                    f"Processing file {i} of {len(file_paths)}: {os.path.basename(file_path)}"
                )
                
                # Process the file
                target_path = self.organizer.store_monthly_file(
                    file_path,
                    year=year if year and year.isdigit() else None,
                    month=month if month and month.isdigit() else None
                )
                
                self.log_text.append(f"✅ Successfully organized: {os.path.basename(file_path)}")
                self.log_text.append(f"   → Moved to: {target_path}\n")
                success_count += 1
                
            except Exception as e:
                error_msg = str(e)
                self.log_text.append(f"❌ Error processing {os.path.basename(file_path)}: {error_msg}")
                self.log_text.append(f"   → {str(e)}\n")
            
            # Process events to update the UI
            QApplication.processEvents()
        
        # Show completion message
        status_msg = (
            f"✅ Successfully organized {success_count} of {len(file_paths)} files. "
            f"Failed: {len(file_paths) - success_count} files."
        )
        self.status_bar.showMessage(status_msg)
        self.log_text.append("\n" + "="*50)
        self.log_text.append(status_msg)
        
        # Scroll to the bottom of the log
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        
        # Show completion dialog
        QMessageBox.information(
            self,
            "File Organization Complete",
            status_msg
        )
    
    def generate_annual_report(self):
        """Generate an annual report for the selected year"""
        year = self.year_combo.currentText()
        
        try:
            self.report_status.clear()
            self.report_status.append(f"Generating annual report for {year}...")
            
            # Generate the report using the organizer
            report_path = self.organizer.generate_annual_report(int(year))
            
            # Update the preview with the report content
            try:
                with open(report_path, 'r') as f:
                    content = f.read()
                self.preview_text.setPlainText(content)
                self.report_status.append("\nReport generated successfully!")
                self.report_status.append(f"Location: {report_path}")
                
                QMessageBox.information(
                    self, 
                    "Success", 
                    f"Annual report for {year} generated successfully!\n\n"
                    f"Location: {report_path}"
                )
                
            except Exception as e:
                self.report_status.append(f"\nError reading report: {str(e)}")
                raise
                
        except Exception as e:
            error_msg = f"Failed to generate annual report:\n{str(e)}"
            self.report_status.append("\nError: " + str(e))
            QMessageBox.critical(self, "Error", error_msg)
    
    def save_report(self):
        """Save the generated report to a file"""
        if not self.preview_text.toPlainText():
            QMessageBox.warning(self, "No Report", "Please generate a report before saving.")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Report",
            f"Telemetry_Report_{self.year_combo.currentText()}.txt",
            "Text Files (*.txt);;CSV Files (*.csv);;All Files (*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    f.write(self.preview_text.toPlainText())
                self.status_bar.showMessage(f"Report saved to {file_path}")
                QMessageBox.information(
                    self,
                    "Success",
                    f"Report saved successfully to:\n{file_path}"
                )
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save report: {str(e)}")


def main():
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle('Fusion')
    
    window = TelemetryAnalysisSuite()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
