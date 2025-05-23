"""
Simple launcher script that directly executes the code from data_organizer and annual_report_generator
without relying on Python's import system
"""
import os
import sys
import subprocess

def run_app():
    """Run the Annual Report Generator application"""
    # Get the directory of this script
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Path to Python executable in the virtual environment
    python_exe = os.path.join(current_dir, "venv", "Scripts", "python.exe")
    
    # Check if Python executable exists
    if not os.path.exists(python_exe):
        print(f"Error: Python executable not found at {python_exe}")
        return
    
    # Run the annual report generator directly
    try:
        # Read in the annual_report_generator.py file content
        with open(os.path.join(current_dir, "annual_report_generator.py"), "r") as f:
            code = f.read()
        
        # Replace the import statement with code that directly loads the TelemetryDataOrganizer class
        # First, read in the data_organizer.py file
        with open(os.path.join(current_dir, "data_organizer.py"), "r") as f:
            data_organizer_code = f.read()
        
        # Create a new temporary file with combined code
        temp_file = os.path.join(current_dir, "temp_app.py")
        with open(temp_file, "w") as f:
            # Remove the data_organizer import line from the annual_report_generator code
            code = code.replace("from data_organizer import TelemetryDataOrganizer", "# TelemetryDataOrganizer is defined above")
            
            # Write the combined code
            f.write(data_organizer_code + "\n\n# Annual Report Generator Code\n\n" + code)
        
        # Run the temporary file
        subprocess.call([python_exe, temp_file])
        
        # Clean up the temporary file
        try:
            os.remove(temp_file)
        except:
            pass
            
    except Exception as e:
        print(f"Error running application: {str(e)}")

if __name__ == "__main__":
    run_app()
