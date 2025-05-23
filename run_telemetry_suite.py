"""
Launcher script for the Telemetry Analysis Suite
This script manually loads the necessary modules to avoid import errors
"""
import os
import sys
import importlib.util

def load_module_from_file(module_name, file_path):
    """Load a module from a file path"""
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None:
        raise ImportError(f"Could not load module {module_name} from {file_path}")
    
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Ensure current directory is in path
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

if __name__ == "__main__":
    try:
        # Explicitly import required modules from files
        data_organizer = load_module_from_file('data_organizer', 
                                               os.path.join(current_dir, 'data_organizer.py'))
        
        # Import and execute the telemetry analysis suite
        telemetry_analysis_suite = load_module_from_file('telemetry_analysis_suite',
                                                         os.path.join(current_dir, 'telemetry_analysis_suite.py'))
        
        # The suite should now be running via the import
        print("Telemetry Analysis Suite has been launched.")
        
    except Exception as e:
        print(f"Error launching Telemetry Analysis Suite: {str(e)}")
        import traceback
        traceback.print_exc()
