import os
import shutil
import pandas as pd
from datetime import datetime
import calendar
import glob

class TelemetryDataOrganizer:
    """
    A class for organizing telemetry data files by year and month,
    and generating annual reports by combining monthly data.
    """
    
    def __init__(self, base_directory="data"):
        """
        Initialize the data organizer with a base directory.
        
        Args:
            base_directory (str): The base directory for storing organized data files
        """
        self.base_directory = base_directory
        
        # Create the base directory if it doesn't exist
        if not os.path.exists(self.base_directory):
            os.makedirs(self.base_directory)
    
    def create_directory_structure(self, year):
        """
        Create directory structure for a specific year with folders for each month.
        
        Args:
            year (str or int): The year to create directories for
        
        Returns:
            dict: Dictionary with paths to each month directory
        """
        year_dir = os.path.join(self.base_directory, str(year))
        
        if not os.path.exists(year_dir):
            os.makedirs(year_dir)
        
        # Create a folder for each month
        month_dirs = {}
        for month_num in range(1, 13):
            month_name = calendar.month_name[month_num]
            month_dir = os.path.join(year_dir, f"{month_num:02d}_{month_name}")
            
            if not os.path.exists(month_dir):
                os.makedirs(month_dir)
            
            month_dirs[month_num] = month_dir
        
        return month_dirs
    
    def store_monthly_file(self, file_path, year=None, month=None, copy_file=True):
        """
        Store a monthly data file in the appropriate year/month directory.
        If year/month not provided, tries to determine from filename or file content.
        
        Args:
            file_path (str): Path to the file to store
            year (str or int, optional): Year to store the file under
            month (str or int, optional): Month to store the file under
            copy_file (bool): If True, copy the file; if False, move it
        
        Returns:
            str: Path to the stored file
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # If year/month not provided, try to determine from filename or content
        if year is None or month is None:
            determined_date = self._determine_date_from_file(file_path)
            year = year or determined_date.get('year')
            month = month or determined_date.get('month')
        
        # Convert to strings for directory path
        year = str(year)
        month = int(month) if month else None
        
        if not year or not month:
            raise ValueError("Could not determine year or month for file")
        
        # Create directory structure for this year if it doesn't exist
        month_dirs = self.create_directory_structure(year)
        target_dir = month_dirs[month]
        
        # Get the file name from the path
        file_name = os.path.basename(file_path)
        
        # Destination path
        destination = os.path.join(target_dir, file_name)
        
        # Copy or move the file
        if copy_file:
            shutil.copy2(file_path, destination)
        else:
            shutil.move(file_path, destination)
        
        return destination
    
    def _determine_date_from_file(self, file_path):
        """
        Attempt to determine year and month from file name or content.
        
        Args:
            file_path (str): Path to the file
        
        Returns:
            dict: Dictionary with 'year' and 'month' if found
        """
        result = {'year': None, 'month': None}
        file_name = os.path.basename(file_path)
        
        # Try to get date from filename first
        # Common formats might be YYYY_MM, YYYY-MM, etc.
        import re
        
        # Pattern for YYYY_MM or YYYY-MM or other common separators
        patterns = [
            r'(\d{4})[_\-\.](\d{1,2})',  # YYYY_MM, YYYY-MM, YYYY.MM
            r'(\d{1,2})[_\-\.](\d{4})',  # MM_YYYY, MM-YYYY, MM.YYYY
        ]
        
        for pattern in patterns:
            match = re.search(pattern, file_name)
            if match:
                groups = match.groups()
                if len(groups[0]) == 4:  # YYYY_MM format
                    result['year'] = groups[0]
                    result['month'] = int(groups[1])
                else:  # MM_YYYY format
                    result['year'] = groups[1]
                    result['month'] = int(groups[0])
                break
        
        # If we couldn't determine from filename, try to read the file if it's Excel
        if (result['year'] is None or result['month'] is None) and file_path.endswith(('.xlsx', '.xls')):
            try:
                # Read the first few rows to look for date columns
                df = pd.read_excel(file_path, nrows=10)
                
                # Look for date columns
                for column in df.columns:
                    if 'date' in column.lower() or 'time' in column.lower():
                        # Check if column has datetime values
                        if pd.api.types.is_datetime64_any_dtype(df[column]):
                            # Get the first valid date
                            first_date = df[column].dropna().iloc[0] if not df[column].dropna().empty else None
                            if first_date:
                                # Convert to datetime if it's not already
                                if not isinstance(first_date, datetime):
                                    try:
                                        first_date = pd.to_datetime(first_date)
                                    except:
                                        continue
                                
                                result['year'] = str(first_date.year)
                                result['month'] = first_date.month
                                break
            except Exception as e:
                print(f"Error reading Excel file: {e}")
        
        return result
    
    def list_files_for_month(self, year, month):
        """
        List all files stored for a specific month.
        
        Args:
            year (str or int): Year to look in
            month (str or int): Month to look in
        
        Returns:
            list: List of file paths for the specified month
        """
        month = int(month)
        year_dir = os.path.join(self.base_directory, str(year))
        month_name = calendar.month_name[month]
        month_dir = os.path.join(year_dir, f"{month:02d}_{month_name}")
        
        if not os.path.exists(month_dir):
            return []
        
        return [os.path.join(month_dir, f) for f in os.listdir(month_dir) if os.path.isfile(os.path.join(month_dir, f))]
    
    def list_files_for_year(self, year):
        """
        List all files stored for a specific year, organized by month.
        
        Args:
            year (str or int): Year to look in
        
        Returns:
            dict: Dictionary with months as keys and lists of file paths as values
        """
        year_dir = os.path.join(self.base_directory, str(year))
        
        if not os.path.exists(year_dir):
            return {}
        
        result = {}
        
        for month_num in range(1, 13):
            month_name = calendar.month_name[month_num]
            month_dir = os.path.join(year_dir, f"{month_num:02d}_{month_name}")
            
            if os.path.exists(month_dir):
                files = [os.path.join(month_dir, f) for f in os.listdir(month_dir) if os.path.isfile(os.path.join(month_dir, f))]
                result[month_num] = files
            else:
                result[month_num] = []
        
        return result
    
    def generate_annual_report(self, year, output_path=None, process_func=None):
        """
        Generate an annual report by combining monthly data files.
        
        Args:
            year (str or int): Year to generate report for
            output_path (str, optional): Path to save the output report
            process_func (callable, optional): Function to process each file before combining
                                              Should take file_path as input and return a DataFrame
        
        Returns:
            tuple: (DataFrame with combined data, path to saved report if output_path provided)
        """
        year_str = str(year)
        year_files = self.list_files_for_year(year_str)
        
        # If no output path provided, create one in the year directory
        if output_path is None:
            year_dir = os.path.join(self.base_directory, year_str)
            output_path = os.path.join(year_dir, f"Annual_Report_{year_str}.xlsx")
        
        # Default processing function if none provided
        if process_func is None:
            from sum_telemetry import process_excel_file
            
            def default_process(file_path):
                # Create a temporary output path
                temp_output = os.path.join(os.path.dirname(file_path), f"temp_{os.path.basename(file_path)}")
                # Process the file
                results = process_excel_file(file_path, temp_output)
                # Load the processed data
                if os.path.exists(temp_output):
                    result_data = pd.read_excel(temp_output)
                    # Clean up
                    os.remove(temp_output)
                    return result_data
                return None
            
            process_func = default_process
        
        # Combine data from all months
        all_data = []
        processed_months = []
        
        for month, files in sorted(year_files.items()):
            for file_path in files:
                try:
                    # Process the file
                    processed_data = process_func(file_path)
                    
                    if processed_data is not None and not processed_data.empty:
                        # Add month information if not already present
                        if 'Month' not in processed_data.columns:
                            processed_data['Month'] = calendar.month_name[month]
                        if 'MonthNum' not in processed_data.columns:
                            processed_data['MonthNum'] = month
                        
                        all_data.append(processed_data)
                        if month not in processed_months:
                            processed_months.append(month)
                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")
        
        # Combine all processed data
        if all_data:
            combined_data = pd.concat(all_data, ignore_index=True)
            
            # Save to the output file if a path is provided
            if output_path:
                # Create directory if it doesn't exist
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                # Save to Excel
                with pd.ExcelWriter(output_path) as writer:
                    combined_data.to_excel(writer, sheet_name='Annual_Summary', index=False)
                    
                    # Add a sheet with month information
                    month_info = pd.DataFrame({
                        'Month': [calendar.month_name[m] for m in processed_months],
                        'MonthNum': processed_months,
                        'Files Processed': [len([f for f in year_files[m]]) for m in processed_months]
                    })
                    month_info.to_excel(writer, sheet_name='Months_Included', index=False)
                
                return combined_data, output_path
            
            return combined_data, None
        
        return pd.DataFrame(), None
    
    def process_new_files(self, input_directory, year=None, month=None, process_immediately=False):
        """
        Process new files from an input directory and organize them by year/month.
        
        Args:
            input_directory (str): Directory containing input files to process
            year (str or int, optional): Year to assign if not determined from files
            month (str or int, optional): Month to assign if not determined from files
            process_immediately (bool): Whether to process files immediately after organizing
        
        Returns:
            dict: Report of processed files
        """
        if not os.path.exists(input_directory):
            raise FileNotFoundError(f"Input directory not found: {input_directory}")
        
        # Get all Excel files in the input directory
        files = glob.glob(os.path.join(input_directory, "*.xlsx")) + glob.glob(os.path.join(input_directory, "*.xls"))
        
        report = {
            'total_files': len(files),
            'organized': [],
            'errors': []
        }
        
        for file_path in files:
            try:
                # Store the file in the appropriate year/month directory
                destination = self.store_monthly_file(file_path, year, month)
                report['organized'].append({
                    'original': file_path,
                    'destination': destination
                })
                
                # Process the file immediately if requested
                if process_immediately:
                    from sum_telemetry import process_excel_file
                    
                    # Generate output path
                    output_dir = os.path.dirname(destination)
                    file_name = os.path.basename(destination)
                    output_path = os.path.join(output_dir, f"processed_{file_name}")
                    
                    # Process the file
                    try:
                        process_excel_file(destination, output_path)
                        report['organized'][-1]['processed'] = output_path
                    except Exception as e:
                        report['organized'][-1]['processing_error'] = str(e)
                
            except Exception as e:
                report['errors'].append({
                    'file': file_path,
                    'error': str(e)
                })
        
        return report


# Example usage
if __name__ == "__main__":
    # Create a data organizer
    organizer = TelemetryDataOrganizer()
    
    # Example 1: Manually create directory structure for a year
    month_dirs = organizer.create_directory_structure(2023)
    print(f"Created directories for 2023: {month_dirs}")
    
    # Example 2: Organize files from a directory
    # report = organizer.process_new_files("path/to/input_files")
    # print(f"Processed {report['total_files']} files")
    
    # Example 3: Generate an annual report
    # combined_data, report_path = organizer.generate_annual_report(2023)
    # if report_path:
    #     print(f"Annual report saved to: {report_path}")
