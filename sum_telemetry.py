import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'  # Silence the Tk deprecation warning

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict

def process_excel_file(file_path, output_path, prefix_length=6):
    # Load the Excel file
    xl = pd.ExcelFile(file_path)
    
    # Group sheet names by first N characters (default=6)
    sheet_groups = defaultdict(list)
    for sheet_name in xl.sheet_names:
        if len(sheet_name) >= prefix_length:
            prefix = sheet_name[:prefix_length]
            sheet_groups[prefix].append(sheet_name)
    
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
                    
                    # Check if 'Raw' column exists
                    if 'Raw' in df.columns:
                        # Ensure there's a timestamp column (assuming it's named 'Timestamp')
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
                        else:
                            print(f"No timestamp column found in sheet '{sheet}'")
                    else:
                        print(f"No 'Raw' column found in sheet '{sheet}'")
                except Exception as e:
                    print(f"Error processing sheet '{sheet}': {e}")
            
            if not combined_data.empty:
                # Group by timestamp and sum the Raw values
                result = combined_data.groupby('Timestamp')['Raw'].sum().reset_index()
                results[prefix] = result
    
    # Create a new Excel file with the results
    with pd.ExcelWriter(output_path) as writer:
        for prefix, df in results.items():
            sheet_name = f"Sum_{prefix}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Results saved to {output_path}")

def create_ui():
    # Create the main window
    root = tk.Tk()
    root.title("Excel Telemetry Processor")
    root.geometry("500x320")  # Increased height for new controls
    root.resizable(True, True)
    
    # Configure styles
    root.configure(bg="#f0f0f0")
    
    # Variables to store file paths and options
    input_file_var = tk.StringVar()
    prefix_length_var = tk.IntVar(value=6)  # Default to 6 characters
    
    # Function to auto-generate output file path
    def auto_generate_output_path(input_path):
        if not input_path:
            return ""
        # Get directory and filename without extension
        dirname = os.path.dirname(input_path)
        basename = os.path.basename(input_path)
        name_without_ext = os.path.splitext(basename)[0]
        # Create output filename
        output_filename = f"{name_without_ext}_SUMMED.xlsx"
        return os.path.join(dirname, output_filename)
    
    # Function to browse for input file
    def browse_input_file():
        file_path = filedialog.askopenfilename(
            parent=root,
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx"), 
                ("Excel 97-2003", "*.xls"),
                ("All files", "*.*")
            ],
            initialdir=os.path.expanduser("~")
        )
        if file_path:
            input_file_var.set(file_path)
            print(f"Selected file: {file_path}")
    
    # We don't need output file browse function anymore as it's auto-generated
    
    # Function to process the files
    def process_files():
        input_file = input_file_var.get()
        
        if not input_file:
            update_status("Please select an input file", True)
            messagebox.showerror("Error", "Please select an input file")
            return
        
        if not os.path.exists(input_file):
            update_status(f"File not found: {input_file}", True)
            messagebox.showerror("Error", f"File not found: {input_file}")
            return
        
        # Check if the input file is a valid Excel file
        if not (input_file.lower().endswith('.xlsx') or input_file.lower().endswith('.xls')):
            update_status(f"Not a valid Excel file: {input_file}", True)
            messagebox.showerror("Error", f"Not a valid Excel file: {input_file}")
            return
        
        # Auto-generate output file path
        output_file = auto_generate_output_path(input_file)
            
        update_status(f"Processing file: {os.path.basename(input_file)}...")
        
        try:
            # Get the selected prefix length
            prefix_length = prefix_length_var.get()
            process_excel_file(input_file, output_file, prefix_length)
            update_status(f"Processing completed! Results saved to {os.path.basename(output_file)}")
            messagebox.showinfo("Success", f"Processing completed! Results saved to {output_file}")
        except Exception as e:
            error_msg = str(e)
            update_status(f"Error: {error_msg}", True)
            messagebox.showerror("Error", f"An error occurred: {error_msg}")
    
    # Create input file selection frame
    input_frame = tk.Frame(root, pady=10)
    input_frame.pack(fill="x", padx=20)
    
    # Create prefix length selection frame
    prefix_frame = tk.Frame(root, pady=5)
    prefix_frame.pack(fill="x", padx=20)
    
    tk.Label(prefix_frame, text="Sheet Name Prefix Length:", anchor="w").pack(side="left")
    prefix_spinner = tk.Spinbox(prefix_frame, from_=1, to=20, width=5, textvariable=prefix_length_var)
    prefix_spinner.pack(side="left", padx=5)
    
    # Add a help label for the prefix length
    tk.Label(prefix_frame, 
             text="(Number of characters to use for grouping worksheet names)",
             fg="#555555",
             anchor="w").pack(side="left", padx=5)
    
    tk.Label(input_frame, text="Input Excel File:", anchor="w").pack(fill="x")
    
    input_entry_frame = tk.Frame(input_frame)
    input_entry_frame.pack(fill="x", pady=5)
    
    tk.Entry(input_entry_frame, textvariable=input_file_var, width=50).pack(side="left", fill="x", expand=True)
    tk.Button(input_entry_frame, text="Browse", command=browse_input_file).pack(side="right", padx=5)
    
    # Create a note about output file
    note_frame = tk.Frame(root, pady=5)
    note_frame.pack(fill="x", padx=20)
    tk.Label(note_frame, 
             text="Note: Output will be saved automatically in the same folder as the input file",
             fg="#555555",
             anchor="w").pack(fill="x")
    
    # Create process button
    process_frame = tk.Frame(root, pady=20)
    process_frame.pack()
    
    tk.Button(
        process_frame, 
        text="Process Files", 
        command=process_files,
        bg="#4CAF50",
        fg="white",
        font=("Arial", 12, "bold"),
        padx=20,
        pady=10
    ).pack()
    
    # Status label
    status_var = tk.StringVar(value="Ready. Please select files.")
    status_label = tk.Label(root, textvariable=status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=10, pady=5)
    status_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    # Function to update status
    def update_status(message, is_error=False):
        status_var.set(message)
        status_label.config(fg="red" if is_error else "black")
        print(message)
    
    # Override the browse function to update status
    original_browse_input = browse_input_file
    def browse_input_with_status():
        original_browse_input()
        file_path = input_file_var.get()
        if file_path:
            # Show what the output will be
            output_path = auto_generate_output_path(file_path)
            update_status(f"Input: {os.path.basename(file_path)} → Output will be: {os.path.basename(output_path)}")
        else:
            update_status("No input file selected", True)
    
    # Replace the command
    for widget in input_entry_frame.winfo_children():
        if isinstance(widget, tk.Button):
            widget.config(command=browse_input_with_status)
    
    # Start the main loop
    root.mainloop()

# Example usage
if __name__ == "__main__":
    create_ui()
