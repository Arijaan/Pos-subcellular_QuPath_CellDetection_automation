import pandas as pd
import os
import numpy as np
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

# Try to import openpyxl, fallback to xlsxwriter if not available
try:
    import openpyxl
    EXCEL_ENGINE = 'openpyxl'
except ImportError:
    try:
        import xlsxwriter
        EXCEL_ENGINE = 'xlsxwriter'
    except ImportError:
        print("Warning: Neither openpyxl nor xlsxwriter found. Excel functionality may be limited.")
        EXCEL_ENGINE = None


class GroupInputDialog:
    """
    A tkinter dialog for dynamically entering group names/tags
    """
    def __init__(self, parent=None):
        self.groups = []
        self.result = None
        
        # Create the main window
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Define Group Tags")
        self.root.geometry("500x400")
        self.root.resizable(True, True)
        
        # Make window modal
        self.root.transient(parent)
        self.root.grab_set()
        
        self.setup_ui()
        
        # Center the window
        self.center_window()
        
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Instructions
        instructions = ttk.Label(main_frame, text=
            "Enter group tags that appear in your CSV filenames.\n"
            "For example: 'control', 't1d', 't2d', 'aab'\n"
            "Files will be grouped based on these tags found in filenames.",
            justify=tk.LEFT)
        instructions.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))
        
        # Input frame
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))
        input_frame.columnconfigure(0, weight=1)
        
        ttk.Label(input_frame, text="Group Tag:").grid(row=0, column=0, sticky=tk.W)
        self.entry = ttk.Entry(input_frame, width=30)
        self.entry.grid(row=0, column=1, padx=(5, 5), sticky=(tk.W, tk.E))
        self.entry.bind('<Return>', self.add_group)
        
        ttk.Button(input_frame, text="Add", command=self.add_group).grid(row=0, column=2)
        
        # Groups listbox with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=2, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Listbox with scrollbar
        self.listbox = tk.Listbox(list_frame, height=8)
        self.listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        # Remove button
        ttk.Button(list_frame, text="Remove Selected", 
                  command=self.remove_group).grid(row=1, column=0, pady=(5, 0))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(10, 0))
        
        ttk.Button(button_frame, text="OK", command=self.ok_clicked).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Cancel", command=self.cancel_clicked).pack(side=tk.LEFT)
        
        # Add some default suggestions
        default_groups = ['control', 't1d', 't2d', 'aab']
        for group in default_groups:
            self.groups.append(group.lower())
            self.listbox.insert(tk.END, group.lower())
        
        # Focus on entry
        self.entry.focus()
        
    def add_group(self, event=None):
        """Add a group to the list"""
        group = self.entry.get().strip().lower()
        if group and group not in self.groups:
            self.groups.append(group)
            self.listbox.insert(tk.END, group)
            self.entry.delete(0, tk.END)
        elif group in self.groups:
            messagebox.showwarning("Duplicate", f"Group '{group}' already exists!")
            
    def remove_group(self):
        """Remove selected group from the list"""
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            group = self.groups.pop(index)
            self.listbox.delete(index)
            
    def ok_clicked(self):
        """Handle OK button click"""
        if not self.groups:
            messagebox.showwarning("No Groups", "Please add at least one group tag!")
            return
        self.result = self.groups.copy()
        self.root.destroy()
        
    def cancel_clicked(self):
        """Handle Cancel button click"""
        self.result = None
        self.root.destroy()
        
    def show(self):
        """Show the dialog and return the result"""
        self.root.wait_window()
        return self.result


def get_user_groups():
    """
    Get group tags from user input dialog
    """
    # Create a temporary root window (hidden)
    temp_root = tk.Tk()
    temp_root.withdraw()
    
    try:
        dialog = GroupInputDialog(temp_root)
        groups = dialog.show()
        return groups
    finally:
        temp_root.destroy()


def create_summary_files():
    """
    Creates one Excel file with separate sheets for each group containing individual summary tables.
    Each sheet will have columns for each file processed, similar to the Box_plot_data.txt structure.
    Groups are now dynamically defined by user input.
    """
    
    # === GET USER INPUT FOR GROUPS ===
    print("Getting group definitions from user...")
    groups = get_user_groups()
    
    if not groups:
        print("No groups defined. Exiting...")
        return
        
    print(f"Using groups: {groups}")
    
    # === VARIABLE DEFINITIONS ===
    # Define the measurements folder path
    measurements_folder = filedialog.askdirectory(title="Select folder containing CSV measurement files")
    
    if not measurements_folder:
        print("No folder selected. Exiting...")
        return
    
    # Dictionary to store data for each group
    group_data = {group: {} for group in groups}
    
    # Create Measurement_summary folder path
    summary_folder = os.path.join(measurements_folder, "Measurement_summary")
    
    # === MAIN PROCESSING ===
    # Check if the folder exists
    if not os.path.exists(measurements_folder):
        print(f"Error: Measurements folder not found at {measurements_folder}")
        return
    
    # Get all CSV files in the measurements folder
    csv_files = [f for f in os.listdir(measurements_folder) if f.endswith('.csv')]
    
    if not csv_files:
        print("No CSV files found in the measurements folder!")
        return
    
    print(f"Found {len(csv_files)} CSV files in the measurements folder")
    
    # Process each CSV file
    for filename in csv_files:
        file_path = os.path.join(measurements_folder, filename)
        
        # Determine which group this file belongs to
        filename_lower = filename.lower()
        group_assigned = None
        
        for group in groups:
            if group.lower() in filename_lower:
                group_assigned = group
                break
        
        if group_assigned is None:
            print(f"Warning: Could not determine group for file {filename}")
            continue
        
        try:
            # Read the CSV file
            df = pd.read_csv(file_path)
            
            # Check if the DataFrame has data
            if df.empty:
                print(f"Warning: File {filename} is empty")
                continue
            
            # DEBUG: Print file structure to understand the data layout
            print(f"\nDebugging {filename}:")
            print(f"  Shape: {df.shape}")
            print(f"  Columns: {list(df.columns)}")
            
            # Look for the SUMMARY column specifically
            summary_col_idx = None
            for i, col in enumerate(df.columns):
                if 'SUMMARY' in str(col):
                    summary_col_idx = i
                    print(f"  Found SUMMARY column at index {i}")
                    break
            
            if summary_col_idx is not None and summary_col_idx + 1 < len(df.columns):
                # Check the next column after SUMMARY for actual data
                data_col_idx = summary_col_idx + 1
                print(f"  Checking data column at index {data_col_idx}: {df.columns[data_col_idx]}")
                print(f"  First 15 values in data column:")
                data_values = df.iloc[:15, data_col_idx].tolist()
                print(f"    {data_values}")
            
            print()
            
            # Extract summary data from the CSV file
            # Look for summary table next to the SUMMARY column
            summary_data = {}
            found_summary = False
            
            # Method 1: Look for SUMMARY column and extract data from the next column
            if summary_col_idx is not None and summary_col_idx + 1 < len(df.columns):
                try:
                    # Get the data column (next to SUMMARY)
                    data_col_idx = summary_col_idx + 1
                    
                    # Look for rows that contain the summary statistics
                    summary_rows = {}
                    for idx, row in df.iterrows():
                        statistic = str(row.iloc[summary_col_idx]).strip()
                        value = row.iloc[data_col_idx]
                        
                        if pd.notna(value) and statistic in ['Num Cells', 'Num Positive', 'Num 1+', 'Num 2+', 'Num 3+', 'Num Negative']:
                            try:
                                summary_rows[statistic] = float(value)
                            except (ValueError, TypeError):
                                pass
                    
                    # If we found the basic cell counts, calculate proportions
                    if 'Num Cells' in summary_rows and summary_rows['Num Cells'] > 0:
                        num_cells = summary_rows['Num Cells']
                        summary_data = {
                            'Num Cells': int(summary_rows.get('Num Cells', 0)),
                            'Num Positive': int(summary_rows.get('Num Positive', 0)),
                            'Num 1+': int(summary_rows.get('Num 1+', 0)),
                            'Num 2+': int(summary_rows.get('Num 2+', 0)),
                            'Num 3+': int(summary_rows.get('Num 3+', 0)),
                            'Num Negative': int(summary_rows.get('Num Negative', 0)),
                            'Prop_1+': (summary_rows.get('Num 1+', 0) / num_cells) * 100,
                            'Prop_2+': (summary_rows.get('Num 2+', 0) / num_cells) * 100,
                            'Prop_3+': (summary_rows.get('Num 3+', 0) / num_cells) * 100,
                            'Prop_Neg': (summary_rows.get('Num Negative', 0) / num_cells) * 100
                        }
                        found_summary = True
                        print(f"  Found summary data from SUMMARY table: {summary_data}")
                
                except Exception as e:
                    print(f"  Error extracting from SUMMARY column: {e}")
            
            # Method 2: If no summary found using SUMMARY column, try other approaches
            if not found_summary:
                print(f"  SUMMARY column method failed, trying alternative extraction...")
                
                # Look for numeric data that could be cell counts
                for col_idx in range(len(df.columns)):
                    try:
                        col_data = df.iloc[:, col_idx].dropna()
                        if len(col_data) >= 6:
                            # Try to extract numeric data that looks like cell counts
                            test_values = []
                            for i in range(min(10, len(col_data))):
                                try:
                                    val = float(col_data.iloc[i])
                                    # Look for values that could be cell counts (typically > 10 and < 10000)
                                    if 10 <= val <= 10000:
                                        test_values.append(val)
                                    elif val == 0:  # Allow zeros
                                        test_values.append(val)
                                except (ValueError, TypeError):
                                    break
                            
                            # If we got reasonable cell count values
                            if len(test_values) >= 6 and test_values[0] > 100:  # First value should be total cells
                                num_cells = test_values[0]
                                summary_data = {
                                    'Num Cells': int(test_values[0]) if len(test_values) > 0 else 0,
                                    'Num Positive': int(test_values[1]) if len(test_values) > 1 else 0,
                                    'Num 1+': int(test_values[2]) if len(test_values) > 2 else 0,
                                    'Num 2+': int(test_values[3]) if len(test_values) > 3 else 0,
                                    'Num 3+': int(test_values[4]) if len(test_values) > 4 else 0,
                                    'Num Negative': int(test_values[5]) if len(test_values) > 5 else 0,
                                    'Prop_1+': (test_values[2] / num_cells * 100) if len(test_values) > 2 else 0.0,
                                    'Prop_2+': (test_values[3] / num_cells * 100) if len(test_values) > 3 else 0.0,
                                    'Prop_3+': (test_values[4] / num_cells * 100) if len(test_values) > 4 else 0.0,
                                    'Prop_Neg': (test_values[5] / num_cells * 100) if len(test_values) > 5 else 0.0
                                }
                                found_summary = True
                                print(f"  Found summary data in column {col_idx}: {summary_data}")
                                break
                    except Exception as e:
                        continue
            
            if not found_summary or not summary_data:
                print(f"Warning: Could not extract summary data from {filename}")
                continue
            
            # Store the data for this group with filename as key
            base_filename = os.path.splitext(filename)[0]
            group_data[group_assigned][base_filename] = summary_data
            
            print(f"Processed {filename} -> {group_assigned} group")
            
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")
            continue
    
    # Check if any data was found before creating Excel file
    total_files_processed = sum(len(group_data[group]) for group in groups)
    
    if total_files_processed == 0:
        print("\nNo summary data was extracted from any files!")
        print("Please check that the CSV files contain the expected summary data structure.")
        return
    
    # Create the Excel file with separate sheets for each group
    os.makedirs(summary_folder, exist_ok=True)
    excel_path = os.path.join(summary_folder, "Combined_Summary.xlsx")
    
    # Define the row labels (first column)
    row_labels = [
        'Num Cells',
        'Num Positive', 
        'Num 1+',
        'Num 2+',
        'Num 3+',
        'Num Negative',
        'Prop_1+',
        'Prop_2+',
        'Prop_3+',
        'Prop_Neg'
    ]
    
    sheets_created = 0
    
    if EXCEL_ENGINE:
        # Create Excel file with multiple sheets
        # Generate dynamic sheet names mapping
        sheet_name_mapping = {}
        for group in groups:
            # Capitalize first letter and clean up the group name for sheet naming
            clean_name = group.replace('_', ' ').replace('-', ' ').title()
            sheet_name_mapping[group] = clean_name
        
        with pd.ExcelWriter(excel_path, engine=EXCEL_ENGINE) as writer:
            # First, create individual group sheets
            for group in groups:
                if not group_data[group]:
                    print(f"No data found for {group} group")
                    continue
                
                # Create DataFrame for this group
                files_in_group = list(group_data[group].keys())
                
                # Get proper sheet name
                sheet_name = sheet_name_mapping.get(group, group.capitalize())
                
                # Create column headers (Count + filenames)
                column_headers = ['Count'] + [f"{sheet_name}{i+1}" for i in range(len(files_in_group))]
                
                # Initialize the data matrix
                data_matrix = []
                
                for row_label in row_labels:
                    row_data = [row_label]  # Start with the row label
                    
                    # Add data from each file
                    for filename in files_in_group:
                        file_data = group_data[group][filename]
                        value = file_data.get(row_label, 0)
                        row_data.append(value)
                    
                    data_matrix.append(row_data)
                
                # Create DataFrame
                sheet_df = pd.DataFrame(data_matrix, columns=column_headers)
                
                # Write to Excel sheet with proper name
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                sheets_created += 1
                
                print(f"Created sheet '{sheet_name}' with {len(files_in_group)} files")
            
            # Create Summary sheet with totals across all conditions
            if sheets_created > 0:
                # Calculate totals for each classification across all groups
                summary_data = []
                
                # Get only the count-based measurements (not proportions)
                count_labels = ['Num Cells', 'Num Positive', 'Num 1+', 'Num 2+', 'Num 3+', 'Num Negative']
                
                for row_label in count_labels:
                    row_data = [row_label]  # Start with the classification name
                    
                    # Calculate total for each group
                    for group in groups:
                        group_total = 0
                        if group_data[group]:  # If group has data
                            for filename in group_data[group]:
                                file_data = group_data[group][filename]
                                value = file_data.get(row_label, 0)
                                group_total += value
                        row_data.append(group_total)
                    
                    # Calculate grand total across all groups
                    grand_total = sum(row_data[1:])
                    row_data.append(grand_total)
                    
                    summary_data.append(row_data)
                
                # Create column headers for summary sheet
                summary_headers = ['Classification'] + [sheet_name_mapping.get(group, group.capitalize()) for group in groups] + ['Grand Total']
                
                # Create DataFrame for summary
                summary_df = pd.DataFrame(summary_data, columns=summary_headers)
                
                # Write summary sheet
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                print(f"Created 'Summary' sheet with totals across all conditions")
        
        if sheets_created > 0:
            print(f"\nCombined Excel file created: {excel_path}")
            print(f"Created {sheets_created} group sheets plus Summary sheet with totals")
        else:
            print("No sheets were created - no valid data found!")
    else:
        # Fallback: Create separate CSV files for each group
        print("Excel libraries not available. Creating CSV files instead...")
        for group in groups:
            if not group_data[group]:
                print(f"No data found for {group} group")
                continue
            
            # Create DataFrame for this group
            files_in_group = list(group_data[group].keys())
            
            # Create column headers (Count + filenames)
            column_headers = ['Count'] + [f"{group.capitalize()}{i+1}" for i in range(len(files_in_group))]
            
            # Initialize the data matrix
            data_matrix = []
            
            for row_label in row_labels:
                row_data = [row_label]  # Start with the row label
                
                # Add data from each file
                for filename in files_in_group:
                    file_data = group_data[group][filename]
                    value = file_data.get(row_label, 0)
                    row_data.append(value)
                
                data_matrix.append(row_data)
            
            # Create DataFrame
            sheet_df = pd.DataFrame(data_matrix, columns=column_headers)
            
            # Save as CSV
            csv_path = os.path.join(summary_folder, f"{group}_summary.csv")
            sheet_df.to_csv(csv_path, index=False)
            sheets_created += 1
            
            print(f"Created CSV file: {group}_summary.csv with {len(files_in_group)} files")
    
    if sheets_created > 0:
        print("Summary file creation completed!")
    else:
        print("No summary files were created - please check your data structure.")

def analyze_folder_structure():
    """
    Helper function to analyze the structure of files in the measurements folder.
    """
    measurements_folder = r"C:\Users\psoor\OneDrive\Desktop\Third QuPath\Measurements"
    
    if not os.path.exists(measurements_folder):
        print(f"Error: Measurements folder not found at {measurements_folder}")
        return
    
    csv_files = [f for f in os.listdir(measurements_folder) if f.endswith('.csv')]
    
    print(f"Analysis of {len(csv_files)} CSV files:")
    print("-" * 50)
    
    for filename in csv_files[:5]:  # Analyze first 5 files
        file_path = os.path.join(measurements_folder, filename)
        try:
            df = pd.read_csv(file_path)
            print(f"\nFile: {filename}")
            print(f"  Shape: {df.shape}")
            print(f"  Columns: {list(df.columns)[:10]}")  # Show first 10 columns
            print(f"  Numeric columns: {len(df.select_dtypes(include=[np.number]).columns)}")
        except Exception as e:
            print(f"  Error reading {filename}: {str(e)}")

if __name__ == "__main__":
    print("QuPath Measurements Summarizer")
    print("=" * 40)
    
    # Uncomment the line below to analyze folder structure first
    # analyze_folder_structure()
    
    # Create the summary files
    create_summary_files()
