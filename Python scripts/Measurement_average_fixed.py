import pandas as pd
import os
import numpy as np
import tkinter as tk
from tkinter import filedialog

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

def create_summary_files():
    """
    Processes subdirectories containing CSV files and creates averaged summary files for each subdirectory.
    Each subdirectory will get its own Excel file with separate sheets for each group containing averaged values.
    """
    
    # Create file dialog for selecting main directory
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Ask user to select the main directory containing subdirectories
    main_directory = filedialog.askdirectory(title="Select main directory containing subdirectories with CSV files")
    
    if not main_directory:
        print("No directory selected. Exiting...")
        return
    
    # Find all subdirectories that contain CSV files
    subdirs_with_csv = []
    for item in os.listdir(main_directory):
        item_path = os.path.join(main_directory, item)
        if os.path.isdir(item_path):
            # Check if this subdirectory contains CSV files
            csv_files_in_subdir = [f for f in os.listdir(item_path) if f.lower().endswith('.csv')]
            if csv_files_in_subdir:
                subdirs_with_csv.append((item_path, csv_files_in_subdir))
    
    if not subdirs_with_csv:
        print("No subdirectories with CSV files found. Exiting...")
        return
    
    print(f"Found {len(subdirs_with_csv)} subdirectories with CSV files:")
    for subdir_path, csv_files in subdirs_with_csv:
        subdir_name = os.path.basename(subdir_path)
        print(f"  - {subdir_name}: {len(csv_files)} CSV files")
    
    # Ask user where to save the summaries
    output_folder = filedialog.askdirectory(title="Select folder to save all summary files")
    
    if not output_folder:
        print("No output folder selected. Exiting...")
        return
    
    # Process each subdirectory
    total_summaries_created = 0
    
    for subdir_path, csv_files_in_subdir in subdirs_with_csv:
        subdir_name = os.path.basename(subdir_path)
        print(f"\n{'='*50}")
        print(f"Processing subdirectory: {subdir_name}")
        print(f"{'='*50}")
        
        # Get full paths for CSV files in this subdirectory
        csv_files = [os.path.join(subdir_path, csv_file) for csv_file in csv_files_in_subdir]
        
        # Create a specific output folder for this subdirectory
        summary_folder = os.path.join(output_folder, f"{subdir_name}_summary")
        os.makedirs(summary_folder, exist_ok=True)
        
        # Process this subdirectory's CSV files
        summaries_created = process_csv_group(csv_files, summary_folder, subdir_name)
        total_summaries_created += summaries_created
        
        print(f"Completed processing {subdir_name}: {summaries_created} summary files created")
    
    print(f"\n{'='*50}")
    print(f"PROCESSING COMPLETE")
    print(f"Total subdirectories processed: {len(subdirs_with_csv)}")
    print(f"Total summary files created: {total_summaries_created}")
    print(f"{'='*50}")


def process_csv_group(csv_files, summary_folder, subdir_name):
    """
    Process a group of CSV files and create averaged summary files.
    Returns the number of summary files created.
    """
    
    # === VARIABLE DEFINITIONS ===
    # Define the groups we're looking for
    groups = ['control', 't1d', 't2d', 'aab']
    
    # Dictionary to store data for each group
    group_data = {group: {} for group in groups}
    
    print(f"Found {len(csv_files)} CSV files to process")
    
    # === MAIN PROCESSING ===
    # Process each CSV file
    for file_path in csv_files:
        filename = os.path.basename(file_path)
        
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
            
            # Extract base filename (without extension and path)
            base_filename = os.path.splitext(filename)[0]
            
            # Look for numerical data that represents cell measurements
            summary_data = None
            
            # Try different column indexes to find the data
            found_data = False
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
                            found_data = True
                            break
                
                except Exception as e:
                    continue
            
            if not found_data:
                print(f"Warning: No suitable numerical data found in {filename}")
                continue
            
            # Store the summary data
            group_data[group_assigned][base_filename] = summary_data
            print(f"Processed {filename} -> {group_assigned} group")
            
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")
            continue
    
    # Print summary of processed files
    total_files_processed = sum(len(group_data[group]) for group in groups)
    print(f"\nProcessed {total_files_processed} files successfully:")
    for group in groups:
        count = len(group_data[group])
        if count > 0:
            print(f"  {group}: {count} files")
    
    if total_files_processed == 0:
        print("No files were processed successfully. Please check your CSV structure.")
        return 0
    
    # === CREATE SUMMARY FILES ===
    # Create the output folder if it doesn't exist
    os.makedirs(summary_folder, exist_ok=True)
    
    # Define Excel file path
    excel_path = os.path.join(summary_folder, f"{subdir_name}_averaged_summary.xlsx")
    
    # Define the row labels
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
        # Define proper sheet names mapping
        sheet_name_mapping = {
            'control': 'Control',
            't1d': 'T1D',
            't2d': 'T2D',
            'aab': 'Aab'
        }
        
        with pd.ExcelWriter(excel_path, engine=EXCEL_ENGINE) as writer:
            # First, create individual group sheets with averaged data
            for group in groups:
                if not group_data[group]:
                    print(f"No data found for {group} group")
                    continue
                
                # Get proper sheet name
                sheet_name = sheet_name_mapping.get(group, group.capitalize())
                
                # Calculate averages for this group
                files_in_group = list(group_data[group].keys())
                num_files = len(files_in_group)
                
                print(f"Creating {sheet_name} sheet with averaged data from {num_files} files...")
                
                # Initialize the data matrix with averaged values
                data_matrix = []
                
                for row_label in row_labels:
                    # Collect all values for this row label across files
                    values = []
                    for filename in files_in_group:
                        file_data = group_data[group][filename]
                        value = file_data.get(row_label, 0)
                        values.append(value)
                    
                    # Calculate average
                    if values:
                        avg_value = np.mean(values)
                        # Format proportions to 2 decimal places, others as integers for counts
                        if row_label.startswith("Prop_"):
                            avg_value = round(avg_value, 2)
                        else:
                            avg_value = int(round(avg_value))
                    else:
                        avg_value = 0
                    
                    # Create row with just the label and averaged value
                    row_data = [row_label, avg_value]
                    data_matrix.append(row_data)
                
                # Create DataFrame with averaged data
                column_headers = ['Count', f'{group.capitalize()}_avg']
                sheet_df = pd.DataFrame(data_matrix, columns=column_headers)
                
                # Write to Excel sheet
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                sheets_created += 1
            
            # Create summary sheet with all groups averaged
            if sheets_created > 0:
                print("Creating summary sheet with all groups...")
                
                # Calculate summary data for all groups
                summary_data = []
                
                for row_label in row_labels:
                    summary_row = [row_label]
                    
                    for group in groups:
                        if group_data[group]:  # If group has data
                            values = []
                            for filename in group_data[group]:
                                file_data = group_data[group][filename]
                                value = file_data.get(row_label, 0)
                                values.append(value)
                            
                            if values:
                                avg_value = np.mean(values)
                                # Format proportions to 2 decimal places, others as integers for counts
                                if row_label.startswith("Prop_"):
                                    avg_value = round(avg_value, 2)
                                else:
                                    avg_value = int(round(avg_value))
                            else:
                                avg_value = 0
                        else:
                            avg_value = 0
                        
                        summary_row.append(avg_value)
                    
                    summary_data.append(summary_row)
                
                # Create summary sheet headers
                summary_headers = ['Classification'] + [f"{sheet_name_mapping.get(group, group.capitalize())}_avg" for group in groups]
                summary_df = pd.DataFrame(summary_data, columns=summary_headers)
                
                # Write summary sheet
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                sheets_created += 1
        
        print(f"Excel file created: {excel_path}")
    
    else:
        # Fallback: Create separate CSV files for each group with averaged data
        print("Excel libraries not available. Creating CSV files with averaged data instead...")
        for group in groups:
            if not group_data[group]:
                print(f"No data found for {group} group")
                continue
            
            # Calculate averages for this group
            files_in_group = list(group_data[group].keys())
            num_files = len(files_in_group)
            
            print(f"Averaging {num_files} files for {group} group...")
            
            # Initialize the data matrix with averaged values
            data_matrix = []
            
            for row_label in row_labels:
                # Collect all values for this row label across files
                values = []
                for filename in files_in_group:
                    file_data = group_data[group][filename]
                    value = file_data.get(row_label, 0)
                    values.append(value)
                
                # Calculate average
                if values:
                    avg_value = np.mean(values)
                    # Format proportions to 2 decimal places, others as integers for counts
                    if row_label.startswith("Prop_"):
                        avg_value = round(avg_value, 2)
                    else:
                        avg_value = int(round(avg_value))
                else:
                    avg_value = 0
                
                # Create row with just the label and averaged value
                row_data = [row_label, avg_value]
                data_matrix.append(row_data)
            
            # Create DataFrame with averaged data
            column_headers = ['Count', f'{group.capitalize()}_avg']
            sheet_df = pd.DataFrame(data_matrix, columns=column_headers)
            
            # Save as CSV
            csv_path = os.path.join(summary_folder, f"{group}_averaged_summary.csv")
            sheet_df.to_csv(csv_path, index=False)
            sheets_created += 1
            
            print(f"Created CSV file: {group}_averaged_summary.csv with averaged data from {num_files} files")
    
    if sheets_created > 0:
        print(f"Summary file creation completed for {subdir_name}!")
    else:
        print(f"No summary files were created for {subdir_name} - please check your data structure.")
    
    return sheets_created


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
    print("QuPath Measurements Summarizer - Directory Version")
    print("=" * 50)
    
    # Uncomment the line below to analyze folder structure first
    # analyze_folder_structure()
    
    # Create the summary files
    create_summary_files()
