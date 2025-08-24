import pandas as pd
import os
import numpy as np
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
    Creates one Excel file with separate sheets for each group containing individual summary tables.
    Each sheet will have columns for each file processed, similar to the Box_plot_data.txt structure.
    """
    
    # === VARIABLE DEFINITIONS ===
    # Define the measurements folder path
    measurements_folder = filedialog.askdirectory()
    
    # Define the groups we're looking for
    groups = ['control', 't1d', 't2d', 'aab']
    
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
        # Define proper sheet names mapping
        sheet_name_mapping = {
            'control': 'Control',
            't1d': 'T1D',
            't2d': 'T2D',
            'aab': 'Aab'
        }
        
        with pd.ExcelWriter(excel_path, engine=EXCEL_ENGINE) as writer:
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
        
        if sheets_created > 0:
            print(f"\nCombined Excel file created: {excel_path}")
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
