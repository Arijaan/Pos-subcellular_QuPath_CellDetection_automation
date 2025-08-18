import pandas as pd
import os
import numpy as np
from tkinter import filedialog


def create_summary_files():
    """
    Creates simplified summary CSV files for control, T1D, and T2D groups.
    Summary table starts at G1 (6th column, first row).
    """
    
    # === VARIABLE DEFINITIONS ===
    # Define the measurements folder path
    measurements_folder = filedialog.askdirectory()
    
    # Define the groups we're looking for
    groups = ['control', 'T1D', 'T2D']
    
    # Dictionary to store data for each group
    group_data = {group: [] for group in groups}
    
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
            
            # Look for both measurement data and summary data
            # Try to find numeric columns that could be measurements
            numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
            
            # Look for summary-style data (like the Groovy script produces)
            summary_data = {}
            
            # Check if this file has the summary structure from Groovy script
            if 'SUMMARY' in df.columns or any('SUMMARY' in str(col) for col in df.columns):
                print(f"Found summary data in {filename}")
                # Look for summary information in the file
                for col in df.columns:
                    if 'Statistic' in str(col) or 'SUMMARY' in str(col):
                        # Find the corresponding values column
                        col_idx = df.columns.get_loc(col)
                        if col_idx + 1 < len(df.columns):
                            stats_col = col
                            values_col = df.columns[col_idx + 1]
                            
                            for idx, row in df.iterrows():
                                stat_name = row[stats_col]
                                stat_value = row[values_col]
                                if pd.notna(stat_name) and pd.notna(stat_value):
                                    if 'Num' in str(stat_name):  # Cell count statistics
                                        try:
                                            summary_data[str(stat_name)] = float(stat_value)
                                        except (ValueError, TypeError):
                                            pass
            
            # If no summary data found, calculate basic statistics from numeric data
            if not summary_data and numeric_columns:
                summary_stats = df[numeric_columns].mean()
            else:
                # If we found summary data, use that; otherwise use column means
                summary_stats = pd.Series(summary_data) if summary_data else df[numeric_columns].mean()
            
            if summary_stats.empty:
                print(f"Warning: No usable data found in {filename}")
                continue
            
            # Store the data for this group
            group_data[group_assigned].append({
                'filename': filename,
                'data': summary_stats
            })
            
            print(f"Processed {filename} -> {group_assigned} group")
            
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")
            continue
    
    # Create summary files for each group
    for group in groups:
        if not group_data[group]:
            print(f"Warning: No data found for {group} group")
            continue
        
        # Combine all data for this group
        all_measurements = []
        for file_data in group_data[group]:
            all_measurements.append(file_data['data'])
        
        if not all_measurements:
            continue
        
        # Create a DataFrame from all measurements
        combined_df = pd.DataFrame(all_measurements)
        
        # Calculate the average for regular measurements and sum for cell counts
        group_averages = combined_df.mean()
        group_sums = combined_df.sum()
        
        # Identify which measurements should be summed (cell counts) vs averaged
        cell_count_keywords = ['Num Cells', 'Num Positive', 'Num 1+', 'Num 2+', 'Num 3+', 'Num Negative', 'Total']
        
        # Create final summary combining sums for cell counts and averages for other measurements
        final_summary = {}
        for measurement in group_averages.index:
            measurement_str = str(measurement)
            if any(keyword in measurement_str for keyword in cell_count_keywords):
                # Use sum for cell counts
                final_summary[measurement] = group_sums[measurement]
            else:
                # Use average for other measurements
                final_summary[measurement] = group_averages[measurement]
        
        group_summary = pd.Series(final_summary)
        
        # Create simplified summary table starting at the first row, columns G and H (6th and 7th columns)
        # Create DataFrame without column names
        summary_df = pd.DataFrame()
        
        # Calculate cell count totals for this group
        total_cells = sum([group_summary.get(key, 0) for key in group_summary.index if 'Num Cells' in str(key)])
        total_positive = sum([group_summary.get(key, 0) for key in group_summary.index if any(x in str(key) for x in ['Num 1+', 'Num 2+', 'Num 3+'])])
        num_1_plus = sum([group_summary.get(key, 0) for key in group_summary.index if 'Num 1+' in str(key)])
        num_2_plus = sum([group_summary.get(key, 0) for key in group_summary.index if 'Num 2+' in str(key)])
        num_3_plus = sum([group_summary.get(key, 0) for key in group_summary.index if 'Num 3+' in str(key)])
        num_negative = sum([group_summary.get(key, 0) for key in group_summary.index if 'Num Negative' in str(key)])
        
        # Create summary data starting at row 0
        summary_data = [
            ['SUMMARY', ''],
            ['Statistic', 'Count'],
            ['Num Cells', int(total_cells) if total_cells > 0 else 0],
            ['Num Positive', int(total_positive) if total_positive > 0 else 0],
            ['Num 1+', int(num_1_plus) if num_1_plus > 0 else 0],
            ['Num 2+', int(num_2_plus) if num_2_plus > 0 else 0],
            ['Num 3+', int(num_3_plus) if num_3_plus > 0 else 0],
            ['Num Negative', int(num_negative) if num_negative > 0 else 0]
        ]
        
        # Moving the data to the columns E and F to match the pos cell measurements format for the final R-plot
        for i, (stat_name, stat_value) in enumerate(summary_data):
            summary_df.loc[i, 0] = ''  # Column A
            summary_df.loc[i, 1] = ''  # Column B
            summary_df.loc[i, 2] = ''  # Column C
            summary_df.loc[i, 3] = ''  # Column D
            summary_df.loc[i, 4] = stat_name   # Column E
            summary_df.loc[i,5] = stat_value # Column F
            
        # Create output filename and path
        output_filename = f"{group}_summary.csv"
        
        # Create Measurement_summary folder if it doesn't exist
        os.makedirs(summary_folder, exist_ok=True)
        
        output_path = os.path.join(summary_folder, output_filename)
        
        # Save the summary file without column headers
        summary_df.to_csv(output_path, index=False, header=False, na_rep='')
        
        print(f"Created summary file: {output_filename}")
        print(f"  - Number of files processed: {len(group_data[group])}")
        print(f"  - Summary saved to: {output_path}")
        print()
    
    print("Summary file creation completed!")

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
