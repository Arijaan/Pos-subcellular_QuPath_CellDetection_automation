"""
File I/O Module
Handles file operations, Excel export, and data saving
"""

import pandas as pd
import os
from pathlib import Path


def save_summary_tables(summarized_data, plots_folder):
    """Save summary tables to Excel files"""
    print("\n" + "="*50)
    print("SAVING SUMMARY TABLES")
    print("="*50)
    
    if not summarized_data:
        print("No data to save")
        return
    
    # Create Excel file with multiple sheets
    excel_path = os.path.join(plots_folder, "Summary_Tables.xlsx")
    
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for condition, df in summarized_data.items():
                # Clean sheet name (Excel sheet names have restrictions)
                sheet_name = "".join(c for c in condition if c.isalnum() or c in (' ', '-', '_'))[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Saved {condition} data to sheet: {sheet_name}")
            
            # Create combined summary sheet
            all_data = []
            for condition, df in summarized_data.items():
                df_copy = df.copy()
                df_copy['Condition'] = condition
                all_data.append(df_copy)
            
            if all_data:
                combined_df = pd.concat(all_data, ignore_index=True)
                
                # Create summary statistics
                numeric_cols = combined_df.select_dtypes(include=[pd.api.types.is_numeric_dtype]).columns.tolist()
                if 'Condition' in numeric_cols:
                    numeric_cols.remove('Condition')
                
                if numeric_cols:
                    summary_stats = combined_df.groupby('Condition')[numeric_cols].agg(['mean', 'std', 'count'])
                    
                    # Flatten column names
                    summary_stats.columns = [f"{col}_{stat}" for col, stat in summary_stats.columns]
                    summary_stats.reset_index(inplace=True)
                    
                    summary_stats.to_excel(writer, sheet_name='Summary_Statistics', index=False)
                    print("Saved summary statistics")
            
        print(f"Excel file saved: {excel_path}")
            
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        
        # Fallback: save individual CSV files
        print("Falling back to individual CSV files...")
        for condition, df in summarized_data.items():
            csv_path = os.path.join(plots_folder, f"{condition}_summary_table.csv")
            df.to_csv(csv_path, index=False)
            print(f"Saved CSV: {csv_path}")


def create_results_summary(base_dir, plots_folder, summarized_data):
    """Create a summary report of the analysis results"""
    summary_path = os.path.join(plots_folder, "Analysis_Summary.txt")
    
    with open(summary_path, 'w') as f:
        f.write("QuPath Measurements Analysis Summary\n")
        f.write("=" * 50 + "\n\n")
        
        f.write(f"Analysis Date: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Base Directory: {base_dir}\n")
        f.write(f"Results Directory: {plots_folder}\n\n")
        
        f.write("Data Summary:\n")
        f.write("-" * 20 + "\n")
        
        if summarized_data:
            for condition, df in summarized_data.items():
                f.write(f"{condition}: {df.shape[0]} samples, {df.shape[1]} variables\n")
                
            # Count numeric variables
            sample_df = list(summarized_data.values())[0]
            numeric_cols = sample_df.select_dtypes(include=[pd.api.types.is_numeric_dtype]).columns.tolist()
            f.write(f"\nNumeric variables analyzed: {len(numeric_cols)}\n")
            
            # List plot files created
            f.write("\nGenerated Files:\n")
            f.write("-" * 20 + "\n")
            
            plot_files = [f for f in os.listdir(plots_folder) if f.endswith('.png')]
            for plot_file in plot_files:
                f.write(f"- {plot_file}\n")
                
            if os.path.exists(os.path.join(plots_folder, "Summary_Tables.xlsx")):
                f.write("- Summary_Tables.xlsx\n")
        else:
            f.write("No summarized data was available for analysis.\n")
    
    print(f"Created analysis summary: {summary_path}")


def validate_file_structure(base_dir):
    """Validate the expected file structure exists"""
    required_folders = ['Measurement_summary']
    optional_folders = ['summarized', 'Plots_and_Analysis']
    
    issues = []
    
    # Check for required folders
    for folder in required_folders:
        folder_path = os.path.join(base_dir, folder)
        if not os.path.exists(folder_path):
            issues.append(f"Missing required folder: {folder}")
    
    # Check for CSV files in measurements folder
    csv_files = [f for f in os.listdir(base_dir) if f.endswith('.csv')]
    if not csv_files:
        issues.append("No CSV files found in base directory")
    
    return issues


def backup_original_files(base_dir):
    """Create backup of original files before processing"""
    backup_dir = os.path.join(base_dir, "Backup_Original")
    os.makedirs(backup_dir, exist_ok=True)
    
    # Backup CSV files
    csv_files = [f for f in os.listdir(base_dir) if f.endswith('.csv')]
    backed_up_count = 0
    
    for csv_file in csv_files:
        source_path = os.path.join(base_dir, csv_file)
        backup_path = os.path.join(backup_dir, csv_file)
        
        if not os.path.exists(backup_path):
            try:
                import shutil
                shutil.copy2(source_path, backup_path)
                backed_up_count += 1
            except Exception as e:
                print(f"Warning: Could not backup {csv_file}: {e}")
    
    if backed_up_count > 0:
        print(f"Backed up {backed_up_count} original files to: {backup_dir}")
    
    return backup_dir


def ensure_directories_exist(base_dir):
    """Ensure all required directories exist"""
    directories = [
        os.path.join(base_dir, "Measurement_summary"),
        os.path.join(base_dir, "Plots_and_Analysis"),
        os.path.join(base_dir, "summarized")
    ]
    
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
    
    print("✓ Directory structure validated")
