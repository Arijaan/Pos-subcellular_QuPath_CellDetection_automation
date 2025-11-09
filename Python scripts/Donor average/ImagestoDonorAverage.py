#!/usr/bin/env python3
"""
Combined Summary to Combined Donor Averager

This script processes combined_summary.xlsx files to average multiple columns
of the same condition into a single averaged column.

Example:
Control1, Control2, Control3, Control4 → Control (averaged)
T1D1, T1D2, T1D3 → T1D (averaged)

Creates output: combined_donor.xlsx

Author: Generated for QuPath measurement analysis
"""

import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import re


def extract_condition_groups(columns):
    """
    Group columns by their base condition name.
    
    Args:
        columns (list): List of column names
        
    Returns:
        dict: Dictionary with condition as key and list of columns as value
    """
    condition_groups = {}
    
    for col in columns:
        if col.lower() in ['count', 'classification']:
            continue
            
        # Extract base condition name (remove numbers at the end)
        base_condition = re.sub(r'\d+$', '', col)
        
        if base_condition not in condition_groups:
            condition_groups[base_condition] = []
        condition_groups[base_condition].append(col)
    
    return condition_groups


def process_sheet(df, sheet_name):
    """
    Process a single sheet to average columns by condition.
    
    Args:
        df (DataFrame): Input sheet data
        sheet_name (str): Name of the sheet
        
    Returns:
        DataFrame: Processed sheet with averaged columns
    """
    print(f"  📄 Processing {sheet_name}...")
    
    if df.empty:
        print(f"    ⚠️  Empty sheet")
        return None
    
    # Identify the row identifier column (Count or Classification)
    row_col = None
    if 'Count' in df.columns:
        row_col = 'Count'
    elif 'Classification' in df.columns:
        row_col = 'Classification'
    else:
        print(f"    ❌ No Count or Classification column found")
        return None
    
    # Get condition groups
    numeric_columns = [col for col in df.columns if col != row_col]
    condition_groups = extract_condition_groups(numeric_columns)
    
    if not condition_groups:
        print(f"    ⚠️  No condition columns found")
        return None
    
    print(f"    📊 Found conditions: {list(condition_groups.keys())}")
    
    # Create new DataFrame with averaged columns
    result_data = {row_col: df[row_col].tolist()}
    
    for condition, columns in condition_groups.items():
        print(f"      🔄 {condition}: averaging {len(columns)} columns")
        
        # Average the columns for this condition
        condition_values = []
        for _, row in df.iterrows():
            values = []
            for col in columns:
                if col in df.columns:
                    try:
                        val = pd.to_numeric(row[col], errors='coerce')
                        if not pd.isna(val):
                            values.append(val)
                    except:
                        pass
            
            # Calculate average for this row
            if values:
                avg_value = np.mean(values)
            else:
                avg_value = 0.0
            
            condition_values.append(avg_value)
        
        result_data[condition] = condition_values
    
    result_df = pd.DataFrame(result_data)
    print(f"    ✅ Created averaged sheet: {len(result_df)} rows × {len(result_df.columns)} columns")
    
    return result_df


def process_combined_summary(file_path):
    """
    Process the entire combined_summary.xlsx file.
    
    Args:
        file_path (str): Path to the input file
        
    Returns:
        dict: Dictionary of processed sheets
    """
    try:
        xl_file = pd.ExcelFile(file_path)
        print(f"📊 Processing: {os.path.basename(file_path)}")
        print(f"📋 Found {len(xl_file.sheet_names)} sheets")
        
        processed_sheets = {}
        
        for sheet_name in xl_file.sheet_names:
            if sheet_name.lower() == 'summary':
                print(f"  📄 {sheet_name} (SKIPPED)")
                continue
            
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                processed_df = process_sheet(df, sheet_name)
                
                if processed_df is not None:
                    processed_sheets[sheet_name] = processed_df
                    
            except Exception as e:
                print(f"  ❌ Error processing {sheet_name}: {e}")
        
        return processed_sheets
        
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return {}


def create_summary_sheet(processed_sheets):
    """
    Create a summary sheet from all processed sheets.
    
    Args:
        processed_sheets (dict): Dictionary of processed DataFrames
        
    Returns:
        DataFrame: Summary sheet
    """
    if not processed_sheets:
        return None
    
    print(f"\n📋 Creating Summary sheet...")
    
    # Get all unique conditions across all sheets
    all_conditions = set()
    row_identifiers = None
    
    for sheet_df in processed_sheets.values():
        # Get row identifiers from first sheet
        if row_identifiers is None:
            row_col = 'Count' if 'Count' in sheet_df.columns else 'Classification'
            row_identifiers = sheet_df[row_col].tolist()
        
        # Collect all condition columns
        for col in sheet_df.columns:
            if col not in ['Count', 'Classification']:
                all_conditions.add(col)
    
    all_conditions = sorted(list(all_conditions))
    
    # Create summary data
    summary_data = {'Count': row_identifiers}
    
    for condition in all_conditions:
        condition_values = []
        
        for row_idx in range(len(row_identifiers)):
            values = []
            
            # Collect values for this condition across all sheets
            for sheet_df in processed_sheets.values():
                if condition in sheet_df.columns:
                    try:
                        val = pd.to_numeric(sheet_df.iloc[row_idx][condition], errors='coerce')
                        if not pd.isna(val):
                            values.append(val)
                    except:
                        pass
            
            # Average across sheets
            if values:
                avg_value = np.mean(values)
            else:
                avg_value = 0.0
            
            condition_values.append(avg_value)
        
        summary_data[condition] = condition_values
    
    summary_df = pd.DataFrame(summary_data)
    print(f"✅ Summary sheet: {len(summary_df)} rows × {len(all_conditions)} conditions")
    
    return summary_df


def save_combined_donor(summary_df, processed_sheets, original_file_path):
    """
    Save the results as combined_(inputfilename).xlsx
    
    Args:
        summary_df (DataFrame): Summary sheet
        processed_sheets (dict): Individual processed sheets
        original_file_path (str): Path to original file
        
    Returns:
        str: Path to output file or None if failed
    """
    try:
        input_dir = os.path.dirname(original_file_path)
        input_filename = os.path.splitext(os.path.basename(original_file_path))[0]
        output_filename = f"combined_{input_filename}.xlsx"
        output_path = os.path.join(input_dir, output_filename)
        
        print(f"\n💾 Saving to: {output_filename}")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write Summary sheet first
            if summary_df is not None:
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                print(f"  📄 Summary: {len(summary_df)} rows")
            
            # Write processed sheets
            for sheet_name, df in processed_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  📄 {sheet_name}: {len(df)} rows")
        
        print(f"✅ Saved: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"❌ Error saving file: {e}")
        return None


def display_preview(summary_df, input_filename=None):
    """Display a preview of the summary results."""
    output_name = f"combined_{input_filename}.xlsx" if input_filename else "combined_donor.xlsx"
    
    print(f"\n{'='*60}")
    print(f"PREVIEW OF {output_name.upper()}")
    print(f"{'='*60}")
    
    if summary_df is not None and not summary_df.empty:
        print("\n📊 SUMMARY SHEET:")
        print(summary_df.to_string(index=False))
    else:
        print("⚠️  No summary data to display")


def main():
    """Main function."""
    print("=" * 60)
    print("COMBINED SUMMARY → COMBINED DONOR CONVERTER")
    print("=" * 60)
    print("Averages multiple condition columns into single columns")
    print("Example: Control1, Control2, Control3 → Control")
    
    # Select input file
    print("\n📂 Select combined_summary.xlsx file...")
    
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="Select combined_summary.xlsx file",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    
    root.destroy()
    
    if not file_path:
        print("❌ No file selected. Exiting...")
        return
    
    print(f"📁 Input: {os.path.basename(file_path)}")
    
    # Extract input filename for output naming
    input_filename = os.path.splitext(os.path.basename(file_path))[0]
    
    # Process the file
    processed_sheets = process_combined_summary(file_path)
    
    if not processed_sheets:
        print("❌ No sheets could be processed. Exiting...")
        return
    
    # Create summary sheet
    summary_df = create_summary_sheet(processed_sheets)
    
    # Display preview
    display_preview(summary_df, input_filename)
    
    # Save results
    output_path = save_combined_donor(summary_df, processed_sheets, file_path)
    
    if output_path:
        output_filename = os.path.basename(output_path)
        # Final summary
        print(f"\n{'='*60}")
        print("✅ CONVERSION COMPLETE!")
        print(f"{'='*60}")
        print(f"📁 Input:  {os.path.basename(file_path)}")
        print(f"📁 Output: {output_filename}")
        print(f"📊 Sheets processed: {len(processed_sheets)}")
        print(f"💾 Location: {os.path.dirname(output_path)}")
        print("\n🎉 Ready for analysis!")
    else:
        print("❌ Failed to save output file.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n❌ Operation cancelled by user.")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
    
    input("\nPress Enter to exit...")
