#!/usr/bin/env python3
"""
Directory-Based Combined Summary to Patient Consolidated Averager

This script processes a directory of combined_summary.xlsx files to:
1. Convert each file: multiple columns → averaged columns (combined_*.xlsx)
2. Consolidate all combined files into one final patient file

Processing Flow:
Directory with: 1.07.xlsx, 2.08.xlsx, etc.
↓
Step 1: Each file → combined_1.07.xlsx, combined_2.08.xlsx 
(Control1, Control2 → Control averaged)
↓
Step 2: All combined files → total_patient_[dates].xlsx
(Control from each file → Control1, Control2, etc.)

Author: Generated for QuPath measurement analysis
"""

import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import re
from datetime import datetime
import glob


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
    Process a single sheet to average columns by condition and calculate proportions.
    
    Args:
        df (DataFrame): Input sheet data
        sheet_name (str): Name of the sheet
        
    Returns:
        DataFrame: Processed sheet with averaged columns and proportions
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
    
    # Calculate proportions and add them as additional rows
    result_df = add_proportion_rows(result_df, row_col)
    
    print(f"    ✅ Created averaged sheet with proportions: {len(result_df)} rows × {len(result_df.columns)} columns")
    
    return result_df


def add_proportion_rows(df, row_col):
    """
    Add proportion calculations as additional rows following the raw data.
    Restructured to handle cell count data with Num/Prop patterns.
    
    Args:
        df (DataFrame): Input DataFrame with raw data
        row_col (str): Name of the row identifier column
        
    Returns:
        DataFrame: DataFrame with proportion rows added
    """
    # Get numeric columns (conditions)
    numeric_columns = [col for col in df.columns if col != row_col]
    
    # Get row identifiers
    row_identifiers = df[row_col].tolist()
    
    # Find indices for different row types
    num_cells_idx = None
    num_positive_indices = {}  # Will store {classification: index}
    prop_indices = {}  # Will store existing proportion rows
    
    for i, identifier in enumerate(row_identifiers):
        identifier_str = str(identifier).strip()
        
        # Find total cell count row
        if identifier_str in ['Num Cells', 'Total Cells', 'Num_Cells']:
            num_cells_idx = i
            
        # Find positive cell count rows (Num 1+, Num 2+, etc.)
        elif identifier_str.startswith('Num ') and ('+' in identifier_str or 'Pos' in identifier_str):
            classification = identifier_str.replace('Num ', '').strip()
            num_positive_indices[classification] = i
            
        # Track existing proportion rows
        elif identifier_str.startswith('Prop_') or 'Prop' in identifier_str:
            prop_indices[identifier_str] = i
    
    print(f"      📊 Found structure:")
    print(f"        - Total cells row: {'Yes' if num_cells_idx is not None else 'No'}")
    print(f"        - Positive count rows: {len(num_positive_indices)}")
    print(f"        - Existing proportion rows: {len(prop_indices)}")
    
    # Calculate new proportions if we have both total and positive counts
    if num_cells_idx is not None and num_positive_indices:
        print(f"      📈 Calculating proportions for {len(num_positive_indices)} categories")
        
        # Create new proportion rows
        new_proportion_data = []
        
        for classification, pos_idx in num_positive_indices.items():
            # Create proportion row identifier
            prop_identifier = f"Prop_{classification}_Calculated"
            
            # Calculate proportions for this classification
            proportion_row = {row_col: prop_identifier}
            
            for condition in numeric_columns:
                pos_value = pd.to_numeric(df.iloc[pos_idx][condition], errors='coerce')
                total_value = pd.to_numeric(df.iloc[num_cells_idx][condition], errors='coerce')
                
                if pd.isna(pos_value) or pd.isna(total_value) or total_value == 0:
                    proportion = 0.0
                else:
                    proportion = (pos_value / total_value) * 100
                
                proportion_row[condition] = round(proportion, 4)
            
            new_proportion_data.append(proportion_row)
            print(f"        ✓ {classification}: calculated proportions")
        
        # Add new proportion rows to the DataFrame
        if new_proportion_data:
            # Create sections: Raw counts, Original proportions (if any), Calculated proportions
            result_data = []
            
            # 1. Add all original raw count rows
            count_rows = []
            original_prop_rows = []
            
            for i, identifier in enumerate(row_identifiers):
                row_data = df.iloc[i].to_dict()
                
                if any(prop_key in str(identifier) for prop_key in ['Prop_', 'Prop ']):
                    original_prop_rows.append(row_data)
                else:
                    count_rows.append(row_data)
            
            # 2. Combine in structured order
            result_data.extend(count_rows)  # Raw counts first
            
            if original_prop_rows:
                # Add separator row for original proportions
                separator_orig = {row_col: "--- Original Proportions ---"}
                for col in numeric_columns:
                    separator_orig[col] = ""
                result_data.append(separator_orig)
                result_data.extend(original_prop_rows)
            
            # Add separator row for calculated proportions
            separator_calc = {row_col: "--- Calculated Proportions ---"}
            for col in numeric_columns:
                separator_calc[col] = ""
            result_data.append(separator_calc)
            
            # 3. Add calculated proportions
            result_data.extend(new_proportion_data)
            
            result_df = pd.DataFrame(result_data)
            print(f"      ✅ Restructured data: {len(count_rows)} count rows + {len(original_prop_rows)} original props + {len(new_proportion_data)} calculated props")
            return result_df
    
    # If we can't calculate proportions, try to reorganize existing data
    elif prop_indices:
        print(f"      🔄 Reorganizing existing data structure")
        
        # Separate count rows from proportion rows
        count_rows = []
        prop_rows = []
        
        for i, identifier in enumerate(row_identifiers):
            row_data = df.iloc[i].to_dict()
            
            if any(prop_key in str(identifier) for prop_key in ['Prop_', 'Prop ']):
                prop_rows.append(row_data)
            else:
                count_rows.append(row_data)
        
        # Restructure with clear sections
        result_data = count_rows.copy()
        
        if prop_rows:
            # Add separator
            separator = {row_col: "--- Proportions (%) ---"}
            for col in numeric_columns:
                separator[col] = ""
            result_data.append(separator)
            result_data.extend(prop_rows)
        
        result_df = pd.DataFrame(result_data)
        print(f"      ✅ Reorganized: {len(count_rows)} count rows + {len(prop_rows)} proportion rows")
        return result_df
    
    print(f"      ⚠️  Could not restructure data - insufficient information for proportion calculation")
    return df
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
    Create a summary sheet from all processed sheets, including proportions.
    
    Args:
        processed_sheets (dict): Dictionary of processed DataFrames
        
    Returns:
        DataFrame: Summary sheet with raw data and proportions
    """
    if not processed_sheets:
        return None

    print(f"\n📋 Creating Summary sheet...")
    
    # Get all unique row identifiers across all sheets
    all_row_identifiers = set()
    all_conditions = set()
    row_col = None
    
    for sheet_df in processed_sheets.values():
        # Get row identifier column name
        if row_col is None:
            row_col = 'Count' if 'Count' in sheet_df.columns else 'Classification'
        
        # Collect all row identifiers
        for identifier in sheet_df[row_col]:
            all_row_identifiers.add(identifier)
        
        # Collect all condition columns
        for col in sheet_df.columns:
            if col not in ['Count', 'Classification']:
                all_conditions.add(col)
    
    all_row_identifiers = sorted(list(all_row_identifiers))
    all_conditions = sorted(list(all_conditions))
    
    # Create summary data
    summary_data = {row_col: all_row_identifiers}
    
    for condition in all_conditions:
        condition_values = []
        
        for row_identifier in all_row_identifiers:
            values = []
            
            # Collect values for this row identifier and condition across all sheets
            for sheet_df in processed_sheets.values():
                # Find the row with this identifier
                matching_rows = sheet_df[sheet_df[row_col] == row_identifier]
                
                if not matching_rows.empty and condition in sheet_df.columns:
                    try:
                        val = pd.to_numeric(matching_rows.iloc[0][condition], errors='coerce')
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
    
    # Count raw data rows vs proportion rows
    proportion_rows = len([row for row in all_row_identifiers if 'Proportion_%' in str(row)])
    raw_rows = len(all_row_identifiers) - proportion_rows
    
    print(f"✅ Summary sheet: {raw_rows} raw data rows, {proportion_rows} proportion rows × {len(all_conditions)} conditions")
    
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
    """Display a preview of the restructured summary results."""
    output_name = f"combined_{input_filename}.xlsx" if input_filename else "combined_donor.xlsx"
    
    print(f"\n{'='*60}")
    print(f"PREVIEW OF {output_name.upper()}")
    print(f"{'='*60}")
    
    if summary_df is not None and not summary_df.empty:
        print("\n📊 SUMMARY SHEET (RESTRUCTURED):")
        
        # Get the row identifier column
        row_col = 'Count' if 'Count' in summary_df.columns else 'Classification'
        
        # Separate different sections
        count_rows = []
        original_prop_rows = []
        calculated_prop_rows = []
        current_section = "counts"
        
        for idx, row in summary_df.iterrows():
            identifier = str(row[row_col])
            
            if "--- Original Proportions ---" in identifier:
                current_section = "original_props"
                continue
            elif "--- Calculated Proportions ---" in identifier:
                current_section = "calculated_props"
                continue
            elif identifier.startswith("---"):
                continue
            
            if current_section == "counts":
                count_rows.append(row)
            elif current_section == "original_props":
                original_prop_rows.append(row)
            elif current_section == "calculated_props":
                calculated_prop_rows.append(row)
        
        # Display sections
        if count_rows:
            print("\n📈 CELL COUNTS:")
            count_df = pd.DataFrame(count_rows)
            print(count_df.to_string(index=False))
        
        if original_prop_rows:
            print("\n📊 ORIGINAL PROPORTIONS (%):")
            orig_df = pd.DataFrame(original_prop_rows)
            # Format proportions for better readability
            numeric_cols = [col for col in orig_df.columns if col != row_col]
            for col in numeric_cols:
                orig_df[col] = pd.to_numeric(orig_df[col], errors='coerce').round(2)
            print(orig_df.to_string(index=False))
        
        if calculated_prop_rows:
            print("\n🧮 CALCULATED PROPORTIONS (%):")
            calc_df = pd.DataFrame(calculated_prop_rows)
            # Format calculated proportions
            numeric_cols = [col for col in calc_df.columns if col != row_col]
            for col in numeric_cols:
                calc_df[col] = pd.to_numeric(calc_df[col], errors='coerce').round(2)
            print(calc_df.to_string(index=False))
        
        # Show data structure summary
        print(f"\n📋 DATA STRUCTURE SUMMARY:")
        print(f"   • Cell count rows: {len(count_rows)}")
        print(f"   • Original proportion rows: {len(original_prop_rows)}")
        print(f"   • Calculated proportion rows: {len(calculated_prop_rows)}")
        
    else:
        print("⚠️  No summary data to display")


def process_directory(directory_path):
    """
    Process all Excel files in a directory.
    
    Args:
        directory_path (str): Path to directory containing Excel files
        
    Returns:
        list: List of paths to generated combined files
    """
    print(f"\n📂 PROCESSING DIRECTORY: {os.path.basename(directory_path)}")
    print("=" * 60)
    
    # Find all Excel files in the directory
    excel_files = []
    for ext in ['*.xlsx', '*.xls']:
        excel_files.extend(glob.glob(os.path.join(directory_path, ext)))
    
    # Filter out already processed files (combined_* and total_*)
    original_files = [f for f in excel_files if not (
        os.path.basename(f).startswith('combined_') or 
        os.path.basename(f).startswith('total_')
    )]
    
    if not original_files:
        print("❌ No original Excel files found in directory")
        return []
    
    print(f"📋 Found {len(original_files)} files to process:")
    for f in original_files:
        print(f"  • {os.path.basename(f)}")
    
    combined_files = []
    
    # Process each file
    for i, file_path in enumerate(original_files, 1):
        print(f"\n{'='*40}")
        print(f"� PROCESSING FILE {i}/{len(original_files)}")
        print(f"{'='*40}")
        
        # Process the file
        processed_sheets = process_combined_summary(file_path)
        
        if not processed_sheets:
            print(f"❌ Failed to process {os.path.basename(file_path)}")
            continue
        
        # Create summary sheet
        summary_df = create_summary_sheet(processed_sheets)
        
        # Save combined file
        output_path = save_combined_donor(summary_df, processed_sheets, file_path)
        
        if output_path:
            combined_files.append(output_path)
            print(f"✅ Created: {os.path.basename(output_path)}")
        else:
            print(f"❌ Failed to save combined file for {os.path.basename(file_path)}")
    
    print(f"\n{'='*60}")
    print(f"📊 DIRECTORY PROCESSING COMPLETE")
    print(f"{'='*60}")
    print(f"✅ Successfully processed: {len(combined_files)}/{len(original_files)} files")
    
    return combined_files


def consolidate_combined_files(combined_files, output_directory):
    """
    Consolidate all combined files into one final patient file with summed totals.
    Structure matches: Control1, Control2, Control3... where each column is sum of all sheets in that file.
    
    Args:
        combined_files (list): List of paths to combined files
        output_directory (str): Directory to save the final consolidated file
        
    Returns:
        str: Path to the final consolidated file
    """
    if not combined_files:
        print("❌ No combined files to consolidate")
        return None
    
    print(f"\n📋 CONSOLIDATING {len(combined_files)} COMBINED FILES")
    print("=" * 60)
    
    # Define the fixed row structure we want
    target_rows = [
        "Num Cells",
        "Num Positive", 
        "Num 1+",
        "Num 2+", 
        "Num 3+",
        "Num Negative",
        "Prop_1+",
        "Prop_2+", 
        "Prop_3+",
        "Prop_Neg"
    ]
    
    # Process each combined file to get summed totals
    file_data = {}
    
    for file_path in combined_files:
        try:
            xl_file = pd.ExcelFile(file_path)
            file_basename = os.path.splitext(os.path.basename(file_path))[0]
            # Extract identifier from filename (remove "combined_" prefix)
            file_identifier = file_basename.replace('combined_', '')
            
            print(f"  📄 Processing {file_identifier}...")
            
            # Initialize sums for this file
            file_totals = {row: 0.0 for row in target_rows}
            sheets_processed = 0
            
            # Sum across all sheets (except Summary)
            for sheet_name in xl_file.sheet_names:
                if sheet_name.lower() == 'summary':
                    continue
                    
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    
                    if df.empty:
                        continue
                    
                    # Get row identifier column
                    row_col = 'Count' if 'Count' in df.columns else 'Classification'
                    
                    # Get condition columns (all non-row identifier columns)
                    condition_columns = [col for col in df.columns if col != row_col]
                    
                    # Sum values across all conditions in this sheet
                    for _, row in df.iterrows():
                        row_identifier = str(row[row_col]).strip()
                        
                        # Match row identifier to our target rows
                        matched_target = None
                        for target_row in target_rows:
                            if (row_identifier == target_row or 
                                row_identifier.replace(' ', '_') == target_row or
                                row_identifier.replace('_', ' ') == target_row):
                                matched_target = target_row
                                break
                        
                        if matched_target:
                            # Sum all condition values for this row
                            row_sum = 0.0
                            for condition_col in condition_columns:
                                try:
                                    val = pd.to_numeric(row[condition_col], errors='coerce')
                                    if not pd.isna(val):
                                        row_sum += val
                                except:
                                    pass
                            
                            file_totals[matched_target] += row_sum
                    
                    sheets_processed += 1
                    
                except Exception as e:
                    print(f"    ⚠️  Error processing sheet {sheet_name}: {e}")
            
            # Calculate proportions for this file
            total_cells = file_totals["Num Cells"]
            if total_cells > 0:
                file_totals["Prop_1+"] = (file_totals["Num 1+"] / total_cells) * 100
                file_totals["Prop_2+"] = (file_totals["Num 2+"] / total_cells) * 100  
                file_totals["Prop_3+"] = (file_totals["Num 3+"] / total_cells) * 100
                file_totals["Prop_Neg"] = (file_totals["Num Negative"] / total_cells) * 100
            
            file_data[file_identifier] = file_totals
            print(f"    ✅ {file_identifier}: {sheets_processed} sheets summed")
            
        except Exception as e:
            print(f"    ❌ Error processing {os.path.basename(file_path)}: {e}")
    
    if not file_data:
        print("❌ No valid data collected from combined files")
        return None
    
    # Create the consolidated DataFrame in the exact structure requested
    consolidated_data = {"Count": target_rows}
    
    # Sort file identifiers for consistent column ordering
    sorted_files = sorted(file_data.keys())
    
    for file_id in sorted_files:
        column_name = f"Control{len(consolidated_data)}"  # Control1, Control2, etc.
        column_values = []
        
        for row_name in target_rows:
            value = file_data[file_id][row_name]
            
            # Format proportions to 2 decimal places, others as integers for counts
            if row_name.startswith("Prop_"):
                column_values.append(round(value, 2))
            else:
                column_values.append(int(value))
        
        consolidated_data[column_name] = column_values
        print(f"    📊 {column_name}: {file_id} data added")
    
    # Create DataFrame
    consolidated_df = pd.DataFrame(consolidated_data)
    
    # Generate output filename
    file_identifiers = sorted_files
    if len(file_identifiers) <= 3:
        filename_part = "-".join(file_identifiers)
    else:
        filename_part = f"{file_identifiers[0]}-to-{file_identifiers[-1]}"
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_filename = f"total_patient_{filename_part}_{timestamp}.xlsx"
    output_path = os.path.join(output_directory, output_filename)
    
    # Save consolidated file
    try:
        print(f"\n💾 Saving consolidated file: {output_filename}")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            consolidated_df.to_excel(writer, sheet_name='Summary', index=False)
            print(f"  📄 Summary: {len(consolidated_df)} rows × {len(consolidated_df.columns)} columns")
        
        print(f"✅ Consolidated file saved: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"❌ Error saving consolidated file: {e}")
        return None
    if not combined_files:
        print("❌ No combined files to consolidate")
        return None
    
    print(f"\n� CONSOLIDATING {len(combined_files)} COMBINED FILES")

def display_consolidation_preview(consolidated_path):
    """Display a preview of the consolidated file with summed structure."""
    if not consolidated_path or not os.path.exists(consolidated_path):
        print("⚠️  No consolidated file to preview")
        return
    
    try:
        xl_file = pd.ExcelFile(consolidated_path)
        filename = os.path.basename(consolidated_path)
        
        print(f"\n{'='*60}")
        print(f"PREVIEW OF {filename.upper()}")
        print(f"{'='*60}")
        
        # Read the summary sheet
        df = pd.read_excel(consolidated_path, sheet_name='Summary')
        
        if not df.empty:
            print(f"\n📊 SUMMARY SHEET - PATIENT TOTALS:")
            print(f"   📋 Structure: {len(df)} rows × {len(df.columns)} columns")
            
            # Get column info
            count_col = df.columns[0]
            control_cols = [col for col in df.columns if col.startswith('Control')]
            
            print(f"   � Files consolidated: {len(control_cols)}")
            print(f"   � Data categories: {len(df)}")
            
            # Show the data structure matching table1.txt format
            print("\n📋 DATA STRUCTURE (matching table1.txt format):")
            print(df.to_string(index=False))
            
            # Show summary statistics
            print(f"\n📊 SUMMARY STATISTICS:")
            for col in control_cols:
                total_cells = df[df[count_col] == 'Num Cells'][col].iloc[0] if not df[df[count_col] == 'Num Cells'].empty else 0
                print(f"   • {col}: {total_cells} total cells")
        else:
            print("   ⚠️  Empty sheet")
                
    except Exception as e:
        print(f"❌ Error previewing consolidated file: {e}")


def main():
    """Main function for directory processing."""
    print("=" * 60)
    print("DIRECTORY → PATIENT CONSOLIDATION PROCESSOR")
    print("=" * 60)
    print("Step 1: Process all Excel files in directory")
    print("        (Control1, Control2 → Control averaged)")
    print("Step 2: Consolidate all combined files")
    print("        (Control from each file → Control_file1, Control_file2)")
    
    # Select input directory
    print("\n📂 Select directory containing Excel files...")
    
    root = tk.Tk()
    root.withdraw()
    
    directory_path = filedialog.askdirectory(
        title="Select directory containing Excel files"
    )
    
    root.destroy()
    
    if not directory_path:
        print("❌ No directory selected. Exiting...")
        return
    
    print(f"📁 Directory: {os.path.basename(directory_path)}")
    
    # Step 1: Process all files in directory
    combined_files = process_directory(directory_path)
    
    if not combined_files:
        print("❌ No files were successfully processed. Exiting...")
        return
    
    # Step 2: Consolidate all combined files
    consolidated_path = consolidate_combined_files(combined_files, directory_path)
    
    if consolidated_path:
        # Display preview
        display_consolidation_preview(consolidated_path)
        
        # Final summary
        print(f"\n{'='*60}")
        print("🎉 PROCESSING COMPLETE!")
        print(f"{'='*60}")
        print(f"📁 Input directory: {os.path.basename(directory_path)}")
        print(f"� Files processed: {len(combined_files)}")
        print(f"� Combined files created: {len(combined_files)}")
        print(f"� Final consolidated file: {os.path.basename(consolidated_path)}")
        print(f"💾 Location: {directory_path}")
        print("\n✨ Ready for patient-level analysis!")
    else:
        print("❌ Failed to create consolidated file.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n❌ Operation cancelled by user.")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
    
    input("\nPress Enter to exit...")
