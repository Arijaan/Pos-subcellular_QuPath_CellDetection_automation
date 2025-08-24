#!/usr/bin/env python3
"""
CSV File Organizer by Patient Code and Condition

This script organizes CSV measurement files by extracting:
1. 4-digit patient codes (e.g., 6249)
2. Condition names (e.g., Control, T1D, T2D)

Creates two organizational structures:
1. Individual patient directories: XXXX-Condition/
2. Condition-based directories: By_Condition/Condition_Name/

Author: Generated for QuPath measurement analysis
"""

import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from collections import defaultdict


def extract_patient_info(filename):
    """
    Extract patient code (4 digits) and condition from filename.
    
    Args:
        filename (str): Name of the CSV file
        
    Returns:
        tuple: (patient_code, condition) or (None, None) if not found
    """
    # Define possible conditions to look for
    conditions = ['Control', 'T1D', 'T2D', 'Aab', 'control', 't1d', 't2d', 'aab']
    
    # Find 4-digit patient code
    patient_match = re.search(r'\b(\d{4})\b', filename)
    patient_code = patient_match.group(1) if patient_match else None
    
    # Find condition (case-insensitive)
    condition = None
    filename_lower = filename.lower()
    
    for cond in conditions:
        if cond.lower() in filename_lower:
            # Standardize condition names
            if cond.lower() in ['control']:
                condition = 'Control'
            elif cond.lower() in ['t1d']:
                condition = 'T1D'
            elif cond.lower() in ['t2d']:
                condition = 'T2D'
            elif cond.lower() in ['aab']:
                condition = 'Aab'
            break
    
    return patient_code, condition


def create_directory_structure(base_dir, patients_data):
    """
    Create organized directory structure.
    
    Args:
        base_dir (str): Base directory containing CSV files
        patients_data (dict): Dictionary with patient information
        
    Returns:
        tuple: (individual_dirs, condition_dirs) - paths created
    """
    individual_dirs = {}
    condition_dirs = {}
    
    # Create individual patient directories
    for (patient_code, condition), files in patients_data.items():
        if patient_code and condition:
            # Create individual patient directory
            patient_dir_name = f"{patient_code}-{condition}"
            patient_dir_path = os.path.join(base_dir, patient_dir_name)
            
            if not os.path.exists(patient_dir_path):
                os.makedirs(patient_dir_path)
                print(f"✅ Created directory: {patient_dir_name}")
            
            individual_dirs[(patient_code, condition)] = patient_dir_path
            
            # Create condition-based directory structure
            condition_base = os.path.join(base_dir, "By_Condition")
            condition_dir_path = os.path.join(condition_base, condition)
            
            if not os.path.exists(condition_dir_path):
                os.makedirs(condition_dir_path)
                print(f"✅ Created condition directory: By_Condition/{condition}")
            
            if condition not in condition_dirs:
                condition_dirs[condition] = condition_dir_path
    
    return individual_dirs, condition_dirs


def organize_files(base_dir, patients_data, individual_dirs, condition_dirs, copy_mode=True):
    """
    Organize CSV files into the created directory structure.
    
    Args:
        base_dir (str): Base directory containing CSV files
        patients_data (dict): Dictionary with patient information
        individual_dirs (dict): Individual patient directories
        condition_dirs (dict): Condition-based directories
        copy_mode (bool): If True, copy files; if False, move files
    """
    operation = "Copying" if copy_mode else "Moving"
    func = shutil.copy2 if copy_mode else shutil.move
    
    print(f"\n{operation} files to organized directories...")
    
    for (patient_code, condition), files in patients_data.items():
        if patient_code and condition:
            # Get destination directories
            individual_dir = individual_dirs.get((patient_code, condition))
            condition_dir = condition_dirs.get(condition)
            
            for filename in files:
                source_path = os.path.join(base_dir, filename)
                
                if os.path.exists(source_path):
                    # Copy/Move to individual patient directory
                    if individual_dir:
                        dest_individual = os.path.join(individual_dir, filename)
                        try:
                            func(source_path, dest_individual)
                            print(f"  📁 {patient_code}-{condition}: {filename}")
                        except Exception as e:
                            print(f"  ❌ Error with {filename}: {e}")
                    
                    # Copy to condition directory (always copy for condition dirs)
                    if condition_dir and copy_mode:
                        dest_condition = os.path.join(condition_dir, filename)
                        try:
                            shutil.copy2(source_path, dest_condition)
                        except Exception as e:
                            print(f"  ❌ Error copying to condition dir {filename}: {e}")


def analyze_directory(directory_path):
    """
    Analyze directory and extract patient information from CSV files.
    
    Args:
        directory_path (str): Path to directory containing CSV files
        
    Returns:
        dict: Dictionary organized by (patient_code, condition)
    """
    csv_files = [f for f in os.listdir(directory_path) 
                 if f.lower().endswith('.csv')]
    
    if not csv_files:
        print("❌ No CSV files found in the selected directory!")
        return {}
    
    print(f"📊 Found {len(csv_files)} CSV files")
    
    # Group files by patient and condition
    patients_data = defaultdict(list)
    unmatched_files = []
    
    for filename in csv_files:
        patient_code, condition = extract_patient_info(filename)
        
        if patient_code and condition:
            patients_data[(patient_code, condition)].append(filename)
            print(f"  ✅ {filename} → Patient: {patient_code}, Condition: {condition}")
        else:
            unmatched_files.append(filename)
            print(f"  ⚠️  {filename} → Could not extract patient code or condition")
    
    # Show summary
    print(f"\n📈 ANALYSIS SUMMARY:")
    print(f"✅ Matched files: {sum(len(files) for files in patients_data.values())}")
    print(f"⚠️  Unmatched files: {len(unmatched_files)}")
    
    if patients_data:
        print(f"\n👥 PATIENTS FOUND:")
        for (patient_code, condition), files in sorted(patients_data.items()):
            print(f"  📋 {patient_code}-{condition}: {len(files)} files")
    
    if unmatched_files:
        print(f"\n❓ UNMATCHED FILES:")
        for filename in unmatched_files[:10]:  # Show first 10
            print(f"  📄 {filename}")
        if len(unmatched_files) > 10:
            print(f"  ... and {len(unmatched_files) - 10} more")
    
    return dict(patients_data)


def get_user_choice():
    """Get user preference for copy vs move operation."""
    root = tk.Tk()
    root.withdraw()  # Hide main window
    
    result = messagebox.askyesnocancel(
        "File Organization Mode",
        "Choose how to organize the files:\n\n"
        "YES = Copy files (keeps originals)\n"
        "NO = Move files (removes originals)\n"
        "CANCEL = Exit without organizing"
    )
    
    root.destroy()
    return result


def main():
    """Main function to run the CSV file organizer."""
    print("=" * 70)
    print("CSV FILE ORGANIZER BY PATIENT CODE AND CONDITION")
    print("=" * 70)
    
    # Select directory containing CSV files
    print("📂 Select directory containing CSV measurement files...")
    
    root = tk.Tk()
    root.withdraw()
    
    directory_path = filedialog.askdirectory(
        title="Select directory containing CSV measurement files"
    )
    
    root.destroy()
    
    if not directory_path:
        print("❌ No directory selected. Exiting...")
        return
    
    print(f"📁 Selected directory: {directory_path}")
    
    # Analyze the directory
    patients_data = analyze_directory(directory_path)
    
    if not patients_data:
        print("❌ No valid patient files found. Exiting...")
        return
    
    # Get user choice for copy vs move
    choice = get_user_choice()
    
    if choice is None:  # Cancel
        print("❌ Operation cancelled by user.")
        return
    
    copy_mode = choice  # True for copy, False for move
    operation_type = "COPYING" if copy_mode else "MOVING"
    
    print(f"\n🔄 {operation_type} FILES...")
    
    # Create directory structure
    individual_dirs, condition_dirs = create_directory_structure(directory_path, patients_data)
    
    # Organize files
    organize_files(directory_path, patients_data, individual_dirs, condition_dirs, copy_mode)
    
    # Final summary
    print(f"\n{'='*70}")
    print("✅ ORGANIZATION COMPLETE!")
    print(f"{'='*70}")
    print(f"📁 Base directory: {directory_path}")
    print(f"📊 Operation: {operation_type} files")
    
    print(f"\n📂 CREATED STRUCTURE:")
    print(f"Individual Patient Directories:")
    for (patient_code, condition), path in individual_dirs.items():
        rel_path = os.path.relpath(path, directory_path)
        file_count = len(patients_data[(patient_code, condition)])
        print(f"  📋 {rel_path} ({file_count} files)")
    
    print(f"\nCondition-Based Directories:")
    for condition, path in condition_dirs.items():
        rel_path = os.path.relpath(path, directory_path)
        file_count = sum(len(files) for (pc, cond), files in patients_data.items() if cond == condition)
        print(f"  📂 {rel_path} ({file_count} files)")
    
    print(f"\n🎉 File organization completed successfully!")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n❌ Operation cancelled by user.")
    except Exception as e:
        print(f"\n❌ Error occurred: {e}")
        print("Please check file permissions and try again.")
    
    input("\nPress Enter to exit...")
