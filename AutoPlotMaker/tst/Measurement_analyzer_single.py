"""
QuPath Measurement Analyzer - Complete Workflow
This script combines CSV processing, Excel creation, and R plotting into one automated workflow.

Workflow:
1. Processes CSV files to create Combined_Summary.xlsx
2. Runs R boxplot analysis
3. Runs R barplot analysis
4. Organizes all outputs in the input directory

Author: AI-Assistant
Date: August 21, 2025
"""

import pandas as pd
import os
import numpy as np
import subprocess
import sys
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk
import shutil
from pathlib import Path

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
    def __init__(self, parent):
        print("DEBUG: Creating GroupInputDialog...")
        self.result = None
        
        try:
            self.root = tk.Toplevel(parent) if parent else tk.Tk()
            self.root.title("Group Selection")
            self.root.geometry("600x700")  # Increased height from 600 to 700
            self.root.resizable(True, True)
            
            # Make sure the window appears on top
            self.root.lift()
            self.root.attributes('-topmost', True)
            self.root.focus_force()
            
            # Add protocol handler for window close
            self.root.protocol("WM_DELETE_WINDOW", self.cancel_clicked)
            
            print("DEBUG: Window created, setting up UI...")
            self.setup_ui()
            
            # Center the window
            self.center_window()
            
            # Force window to be visible
            self.root.update()
            self.root.deiconify()
            
            print("DEBUG: UI setup complete, window should be visible")
            
        except Exception as e:
            print(f"DEBUG: Error creating dialog: {e}")
            raise

    def center_window(self):
        try:
            self.root.update_idletasks()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            x = (self.root.winfo_screenwidth() // 2) - (width // 2)
            y = (self.root.winfo_screenheight() // 2) - (height // 2)
            self.root.geometry(f'{width}x{height}+{x}+{y}')
        except Exception as e:
            print(f"DEBUG: Error centering window: {e}")

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Select Groups for Analysis", 
                               font=('Arial', 14, 'bold'))  # Increased font size
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                                text="Choose to use default groups or define custom groups for your analysis:",
                                wraplength=500, font=('Arial', 10))  # Increased wraplength
        instructions.grid(row=1, column=0, columnspan=2, pady=(0, 20))
        
        # Default groups option
        self.use_default = tk.BooleanVar(value=True)
        default_check = ttk.Checkbutton(main_frame, text="Use default groups", 
                                       variable=self.use_default,
                                       command=self.toggle_custom_groups)
        default_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 15))
        
        # Default groups display
        default_frame = ttk.LabelFrame(main_frame, text="Default Groups", padding="15")
        default_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 25))
        default_frame.columnconfigure(0, weight=1)
        
        default_text = "• control\n• t1d\n• t2d\n• aab"
        ttk.Label(default_frame, text=default_text, font=('Arial', 10)).grid(row=0, column=0, sticky=tk.W)
        
        # Custom groups section
        self.custom_frame = ttk.LabelFrame(main_frame, text="Custom Groups", padding="15")
        self.custom_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 25))
        self.custom_frame.columnconfigure(1, weight=1)
        
        # Group entries
        self.group_vars = []
        self.group_entries = []
        
        for i in range(6):  # Increased from 4 to 6 initial entries
            ttk.Label(self.custom_frame, text=f"Group {i+1}:").grid(row=i, column=0, sticky=tk.W, pady=3)
            var = tk.StringVar()
            entry = ttk.Entry(self.custom_frame, textvariable=var, width=35)
            entry.grid(row=i, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=3)
            self.group_vars.append(var)
            self.group_entries.append(entry)
        
        # Add more groups button
        ttk.Button(self.custom_frame, text="Add More Groups", 
                  command=self.add_group_entry).grid(row=6, column=0, columnspan=2, pady=(15, 0))
        
        # Initially disable custom groups
        self.toggle_custom_groups()
        
        # Buttons - larger and more spacing
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=(30, 0))
        
        ttk.Button(button_frame, text="OK", command=self.ok_clicked, width=15).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(button_frame, text="Cancel", command=self.cancel_clicked, width=15).pack(side=tk.LEFT)

    def toggle_custom_groups(self):
        state = tk.DISABLED if self.use_default.get() else tk.NORMAL
        for entry in self.group_entries:
            entry.configure(state=state)

    def add_group_entry(self):
        if self.use_default.get():
            return
            
        row = len(self.group_vars)
        ttk.Label(self.custom_frame, text=f"Group {row+1}:").grid(row=row, column=0, sticky=tk.W, pady=2)
        var = tk.StringVar()
        entry = ttk.Entry(self.custom_frame, textvariable=var, width=30)
        entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        self.group_vars.append(var)
        self.group_entries.append(entry)

    def ok_clicked(self):
        print("DEBUG: OK button clicked")
        try:
            if self.use_default.get():
                self.result = ['control', 't1d', 't2d', 'aab']
            else:
                groups = [var.get().strip() for var in self.group_vars if var.get().strip()]
                if not groups:
                    messagebox.showerror("Error", "Please enter at least one group name.")
                    return
                self.result = groups
            print(f"DEBUG: Result set to: {self.result}")
            
            # Try to destroy the window safely
            print("DEBUG: Attempting to destroy dialog window...")
            try:
                self.root.quit()  # Exit mainloop first
                print("DEBUG: Mainloop quit successful")
            except Exception as e:
                print(f"DEBUG: Error quitting mainloop: {e}")
            
            try:
                self.root.destroy()  # Then destroy window
                print("DEBUG: Window destroy successful")
            except Exception as e:
                print(f"DEBUG: Error destroying window: {e}")
                
        except Exception as e:
            print(f"DEBUG: Error in ok_clicked: {e}")
            import traceback
            traceback.print_exc()

    def cancel_clicked(self):
        print("DEBUG: Cancel button clicked")
        try:
            self.result = None
            print("DEBUG: Attempting to destroy dialog window...")
            try:
                self.root.quit()  # Exit mainloop first
                print("DEBUG: Mainloop quit successful")
            except Exception as e:
                print(f"DEBUG: Error quitting mainloop: {e}")
            
            try:
                self.root.destroy()  # Then destroy window
                print("DEBUG: Window destroy successful")
            except Exception as e:
                print(f"DEBUG: Error destroying window: {e}")
                
        except Exception as e:
            print(f"DEBUG: Error in cancel_clicked: {e}")
            import traceback
            traceback.print_exc()


def get_user_groups():
    """Get group names from user with improved dialog handling"""
    print("DEBUG: Starting get_user_groups function...")
    
    # Try multiple approaches to ensure we get user input
    for attempt in range(3):
        print(f"DEBUG: Attempt {attempt + 1} to show dialog...")
        
        try:
            # Create a temporary root window if none exists
            if not hasattr(get_user_groups, 'root') or not get_user_groups.root:
                print("DEBUG: Creating new root window...")
                get_user_groups.root = tk.Tk()
                get_user_groups.root.withdraw()  # Hide the root window
                print("DEBUG: Root window created and hidden")
            
            # Create and show the dialog
            print("DEBUG: Creating GroupInputDialog...")
            dialog = GroupInputDialog(get_user_groups.root)
            print("DEBUG: Dialog created successfully")
            
            # Make sure the dialog is visible with multiple attempts
            for visibility_attempt in range(3):
                try:
                    dialog.root.deiconify()
                    dialog.root.lift()
                    dialog.root.focus_force()
                    dialog.root.attributes('-topmost', True)
                    dialog.root.update()
                    print(f"DEBUG: Dialog visibility attempt {visibility_attempt + 1} completed")
                    break
                except Exception as ve:
                    print(f"DEBUG: Visibility attempt {visibility_attempt + 1} failed: {ve}")
                    if visibility_attempt == 2:
                        raise
            
            print("DEBUG: Starting dialog mainloop...")
            
            # Run the dialog with timeout protection
            dialog.root.mainloop()
            print("DEBUG: Dialog mainloop ended")
            
            result = dialog.result
            print(f"DEBUG: Dialog result: {result}")
            
            if result is not None:
                print(f"Selected groups: {result}")
                return result
            else:
                print("DEBUG: No result from dialog, trying again...")
                continue
                
        except Exception as e:
            print(f"DEBUG: Attempt {attempt + 1} failed with error: {e}")
            print(f"DEBUG: Error type: {type(e)}")
            import traceback
            traceback.print_exc()
            
            # If this is the last attempt, fall back to console
            if attempt == 2:
                break
            
            # Wait a bit before retrying
            import time
            time.sleep(0.5)
    
    # Fallback to console input if all GUI attempts failed
    print("\n" + "="*50)
    print("GUI dialog failed. Falling back to console input...")
    print("="*50)
    print("Available options:")
    print("1. Use default groups (control, t1d, t2d, aab)")
    print("2. Enter custom groups")
    
    while True:
        try:
            choice = input("Enter choice (1 or 2): ").strip()
            if choice == '1':
                return ['control', 't1d', 't2d', 'aab']
            elif choice == '2':
                groups = []
                print("Enter group names (press Enter with empty line to finish):")
                while True:
                    group = input(f"Group {len(groups)+1}: ").strip()
                    if not group:
                        break
                    groups.append(group)
                if groups:
                    return groups
                else:
                    print("No groups entered. Using default groups.")
                    return ['control', 't1d', 't2d', 'aab']
            else:
                print("Invalid choice. Please enter 1 or 2.")
        except (EOFError, KeyboardInterrupt):
            print("\nUsing default groups...")
            return ['control', 't1d', 't2d', 'aab']


def cleanup_tkinter():
    """Clean up tkinter root window"""
    if hasattr(get_user_groups, 'root') and get_user_groups.root:
        try:
            get_user_groups.root.destroy()
            get_user_groups.root = None
            print("DEBUG: Tkinter root window cleaned up")
        except:
            pass


def create_combined_summary(measurements_folder, groups):
    """
    Creates Combined_Summary.xlsx from CSV files
    """
    print("=" * 50)
    print("STEP 1: Creating Combined Summary")
    print("=" * 50)
    
    # Dictionary to store data for each group
    group_data = {group: {} for group in groups}
    
    # Create Measurement_summary folder path
    summary_folder = os.path.join(measurements_folder, "Measurement_summary")
    
    # Get all CSV files in the measurements folder
    csv_files = [f for f in os.listdir(measurements_folder) if f.endswith('.csv')]
    
    if not csv_files:
        print("No CSV files found in the measurements folder!")
        return None
        
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
            
            # Look for the SUMMARY column specifically
            summary_col_idx = None
            for i, col in enumerate(df.columns):
                if 'SUMMARY' in str(col):
                    summary_col_idx = i
                    break
            
            # Extract summary data from the CSV file
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
                
                except Exception as e:
                    print(f"  Error extracting from SUMMARY column: {e}")
            
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
        return None
    
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
        # Generate dynamic sheet names mapping
        sheet_name_mapping = {}
        for group in groups:
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
            return excel_path
        else:
            print("No sheets were created - no valid data found!")
            return None
    else:
        print("Excel libraries not available!")
        return None


def run_r_script(script_path, combined_summary_path, output_folder, analysis_name):
    """
    Run an R script with the combined summary file
    """
    try:
        # Find R executable
        r_executable = None
        possible_r_paths = [
            "Rscript",
            "C:\\Program Files\\R\\R-4.3.0\\bin\\Rscript.exe",
            "C:\\Program Files\\R\\R-4.2.0\\bin\\Rscript.exe",
            "C:\\Program Files\\R\\R-4.1.0\\bin\\Rscript.exe",
            "C:\\Program Files\\R\\R-4.0.0\\bin\\Rscript.exe"
        ]
        
        for r_path in possible_r_paths:
            try:
                result = subprocess.run([r_path, "--version"], capture_output=True, text=True, timeout=10)
                if result.returncode == 0:
                    r_executable = r_path
                    break
            except:
                continue
        
        if r_executable is None:
            print(f"Warning: R not found. Skipping {analysis_name} analysis.")
            return False
        
        print(f"\nRunning {analysis_name} analysis...")
        print(f"R executable: {r_executable}")
        print(f"Script: {script_path}")
        print(f"Input file: {combined_summary_path}")
        print(f"Output folder: {output_folder}")
        
        # Create R script command with arguments
        cmd = [
            r_executable,
            script_path,
            combined_summary_path.replace('\\', '/')  # Pass file path as argument, normalize path
        ]
        
        # Run the R script
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        
        if result.returncode == 0:
            print(f"{analysis_name} analysis completed successfully!")
            return True
        else:
            print(f"Error running {analysis_name} analysis:")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            return False
            
    except Exception as e:
        print(f"Error running {analysis_name} analysis: {str(e)}")
        return False


def process_all_summary_files(directory_path, groups):
    """
    Process all Combined_Summary.xlsx files in a directory and its subdirectories
    
    Args:
        directory_path (str): Path to the directory to search
        groups (list): List of group names for analysis
    
    Returns:
        list: List of tuples (file_path, success_status, output_dir)
    """
    results = []
    
    # Find all Combined_Summary.xlsx files
    summary_files = []
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith('.xlsx') and ('summary' in file.lower() or 'combined' in file.lower()):
                full_path = os.path.join(root, file)
                summary_files.append(full_path)
    
    if not summary_files:
        print(f"No summary Excel files found in {directory_path}")
        return results
    
    print(f"Found {len(summary_files)} summary file(s) to process:")
    for i, file_path in enumerate(summary_files, 1):
        rel_path = os.path.relpath(file_path, directory_path)
        print(f"  {i}. {rel_path}")
    
    # Set up R script paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    boxplot_script = os.path.join(script_dir, "T-AutoBoxPlotMaker-py.R")
    barplot_script = os.path.join(script_dir, "T-BarPlotMaker-py.R")
    
    # Check if R scripts exist
    scripts_to_run = []
    if os.path.exists(boxplot_script):
        scripts_to_run.append(("BoxPlot", boxplot_script))
        print(f"✅ Found BoxPlot script: {boxplot_script}")
    else:
        print(f"⚠️  Warning: BoxPlot script not found at {boxplot_script}")
    
    if os.path.exists(barplot_script):
        scripts_to_run.append(("BarPlot", barplot_script))
        print(f"✅ Found BarPlot script: {barplot_script}")
    else:
        print(f"⚠️  Warning: BarPlot script not found at {barplot_script}")
    
    if not scripts_to_run:
        print("No R scripts found. Cannot proceed with analysis.")
        return results
    
    # Process each summary file
    for file_path in summary_files:
        print(f"\n{'='*80}")
        print(f"PROCESSING: {os.path.basename(file_path)}")
        print(f"Location: {os.path.dirname(file_path)}")
        print(f"{'='*80}")
        
        # Create output directory in the same location as the summary file
        file_dir = os.path.dirname(file_path)
        output_subdir = os.path.join(file_dir, "AutoPlot_Analysis")
        
        if not os.path.exists(output_subdir):
            os.makedirs(output_subdir)
            print(f"📁 Created output directory: {output_subdir}")
        else:
            print(f"📁 Using existing output directory: {output_subdir}")
        
        file_success = True
        
        # Run each R analysis script
        for analysis_name, script_path in scripts_to_run:
            print(f"\n{'-'*50}")
            print(f"Running {analysis_name} Analysis")
            print(f"{'-'*50}")
            
            success = run_r_script(script_path, file_path, file_dir, analysis_name)
            
            if success:
                print(f"✅ {analysis_name} analysis completed!")
            else:
                print(f"❌ {analysis_name} analysis failed!")
                file_success = False
        
        # Move analysis results to the output subdirectory
        try:
            move_analysis_results_to_subdir(file_dir, output_subdir)
            print(f"📦 Moved analysis results to: {output_subdir}")
        except Exception as e:
            print(f"⚠️  Warning: Could not move some results to subdirectory: {e}")
        
        results.append((file_path, file_success, output_subdir))
        
        if file_success:
            print(f"✅ Successfully processed: {os.path.basename(file_path)}")
        else:
            print(f"❌ Some analyses failed for: {os.path.basename(file_path)}")
    
    return results


def move_analysis_results_to_subdir(source_dir, target_dir):
    """
    Move analysis result directories to a subdirectory
    
    Args:
        source_dir (str): Directory containing analysis results
        target_dir (str): Target subdirectory for organizing results
    """
    analysis_dirs = ["BoxPlot_Analysis", "Weighted_BarPlot_Analysis", "BarPlot_Analysis"]
    
    for analysis_dir in analysis_dirs:
        source_path = os.path.join(source_dir, analysis_dir)
        if os.path.exists(source_path):
            target_path = os.path.join(target_dir, analysis_dir)
            
            # If target exists, remove it first
            if os.path.exists(target_path):
                shutil.rmtree(target_path)
            
            # Move the directory
            shutil.move(source_path, target_path)
            print(f"   Moved {analysis_dir} to subdirectory")


def main():
    """
    Main workflow function - Single file processing
    """
    print("QuPath Measurement Analyzer - Single File Workflow")
    print("=" * 60)
    
    # === STEP 1: Get user input ===
    print("Getting group definitions from user...")
    groups = get_user_groups()
    print("DEBUG: get_user_groups() returned successfully")
    
    if not groups:
        print("No groups defined. Exiting...")
        cleanup_tkinter()
        return
        
    print(f"Using groups: {groups}")
    
    # === STEP 2: Choose processing mode ===
    print("\nChoose processing mode:")
    print("1. Process CSV files to create Combined_Summary.xlsx and analyze")
    print("2. Process an existing summary Excel file")
    
    try:
        mode = input("Enter choice (1 or 2): ").strip()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        cleanup_tkinter()
        return
    
    if mode == "1":
        # Original workflow: Process CSV files
        print("DEBUG: About to show folder selection dialog...")
        
        # Select measurements folder
        print("\nSelect folder containing CSV measurement files...")
        if hasattr(get_user_groups, 'root') and get_user_groups.root:
            print("DEBUG: Using existing root window for file dialog")
            measurements_folder = filedialog.askdirectory(
                parent=get_user_groups.root,
                title="Select folder containing CSV measurement files"
            )
        else:
            print("DEBUG: Creating new root window for file dialog")
            root = tk.Tk()
            root.withdraw()
            measurements_folder = filedialog.askdirectory(title="Select folder containing CSV measurement files")
            root.destroy()
        
        print("DEBUG: File dialog completed")
        
        if not measurements_folder:
            print("No folder selected. Exiting...")
            cleanup_tkinter()
            return

        print(f"Measurements folder: {measurements_folder}")
        print("DEBUG: Starting create_combined_summary...")
        
        # === STEP 3: Create Combined Summary ===
        combined_summary_path = create_combined_summary(measurements_folder, groups)
        
        if not combined_summary_path:
            print("Failed to create Combined Summary. Exiting...")
            cleanup_tkinter()
            return
        
        # Use the directory containing the CSV files as output directory
        output_base = measurements_folder
        excel_file_path = combined_summary_path
        
    elif mode == "2":
        # New workflow: Process existing summary file
        print("\nSelect an existing summary Excel file...")
        if hasattr(get_user_groups, 'root') and get_user_groups.root:
            excel_file_path = filedialog.askopenfilename(
                parent=get_user_groups.root,
                title="Select summary Excel file",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
        else:
            root = tk.Tk()
            root.withdraw()
            excel_file_path = filedialog.askopenfilename(
                title="Select summary Excel file",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            root.destroy()
        
        if not excel_file_path:
            print("No file selected. Exiting...")
            cleanup_tkinter()
            return
        
        print(f"Processing Excel file: {excel_file_path}")
        
        # Use the directory containing the Excel file as output directory
        output_base = os.path.dirname(excel_file_path)
        
    else:
        print("Invalid choice. Please run the script again and choose 1 or 2.")
        cleanup_tkinter()
        return
    
    # === STEP 4: Set up R script paths ===
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    boxplot_script = os.path.join(script_dir, "T-AutoBoxPlotMaker-py.R")
    barplot_script = os.path.join(script_dir, "T-BarPlotMaker-py.R")
    
    # Check if R scripts exist
    scripts_to_run = []
    if os.path.exists(boxplot_script):
        scripts_to_run.append(("BoxPlot", boxplot_script))
        print(f"✅ Found BoxPlot script: {boxplot_script}")
    else:
        print(f"⚠️  Warning: BoxPlot script not found at {boxplot_script}")
    
    if os.path.exists(barplot_script):
        scripts_to_run.append(("BarPlot", barplot_script))
        print(f"✅ Found BarPlot script: {barplot_script}")
    else:
        print(f"⚠️  Warning: BarPlot script not found at {barplot_script}")
    
    if not scripts_to_run:
        print("No R scripts found. Analysis complete with input file only.")
        cleanup_tkinter()
        return
    
    # === STEP 5: Run R analyses ===
    for analysis_name, script_path in scripts_to_run:
        print("\n" + "=" * 50)
        print(f"STEP: Running {analysis_name} Analysis")
        print("=" * 50)
        
        success = run_r_script(script_path, excel_file_path, output_base, analysis_name)
        
        if success:
            print(f"✅ {analysis_name} analysis completed!")
        else:
            print(f"❌ {analysis_name} analysis failed!")
    
    # === STEP 6: Final summary ===
    print("\n" + "=" * 60)
    print("SINGLE FILE WORKFLOW COMPLETE!")
    print("=" * 60)
    print(f"📁 Output directory: {output_base}")
    print(f"📊 Processed file: {os.path.basename(excel_file_path)}")
    
    # List generated files
    analysis_dirs = ["BoxPlot_Analysis", "Weighted_BarPlot_Analysis", "BarPlot_Analysis"]
    print(f"📈 Analysis results in: {output_base}")
    
    for analysis_dir in analysis_dirs:
        analysis_path = os.path.join(output_base, analysis_dir)
        if os.path.exists(analysis_path):
            files = os.listdir(analysis_path)
            print(f"   � {analysis_dir}:")
            for file in files:
                print(f"      - {file}")
    
    print("\n🎉 Single file analysis completed successfully!")
    
    # Clean up tkinter
    cleanup_tkinter()


if __name__ == "__main__":
    main()
