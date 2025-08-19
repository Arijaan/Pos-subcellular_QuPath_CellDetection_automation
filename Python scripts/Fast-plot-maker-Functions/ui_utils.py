"""
User Interface Module
Handles GUI interactions, directory selection, and user prompts
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os


def select_base_directory(title="Select base directory containing measurements"):
    """Select the base directory containing measurements"""
    root = tk.Tk()
    root.withdraw()
    
    base_dir = filedialog.askdirectory(title=title)
    
    if not base_dir:
        messagebox.showinfo("Cancelled", "No directory selected. Exiting.")
        return None
        
    print(f"Selected base directory: {base_dir}")
    return base_dir


def show_completion_message(plots_folder, success=True):
    """Show completion notification to user"""
    root = tk.Tk()
    root.withdraw()
    
    if success:
        messagebox.showinfo(
            "Analysis Complete", 
            f"Analysis pipeline completed successfully!\n\n"
            f"Results saved in:\n{plots_folder}\n\n"
            f"Check the folder for:\n"
            f"- Individual variable plots\n"
            f"- Summary comparison plots\n"
            f"- Excel summary tables\n"
            f"- Analysis summary report"
        )
    else:
        messagebox.showerror(
            "Analysis Failed", 
            f"Analysis pipeline encountered errors.\n\n"
            f"Please check the console output for details."
        )


def confirm_operation(title, message):
    """Show confirmation dialog"""
    root = tk.Tk()
    root.withdraw()
    
    return messagebox.askyesno(title, message)


def show_error_message(title, message):
    """Show error message dialog"""
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showerror(title, message)


def show_info_message(title, message):
    """Show information message dialog"""
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo(title, message)


def get_user_preferences():
    """Get user preferences for analysis options"""
    root = tk.Tk()
    root.withdraw()
    
    preferences = {
        'create_advanced_plots': False,
        'backup_files': True,
        'validate_r_scripts': True
    }
    
    # Simple yes/no dialogs for preferences
    preferences['create_advanced_plots'] = messagebox.askyesno(
        "Plot Options", 
        "Create advanced plots (correlation matrices, etc.)?\n\n"
        "This may take additional time but provides more comprehensive analysis."
    )
    
    preferences['backup_files'] = messagebox.askyesno(
        "Backup Options", 
        "Create backup of original files before processing?\n\n"
        "Recommended for safety."
    )
    
    return preferences


def display_progress(step, total_steps, description):
    """Display progress information"""
    progress_bar = "█" * int(20 * step / total_steps)
    empty_bar = "░" * (20 - int(20 * step / total_steps))
    percentage = int(100 * step / total_steps)
    
    print(f"\rProgress: [{progress_bar}{empty_bar}] {percentage}% - {description}", end="", flush=True)
    
    if step == total_steps:
        print()  # New line when complete


class ProgressTracker:
    """Simple progress tracking class"""
    
    def __init__(self, total_steps):
        self.total_steps = total_steps
        self.current_step = 0
        
    def update(self, description=""):
        self.current_step += 1
        display_progress(self.current_step, self.total_steps, description)
        
    def reset(self):
        self.current_step = 0
