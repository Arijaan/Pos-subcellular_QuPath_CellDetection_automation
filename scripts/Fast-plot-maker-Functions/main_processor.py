"""
Main Processor Module
Contains the main MeasurementProcessor class that orchestrates the entire pipeline
"""

import os
import sys
from pathlib import Path

# Add the Functions directory to the path
functions_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Fast-plot-maker-Functions')
if functions_dir not in sys.path:
    sys.path.insert(0, functions_dir)

# Import all the modules
from data_processing import create_summary_files_for_pipeline, load_summarized_data
from r_integration import run_r_scripts, check_r_installation, validate_r_scripts
from plotting import create_plots, create_advanced_plots, setup_plot_style
from file_io import save_summary_tables, create_results_summary, ensure_directories_exist, backup_original_files
from ui_utils import select_base_directory, show_completion_message, get_user_preferences, ProgressTracker


class MeasurementProcessor:
    """
    Complete measurement processing pipeline including summarization, 
    R script execution, and plotting.
    """
    
    def __init__(self):
        self.base_dir = None
        self.summary_folder = None
        self.plots_folder = None
        self.r_script_dir = None
        self.preferences = {}
        
    def initialize_directories(self):
        """Initialize and set up directory structure"""
        if not self.base_dir:
            return False
            
        # Set up folder paths
        self.summary_folder = os.path.join(self.base_dir, "Measurement_summary")
        self.plots_folder = os.path.join(self.base_dir, "Plots_and_Analysis")
        
        # Look for R scripts in the main scripts directory (parent of Functions directory)
        functions_dir = os.path.dirname(os.path.abspath(__file__))
        script_dir = os.path.dirname(functions_dir)  # Go up one level to main scripts directory
        self.r_script_dir = script_dir
        
        # Ensure directories exist
        ensure_directories_exist(self.base_dir)
        
        print(f"Summary folder: {self.summary_folder}")
        print(f"Plots folder: {self.plots_folder}")
        print(f"R script directory: {self.r_script_dir}")
        
        return True
    
    def validate_prerequisites(self):
        """Validate that all prerequisites are met"""
        print("\n" + "="*50)
        print("VALIDATING PREREQUISITES")
        print("="*50)
        
        issues = []
        
        # Check R installation
        if not check_r_installation():
            issues.append("R is not properly installed or not in PATH")
        
        # Check R scripts if validation requested
        if self.preferences.get('validate_r_scripts', True):
            if not validate_r_scripts(self.r_script_dir):
                issues.append("Required R scripts are missing")
        
        # Check for CSV files
        csv_files = [f for f in os.listdir(self.base_dir) if f.endswith('.csv')]
        if not csv_files:
            issues.append("No CSV files found in base directory")
        
        if issues:
            print("⚠️  Issues found:")
            for issue in issues:
                print(f"   - {issue}")
            return False
        else:
            print("✓ All prerequisites validated")
            return True
    
    def run_complete_pipeline(self):
        """Run the complete analysis pipeline"""
        print("QuPath Measurements Complete Analysis Pipeline")
        print("=" * 60)
        
        # Initialize progress tracker
        progress = ProgressTracker(7)
        
        # Step 0: Get user preferences
        self.preferences = get_user_preferences()
        
        # Step 1: Select directory
        progress.update("Selecting directory")
        self.base_dir = select_base_directory()
        if not self.base_dir:
            return
        
        # Initialize directory structure
        if not self.initialize_directories():
            show_completion_message(self.plots_folder, success=False)
            return
        
        # Step 2: Validate prerequisites
        progress.update("Validating prerequisites")
        if not self.validate_prerequisites():
            print("\n⚠️  Prerequisites not met. Please resolve issues and try again.")
            show_completion_message(self.plots_folder, success=False)
            return
        
        # Step 3: Backup files if requested
        progress.update("Creating backups")
        if self.preferences.get('backup_files', True):
            backup_original_files(self.base_dir)
        
        # Step 4: Create initial summaries
        progress.update("Creating initial summaries")
        print("\n" + "="*50)
        print("STEP 1: CREATING INITIAL SUMMARIES")
        print("="*50)
        
        original_cwd = os.getcwd()
        os.chdir(self.base_dir)
        
        try:
            create_summary_files_for_pipeline(self.base_dir)
        except Exception as e:
            print(f"Error in initial summary creation: {e}")
            os.chdir(original_cwd)
            show_completion_message(self.plots_folder, success=False)
            return
        
        os.chdir(original_cwd)
        
        # Step 5: Run R scripts
        progress.update("Running R scripts")
        print("\n" + "="*50)
        print("STEP 2: RUNNING R SCRIPTS")
        print("="*50)
        run_r_scripts(self.base_dir, self.r_script_dir)
        
        # Step 6: Load and process data
        progress.update("Loading and processing data")
        summarized_data = load_summarized_data(self.base_dir)
        
        # Step 7: Create plots and save results
        progress.update("Creating plots and saving results")
        if summarized_data:
            # Set up plotting style
            setup_plot_style()
            
            # Create standard plots
            create_plots(summarized_data, self.plots_folder)
            
            # Create advanced plots if requested
            if self.preferences.get('create_advanced_plots', False):
                create_advanced_plots(summarized_data, self.plots_folder)
            
            # Save summary tables
            save_summary_tables(summarized_data, self.plots_folder)
            
            # Create results summary
            create_results_summary(self.base_dir, self.plots_folder, summarized_data)
        
        progress.update("Complete!")
        
        print("\n" + "="*60)
        print("PIPELINE COMPLETE!")
        print("="*60)
        print(f"Results saved in: {self.plots_folder}")
        
        # Show completion message
        show_completion_message(self.plots_folder, success=bool(summarized_data))


def run_pipeline():
    """Convenience function to run the complete pipeline"""
    processor = MeasurementProcessor()
    processor.run_complete_pipeline()
