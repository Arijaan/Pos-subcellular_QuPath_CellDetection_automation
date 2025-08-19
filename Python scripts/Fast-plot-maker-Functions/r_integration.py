"""
R Scripts Integration Module
Handles execution of R scripts and interaction with R environment
"""

import subprocess
import os


def run_r_scripts(base_dir, r_script_dir):
    """Run the two R scripts with automatic file detection and custom plot names"""
    print("\n" + "="*50)
    print("RUNNING R SCRIPTS")
    print("="*50)
    
    # Automatically find the summary files
    summary_dir = os.path.join(base_dir, "Measurement_summary")
    if not os.path.exists(summary_dir):
        print(f"Error: Summary directory not found: {summary_dir}")
        return False
    
    # Look for control, T1D, and T2D summary files
    control_file = None
    t1d_file = None
    t2d_file = None
    
    for filename in os.listdir(summary_dir):
        if filename.endswith('.csv'):
            filename_lower = filename.lower()
            if 'control' in filename_lower:
                control_file = os.path.join(summary_dir, filename)
            elif 't1d' in filename_lower:
                t1d_file = os.path.join(summary_dir, filename)
            elif 't2d' in filename_lower:
                t2d_file = os.path.join(summary_dir, filename)
    
    # Validate that all required files were found
    missing_files = []
    if not control_file:
        missing_files.append("control")
    if not t1d_file:
        missing_files.append("T1D")
    if not t2d_file:
        missing_files.append("T2D")
    
    if missing_files:
        print(f"Error: Missing summary files for: {', '.join(missing_files)}")
        return False
    
    print(f"✓ Found control file: {os.path.basename(control_file)}")
    print(f"✓ Found T1D file: {os.path.basename(t1d_file)}")
    print(f"✓ Found T2D file: {os.path.basename(t2d_file)}")
    
    # Define R scripts with their corresponding output plot names
    r_scripts = [
        {
            "script": "F-OnlyPosProp-Ctrl-T1D-T2D.R",
            "plot_name": "Classifications"
        },
        {
            "script": "F-Weighted-Ctrl-T1D-T2D.R", 
            "plot_name": "Weighted"
        }
    ]
    
    success_count = 0
    
    for script_info in r_scripts:
        script_name = script_info["script"]
        plot_name = script_info["plot_name"]
        script_path = os.path.join(r_script_dir, script_name)
        
        if not os.path.exists(script_path):
            print(f"Warning: R script not found: {script_path}")
            continue
            
        print(f"\nRunning {script_name} (Output: {plot_name})...")
        
        try:
            # Change to the base directory so R script can find the summary files
            original_dir = os.getcwd()
            os.chdir(base_dir)
            
            # Run R script with arguments for file paths and output name
            cmd_args = [
                "Rscript", 
                script_path,
                control_file,
                t1d_file, 
                t2d_file,
                plot_name  # Pass the custom plot name
            ]
            
            result = subprocess.run(
                cmd_args, 
                capture_output=True, 
                text=True,
                timeout=300  # 5 minute timeout
            )
            
            # Restore original directory
            os.chdir(original_dir)
            
            if result.returncode == 0:
                print(f"✓ {script_name} completed successfully")
                print(f"  → Plot saved as: {plot_name}")
                if result.stdout:
                    print(f"  Output: {result.stdout}")
                success_count += 1
            else:
                print(f"✗ {script_name} failed with return code {result.returncode}")
                if result.stderr:
                    print(f"  Error: {result.stderr}")
                    
        except subprocess.TimeoutExpired:
            print(f"✗ {script_name} timed out after 5 minutes")
        except FileNotFoundError:
            print(f"✗ Rscript not found. Please ensure R is installed and in PATH")
        except Exception as e:
            print(f"✗ Error running {script_name}: {e}")
    
    print(f"\nR Scripts Summary: {success_count}/{len(r_scripts)} completed successfully")
    return success_count > 0


def check_r_installation():
    """Check if R is properly installed and accessible"""
    try:
        result = subprocess.run(
            ["Rscript", "--version"], 
            capture_output=True, 
            text=True,
            timeout=10
        )
        if result.returncode == 0:
            print("✓ R installation found")
            return True
        else:
            print("✗ R installation issue")
            return False
    except FileNotFoundError:
        print("✗ R/Rscript not found in PATH")
        return False
    except Exception as e:
        print(f"✗ Error checking R installation: {e}")
        return False


def validate_r_scripts(r_script_dir):
    """Validate that required R scripts exist"""
    r_scripts = [
        "F-OnlyPosProp-Ctrl-T1D-T2D.R",
        "F-Weighted-Ctrl-T1D-T2D.R"
    ]
    
    missing_scripts = []
    for script_name in r_scripts:
        script_path = os.path.join(r_script_dir, script_name)
        if not os.path.exists(script_path):
            missing_scripts.append(script_name)
    
    if missing_scripts:
        print(f"Warning: Missing R scripts: {', '.join(missing_scripts)}")
        return False
    else:
        print("✓ All required R scripts found")
        return True


def create_r_script_template(script_path, plot_name):
    """Create a template R script that accepts command line arguments"""
    template = f'''#!/usr/bin/env Rscript

# {os.path.basename(script_path)}
# Auto-generated template for {plot_name} analysis

# Get command line arguments
args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 4) {{
    cat("Usage: Rscript {os.path.basename(script_path)} <control_file> <t1d_file> <t2d_file> <plot_name>\\n")
    quit(status = 1)
}}

control_file <- args[1]
t1d_file <- args[2] 
t2d_file <- args[3]
plot_name <- args[4]

cat("Processing files:\\n")
cat("  Control:", control_file, "\\n")
cat("  T1D:", t1d_file, "\\n") 
cat("  T2D:", t2d_file, "\\n")
cat("  Output plot name:", plot_name, "\\n")

# Load required libraries
library(ggplot2)
library(readr)

# Read the CSV files
tryCatch({{
    control_data <- read_csv(control_file, show_col_types = FALSE)
    t1d_data <- read_csv(t1d_file, show_col_types = FALSE)
    t2d_data <- read_csv(t2d_file, show_col_types = FALSE)
    
    cat("✓ Successfully loaded all data files\\n")
    
    # TODO: Add your specific analysis code here
    # This is a template - replace with your actual analysis
    
    # Example: Create a simple plot
    # combined_data <- rbind(
    #     data.frame(Group = "Control", Value = control_data$SomeColumn),
    #     data.frame(Group = "T1D", Value = t1d_data$SomeColumn),
    #     data.frame(Group = "T2D", Value = t2d_data$SomeColumn)
    # )
    # 
    # p <- ggplot(combined_data, aes(x = Group, y = Value)) +
    #     geom_boxplot() +
    #     labs(title = plot_name)
    # 
    # ggsave(paste0(plot_name, ".png"), plot = p, width = 8, height = 6, dpi = 300)
    
    cat("✓ Analysis completed successfully\\n")
    cat("✓ Plot saved as:", paste0(plot_name, ".png"), "\\n")
    
}}, error = function(e) {{
    cat("✗ Error processing files:", conditionMessage(e), "\\n")
    quit(status = 1)
}})
'''
    
    return template
