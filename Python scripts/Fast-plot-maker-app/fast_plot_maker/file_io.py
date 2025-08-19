import os, pandas as pd

def ensure_dirs(base_dir):
    summary = os.path.join(base_dir,'Measurement_summary')
    plots = os.path.join(base_dir,'Plots_and_Analysis')
    os.makedirs(summary, exist_ok=True)
    os.makedirs(plots, exist_ok=True)
    return summary, plots

