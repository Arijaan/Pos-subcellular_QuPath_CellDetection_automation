"""
Plotting Module
Handles all plotting and visualization functionality
"""

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os


def create_plots(summarized_data, plots_folder):
    """Create plots from the summarized data"""
    print("\n" + "="*50)
    print("CREATING PLOTS")
    print("="*50)
    
    # Create plots folder
    os.makedirs(plots_folder, exist_ok=True)
    
    if not summarized_data:
        print("No summarized data available for plotting")
        return
    
    # Set up plotting style
    plt.style.use('default')
    
    # Create combined dataset for plotting
    all_data = []
    for condition, df in summarized_data.items():
        df_copy = df.copy()
        df_copy['Condition'] = condition
        all_data.append(df_copy)
    
    if not all_data:
        print("No data available for plotting")
        return
        
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Find numeric columns for plotting
    numeric_cols = combined_df.select_dtypes(include=[np.number]).columns.tolist()
    
    # Remove condition column if it exists
    if 'Condition' in numeric_cols:
        numeric_cols.remove('Condition')
    
    print(f"Creating plots for {len(numeric_cols)} numeric variables...")
    
    # Create individual plots for each numeric variable
    for col in numeric_cols:
        if combined_df[col].notna().sum() == 0:
            continue
            
        plt.figure(figsize=(10, 6))
        
        # Create box plot
        plt.subplot(1, 2, 1)
        conditions = combined_df['Condition'].unique()
        data_for_box = [combined_df[combined_df['Condition'] == cond][col].dropna() 
                       for cond in conditions]
        plt.boxplot(data_for_box, labels=conditions)
        plt.title(f'{col} - Box Plot')
        plt.ylabel(col)
        plt.xticks(rotation=45)
        
        # Create bar plot with means
        plt.subplot(1, 2, 2)
        means = combined_df.groupby('Condition')[col].mean()
        stds = combined_df.groupby('Condition')[col].std()
        
        bars = plt.bar(means.index, means.values, yerr=stds.values, capsize=5)
        plt.title(f'{col} - Mean ± SD')
        plt.ylabel(col)
        plt.xticks(rotation=45)
        
        # Add value labels on bars
        for bar, mean_val in zip(bars, means.values):
            plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + bar.get_height()*0.01,
                    f'{mean_val:.2f}', ha='center', va='bottom')
        
        plt.tight_layout()
        
        # Save plot
        safe_col_name = "".join(c for c in col if c.isalnum() or c in (' ', '-', '_')).rstrip()
        plot_filename = f"{safe_col_name}_comparison.png"
        plot_path = os.path.join(plots_folder, plot_filename)
        plt.savefig(plot_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"Created plot: {plot_filename}")
    
    # Create summary comparison plot
    create_summary_plot(combined_df, plots_folder)


def create_summary_plot(combined_df, plots_folder):
    """Create a comprehensive summary plot"""
    # Look for cell count columns
    cell_count_cols = [col for col in combined_df.columns 
                      if any(keyword in col.lower() for keyword in 
                           ['num cells', 'num positive', 'num 1+', 'num 2+', 'num 3+', 'negative'])]
    
    if not cell_count_cols:
        print("No cell count columns found for summary plot")
        return
        
    plt.figure(figsize=(15, 10))
    
    # Create subplot for each cell count type
    n_plots = len(cell_count_cols)
    n_cols = min(3, n_plots)
    n_rows = (n_plots + n_cols - 1) // n_cols
    
    for i, col in enumerate(cell_count_cols):
        plt.subplot(n_rows, n_cols, i + 1)
        
        conditions = combined_df['Condition'].unique()
        means = combined_df.groupby('Condition')[col].mean()
        stds = combined_df.groupby('Condition')[col].std()
        
        bars = plt.bar(means.index, means.values, yerr=stds.values, capsize=5)
        plt.title(col)
        plt.ylabel('Count')
        plt.xticks(rotation=45)
        
        # Add value labels
        for bar, mean_val in zip(bars, means.values):
            plt.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                    f'{mean_val:.1f}', ha='center', va='bottom')
    
    plt.tight_layout()
    summary_plot_path = os.path.join(plots_folder, "Summary_Cell_Counts.png")
    plt.savefig(summary_plot_path, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"Created summary plot: Summary_Cell_Counts.png")


def create_advanced_plots(summarized_data, plots_folder):
    """Create additional advanced plots"""
    if not summarized_data:
        return
    
    # Create combined dataset
    all_data = []
    for condition, df in summarized_data.items():
        df_copy = df.copy()
        df_copy['Condition'] = condition
        all_data.append(df_copy)
    
    if not all_data:
        return
        
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Create correlation heatmap if multiple numeric variables exist
    numeric_cols = combined_df.select_dtypes(include=[np.number]).columns.tolist()
    if 'Condition' in numeric_cols:
        numeric_cols.remove('Condition')
    
    if len(numeric_cols) > 1:
        plt.figure(figsize=(10, 8))
        correlation_matrix = combined_df[numeric_cols].corr()
        
        # Create heatmap
        im = plt.imshow(correlation_matrix, cmap='coolwarm', aspect='auto', vmin=-1, vmax=1)
        plt.colorbar(im)
        
        # Add labels
        plt.xticks(range(len(numeric_cols)), numeric_cols, rotation=45, ha='right')
        plt.yticks(range(len(numeric_cols)), numeric_cols)
        plt.title('Correlation Matrix of Measurements')
        
        # Add correlation values to the plot
        for i in range(len(numeric_cols)):
            for j in range(len(numeric_cols)):
                plt.text(j, i, f'{correlation_matrix.iloc[i, j]:.2f}', 
                        ha='center', va='center', color='black' if abs(correlation_matrix.iloc[i, j]) < 0.5 else 'white')
        
        plt.tight_layout()
        correlation_plot_path = os.path.join(plots_folder, "Correlation_Matrix.png")
        plt.savefig(correlation_plot_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        print("Created correlation matrix plot")


def setup_plot_style():
    """Set up consistent plotting style"""
    plt.rcParams['figure.dpi'] = 100
    plt.rcParams['savefig.dpi'] = 300
    plt.rcParams['font.size'] = 10
    plt.rcParams['axes.titlesize'] = 12
    plt.rcParams['axes.labelsize'] = 10
    plt.rcParams['xtick.labelsize'] = 9
    plt.rcParams['ytick.labelsize'] = 9
    plt.rcParams['legend.fontsize'] = 9
