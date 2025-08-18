import os, shutil, pandas as pd, numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

PRIMARY_COLORS = ['#2563eb','#dc2626','#0d9488','#7c3aed','#d97706']

def create_basic_plots(summary_dir: str, plots_dir: str):
    """Collect R-generated plots (if any) & place copies in plots_dir for consistency."""
    generated = []
    for name in ['Classifications.png','Weighted.png']:
        src = os.path.join(summary_dir, name)
        if os.path.exists(src):
            dst = os.path.join(plots_dir, name)
            try:
                if src != dst:
                    shutil.copyfile(src, dst)
                generated.append(dst)
            except Exception:
                generated.append(src)  # fallback
    return generated

def quick_overview_table(summary_dir: str, plots_dir: str):
    # compile counts into a table plot if available
    dfs = []
    for f in os.listdir(summary_dir):
        if f.lower().endswith('_summary.csv'):
            path = os.path.join(summary_dir,f)
            try:
                df = pd.read_csv(path, header=None)
                df['Group'] = f.split('_')[0]
                dfs.append(df)
            except Exception:
                pass
    if not dfs:
        return None
    combined = pd.concat(dfs)
    # filter statistic rows
    stats = combined[combined[4].isin(['Num Cells','Num Positive','Num 1+','Num 2+','Num 3+','Num Negative'])]
    if stats.empty:
        return None
    # Coerce counts to numeric
    stats = stats.copy()
    stats.loc[:,5] = pd.to_numeric(stats[5], errors='coerce')
    pivot = stats.pivot_table(index=4, columns='Group', values=5, aggfunc='first')
    fig, ax = plt.subplots(figsize=(8,4))
    ax.axis('off')
    tbl = ax.table(cellText=pivot.fillna('').values, rowLabels=pivot.index, colLabels=pivot.columns, loc='center')
    fig.tight_layout()
    out_path = os.path.join(plots_dir,'Summary_Table.png')
    fig.savefig(out_path, dpi=200)
    plt.close(fig)
    return out_path
