import os, re, pandas as pd, numpy as np
from typing import Dict, List, Callable, Optional

GROUPS = ['control','t1d','t2d']
GROUP_LABELS = {
    'control': 'control',   # keep lowercase for consistency with legacy scripts
    't1d': 'T1D',
    't2d': 'T2D'
}

SUMMARY_KEYS = [
    'Num Cells','Num Positive','Num 1+','Num 2+','Num 3+','Num Negative','Total'
]

SUMMARY_KEY_NORMALIZE = {k.lower(): k for k in SUMMARY_KEYS}

def _clean_number(val):
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return val
    s = str(val).strip()
    if not s:
        return None
    # Keep digits, minus, dot
    s2 = re.sub(r"[^0-9eE+\-\.]", "", s)
    try:
        return float(s2)
    except Exception:
        return None

def _extract_summary_like(df: pd.DataFrame) -> Dict[str, float]:
    """Attempt to extract QuPath summary style stats from a loosely-formatted DataFrame.
    Looks for known labels in any cell & takes the immediate next column value in the same row
    or the next non-empty cell to the right.
    """
    found = {}
    if df is None or df.empty:
        return found
    # Iterate rows
    for ridx, row in df.iterrows():
        for cidx, v in enumerate(row):
            if pd.isna(v):
                continue
            label = str(v).strip()
            key_norm = label.lower()
            if key_norm in SUMMARY_KEY_NORMALIZE:
                # find value in next columns
                val = None
                for look_c in range(cidx+1, len(row)):
                    raw = row.iloc[look_c]
                    val_clean = _clean_number(raw)
                    if val_clean is not None:
                        val = val_clean
                        break
                if val is not None:
                    found[SUMMARY_KEY_NORMALIZE[key_norm]] = val
    return found

def summarize_measurements(base_dir: str, progress=None, logger: Optional[Callable[[str,str],None]] = None):
    summary_dir = os.path.join(base_dir, 'Measurement_summary')
    os.makedirs(summary_dir, exist_ok=True)
    csvs = [f for f in os.listdir(base_dir) if f.lower().endswith('.csv')]
    group_data = {g: [] for g in GROUPS}
    diagnostics = []
    for f in csvs:
        lower = f.lower()
        assigned = None
        for g in GROUPS:
            if g in lower:
                assigned = g
                break
        if not assigned:
            diagnostics.append((f, 'SKIP', 'No group keyword'))
            continue
        path = os.path.join(base_dir, f)
        try:
            df = pd.read_csv(path, low_memory=False)
        except Exception as e:
            diagnostics.append((f, 'ERROR', f'read_csv failed: {e}'))
            continue
        if df.empty:
            diagnostics.append((f, 'SKIP', 'Empty dataframe'))
            continue
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        summary_stats = None
        reason = ''

        if numeric_cols:
            summary_stats = df[numeric_cols].mean()
            reason = f'mean of {len(numeric_cols)} numeric cols'
        # Fallback: attempt summary style extraction if means empty or too few numeric cols
        if (summary_stats is None or summary_stats.empty or len(numeric_cols) < 2):
            extracted = _extract_summary_like(df)
            if extracted:
                summary_stats = pd.Series(extracted)
                reason = f'extracted summary keys ({len(extracted)})'
        # Second fallback: try reading without header if still empty
        if (summary_stats is None or summary_stats.empty):
            try:
                df2 = pd.read_csv(path, header=None, low_memory=False)
                extracted2 = _extract_summary_like(df2)
                if extracted2:
                    summary_stats = pd.Series(extracted2)
                    reason = f'extracted summary (no header) ({len(extracted2)})'
            except Exception:
                pass

        if summary_stats is None or summary_stats.empty:
            diagnostics.append((f, 'SKIP', 'No numeric or summary data'))
            continue
        group_data[assigned].append(summary_stats)
        diagnostics.append((f, 'OK', reason))
        if logger:
            logger('debug', f"{f}: {reason}")
    # build group summaries
    for g, arr in group_data.items():
        if not arr:
            diagnostics.append((g, 'GROUP_SKIP', 'No files aggregated'))
            continue
        combined = pd.DataFrame(arr)
        means = combined.mean()
        sums = combined.sum()
        cell_count_keywords = ['Num Cells','Num Positive','Num 1+','Num 2+','Num 3+','Num Negative','Total']
        final = {}
        for k in means.index:
            if any(kw.lower() in str(k).lower() for kw in cell_count_keywords):
                final[k] = sums[k]
            else:
                final[k] = means[k]
        # create summary format
        total_cells = sum(v for name,v in final.items() if 'num cells' in str(name).lower())
        total_positive = sum(v for name,v in final.items() if any(x in str(name).lower() for x in ['num 1+','num 2+','num 3+']))
        num_1 = sum(v for name,v in final.items() if 'num 1+' in str(name).lower())
        num_2 = sum(v for name,v in final.items() if 'num 2+' in str(name).lower())
        num_3 = sum(v for name,v in final.items() if 'num 3+' in str(name).lower())
        num_neg = sum(v for name,v in final.items() if 'num negative' in str(name).lower())
        summary_rows = [
            ['SUMMARY',''], ['Statistic','Count'], ['Num Cells', int(total_cells)], ['Num Positive', int(total_positive)],
            ['Num 1+', int(num_1)], ['Num 2+', int(num_2)], ['Num 3+', int(num_3)], ['Num Negative', int(num_neg)]
        ]
        out_df = pd.DataFrame(columns=[0,1,2,3,4,5])
        for i,(a,b) in enumerate(summary_rows):
            out_df.loc[i,4] = a
            out_df.loc[i,5] = b
        # Use stable naming so other components (R scripts, downstream plotting) can find them
        out_path = os.path.join(summary_dir, f"{GROUP_LABELS.get(g,g)}_summary.csv")
        out_df.to_csv(out_path, index=False, header=False)
        diagnostics.append((out_path, 'WRITE', f'Rows={len(out_df)}'))
    # Write diagnostics CSV
    if diagnostics:
        diag_path = os.path.join(summary_dir, '_processing_diagnostics.csv')
        pd.DataFrame(diagnostics, columns=['Item','Status','Detail']).to_csv(diag_path, index=False)
        if logger:
            logger('info', f'Diagnostics written: {diag_path}')
    return summary_dir
