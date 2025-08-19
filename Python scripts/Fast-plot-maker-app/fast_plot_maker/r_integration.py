import os, subprocess
from typing import Tuple

def check_r_installation() -> bool:
    try:
        r = subprocess.run(["Rscript", "--version"], capture_output=True, text=True, timeout=10)
        return r.returncode == 0
    except Exception:
        return False

def run_r_scripts(base_dir: str, summary_dir: str) -> list:
    """Run R scripts located in the parent 'scripts' directory (outside the app folder)."""
    scripts = [
        ("F-OnlyPosProp-Ctrl-T1D-T2D.R", "Classifications"),
        ("F-Weighted-Ctrl-T1D-T2D.R", "Weighted")
    ]
    results = []
    # Determine probable scripts directory (parent of app root)
    app_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    scripts_dir = os.path.abspath(os.path.join(app_root, '..'))

    control = _find_file(summary_dir, 'control')
    t1d = _find_file(summary_dir, 't1d')
    t2d = _find_file(summary_dir, 't2d')
    csv_diag = f"CSV found -> control: {bool(control)}, T1D: {bool(t1d)}, T2D: {bool(t2d)} (dir={summary_dir})"

    for script, outname in scripts:
        script_path = os.path.join(scripts_dir, script)
        if not os.path.exists(script_path):
            results.append((script, False, f"Missing R script at {script_path}"))
            continue
        if not all([control, t1d, t2d]):
            results.append((script, False, f"Missing one or more summary CSVs. {csv_diag}"))
            continue
        try:
            proc = subprocess.run([
                'Rscript', script_path,
                control, t1d, t2d, outname
            ], capture_output=True, text=True, cwd=summary_dir, timeout=300)
            ok = proc.returncode == 0
            msg = (proc.stdout + '\n' + proc.stderr).strip()
            results.append((script, ok, msg if msg else ''))
        except Exception as e:
            results.append((script, False, f"Execution error: {e}"))
    return results

def _find_file(folder: str, keyword: str):
    for f in os.listdir(folder):
        if f.lower().endswith('.csv') and keyword in f.lower():
            return os.path.join(folder, f)
    return None
