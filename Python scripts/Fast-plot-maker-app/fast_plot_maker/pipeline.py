import os
from .data_processing import summarize_measurements
from .file_io import ensure_dirs
from .r_integration import run_r_scripts, check_r_installation
from .plotting import create_basic_plots, quick_overview_table

class Pipeline:
    def __init__(self, base_dir: str, progress=None, logger=None):
        self.base_dir = base_dir
        self.progress = progress
        self.logger = logger or (lambda lvl,msg: None)
        self.summary_dir = None
        self.plots_dir = None

    def log(self, lvl, msg):
        self.logger(lvl, msg)

    def run(self):
        self.log('info','Ensuring directories')
        self.summary_dir, self.plots_dir = ensure_dirs(self.base_dir)
        if self.progress:
            self.progress.step('Directories ready')

        self.log('info','Summarizing measurements')
        self.summary_dir = summarize_measurements(self.base_dir, logger=self.log)
        if self.progress:
            self.progress.step('Summaries created')

        if not check_r_installation():
            self.log('warn','R not found; skipping R scripts')
        else:
            self.log('info','Running R scripts')
            r_results = run_r_scripts(self.base_dir, self.summary_dir)
            for script, ok, msg in r_results:
                self.log('info' if ok else 'error', f"{script}: {'OK' if ok else 'FAIL'}")
                if msg:
                    self.log('debug', msg.strip()[:500])
        if self.progress:
            self.progress.step('R scripts')

        self.log('info','Collecting plots')
        generated = create_basic_plots(self.summary_dir, self.plots_dir)
        table = quick_overview_table(self.summary_dir, self.plots_dir)
        if table:
            generated.append(table)
        if self.progress:
            self.progress.step('Plots ready')

        return {
            'summary_dir': self.summary_dir,
            'plots_dir': self.plots_dir,
            'plots': generated
        }
