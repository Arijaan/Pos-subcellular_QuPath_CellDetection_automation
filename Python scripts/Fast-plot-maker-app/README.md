# Fast Plot Maker App

Modern GUI application to run the QuPath measurement summarization + R analysis + plotting pipeline.

## Features
- Point & click selection of measurement directory
- Automatic detection of Control / T1D / T2D summary CSV files
- Status timeline with live progress & logs
- Run legacy summary-only mode or full pipeline
- Auto execution of R scripts with argument passing
- Displays generated plots inline (preview) and links to output folder
- Modular architecture

## Quick Start
1. Install dependencies:
```
pip install -r requirements.txt
```
2. Run the app:
```
python -m fast_plot_maker
```

## Requirements
- Python 3.9+
- R installed and Rscript in PATH
- Packages: pandas, numpy, matplotlib, PySide6

## Structure
```
fast_plot_maker/
  __init__.py
  main.py
  app.py
  pipeline.py
  r_integration.py
  data_processing.py
  plotting.py
  file_io.py
  utils.py
  ui/
    style.qss
    icons.py
```

## Outputs
Generated under selected base folder:
- Measurement_summary/
- Plots_and_Analysis/

## License
Internal use.
