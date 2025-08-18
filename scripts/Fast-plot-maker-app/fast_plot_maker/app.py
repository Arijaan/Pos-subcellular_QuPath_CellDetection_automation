import os, sys, webbrowser
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QHBoxLayout,
    QLineEdit, QTextEdit, QListWidget, QListWidgetItem, QProgressBar, QFrame, QMessageBox,
    QSizePolicy, QSplitter
)
from PySide6.QtCore import Qt, QThread, Signal
from .utils import ProgressBus, LogEvent
from .pipeline import Pipeline

class Worker(QThread):
    progress_update = Signal(int, str)
    log_event = Signal(str, str)
    finished = Signal(dict)

    def __init__(self, base_dir):
        super().__init__()
        self.base_dir = base_dir

    def run(self):
        bus = ProgressBus(5)
        bus.on_update(lambda pct,label: self.progress_update.emit(pct,label))
        def logger(level, msg):
            self.log_event.emit(level, msg)
        pipeline = Pipeline(self.base_dir, progress=bus, logger=logger)
        result = pipeline.run()
        self.finished.emit(result)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Fast Plot Maker')
        self.setMinimumSize(1100, 680)
        self.base_dir = ''
        self.worker = None

        with open(os.path.join(os.path.dirname(__file__), 'ui', 'style.qss'),'r') as f:
            self.setStyleSheet(f.read())

        layout = QVBoxLayout(self)
        header = QLabel('Fast Plot Maker')
        header.setObjectName('Heading')
        sub = QLabel('QuPath measurement summarization • R analysis • Plot generation')
        sub.setObjectName('Sub')
        layout.addWidget(header)
        layout.addWidget(sub)

        # Directory chooser
        dir_frame = QFrame(); dir_frame.setObjectName('Card')
        dir_layout = QHBoxLayout(dir_frame)
        self.dir_edit = QLineEdit(); self.dir_edit.setPlaceholderText('Select base directory containing measurement CSV files...')
        browse_btn = QPushButton('Browse')
        browse_btn.clicked.connect(self.choose_dir)
        dir_layout.addWidget(self.dir_edit)
        dir_layout.addWidget(browse_btn)
        layout.addWidget(dir_frame)

        # Controls
        control_frame = QFrame(); control_frame.setObjectName('Card')
        ctrl_layout = QHBoxLayout(control_frame)
        self.run_btn = QPushButton('Run Full Pipeline'); self.run_btn.clicked.connect(self.run_pipeline)
        self.run_btn.setEnabled(False)
        self.open_summary_btn = QPushButton('Open Output Folder'); self.open_summary_btn.clicked.connect(self.open_output)
        self.open_summary_btn.setEnabled(False)
        ctrl_layout.addWidget(self.run_btn)
        ctrl_layout.addWidget(self.open_summary_btn)
        layout.addWidget(control_frame)

        # Splitter for logs and previews
        splitter = QSplitter(Qt.Horizontal)

        # Left side: logs & progress
        left_frame = QFrame(); left_frame.setObjectName('Card')
        left_layout = QVBoxLayout(left_frame)
        self.progress_bar = QProgressBar(); self.progress_bar.setValue(0)
        self.progress_label = QLabel('Idle')
        left_layout.addWidget(self.progress_bar)
        left_layout.addWidget(self.progress_label)
        self.log_view = QListWidget(); left_layout.addWidget(self.log_view,1)
        splitter.addWidget(left_frame)

        # Right side: previews
        right_frame = QFrame(); right_frame.setObjectName('Card')
        right_layout = QVBoxLayout(right_frame)
        self.preview_label = QLabel('Generated Plots'); self.preview_label.setObjectName('Sub')
        right_layout.addWidget(self.preview_label)
        self.preview_list = QListWidget(); right_layout.addWidget(self.preview_list,1)
        splitter.addWidget(right_frame)

        layout.addWidget(splitter,1)

    def choose_dir(self):
        d = QFileDialog.getExistingDirectory(self, 'Select Measurements Base Directory')
        if d:
            self.base_dir = d
            self.dir_edit.setText(d)
            self.run_btn.setEnabled(True)
            self.log('info', f'Selected directory: {d}')

    def run_pipeline(self):
        if not self.base_dir:
            QMessageBox.warning(self,'Missing','Please select a base directory first')
            return
        self.log_view.clear(); self.preview_list.clear()
        self.progress_bar.setValue(0)
        self.worker = Worker(self.base_dir)
        self.worker.progress_update.connect(self.on_progress)
        self.worker.log_event.connect(self.log)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()
        self.run_btn.setEnabled(False)
        self.log('info','Pipeline started')

    def on_progress(self, pct, label):
        self.progress_bar.setValue(pct)
        self.progress_label.setText(label)

    def on_finished(self, result):
        self.log('info','Pipeline finished')
        self.run_btn.setEnabled(True)
        self.open_summary_btn.setEnabled(True)
        for plot in result.get('plots', []):
            item = QListWidgetItem(os.path.basename(plot))
            item.setToolTip(plot)
            self.preview_list.addItem(item)
        self.summary_dir = result.get('summary_dir')
        self.plots_dir = result.get('plots_dir')

    def log(self, level, message):
        item = QListWidgetItem(f"[{level.upper()}] {message}")
        if level == 'error':
            item.setForeground(Qt.red)
        elif level == 'warn':
            item.setForeground(Qt.yellow)
        elif level == 'debug':
            item.setForeground(Qt.gray)
        self.log_view.addItem(item)
        self.log_view.scrollToBottom()

    def open_output(self):
        if hasattr(self,'plots_dir') and self.plots_dir and os.path.exists(self.plots_dir):
            webbrowser.open(self.plots_dir)


def launch():
    app = QApplication(sys.argv)
    w = MainWindow(); w.show()
    sys.exit(app.exec())
