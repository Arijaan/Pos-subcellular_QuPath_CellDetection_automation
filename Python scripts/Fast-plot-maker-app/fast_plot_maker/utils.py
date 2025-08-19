import os, sys, time
from dataclasses import dataclass
from typing import List, Optional

@dataclass
class LogEvent:
    level: str
    message: str
    timestamp: float

    def format(self) -> str:
        return f"[{time.strftime('%H:%M:%S', time.localtime(self.timestamp))}] {self.level.upper()}: {self.message}"

def add_path(p):
    if p not in sys.path:
        sys.path.insert(0, p)

class ProgressBus:
    def __init__(self, total_steps: int):
        self.total = total_steps
        self.current = 0
        self.callbacks = []

    def step(self, label: str):
        self.current += 1
        pct = min(100, int(self.current / self.total * 100))
        for cb in self.callbacks:
            cb(pct, label)

    def on_update(self, cb):
        self.callbacks.append(cb)

