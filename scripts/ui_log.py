"""Bounded UI history backed by a complete UTF-8 session log."""
from collections import deque
from pathlib import Path
import threading


class SessionLog:
    def __init__(self, path, max_lines=500):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.lines = deque(maxlen=max_lines)
        self._lock = threading.RLock()

    @staticmethod
    def normalize(text):
        return (text or "").replace("\r\n", "\n").replace("\r", "\n")

    def reset(self, text, *, persist=True):
        with self._lock:
            self.lines.clear()
            self.append(text, persist=persist)

    def append(self, text, *, persist=True):
        normalized = self.normalize(text)
        with self._lock:
            if persist:
                with self.path.open("a", encoding="utf-8") as stream:
                    stream.write(normalized)
                    if not normalized.endswith("\n"):
                        stream.write("\n")
            # Match the existing UI line splitting, but cap memory and rendered controls.
            self.lines.extend(normalized.split("\n"))

    def full_text(self):
        with self._lock:
            return self.path.read_text(encoding="utf-8") if self.path.exists() else ""
