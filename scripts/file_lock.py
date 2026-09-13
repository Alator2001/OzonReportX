"""Non-blocking, reentrant file locks shared by GUI and CLI writers."""
from contextlib import contextmanager
from functools import wraps
import os
from pathlib import Path
import sys
import threading

# Both script and package imports must share the same in-process lock registry.
sys.modules.setdefault("file_lock", sys.modules[__name__])
sys.modules.setdefault("scripts.file_lock", sys.modules[__name__])


class FileBusyError(RuntimeError):
    pass


_registry_guard = threading.Lock()
_registry = {}


@contextmanager
def exclusive_file(destination):
    destination = Path(destination).resolve()
    key = os.path.normcase(str(destination))
    with _registry_guard:
        state = _registry.setdefault(key, {"lock": threading.RLock(), "depth": 0})
    if not state["lock"].acquire(blocking=False):
        raise FileBusyError(f"Файл {destination.name} изменяется другой операцией. Повторите после её завершения.")
    handle = None
    try:
        if state["depth"] == 0:
            lock_path = destination.with_name(f".{destination.name}.ozon-lock")
            handle = open(lock_path, "a+b")
            if os.fstat(handle.fileno()).st_size == 0:
                handle.write(b"0")
                handle.flush()
            handle.seek(0)
            try:
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError as exc:
                raise FileBusyError(f"Файл {destination.name} используется другим экземпляром приложения.") from exc
        state["depth"] += 1
        try:
            yield
        finally:
            state["depth"] -= 1
    finally:
        # Closing the descriptor releases the OS lock, including after exceptions.
        if handle is not None:
            handle.close()
        state["lock"].release()


def locked_costs(function):
    """Hold the costs lock through the entire read/compute/write operation."""
    @wraps(function)
    def wrapped(repo_root, *args, **kwargs):
        with exclusive_file(Path(repo_root) / "costs.xlsx"):
            return function(repo_root, *args, **kwargs)
    return wrapped


def file_signature(path):
    stat = Path(path).stat()
    return stat.st_mtime_ns, stat.st_size, stat.st_ino
