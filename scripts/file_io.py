"""Atomic file writes shared by report generation and the UI."""
from contextlib import contextmanager
import os
from pathlib import Path
import shutil
import tempfile
try:
    from scripts.file_lock import exclusive_file
except ModuleNotFoundError:
    from file_lock import exclusive_file


@contextmanager
def atomic_output_path(destination):
    """Publish a completed file only after the writer succeeds."""
    destination = Path(destination)
    with exclusive_file(destination):
        with _temporary_output_path(destination) as temporary:
            yield temporary


@contextmanager
def _temporary_output_path(destination):
    descriptor, name = tempfile.mkstemp(
        prefix="~tmp_", suffix=destination.suffix, dir=destination.parent,
    )
    os.close(descriptor)
    temporary = Path(name)
    try:
        yield temporary
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def save_workbook_atomic(workbook, destination):
    with atomic_output_path(destination) as temporary:
        workbook.save(temporary)


@contextmanager
def excel_writer(destination, **options):
    """Save only on success and always release the owned file handle."""
    import pandas as pd

    mode = "r+b" if options.get("mode") == "a" else "w+b"
    with open(destination, mode) as stream:
        writer = pd.ExcelWriter(stream, engine="openpyxl", **options)
        try:
            yield writer
            writer.close()
        finally:
            writer.book.close()


@contextmanager
def costs_excel_writer(destination):
    """Replace selected sheets while preserving other sheets and the old file on failure."""
    import pandas as pd

    destination = Path(destination)
    with atomic_output_path(destination) as temporary:
        exists = destination.exists()
        if exists:
            shutil.copyfile(destination, temporary)
        options = {"mode": "a", "if_sheet_exists": "replace"} if exists else {"mode": "w"}
        with excel_writer(temporary, **options) as writer:
            # Migrate the legacy first sheet instead of leaving a stale copy ahead of Основной.
            if "Основной" not in writer.book.sheetnames and writer.book.worksheets:
                writer.book.worksheets[0].title = "Основной"
            yield writer


def write_costs_dataframe(dataframe, destination):
    with costs_excel_writer(destination) as writer:
        dataframe.to_excel(writer, sheet_name="Основной", index=False)


def read_costs_dataframe(source):
    """Use the same costs sheet in calculations and the editor, including legacy workbooks."""
    import pandas as pd

    with pd.ExcelFile(source) as workbook:
        sheet_name = "Основной" if "Основной" in workbook.sheet_names else workbook.sheet_names[0]
        return workbook.parse(sheet_name)
