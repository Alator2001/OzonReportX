# merge_excel_columns.py
import os
import sys
import argparse
import warnings
from typing import Dict, List, Optional

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)

# Канонические названия столбцов
CANON = {
    "artikul": "Артикул",
    "tsena_prodazhi": "Цена продажи",
    "kolichestvo_sht": "Количество шт.",
    "pribyl": "Прибыль",
    "data_otgruzki": "Дата отгрузки",
}

# Возможные варианты написания (регистр/пробелы/точки/синонимы)
ALIASES: Dict[str, List[str]] = {
    "artikul": ["артикул", "артикулы", "sku", "код", "код товара", "код/артикул"],
    "tsena_prodazhi": ["цена продажи", "цена", "продажная цена", "стоимость", "sale price"],
    "kolichestvo_sht": ["количество шт.", "кол-во", "количество", "шт", "шт.", "кол-во шт.", "qty"],
    "pribyl": ["прибыль", "маржа", "доход", "profit"],
    "data_otgruzki": ["дата отгрузки", "отгрузка", "дата поставки", "ship date", "дата"],
}

# Расширения Excel
EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}

def norm(s: str) -> str:
    """Нормализация заголовка: нижний регистр, убираем пробелы по краям и двойные пробелы, точки."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = " ".join(s.split())       # сжать повторные пробелы
    s = s.replace(".", "")        # убрать точки (часто пишут "шт.")
    return s

def build_reverse_map() -> Dict[str, str]:
    """Карта из нормализованного псевдонима к ключу CANON."""
    rmap = {}
    for key, variants in ALIASES.items():
        for v in variants:
            rmap[norm(v)] = key
    # добавим сами канонические названия
    for key, title in CANON.items():
        rmap[norm(title)] = key
    return rmap

REV = build_reverse_map()

def find_columns(df: pd.DataFrame) -> Dict[str, str]:
    """
    Возвращает сопоставление {канонический_ключ -> реальное_имя_столбца_в_df}
    Например: {"artikul": "Артикул", ...}
    """
    mapping = {}
    # Если вдруг MultiIndex в колонках — сплющим
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join([str(x) for x in tup if pd.notna(x)]).strip() for tup in df.columns]

    for col in df.columns:
        k = REV.get(norm(col))
        if k and k not in mapping:
            mapping[k] = col
    return mapping

def read_all_sheets(path: str) -> Dict[str, pd.DataFrame]:
    """Читает ВСЕ листы книги в dict {sheet_name: df}. Для .xls нужна библиотека xlrd."""
    try:
        xls = pd.ExcelFile(path, engine=None)  # pandas сам подберёт движок (openpyxl/xlrd)
        dfs = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
        return dfs
    except Exception as e:
        raise RuntimeError(f"Не удалось открыть файл '{path}': {e}")

def merge_folder(input_dir: str, output_path: str) -> None:
    rows = []
    report_missing = []   # листы, где нет всех нужных колонок
    report_partial = []   # листы, где нашли часть колонок

    files = sorted(
        f for f in os.listdir(input_dir)
        if os.path.isfile(os.path.join(input_dir, f))
        and os.path.splitext(f)[1].lower() in EXCEL_EXT
        and not f.startswith("~$")
    )

    if not files:
        print("В папке не найдено Excel-файлов.")
        return

    for fname in files:
        fpath = os.path.join(input_dir, fname)
        try:
            sheets = read_all_sheets(fpath)
        except Exception as e:
            print(e)
            continue

        for sheet_name, df in sheets.items():
            if df is None or df.empty:
                report_missing.append((fname, sheet_name, "лист пустой"))
                continue

            col_map = find_columns(df)
            found_keys = set(col_map.keys())
            required_keys = set(CANON.keys())

            if not found_keys:
                report_missing.append((fname, sheet_name, "ни один столбец не найден"))
                continue

            if found_keys != required_keys:
                missing = required_keys - found_keys
                # Если вообще ничего не найдено — уже учтено выше; здесь частичное совпадение
                if missing:
                    report_partial.append((fname, sheet_name, f"нет столбцов: {', '.join(CANON[k] for k in missing)}"))

            # Берём только найденные столбцы, переименовываем в канон
            use_cols = {col_map[k]: CANON[k] for k in found_keys}
            sub = df[list(use_cols.keys())].rename(columns=use_cols)

            # Приведём типы слегка (опционально)
            # Даты
            if "Дата отгрузки" in sub.columns:
                sub["Дата отгрузки"] = pd.to_datetime(sub["Дата отгрузки"], errors="coerce").dt.date
            # Числа
            for num_col in ["Цена продажи", "Количество шт.", "Прибыль"]:
                if num_col in sub.columns:
                    sub[num_col] = pd.to_numeric(sub[num_col], errors="coerce")

            # Добавим источник
            sub["Источник файл"] = fname
            sub["Лист"] = sheet_name

            rows.append(sub)

    if not rows:
        print("Нечего объединять — нужные столбцы не найдены ни в одном листе.")
        if report_missing:
            print("\nОтчёт по пропускам:")
            for f, s, msg in report_missing:
                print(f"- {f} / {s}: {msg}")
        return

    merged = pd.concat(rows, ignore_index=True, sort=False)

    # Обеспечим порядок колонок: сначала канонические, затем служебные, затем возможные другие (на всякий)
    ordered = [CANON["artikul"], CANON["tsena_prodazhi"], CANON["kolichestvo_sht"],
               CANON["pribyl"], CANON["data_otgruzki"], "Источник файл", "Лист"]
    other_cols = [c for c in merged.columns if c not in ordered]
    merged = merged[ [c for c in ordered if c in merged.columns] + other_cols ]

    # Сохраняем
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Объединено", index=False)

        # Отчёт на второй лист
        report_rows = []
        for f, s, msg in report_partial:
            report_rows.append({"Файл": f, "Лист": s, "Комментарий": msg, "Тип": "Частично найдено"})
        for f, s, msg in report_missing:
            report_rows.append({"Файл": f, "Лист": s, "Комментарий": msg, "Тип": "Не найдено"})
        if report_rows:
            rep_df = pd.DataFrame(report_rows)
            rep_df.to_excel(writer, sheet_name="Отчёт", index=False)

    print(f"Готово. Сохранено: {output_path}")

def main(argv: Optional[List[str]] = None):
    parser = argparse.ArgumentParser(description="Собрать нужные столбцы из всех Excel-файлов папки в один файл.")
    parser.add_argument("-i", "--input_dir", default="./reports", help="Путь к папке с входными Excel отчётами (по умолчанию merged.xlsx)")
    parser.add_argument("-o", "--output", default="merged.xlsx", help="Путь к выходному Excel (по умолчанию merged.xlsx)")
    args = parser.parse_args(argv)

    if not os.path.isdir(args.input_dir):
        print(f"Папка не найдена: {args.input_dir}")
        sys.exit(1)

    merge_folder(args.input_dir, args.output)

if __name__ == "__main__":
    main()
