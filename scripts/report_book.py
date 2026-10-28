"""
Где в месячном отчёте Ozon лежат данные — в отчёте старого и нового вида.

Старый вид (до 3.5.0): данные на видимых листах «Заказы» (заказы A–M, показатели P:Q, сверка выплат S:Y),
«Кампании» и «Расходы вне заказов».
Новый вид: те же листы с той же раскладкой, но скрытые и переименованные («Данные», «Данные кампаний»,
«Данные расходов»), а видимые «Сводка», «Заказы», «Реклама», «Расходы вне заказов» — оформление для человека
(scripts/report_design.py), которое программа не читает.

Все, кто читает отчёт, берут лист через data_sheet / data_sheet_name: есть лист нового вида — он, иначе старый.
Отчёты WB устроены так же для «Заказов»: прежний лист с показателями — скрытый «Данные», видимые «Сводка»
и «Заказы» строит scripts/wb_report_design.py. Отчёт WB старого вида без «Данных» читается по «Заказам», как раньше.
"""

from typing import Iterable, Optional

ORDERS = "Заказы"
CAMPAIGNS = "Кампании"
CHARGES = "Расходы вне заказов"
DATA_SHEETS = {ORDERS: "Данные", CAMPAIGNS: "Данные кампаний", CHARGES: "Данные расходов"}


def data_sheet_name(sheetnames: Iterable[str], sheet: str = ORDERS) -> Optional[str]:
    """Имя листа с данными: «Данные…» в отчёте нового вида, прежнее имя в старом; None — листа нет."""
    names = list(sheetnames)
    for name in (DATA_SHEETS[sheet], sheet):
        if name in names:
            return name
    return None


def data_sheet(workbook, sheet: str = ORDERS):
    """Лист с данными или None. Для «Заказов» без такого листа — первый лист, как раньше читали отчёты."""
    name = data_sheet_name(workbook.sheetnames, sheet)
    if name:
        return workbook[name]
    return workbook.active if sheet == ORDERS else None


def is_designed(sheetnames: Iterable[str]) -> bool:
    """Отчёт нового вида: данные уже на скрытом листе «Данные»."""
    return DATA_SHEETS[ORDERS] in list(sheetnames)
