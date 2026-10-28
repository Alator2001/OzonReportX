"""
Сверка выплат WB: сколько денег реально ушло вам на счёт и какие месяцы этим закрыты — как сверка выплат Ozon
(payout_settlement.py).

Списка выводов денег WB в API не отдаёт. Но всё, что WB начислил по финансовым отчётам с начала магазина,
либо лежит на балансе, либо уже выведено, поэтому:

    выведено = Σ «Итого к оплате» всех ежедневных отчётов с начала магазина − текущий баланс WB.

Проверено на данных магазина 27.10.2026: других списаний с баланса (реклама с баланса и т. п.) мимо отчётов нет —
удержания (штрафы, плата за досрочный вывод) уже внутри «Итого к оплате». Если такие списания появятся,
выведенная сумма окажется завышенной.

Начисления каждого месяца — «долг» WB: сначала себестоимость, затем прибыль. Выведенные деньги гасят его
по порядку, с самого старого месяца (FIFO). Месяц с отрицательным итогом (только хранение и т. п.) долга не создаёт,
а уменьшает то, что WB должен по прошлым месяцам. Себестоимость известна только для месяцев, загруженных в
«Бизнес-сводку» (wb reports/<ГГГГ-ММ>.json) целиком; для остальных месяц гасится одной суммой «Начислено».

Итоги отчётов за всю историю и баланс сохраняются в wb reports/settlement.json (refresh, два запроса к WB),
распределение считается заново при каждом построении отчёта (settle, без запросов).
"""
import json
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from scripts import wb_finance as finance
from scripts.file_io import atomic_output_path
from scripts.marketplace_settings import read_settings

SETTLEMENT_FILE = Path("wb reports") / "settlement.json"
HISTORY_START = "2024-01-29"  # раньше WB финансовые отчёты по API не отдаёт
COST, PROFIT, ACCRUED = "Себестоимость", "Прибыль", "Начислено"
STALE_TOLERANCE = 100  # копеек: расхождение итога месяца в сохранённом отчёте и в списке отчётов WB


def _kopecks(value) -> int:
    return int(round(finance.number(value or 0) * 100))


def _path(root) -> Path:
    return Path(root) / SETTLEMENT_FILE


def refresh(root, client, today=None) -> dict:
    """Загрузить итоги всех отчётов WB и баланс и сохранить. client — wb_finance.FinanceClient."""
    today = today or date.today()
    reports = client.list_reports(HISTORY_START, today.isoformat())
    balance = client.balance()
    months = defaultdict(int)
    for report in reports:
        months[str(report.get("dateFrom") or "")[:7]] += _kopecks(report.get("bankPaymentSum"))
    data = {
        "version": 1, "account": finance.identity(client.token),
        "checked_at": datetime.now().isoformat(timespec="seconds"),
        "balance": _kopecks(balance["current"]), "for_withdraw": _kopecks(balance["for_withdraw"]),
        "months": {key: value for key, value in sorted(months.items()) if key},
    }
    path = _path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    with atomic_output_path(path) as temporary:
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    return data


def load(root):
    """Сохранённая сверка этого же токена WB или None."""
    path = _path(root)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    token = read_settings(root).get("WB_API_TOKEN") or ""
    if data.get("version") != 1 or data.get("account") != finance.identity(token):
        return None
    return data


def _month_cost(root, key, payout, costs):
    """Себестоимость месяца в копейках или (None, причина)."""
    year, month = int(key[:4]), int(key[5:7])
    data = finance.load_report(root, year, month)
    if not data or not data.get("reports"):
        return None, "месяц не загружен в «Бизнес-сводку» — себестоимость неизвестна"
    saved = sum(_kopecks(report.get("bankPaymentSum")) for report in data["reports"])
    if abs(saved - payout) > STALE_TOLERANCE:
        return None, "с момента загрузки месяца WB добавил отчёты — загрузите месяц заново в «Бизнес-сводке»"
    summary = finance.summarize(data, costs)
    if summary.get("costs") is None:
        return None, "не заполнена себестоимость или есть возвраты прошлых периодов"
    return int(round(summary["costs"] * 100)), None


def settle(root, costs, saved=None) -> dict | None:
    """Распределить выведенные деньги по месяцам. None — сверка ещё не проводилась."""
    saved = saved or load(root)
    if not saved:
        return None
    months, buckets, negatives = {}, [], 0
    for key, payout in sorted(saved["months"].items()):
        cost, reason = _month_cost(root, key, payout, costs)
        if payout < 0:
            negatives -= payout
            due = {}
        elif cost is None:
            due = {ACCRUED: payout}
        else:
            part = min(max(cost, 0), payout)  # себестоимость < 0 — в месяце одни возвраты прошлых продаж
            due = {COST: part, PROFIT: payout - part}
        months[key] = {"payout": payout, "cost": cost, "profit": payout - cost if cost is not None else None,
                       "due": due, "closed": {kind: 0 for kind in due}, "reason": reason}
        buckets += [[key, kind, amount] for kind, amount in due.items()]
    withdrawn = max(sum(saved["months"].values()) - saved["balance"], 0)
    left = withdrawn + negatives
    for key, kind, amount in buckets:
        part = min(left, amount)
        months[key]["closed"][kind] += part
        left -= part
    for month in months.values():
        month["left"] = sum(month["due"].values()) - sum(month["closed"].values())
        month["done"] = month["left"] == 0
    return {"checked_at": saved["checked_at"], "balance": saved["balance"], "for_withdraw": saved["for_withdraw"],
            "withdrawn": withdrawn, "over": left, "months": months}
