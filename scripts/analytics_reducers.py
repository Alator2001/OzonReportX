# -*- coding: utf-8 -*-
import re
from typing import Any, Dict, List, Optional

import pandas as pd

from analytics_models import ArtikulAggregateState, YearlyArtikulSummaryState


def to_float_or_none(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip().replace(",", ".")
        if not text or text == "-":
            return None
        value = text
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return None
    if pd.isna(numeric):
        return None
    return float(numeric)


def normalize_status_name(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def normalize_artikul_identifier(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def extract_year_from_text(text: str, current_year: int) -> Optional[int]:
    if not text:
        return None
    match = re.search(r"\b(20\d{2})\b", text)
    if match:
        return int(match.group(1))
    if "этот год" in text.lower():
        return current_year
    return None


def is_yearly_artikul_report_request_core(
    user_message: Optional[str],
    results: Dict[str, Any],
    current_year: int,
) -> bool:
    month_reports = results.get("get_month_report_full_data")
    if not isinstance(month_reports, list) or len(month_reports) < 2:
        return False
    if user_message:
        text = user_message.lower()
        year_markers = ("за этот год", "за год", "годовой", "по году")
        artikul_markers = ("артикул", "артикулов", "товар", "товаров")
        metric_markers = ("прибыл", "продаж", "выручк", "заказ")
        if (
            any(marker in text for marker in year_markers)
            and any(marker in text for marker in artikul_markers)
            and any(marker in text for marker in metric_markers)
        ):
            return True
        return False

    periods: List[str] = []
    years = set()
    for report in month_reports:
        if not isinstance(report, dict):
            continue
        period = str(report.get("Период отчёта") or "").strip()
        if period:
            periods.append(period)
            year = extract_year_from_text(period, current_year)
            if year is not None:
                years.add(year)

    return len(periods) >= 3 and len(years) <= 1


def build_yearly_artikul_profit_summary_core(
    results: Dict[str, Any],
    user_message: Optional[str],
    current_year: int,
) -> Optional[Dict[str, Any]]:
    month_reports = results.get("get_month_report_full_data")
    if not isinstance(month_reports, list) or not month_reports:
        return None

    requested_year = extract_year_from_text(user_message or "", current_year)
    artikul_map: Dict[str, ArtikulAggregateState] = {}
    months_included: List[str] = []
    months_skipped: List[Dict[str, str]] = []
    total_rows = 0
    finalized_rows = 0
    in_progress_rows = 0
    totals_profit = 0.0
    totals_revenue = 0.0
    total_status_counts: Dict[str, int] = {}

    for report in month_reports:
        if not isinstance(report, dict):
            continue
        period = str(report.get("Период отчёта") or "").strip()
        period_year = extract_year_from_text(period, current_year)
        if requested_year is not None and period_year is not None and period_year != requested_year:
            months_skipped.append({"period": period, "reason": "Период вне запрошенного года"})
            continue

        rows = report.get("Строки отчёта")
        if not isinstance(rows, list):
            months_skipped.append({"period": period or "Неизвестный период", "reason": "Нет строк отчёта"})
            continue

        month_had_finalized_rows = False
        for row in rows:
            if not isinstance(row, dict):
                continue

            artikul = normalize_artikul_identifier(row.get("Артикул"))
            if not artikul or artikul.lower() == "nan":
                continue

            status = normalize_status_name(row.get("Статус"))
            total_rows += 1
            if status:
                total_status_counts[status] = total_status_counts.get(status, 0) + 1

            entry = artikul_map.get(artikul)
            if entry is None:
                entry = ArtikulAggregateState(artikul=artikul)
                artikul_map[artikul] = entry

            entry.order_rows += 1
            if period and period not in entry.periods:
                entry.periods.append(period)

            product_name = str(row.get("Название товара") or "").strip()
            if product_name and product_name not in entry.product_names:
                entry.product_names.append(product_name)

            if status == "delivered":
                entry.delivered_rows += 1
            elif status == "cancelled":
                entry.cancelled_rows += 1
            elif status == "returned":
                entry.returned_rows += 1
            elif status in {"delivering", "awaiting_deliver", "awaiting_packaging", "processing", "pending"}:
                entry.in_progress_rows += 1
                in_progress_rows += 1

            profit = to_float_or_none(row.get("Прибыль"))
            if profit is None:
                continue

            month_had_finalized_rows = True
            entry.finalized_rows += 1
            finalized_rows += 1
            entry.total_profit += profit
            totals_profit += profit

            quantity = to_float_or_none(row.get("Количество шт."))
            if quantity is not None:
                entry.total_quantity += quantity

            revenue = to_float_or_none(row.get("Цена продажи"))
            if revenue is not None:
                entry.total_revenue += revenue
                totals_revenue += revenue

            cost = to_float_or_none(row.get("Себестоимость"))
            if cost is not None:
                entry.total_cost += cost

        if month_had_finalized_rows:
            if period:
                months_included.append(period)
        else:
            months_skipped.append(
                {"period": period or "Неизвестный период", "reason": "Нет завершённых финансовых строк, пригодных для агрегации"}
            )

    if not artikul_map:
        return None

    artikuls = sorted(artikul_map.values(), key=lambda item: item.total_profit, reverse=True)
    for item in artikuls:
        if item.finalized_rows > 0:
            item.avg_profit_per_finalized_row = item.total_profit / item.finalized_rows
        item.periods.sort()

    inferred_year = requested_year
    if inferred_year is None:
        derived_years = {
            extract_year_from_text(period, current_year)
            for period in months_included
            if extract_year_from_text(period, current_year) is not None
        }
        if len(derived_years) == 1:
            inferred_year = next(iter(derived_years))

    summary = YearlyArtikulSummaryState(
        year=inferred_year,
        requested_scope="yearly_artikul_profit_report",
        months_included=months_included,
        months_skipped=months_skipped,
        artikuls=artikuls,
        top_profit=[
            {
                "artikul": item.artikul,
                "profit": round(item.total_profit, 2),
                "finalized_rows": item.finalized_rows,
                "product_name": item.product_names[0] if item.product_names else "",
            }
            for item in artikuls[:10]
        ],
        top_orders=[
            {
                "artikul": item.artikul,
                "order_rows": item.order_rows,
                "profit": round(item.total_profit, 2),
                "product_name": item.product_names[0] if item.product_names else "",
            }
            for item in sorted(artikuls, key=lambda item: item.order_rows, reverse=True)[:10]
        ],
        totals={
            "artikuls_count": len(artikuls),
            "total_rows_seen": total_rows,
            "finalized_rows": finalized_rows,
            "in_progress_rows": in_progress_rows,
            "total_profit": round(totals_profit, 2),
            "total_revenue": round(totals_revenue, 2),
            "status_counts": total_status_counts,
        },
        notes=[],
    )

    if months_skipped:
        skipped_periods = [item["period"] for item in months_skipped if item.get("period")]
        if skipped_periods:
            summary.notes.append("Из финансовой агрегации исключены периоды без завершённых строк: " + ", ".join(skipped_periods))

    return {
        "yearly_artikul_profit_summary": {
            "year": summary.year,
            "requested_scope": summary.requested_scope,
            "months_included": summary.months_included,
            "months_skipped": summary.months_skipped,
            "totals": summary.totals,
            "top_profit": summary.top_profit,
            "top_orders": summary.top_orders,
            "artikuls": [
                {
                    "artikul": item.artikul,
                    "product_name": item.product_names[0] if item.product_names else "",
                    "periods": item.periods,
                    "order_rows": item.order_rows,
                    "finalized_rows": item.finalized_rows,
                    "delivered_rows": item.delivered_rows,
                    "cancelled_rows": item.cancelled_rows,
                    "returned_rows": item.returned_rows,
                    "in_progress_rows": item.in_progress_rows,
                    "total_quantity": round(item.total_quantity, 2),
                    "total_revenue": round(item.total_revenue, 2),
                    "total_profit": round(item.total_profit, 2),
                    "total_cost": round(item.total_cost, 2),
                    "avg_profit_per_finalized_row": round(item.avg_profit_per_finalized_row, 2),
                }
                for item in artikuls
            ],
            "notes": summary.notes,
        }
    }


def build_workflow_fallback_answer_core(
    user_message: Optional[str],
    results: Dict[str, Any],
    current_year: int,
) -> Optional[str]:
    text = (user_message or "").lower()
    if "артикул" in text and "get_artikul_stats" in results:
        return build_artikul_drilldown_fallback_answer_core(results)
    if "get_top_profit" in results:
        return build_top_profit_fallback_answer_core(results)
    if "get_top_orders" in results:
        return build_top_orders_fallback_answer_core(results)
    if "get_month_summary" in results:
        return build_month_summary_fallback_answer_core(results)

    compact_summary = build_yearly_artikul_profit_summary_core(results, user_message, current_year)
    if not compact_summary:
        return None

    summary = compact_summary.get("yearly_artikul_profit_summary", {})
    if not isinstance(summary, dict):
        return None

    year = summary.get("year") or "текущий год"
    months_included = summary.get("months_included") or []
    totals = summary.get("totals") or {}
    top_profit = summary.get("top_profit") or []
    top_orders = summary.get("top_orders") or []
    notes = summary.get("notes") or []

    lines: List[str] = []
    included_text = ", ".join(months_included) if months_included else "доступные месяцы"
    lines.append(f"Отчёт по прибыли артикулов за {year} год построен по периодам: {included_text}.")

    total_profit = totals.get("total_profit")
    artikuls_count = totals.get("artikuls_count")
    finalized_rows = totals.get("finalized_rows")
    if total_profit is not None and artikuls_count is not None and finalized_rows is not None:
        lines.append(
            f"Итоговая прибыль: {float(total_profit):,.2f} ₽. В расчёт вошло {int(finalized_rows)} завершённых строк по {int(artikuls_count)} артикулам."
            .replace(",", " ")
        )

    if top_profit:
        leaders = []
        for item in top_profit[:5]:
            if isinstance(item, dict) and item.get("artikul") is not None and item.get("profit") is not None:
                leaders.append(f"{item['artikul']} — {float(item['profit']):,.2f} ₽".replace(",", " "))
        if leaders:
            lines.append("Топ по прибыли: " + "; ".join(leaders) + ".")

    if top_orders:
        volume_leaders = []
        for item in top_orders[:5]:
            if isinstance(item, dict) and item.get("artikul") is not None and item.get("order_rows") is not None:
                volume_leaders.append(f"{item['artikul']} — {int(item['order_rows'])} строк")
        if volume_leaders:
            lines.append("Лидеры по объёму: " + "; ".join(volume_leaders) + ".")

    if notes:
        lines.append("Примечание: " + " ".join(str(note) for note in notes if note))

    return " ".join(lines).strip()


def build_month_summary_payload_core(results: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    summary = results.get("get_month_summary")
    if not isinstance(summary, dict) or summary.get("error"):
        return None

    period = summary.get("Период отчёта")
    payload = {
        "month_summary_payload": {
            "period": period,
            "financials": {
                "revenue": summary.get("Общая выручка"),
                "net_profit": summary.get("Чистая прибыль"),
                "cost_of_goods": summary.get("Итоговая себестоимость"),
                "gross_profit": summary.get("COGS (валовая прибыль)"),
                "operating_expenses": summary.get("Операционные расходы"),
            },
            "orders": {
                "total": summary.get("Общее количество заказов"),
                "delivered": summary.get("Количество доставленных заказов"),
                "cancelled": summary.get("Количество отменённых заказов"),
                "returned": summary.get("Количество возвращённых заказов"),
                "delivering": summary.get("Количество заказов в доставке"),
                "average_check": summary.get("Средний чек"),
            },
            "margins": {
                "net_margin_pct": summary.get("Рентабельность по чистой прибыли (Net Profit Margin) %"),
                "gross_margin_pct": summary.get("Gross Profit Margin Рентабельность по валовой прибыли %"),
                "commission_pct": summary.get("Комиссии Ozon %"),
                "logistics_pct": summary.get("Логистика %"),
            },
            "costs": {
                "promotion_ozon": summary.get("Продвижение Ozon"),
                "star_goods": summary.get("Звёздные товары"),
                "external_marketing": summary.get("Внешний маркетинг"),
                "commission_amount": summary.get("Комиссии Ozon сумма"),
                "logistics_amount": summary.get("Логистика сумма"),
                "fbo_storage": summary.get("Расход хранения FBO"),
            },
        }
    }
    return payload


def build_artikul_drilldown_payload_core(results: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    payload = results.get("get_artikul_stats")
    if not isinstance(payload, dict) or payload.get("error"):
        return None
    stats = payload.get("stats")
    if not isinstance(stats, dict):
        return None

    compact = {
        "artikul_drilldown_payload": {
            "period": payload.get("period"),
            "artikul": payload.get("artikul"),
            "orders_count": stats.get("Количество заказов"),
            "profit_sum": stats.get("Сумма Прибыль"),
            "profit_avg": stats.get("Средняя Прибыль"),
            "revenue_sum": stats.get("Сумма Цена продажи"),
            "quantity_sum": stats.get("Сумма Количество шт."),
            "status_distribution": stats.get("Распределение по статусам"),
            "schema_distribution": stats.get("Распределение по схемам"),
            "sample_rows": stats.get("Строки отчёта", [])[:5] if isinstance(stats.get("Строки отчёта"), list) else [],
        }
    }
    return compact


def build_month_summary_fallback_answer_core(results: Dict[str, Any]) -> Optional[str]:
    payload = build_month_summary_payload_core(results)
    if not payload:
        return None
    summary = payload["month_summary_payload"]
    financials = summary.get("financials", {})
    orders = summary.get("orders", {})
    margins = summary.get("margins", {})

    period = summary.get("period") or "выбранный период"
    parts = [f"Сводка за {period}."]
    if financials.get("revenue") is not None and financials.get("net_profit") is not None:
        parts.append(
            f"Выручка: {float(financials['revenue']):,.2f} ₽, чистая прибыль: {float(financials['net_profit']):,.2f} ₽."
            .replace(",", " ")
        )
    if orders.get("total") is not None and orders.get("delivered") is not None:
        parts.append(
            f"Заказов: {int(orders['total'])}, доставлено: {int(orders['delivered'])}, отменено: {int(orders.get('cancelled') or 0)}."
        )
    if margins.get("net_margin_pct") is not None:
        parts.append(f"Рентабельность по чистой прибыли: {float(margins['net_margin_pct']):.2f}%.")
    return " ".join(parts)


def build_artikul_drilldown_fallback_answer_core(results: Dict[str, Any]) -> Optional[str]:
    payload = build_artikul_drilldown_payload_core(results)
    if not payload:
        return None
    info = payload["artikul_drilldown_payload"]
    parts = [f"Артикул {info.get('artikul')} за период {info.get('period')}."]
    if info.get("orders_count") is not None:
        parts.append(f"Строк отчёта: {int(info['orders_count'])}.")
    if info.get("profit_sum") is not None:
        parts.append(f"Суммарная прибыль: {float(info['profit_sum']):,.2f} ₽.".replace(",", " "))
    if info.get("profit_avg") is not None:
        parts.append(f"Средняя прибыль на строку: {float(info['profit_avg']):,.2f} ₽.".replace(",", " "))
    status_distribution = info.get("status_distribution")
    if isinstance(status_distribution, dict) and status_distribution:
        status_text = ", ".join(f"{key}: {value}" for key, value in status_distribution.items())
        parts.append(f"Статусы: {status_text}.")
    return " ".join(parts)


def build_top_profit_payload_core(results: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    payload = results.get("get_top_profit")
    if not isinstance(payload, dict) or payload.get("error"):
        return None
    top_profit = payload.get("top_profit")
    if not isinstance(top_profit, dict):
        return None
    ranking = [
        {"artikul": str(artikul), "profit": float(profit)}
        for artikul, profit in top_profit.items()
    ]
    return {
        "top_profit_payload": {
            "period": payload.get("period"),
            "ranking": ranking,
        }
    }


def build_top_orders_payload_core(results: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    payload = results.get("get_top_orders")
    if not isinstance(payload, dict) or payload.get("error"):
        return None
    top_orders = payload.get("top_orders")
    if not isinstance(top_orders, dict):
        return None
    ranking = [
        {"artikul": str(artikul), "orders": int(count)}
        for artikul, count in top_orders.items()
    ]
    return {
        "top_orders_payload": {
            "period": payload.get("period"),
            "ranking": ranking,
        }
    }


def build_top_profit_fallback_answer_core(results: Dict[str, Any]) -> Optional[str]:
    payload = build_top_profit_payload_core(results)
    if not payload:
        return None
    info = payload["top_profit_payload"]
    ranking = info.get("ranking") or []
    period = info.get("period") or "выбранный период"
    if not ranking:
        return None
    leaders = "; ".join(
        f"{item['artikul']} — {item['profit']:,.2f} ₽".replace(",", " ")
        for item in ranking[:5]
    )
    return f"Топ артикулов по прибыли за {period}: {leaders}."


def build_top_orders_fallback_answer_core(results: Dict[str, Any]) -> Optional[str]:
    payload = build_top_orders_payload_core(results)
    if not payload:
        return None
    info = payload["top_orders_payload"]
    ranking = info.get("ranking") or []
    period = info.get("period") or "выбранный период"
    if not ranking:
        return None
    leaders = "; ".join(f"{item['artikul']} — {item['orders']} заказов" for item in ranking[:5])
    return f"Топ артикулов по количеству заказов за {period}: {leaders}."


def compact_results_for_model_core(
    results: Dict[str, Any],
    max_chars: int = 50000,
) -> Dict[str, Any]:
    try:
        import json
        raw_text = json.dumps(results, ensure_ascii=False, separators=(",", ":"))
        if len(raw_text) <= max_chars:
            return results
    except Exception:
        return results

    compacted: Dict[str, Any] = {}
    for key, value in results.items():
        if key != "get_month_report_full_data":
            compacted[key] = value
            continue

        reports = value if isinstance(value, list) else [value]
        compact_reports: List[Dict[str, Any]] = []
        for report in reports:
            if not isinstance(report, dict):
                compact_reports.append(report)
                continue

            compact_report: Dict[str, Any] = {}
            for field in (
                "Период отчёта",
                "Всего строк отчёта",
                "Сводка отчёта",
                "Общее распределение по статусам",
                "Общее распределение по схемам",
                "Топ-10 артикулов по прибыли",
                "Топ-10 артикулов по количеству заказов",
            ):
                if field in report:
                    compact_report[field] = report[field]

            artikul_stats = report.get("Статистика по артикулам")
            if isinstance(artikul_stats, list):
                compact_stats: List[Dict[str, Any]] = []
                for item in artikul_stats[:15]:
                    if not isinstance(item, dict):
                        continue
                    compact_item = {
                        "Артикул": item.get("Артикул"),
                        "Количество заказов": item.get("Количество заказов"),
                        "Сумма Прибыль": item.get("Сумма Прибыль"),
                        "Средняя Прибыль": item.get("Средняя Прибыль"),
                        "Распределение по статусам": item.get("Распределение по статусам"),
                        "Распределение по схемам": item.get("Распределение по схемам"),
                    }
                    compact_stats.append(compact_item)
                compact_report["Статистика по артикулам"] = compact_stats

            compact_reports.append(compact_report)

        compacted[key] = compact_reports if isinstance(value, list) else compact_reports[0]

    return compacted
