# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Optional, Tuple

from analytics_models import WorkflowState
from analytics_reducers import build_yearly_artikul_profit_summary_core, extract_year_from_text


YEAR_MARKERS = (
    "\u0437\u0430 \u044d\u0442\u043e\u0442 \u0433\u043e\u0434",
    "\u0437\u0430 \u0433\u043e\u0434",
    "\u0433\u043e\u0434\u043e\u0432\u043e\u0439",
    "\u043f\u043e \u0433\u043e\u0434\u0443",
)
ARTIKUL_MARKERS = (
    "\u0430\u0440\u0442\u0438\u043a\u0443\u043b",
    "\u0430\u0440\u0442\u0438\u043a\u0443\u043b\u043e\u0432",
    "\u0442\u043e\u0432\u0430\u0440",
    "\u0442\u043e\u0432\u0430\u0440\u043e\u0432",
)
METRIC_MARKERS = (
    "\u043f\u0440\u0438\u0431\u044b\u043b",
    "\u043f\u0440\u043e\u0434\u0430\u0436",
    "\u0432\u044b\u0440\u0443\u0447\u043a",
    "\u0437\u0430\u043a\u0430\u0437",
)
MONTH_SUMMARY_MARKERS = (
    "\u0447\u0442\u043e \u043f\u0440\u043e\u0438\u0437\u043e\u0448\u043b\u043e",
    "\u0441\u0432\u043e\u0434\u043a",
    "\u0438\u0442\u043e\u0433\u0438",
    "\u043e\u0431\u0437\u043e\u0440",
    "\u0437\u0430 \u043c\u0435\u0441\u044f\u0446",
    "\u0447\u0442\u043e \u0431\u044b\u043b\u043e",
    "\u0447\u0442\u043e \u0432",
)
TOP_PROFIT_MARKERS = (
    "\u0442\u043e\u043f \u043f\u043e \u043f\u0440\u0438\u0431\u044b\u043b\u0438",
    "\u0442\u043e\u043f \u0430\u0440\u0442\u0438\u043a\u0443\u043b\u043e\u0432 \u043f\u043e \u043f\u0440\u0438\u0431\u044b\u043b\u0438",
    "\u0442\u043e\u043f \u043f\u0440\u0438\u0431\u044b\u043b\u0438",
)
TOP_ORDER_MARKERS = (
    "\u0442\u043e\u043f \u043f\u043e \u0437\u0430\u043a\u0430\u0437",
    "\u0442\u043e\u043f \u0430\u0440\u0442\u0438\u043a\u0443\u043b\u043e\u0432 \u043f\u043e \u0437\u0430\u043a\u0430\u0437",
    "\u0442\u043e\u043f \u0437\u0430\u043a\u0430\u0437",
    "\u0442\u043e\u043f \u043f\u043e \u0437\u0430\u043a\u0430\u0437\u0430\u043c",
)
MONTH_NAME_STEMS = (
    "\u044f\u043d\u0432\u0430\u0440",
    "\u0444\u0435\u0432\u0440\u0430\u043b",
    "\u043c\u0430\u0440\u0442",
    "\u0430\u043f\u0440\u0435\u043b",
    "\u043c\u0430\u0439",
    "\u0438\u044e\u043d",
    "\u0438\u044e\u043b",
    "\u0430\u0432\u0433\u0443\u0441\u0442",
    "\u0441\u0435\u043d\u0442\u044f\u0431\u0440",
    "\u043e\u043a\u0442\u044f\u0431\u0440",
    "\u043d\u043e\u044f\u0431\u0440",
    "\u0434\u0435\u043a\u0430\u0431\u0440",
)


def detect_workflow_core(user_message: Optional[str], current_year: int) -> WorkflowState:
    text = (user_message or "").lower()
    requested_year = extract_year_from_text(user_message or "", current_year)

    if (
        any(marker in text for marker in YEAR_MARKERS)
        and any(marker in text for marker in ARTIKUL_MARKERS)
        and any(marker in text for marker in METRIC_MARKERS)
    ):
        return WorkflowState(name="yearly_artikul_profit_report", requested_year=requested_year)

    if any(marker in text for marker in TOP_PROFIT_MARKERS):
        return WorkflowState(name="top_profit", requested_year=requested_year)
    if any(marker in text for marker in TOP_ORDER_MARKERS):
        return WorkflowState(name="top_orders", requested_year=requested_year)
    if any(marker in text for marker in MONTH_SUMMARY_MARKERS) or (
        requested_year is not None
        and any(month in text for month in MONTH_NAME_STEMS)
        and "\u0430\u0440\u0442\u0438\u043a\u0443\u043b" not in text
        and "\u0442\u043e\u043f" not in text
    ):
        return WorkflowState(name="single_month_summary", requested_year=requested_year)
    if "\u0430\u0440\u0442\u0438\u043a\u0443\u043b" in text:
        return WorkflowState(name="single_artikul_drilldown", requested_year=requested_year)
    return WorkflowState(name="generic", requested_year=requested_year)


def extract_available_year_periods_from_list_reports_core(
    results: Dict[str, Any],
    year: Optional[int],
    months_ru: List[str],
    current_year: int,
) -> List[str]:
    list_reports_result = results.get("list_reports")
    if not isinstance(list_reports_result, dict):
        return []
    reports = list_reports_result.get("reports")
    if not isinstance(reports, list):
        return []

    periods_with_sort_keys: List[Tuple[Tuple[int, int], str]] = []
    for item in reports:
        if not isinstance(item, dict):
            continue
        period = str(item.get("period") or "").strip()
        if not period:
            continue
        period_year = extract_year_from_text(period, current_year)
        if year is not None and period_year != year:
            continue

        month_number = 0
        for idx, month_name in enumerate(months_ru, start=1):
            if period.lower().startswith(month_name.lower()):
                month_number = idx
                break
        if period_year is None:
            period_year = year or 0
        periods_with_sort_keys.append(((period_year, month_number), period))

    periods_with_sort_keys.sort()
    return [period for _, period in periods_with_sort_keys]


def build_workflow_state_core(
    user_message: Optional[str],
    results: Dict[str, Any],
    current_year: int,
    months_ru: List[str],
) -> WorkflowState:
    workflow = detect_workflow_core(user_message, current_year)

    if workflow.name == "yearly_artikul_profit_report":
        available_periods = extract_available_year_periods_from_list_reports_core(
            results, workflow.requested_year, months_ru, current_year
        )
        has_month_reports = isinstance(results.get("get_month_report_full_data"), list) and bool(
            results.get("get_month_report_full_data")
        )
        compact_summary_ready = (
            build_yearly_artikul_profit_summary_core(results, user_message, current_year) is not None
        )

        if not results:
            workflow.stage = "planning"
        elif available_periods and not has_month_reports:
            workflow.stage = "periods_resolved"
        elif has_month_reports and compact_summary_ready:
            workflow.stage = "aggregation_ready"
            workflow.force_compact_payload = True
            workflow.force_final_answer = True
        elif has_month_reports:
            workflow.stage = "month_reports_loaded"
            workflow.force_compact_payload = True
        else:
            workflow.stage = "planning"

    elif workflow.name in {"single_month_summary", "single_artikul_drilldown", "top_profit", "top_orders"}:
        if not results:
            workflow.stage = "planning"
        elif (
            "get_month_summary" in results
            or "get_artikul_stats" in results
            or "get_top_profit" in results
            or "get_top_orders" in results
        ):
            workflow.stage = "aggregation_ready"
            workflow.force_compact_payload = True
            workflow.force_final_answer = True

    return workflow


def get_workflow_followup_needs_core(
    user_message: Optional[str],
    results: Dict[str, Any],
    current_year: int,
    months_ru: List[str],
) -> List[Dict[str, Any]]:
    workflow = build_workflow_state_core(user_message, results, current_year, months_ru)
    if workflow.name != "yearly_artikul_profit_report" or workflow.stage != "periods_resolved":
        return []

    periods = extract_available_year_periods_from_list_reports_core(
        results, workflow.requested_year, months_ru, current_year
    )
    return [{"tool": "get_month_report_full_data", "args": {"period": period}} for period in periods]
