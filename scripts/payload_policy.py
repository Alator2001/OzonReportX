# -*- coding: utf-8 -*-
import json
from typing import Any, Dict, Optional

from analytics_models import WorkflowState
from analytics_reducers import (
    build_artikul_drilldown_payload_core,
    build_month_summary_payload_core,
    build_top_orders_payload_core,
    build_top_profit_payload_core,
    build_yearly_artikul_profit_summary_core,
    compact_results_for_model_core,
    is_yearly_artikul_report_request_core,
)


def build_model_payload_core(
    results: Dict[str, Any],
    user_message: Optional[str],
    workflow_state: Optional[WorkflowState],
    current_year: int,
) -> Dict[str, Any]:
    if workflow_state and workflow_state.force_compact_payload:
        if workflow_state.name == "single_month_summary":
            compact_summary = build_month_summary_payload_core(results)
            if compact_summary:
                return compact_summary
        if workflow_state.name == "single_artikul_drilldown":
            compact_summary = build_artikul_drilldown_payload_core(results)
            if compact_summary:
                return compact_summary
        if workflow_state.name == "top_profit":
            compact_summary = build_top_profit_payload_core(results)
            if compact_summary:
                return compact_summary
        if workflow_state.name == "top_orders":
            compact_summary = build_top_orders_payload_core(results)
            if compact_summary:
                return compact_summary
        compact_summary = build_yearly_artikul_profit_summary_core(results, user_message, current_year)
        if compact_summary:
            return compact_summary

    if is_yearly_artikul_report_request_core(user_message, results, current_year):
        compact_summary = build_yearly_artikul_profit_summary_core(results, user_message, current_year)
        if compact_summary:
            return compact_summary

    return compact_results_for_model_core(results)


def format_model_payload_core(
    results: Dict[str, Any],
    user_message: Optional[str],
    workflow_state: Optional[WorkflowState],
    current_year: int,
) -> str:
    payload = build_model_payload_core(results, user_message, workflow_state, current_year)
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
