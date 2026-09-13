# -*- coding: utf-8 -*-
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class ArtikulAggregateState:
    artikul: str
    product_names: List[str] = field(default_factory=list)
    periods: List[str] = field(default_factory=list)
    order_rows: int = 0
    finalized_rows: int = 0
    delivered_rows: int = 0
    cancelled_rows: int = 0
    returned_rows: int = 0
    in_progress_rows: int = 0
    total_quantity: float = 0.0
    total_revenue: float = 0.0
    total_profit: float = 0.0
    total_cost: float = 0.0
    avg_profit_per_finalized_row: float = 0.0


@dataclass
class YearlyArtikulSummaryState:
    year: Optional[int]
    requested_scope: str
    months_included: List[str] = field(default_factory=list)
    months_skipped: List[Dict[str, str]] = field(default_factory=list)
    artikuls: List[ArtikulAggregateState] = field(default_factory=list)
    top_profit: List[Dict[str, Any]] = field(default_factory=list)
    top_orders: List[Dict[str, Any]] = field(default_factory=list)
    totals: Dict[str, Any] = field(default_factory=dict)
    notes: List[str] = field(default_factory=list)


@dataclass
class WorkflowState:
    name: str
    requested_year: Optional[int] = None
    stage: str = "planning"
    force_compact_payload: bool = False
    force_final_answer: bool = False
