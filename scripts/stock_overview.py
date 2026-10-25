# -*- coding: utf-8 -*-
"""
Раздел «Остатки»: одна таблица на оба магазина — остатки на своём складе (FBS), на складах Ozon (FBO)
и Wildberries (FBW) и сколько нужно держать на каждом из них, чтобы хватило на ближайшие 30 дней.

«Нужно держать» считается ровно так же, как в планах поставок (те же функции), только без срока поставки:
окно — [сегодня, сегодня + 30 дней).
- FBO: fbo_supply_plan.build_plan — прогноз магазина (то же окно год назад × рост) × доля артикула × доля зоны
  склада; по каждому складу FBO цель = квантиль Пуассона (уровень сервиса); «нужно держать FBO» — сумма целей.
- FBW: wb_supply_plan.build_wb_plan — темп заказов WB за 90 дней × сезонность × доля артикула; цель FBW только
  для артикулов, которые продаются не меньше wb_fbw_min_monthly шт. в месяц.
- FBS: один свой склад на оба магазина. Спрос FBS = спрос Ozon, который не покроет FBO (регионы без склада FBO,
  мелкий спрос зон) + спрос WB, который не покроет FBW; нужно = квантиль Пуассона от суммы.
Только чтение: ничего не отправляется и отчёты не пишутся (кэши заказов — общие с планами поставок).
"""

import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

script_dir = Path(__file__).resolve().parent
code_root = script_dir.parent
for path in (str(code_root), str(script_dir)):
    if path not in sys.path:
        sys.path.insert(0, path)
from app_paths import DATA_ROOT  # noqa: E402
repo_root = DATA_ROOT

import fbo_supply_plan as ozon_plan  # noqa: E402
import wb_supply_plan as wb_plan  # noqa: E402

HORIZON_DAYS = 30


def horizon_settings(settings: Dict[str, Any]) -> Dict[str, Any]:
    """Настройки плана поставок без срока поставки: «сколько держать» на ближайшие 30 дней."""
    return {**settings, "lead_days": 0, "cover_days": HORIZON_DAYS}


def ozon_needs(plan: Dict[str, Any]) -> Dict[str, Dict[str, float]]:
    """Из плана Ozon: {артикул: {need_fbo, fbs_demand}}."""
    result: Dict[str, Dict[str, float]] = {}
    details = plan["details"]
    if len(details):
        for article, need in details.groupby("Артикул")["Цель"].sum().items():
            result.setdefault(str(article), {})["need_fbo"] = int(need)
    for _, row in plan["fbs"].iterrows():
        result.setdefault(str(row["Артикул"]), {})["fbs_demand"] = float(row["Спрос на FBS, шт."])
    return result


def wb_needs(plan: Dict[str, Any]) -> Dict[str, Dict[str, float]]:
    """Из плана WB: {артикул: {need_fbw, fbs_demand}}."""
    result: Dict[str, Dict[str, float]] = {}
    for _, row in plan["details"].iterrows():
        result[str(row["Артикул"])] = {"need_fbw": int(row["Цель FBW"]), "fbs_demand": float(row["Спрос на своём складе, шт."])}
    return result


def build_rows(ozon: Optional[Dict[str, Any]], wb: Optional[Dict[str, Any]], level: float) -> List[Dict[str, Any]]:
    """
    ozon / wb: {"needs": ozon_needs/wb_needs, "own": {артикул: FBS}, "stock": {артикул: на складе}, "transit": {артикул: в пути}}
    или None, если магазин не загрузился (его колонки будут пустыми).
    FBS — один склад: остаток — наибольший из показанных Ozon и WB (это один и тот же товар).
    """
    articles = set()
    for source in (ozon, wb):
        if source:
            articles |= set(source["needs"]) | set(source["own"]) | set(source["stock"]) | set(source["transit"])
    rows = []
    for article in sorted(articles, key=ozon_plan._article_sort_key):
        o_need = (ozon or {}).get("needs", {}).get(article, {})
        w_need = (wb or {}).get("needs", {}).get(article, {})
        fbs_by_store = {name: int(source["own"][article]) for name, source in (("Ozon", ozon), ("WB", wb))
                        if source and source["own"].get(article) is not None}
        fbs_demand = float(o_need.get("fbs_demand", 0.0)) + float(w_need.get("fbs_demand", 0.0))
        row = {
            "article": article,
            "fbs": max(fbs_by_store.values()) if fbs_by_store else 0,
            "fbs_by_store": fbs_by_store,
            "fbo": int(ozon["stock"].get(article, 0)) if ozon else None,
            "fbo_transit": int(ozon["transit"].get(article, 0)) if ozon else 0,
            "fbw": int(wb["stock"].get(article, 0)) if wb else None,
            "fbw_transit": int(wb["transit"].get(article, 0)) if wb else 0,
            "need_fbo": int(o_need.get("need_fbo", 0)) if ozon else None,
            "need_fbw": int(w_need.get("need_fbw", 0)) if wb else None,
            "fbs_demand": round(fbs_demand, 1),
            "need_fbs": ozon_plan.poisson_quantile(fbs_demand, level) if fbs_demand > 0 else 0,
        }
        row["short_fbo"] = max(0, (row["need_fbo"] or 0) - (row["fbo"] or 0) - row["fbo_transit"])
        row["short_fbw"] = max(0, (row["need_fbw"] or 0) - (row["fbw"] or 0) - row["fbw_transit"])
        row["short_fbs"] = max(0, row["need_fbs"] - row["fbs"])
        if not any(row[key] for key in ("fbs", "fbo", "fbo_transit", "fbw", "fbw_transit", "need_fbo", "need_fbw", "need_fbs")):
            continue
        rows.append(row)
    return rows


def forecast_note(store: str, units: float, last_30: float, method: str = "", last_year: float = 0.0, factor: float = 1.0) -> str:
    """Откуда прогноз: при сезонном методе он может сильно отличаться от продаж последнего месяца."""
    text = f"{store}: прогноз на 30 дней — {units:.0f} шт. (за последние 30 дней {last_30:.0f})"
    factor_text = f"{factor:.2f}".replace(".", ",")
    if method == "season" and store == "Ozon":
        text += f", год назад в эти дни {last_year:.0f} шт. × темп {factor_text}"
    elif method == "season":
        text += f", сезонность ×{factor_text}"
    return text


def _costs_articles(root: Path) -> List[str]:
    from recommended_prices import COSTS_FILENAME, _artikul_normalize, _normalize_offer_id, load_costs_df

    path = root / COSTS_FILENAME
    if not path.exists():
        return []
    df, key_col, _cost_col = load_costs_df(path)
    return [_normalize_offer_id(article) for article in (_artikul_normalize(value) for value in df[key_col]) if article]


def load_ozon(root: Path, today: date) -> Dict[str, Any]:
    """Остатки Ozon и план «сколько держать» на 30 дней. Долго только в первый раз: история заказов кэшируется."""
    from recommended_prices import get_product_id_mapping, get_product_skus

    settings = horizon_settings(ozon_plan.load_settings(root))
    orders = ozon_plan.load_orders(root, today)
    articles = sorted(set(_costs_articles(root)) | set(orders["offer_id"]), key=ozon_plan._article_sort_key)
    pid_to_offer = get_product_id_mapping(articles)
    skus = get_product_skus(sorted(pid_to_offer))
    stock, _ads = ozon_plan.load_cluster_stock({pid_to_offer[pid]: sku for pid, sku in skus.items() if pid in pid_to_offer})
    own = ozon_plan.load_own_stock(articles)
    plan = ozon_plan.build_plan(orders, stock, settings, today, articles=articles, own_stock=own)
    by_article = stock.groupby("offer_id")[["available", "transit", "requested"]].sum() if len(stock) else None
    forecast = plan["window_forecast"]
    return {
        "needs": ozon_needs(plan),
        "own": own,
        "stock": {a: int(r["available"]) for a, r in by_article.iterrows()} if by_article is not None else {},
        "transit": {a: int(r["transit"] + r["requested"]) for a, r in by_article.iterrows()} if by_article is not None else {},
        "level": float(settings["service_level"]),
        "note": forecast_note("Ozon", forecast["units"], ozon_plan.units_between(plan["daily"], today - timedelta(days=30), today),
                              forecast["method"], forecast.get("last_year_units") or 0.0, forecast.get("growth") or 1.0),
    }


def load_wb(root: Path, today: date) -> Dict[str, Any]:
    """Остатки WB и план «сколько держать» на 30 дней. Первый раз долго: выкупы из финансовых отчётов (1 запрос в минуту)."""
    from scripts.marketplace_settings import read_settings

    token = read_settings(root).get("WB_API_TOKEN")
    if not token:
        raise wb_plan.WbSupplyError("Добавьте API-токен Wildberries в настройках.")
    settings = horizon_settings(wb_plan.load_settings(root))
    orders = wb_plan.load_recent_orders(token, today)
    history = wb_plan.load_sales_history(root, token, today, progress=lambda message: None)
    cards = wb_plan.load_cards(token)
    own = wb_plan.load_own_stock(token, cards)
    wb_stock = wb_plan.load_wb_stock(token, cards)
    plan = wb_plan.build_wb_plan(orders, history, own, wb_stock, settings, today, articles=list(cards))
    live = orders[~orders["cancelled"]] if len(orders) else orders
    last_30 = float(live[live["date"] >= pd.Timestamp(today - timedelta(days=30))]["qty"].sum()) if len(live) else 0.0
    season = plan["season_window"]
    return {
        "needs": wb_needs(plan),
        "own": own,
        "stock": {article: entry["quantity"] for article, entry in wb_stock.items()},
        "transit": {article: entry["from_client"] for article, entry in wb_stock.items()},
        "level": float(settings["service_level"]),
        "note": forecast_note("WB", plan["window_units"], last_30, season["method"], factor=season["factor"]),
    }


def load(root: Path, today: Optional[date] = None) -> Dict[str, Any]:
    """Оба магазина. Магазин, который не загрузился, даёт ошибку в errors, а его колонки остаются пустыми."""
    today = today or date.today()
    sources: Dict[str, Optional[Dict[str, Any]]] = {}
    errors: Dict[str, str] = {}
    for name, loader in (("Ozon", load_ozon), ("WB", load_wb)):
        try:
            sources[name] = loader(root, today)
        except Exception as exc:
            sources[name] = None
            errors[name] = str(exc)
    levels = [source["level"] for source in sources.values() if source]
    level = max(levels) if levels else 0.9
    return {
        "loaded": datetime.now().isoformat(timespec="seconds"),
        "level": level,
        "rows": build_rows(sources["Ozon"], sources["WB"], level),
        "notes": [source["note"] for source in sources.values() if source],
        "errors": errors,
    }
