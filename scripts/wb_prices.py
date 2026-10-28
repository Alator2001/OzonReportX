# -*- coding: utf-8 -*-
"""
Цены WB: текущая цена, выплата, маржа, минимальная и желательная цена. Только чтение — в кабинете ничего не меняет.

Как считается (проверено на всех продажах отчётов 04–10.2026, см. docs/dev/19-wb-prices.md):
- Деньги продавца считаются от цены со скидкой продавца p (discountedPrice). Скидку WB (СПП) оплачивает WB
  из своей комиссии, на выплату она не влияет.
- Выплата = p × (1 − комиссия) − эквайринг. Сверх этого на продажу ложится логистика (доставка и возвраты).
- Комиссия — фактическая доля по продажам в отчётах (commissionPercent): свой артикул, иначе серия. Продаж
  артикула и серии нет — тариф WB по предмету (/api/v1/tariffs/commission), бо́льший из FBS (kgvpMarketplace) и FBW
  (paidStorageKgvp); тарифа нет — средняя по магазину. Тариф FBS у магазина выше фактической на 4–5 пунктов.
- Эквайринг — доля от p, средняя по продажам магазина в отчётах.
- Логистика — рубли за проданную штуку по отчётам последних месяцев: свой артикул, иначе серия
  (первые две цифры артикула), иначе средняя по магазину.
- Цена при марже m: p = (себестоимость + логистика) / (1 − комиссия − эквайринг − m).
- Снимок цен из WB (sync) сохраняется в wb reports/prices.json, таблица считается по нему без запросов (compute).
"""

import json
import math
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import requests

script_dir = Path(__file__).resolve().parent
code_root = script_dir.parent
for path in (str(code_root), str(script_dir)):
    if path not in sys.path:
        sys.path.insert(0, path)

from scripts.file_io import atomic_output_path  # noqa: E402
from scripts.wb_finance import load_report  # noqa: E402

PRICES_URL = "https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter"
CARDS_URL = "https://content-api.wildberries.ru/content/v2/get/cards/list"
COMMISSION_URL = "https://common-api.wildberries.ru/api/v1/tariffs/commission"
SETTINGS_FILENAME = "wb_price_settings.json"
COMMISSION_CACHE = Path("wb reports") / "commission.json"
PRICES_SNAPSHOT = Path("wb reports") / "prices.json"
DEFAULTS = {"min_margin": 0.25, "desired_margin": 0.30}
MODELS = {"fbs": "kgvpMarketplace", "fbw": "paidStorageKgvp"}
REPORT_MONTHS = 3            # логистика и эквайринг — по отчётам стольких последних месяцев
LOGISTICS_MIN_SALES = 2      # меньше продаж у артикула — берём серию
SERIES_MIN_SALES = 3         # меньше продаж у серии — берём магазин
TIMEOUT = 60


class WbPricesError(RuntimeError):
    pass


# ---------------------------------------------------------------- настройки маржи (свои, не общие с Ozon)

def validate_margins(min_margin: Any, desired_margin: Any) -> Dict[str, float]:
    try:
        low, high = float(min_margin), float(desired_margin)
    except (TypeError, ValueError):
        raise WbPricesError("Маржа WB — число, например 0,25.") from None
    if not (0 < low < 1 and 0 < high < 1):
        raise WbPricesError("Маржа WB — доля от 0 до 1, например 0,25 для 25%.")
    if low >= high:
        raise WbPricesError("Минимальная маржа WB должна быть меньше желательной.")
    return {"min_margin": low, "desired_margin": high}


def load_settings(root: Path) -> Dict[str, float]:
    path = Path(root) / SETTINGS_FILENAME
    if not path.exists():
        return dict(DEFAULTS)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return validate_margins(data.get("min_margin"), data.get("desired_margin"))
    except (OSError, ValueError, WbPricesError) as exc:
        print(f"⚠️ Не удалось прочитать {SETTINGS_FILENAME}, использую маржу по умолчанию: {exc}")
        return dict(DEFAULTS)


def save_settings(root: Path, min_margin: Any, desired_margin: Any) -> Dict[str, float]:
    settings = validate_margins(min_margin, desired_margin)
    with atomic_output_path(Path(root) / SETTINGS_FILENAME) as temporary:
        temporary.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
    return settings


# ---------------------------------------------------------------- загрузка из WB (только чтение)

def _request(method: str, url: str, token: str, **kwargs) -> Any:
    response = requests.request(method, url, headers={"Authorization": token}, timeout=TIMEOUT, **kwargs)
    if response.status_code != 200:
        raise WbPricesError(f"WB вернул HTTP {response.status_code} на {url.split('.ru')[-1]}")
    return response.json()


def load_prices(token: str) -> Dict[int, Dict[str, Any]]:
    """nmId -> артикул, цена, скидка, цена со скидкой. Цена одна на все размеры (у магазина нет поразмерных цен)."""
    goods: Dict[int, Dict[str, Any]] = {}
    offset = 0
    while True:
        body = _request("GET", PRICES_URL, token, params={"limit": 1000, "offset": offset})
        page = (body.get("data") or {}).get("listGoods") or []
        if not page:
            return goods
        for item in page:
            sizes = item.get("sizes") or []
            if not sizes:
                continue
            goods[int(item["nmID"])] = {
                "article": str(item.get("vendorCode") or ""),
                "price": float(sizes[0]["price"]),
                "discount": int(item.get("discount") or 0),
                "discounted_price": float(sizes[0]["discountedPrice"]),
            }
        offset += len(page)


def load_subjects(token: str) -> Dict[int, Dict[str, Any]]:
    """nmId -> предмет товара (subjectID и название) из карточек."""
    subjects: Dict[int, Dict[str, Any]] = {}
    cursor: Dict[str, Any] = {"limit": 100}
    while True:
        body = _request("POST", CARDS_URL, token, json={"settings": {"cursor": cursor, "filter": {"withPhoto": -1}}})
        page = body.get("cards") or []
        for card in page:
            subjects[int(card["nmID"])] = {"id": card.get("subjectID"), "name": card.get("subjectName") or ""}
        next_cursor = body.get("cursor") or {}
        if len(page) < cursor["limit"] or not next_cursor.get("nmID"):
            return subjects
        cursor = {"limit": 100, "updatedAt": next_cursor.get("updatedAt"), "nmID": next_cursor.get("nmID")}


def parse_commission(report: Iterable[Dict[str, Any]]) -> Dict[int, Dict[str, float]]:
    """subjectID -> комиссия FBS и FBW, доли."""
    result = {}
    for row in report:
        rates = {model: row.get(field) for model, field in MODELS.items()}
        if row.get("subjectID") is None or any(value is None for value in rates.values()):
            continue
        result[int(row["subjectID"])] = {model: float(value) / 100 for model, value in rates.items()}
    return result


def load_commission(root: Path, token: str, today: Optional[date] = None) -> Dict[int, Dict[str, float]]:
    """Тарифы комиссии по предметам. Метод — запрос в минуту, поэтому кэш на день."""
    today = today or date.today()
    path = Path(root) / COMMISSION_CACHE
    if path.exists():
        try:
            cached = json.loads(path.read_text(encoding="utf-8"))
            if cached.get("date") == today.isoformat():
                return parse_commission(cached["report"])
        except (OSError, ValueError, KeyError):
            pass
    report = _request("GET", COMMISSION_URL, token, params={"locale": "ru"}).get("report") or []
    if not report:
        raise WbPricesError("WB не вернул тарифы комиссии.")
    path.parent.mkdir(parents=True, exist_ok=True)
    with atomic_output_path(path) as temporary:
        temporary.write_text(json.dumps({"date": today.isoformat(), "report": report}, ensure_ascii=False), encoding="utf-8")
    return parse_commission(report)


def recent_details(root: Path, today: Optional[date] = None, months: int = REPORT_MONTHS) -> Dict[str, Any]:
    """Строки финансовых отчётов за последние месяцы, которые уже загружены (сводкой WB) этим же токеном."""
    today = today or date.today()
    year, month = today.year, today.month
    rows: List[Dict[str, Any]] = []
    used: List[str] = []
    for _ in range(months):
        data = load_report(root, year, month)
        if data:
            rows += data.get("details") or []
            used.append(f"{month:02d}.{year}")
        year, month = (year - 1, 12) if month == 1 else (year, month - 1)
    return {"rows": rows, "months": used}


# ---------------------------------------------------------------- расчёт

def _num(value: Any) -> float:
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return 0.0


def _series(article: str) -> str:
    return article[:2] if len(article) >= 4 and article[:2].isdigit() else ""


def report_costs(rows: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    """По отчётам: логистика и комиссия по артикулам (суммы на проданные штуки), доли магазина — эквайринг и СПП."""
    logistics: Dict[str, float] = defaultdict(float)
    commission: Dict[str, float] = defaultdict(float)
    sold: Dict[str, float] = defaultdict(float)
    acquiring = base = spp_sum = spp_count = 0.0
    for row in rows:
        article = str(row.get("vendorCode") or "")
        logistics[article] += _num(row.get("deliveryService"))
        if row.get("sellerOperName") != "Продажа":
            continue
        quantity = _num(row.get("quantity"))
        sold[article] += quantity
        commission[article] += _num(row.get("commissionPercent")) / 100 * quantity
        acquiring += _num(row.get("acquiringFee"))
        base += _num(row.get("retailPriceWithDisc")) * quantity
        spp_sum += _num(row.get("spp")) * quantity
        spp_count += quantity
    return {
        "logistics": {article: logistics.get(article, 0.0) for article in sold},
        "commission": dict(commission),
        "sold": dict(sold),
        "acquiring": acquiring / base if base else None,
        "spp": spp_sum / spp_count / 100 if spp_count else None,
    }


def _per_unit(article: str, sold: Dict[str, float], totals: Dict[str, float], store: bool = True) -> Optional[Dict[str, Any]]:
    """Среднее на проданную штуку: свой артикул, иначе серия, иначе (store) магазин. Нет продаж в отчётах — None."""
    if sold.get(article, 0) >= LOGISTICS_MIN_SALES:
        return {"value": totals.get(article, 0.0) / sold[article], "source": "артикул"}
    series = _series(article)
    if series:
        members = [key for key in sold if _series(key) == series]
        units = sum(sold[key] for key in members)
        if units >= SERIES_MIN_SALES:
            return {"value": sum(totals.get(key, 0.0) for key in members) / units, "source": "серия"}
    units = sum(sold.values())
    if store and units:
        return {"value": sum(totals.get(key, 0.0) for key in sold) / units, "source": "магазин"}
    return None


def logistics_per_unit(article: str, costs: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Логистика за проданную штуку, ₽ (в неё входят и возвраты, и отмены — всё, что WB списал за доставку)."""
    return _per_unit(article, costs["sold"], costs["logistics"])


def commission_share(article: str, costs: Dict[str, Any], tariff: Optional[Dict[str, float]]) -> Optional[Dict[str, Any]]:
    """Комиссия WB, доля от цены со скидкой: фактическая по продажам артикула или серии, иначе тариф WB по предмету
    (бо́льший из FBS и FBW; у магазина тариф FBS выше фактической на 4–5 пунктов), иначе средняя по магазину.
    Тариф раньше средней магазина: комиссия зависит от категории, а продажи магазина — почти одни доски."""
    sold, totals = costs["sold"], costs.get("commission") or {}
    actual = _per_unit(article, sold, totals, store=False)
    if actual:
        return actual
    if tariff:
        return {"value": max(tariff.values()), "source": "тариф WB"}
    return _per_unit(article, sold, totals)


def price_for_margin(cost: float, logistics: float, commission: float, acquiring: float, margin: float) -> Optional[int]:
    share = 1 - commission - acquiring - margin
    if cost <= 0 or share <= 1e-9:  # 1 − 0,70 − 0,05 − 0,25 в float — не ноль, а 5e-17
        return None
    return round((cost + logistics) / share)


def discount_for(price_before: float, target: Optional[float]) -> Optional[int]:
    """Наибольшая целая скидка продавца, %, при которой цена со скидкой не ниже target.
    WB принимает только целые проценты; округление вниз держит цену не ниже цели. None — даже без скидки
    цена до скидки ниже цели (поднимать надо её) или цели нет."""
    if not target or price_before <= 0 or price_before < target:
        return None
    return min(99, math.floor((1 - target / price_before) * 100 + 1e-9))


def unit_profit(price: float, cost: float, logistics: float, commission: float, acquiring: float) -> float:
    return price * (1 - commission - acquiring) - logistics - cost


def build_rows(
    prices: Dict[int, Dict[str, Any]],
    unit_costs: Dict[str, Any],
    subjects: Dict[int, Dict[str, Any]],
    commission: Dict[int, Dict[str, float]],
    costs: Dict[str, Any],
    settings: Dict[str, float],
) -> List[Dict[str, Any]]:
    """Строка на товар. Пустые поля (None) — не хватает данных, причина в notes."""
    acquiring = costs["acquiring"] or 0.0
    spp = costs["spp"]
    rows = []
    for nm_id, item in prices.items():
        price = item["discounted_price"]
        cost = unit_costs.get(str(nm_id))
        cost = float(cost) if cost is not None else None
        subject = subjects.get(nm_id) or {}
        tariff = commission.get(subject.get("id")) if subject.get("id") is not None else None
        rate = commission_share(item["article"], costs, tariff)
        logistics = logistics_per_unit(item["article"], costs)
        notes = []
        if cost is None:
            notes.append("нет себестоимости в wb_costs.xlsx")
        if rate is None:
            notes.append("нет ни продаж, ни тарифа комиссии для предмета")
        if logistics is None:
            notes.append("нет продаж в отчётах — логистику не из чего взять")
        row: Dict[str, Any] = {
            "nm_id": nm_id, "article": item["article"], "subject": subject.get("name") or "",
            "price": item["price"], "discount": item["discount"], "discounted_price": price,
            "buyer_price": round(price * (1 - spp)) if spp is not None else None,
            "cost": cost, "acquiring": acquiring,
            "commission": rate["value"] if rate else None, "commission_source": rate["source"] if rate else None,
            "logistics": logistics["value"] if logistics else None,
            "logistics_source": logistics["source"] if logistics else None,
            "profit": None, "margin": None, "min_price": None, "desired_price": None, "notes": notes,
            "discount_max": None, "discount_desired": None,
        }
        if not notes:
            row["profit"] = unit_profit(price, cost, logistics["value"], rate["value"], acquiring)
            row["margin"] = row["profit"] / price if price else None
            row["min_price"] = price_for_margin(cost, logistics["value"], rate["value"], acquiring, settings["min_margin"])
            row["desired_price"] = price_for_margin(cost, logistics["value"], rate["value"], acquiring, settings["desired_margin"])
        # Какую скидку можно ставить при текущей цене до скидки: от скидки для желательной цены до скидки для минимальной.
        row["discount_max"] = discount_for(item["price"], row["min_price"])
        row["discount_desired"] = discount_for(item["price"], row["desired_price"])
        rows.append(row)
    return sorted(rows, key=lambda row: row["article"])


# ── снимок цен: загрузка из WB отдельно от расчёта ─────────────────────────

def sync(root: Path, token: str, today: Optional[date] = None) -> Dict[str, Any]:
    """Загрузить из WB цены, предметы товаров и тарифы комиссии и сохранить снимок wb reports/prices.json."""
    from scripts.wb_finance import identity

    if not token:
        raise WbPricesError("Добавьте API-токен WB в настройках.")
    snapshot = {
        "version": 1, "account": identity(token), "fetched_at": datetime.now().isoformat(timespec="seconds"),
        "prices": {str(nm): item for nm, item in load_prices(token).items()},
        "subjects": {str(nm): item for nm, item in load_subjects(token).items()},
    }
    load_commission(root, token, today)  # кэш тарифов на день — для товаров без продаж
    from scripts.wb_promotions import record_history
    record_history(root, snapshot["prices"])  # история скидок — по ней раздел «Акции» видит входы и выходы из акций
    path = Path(root) / PRICES_SNAPSHOT
    path.parent.mkdir(parents=True, exist_ok=True)
    with atomic_output_path(path) as temporary:
        temporary.write_text(json.dumps(snapshot, ensure_ascii=False), encoding="utf-8")
    return snapshot


def load_snapshot(root: Path) -> Optional[Dict[str, Any]]:
    """Сохранённый снимок цен этого же токена или None."""
    from scripts.marketplace_settings import read_settings
    from scripts.wb_finance import identity

    try:
        snapshot = json.loads((Path(root) / PRICES_SNAPSHOT).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    token = read_settings(root).get("WB_API_TOKEN") or ""
    return snapshot if snapshot.get("version") == 1 and snapshot.get("account") == identity(token) else None


def compute(root: Path, snapshot: Dict[str, Any], today: Optional[date] = None) -> Dict[str, Any]:
    """Таблица цен по снимку, отчётам WB и wb_costs.xlsx — без запросов к WB."""
    from scripts.wb_finance import read_costs

    details = recent_details(root, today)
    costs = report_costs(details["rows"])
    settings = load_settings(root)
    commission = {}
    try:
        cached = json.loads((Path(root) / COMMISSION_CACHE).read_text(encoding="utf-8"))
        commission = parse_commission(cached.get("report") or [])
    except (OSError, ValueError):
        pass
    rows = build_rows({int(nm): item for nm, item in snapshot["prices"].items()}, read_costs(root),
                      {int(nm): item for nm, item in snapshot["subjects"].items()}, commission, costs, settings)
    return {"rows": rows, "months": details["months"], "acquiring": costs["acquiring"], "spp": costs["spp"],
            "settings": settings, "fetched_at": snapshot["fetched_at"]}


def run(root: Path, token: str, today: Optional[date] = None) -> Dict[str, Any]:
    """Загрузить снимок из WB и посчитать таблицу цен."""
    return compute(root, sync(root, token, today), today)
