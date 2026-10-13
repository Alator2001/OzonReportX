"""
Автодобавление в акции Ozon и учёт минимальной цены — чтение состояния по товарам и изменения с проверкой.

Ozon (Seller API):
- /v5/product/info/prices — price.auto_add_to_ozon_actions_list_enabled (автодобавление в акции),
  price.auto_action_enabled (автоприменение акций), price.min_price;
- /v1/product/action/timer/status — min_price_for_auto_actions_enabled (учитывает ли Ozon минимальную цену
  при добавлении в акции) и expired_at — когда учёт сбросится;
- /v1/product/action/timer/update — продлить учёт минимальной цены (цены не меняются);
- /v1/product/import/prices — включить/выключить автодобавление и учёт минимальной цены.

У самого автодобавления срока нет, поэтому программа пишет журнал своих изменений и сверяет его с Ozon:
если Ozon вернул автодобавление сам, это видно как «Ozon включил обратно».
"""
from __future__ import annotations

import json
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

try:
    from scripts.recommended_prices import OZON_HEADERS, _normalize_offer_id
except ModuleNotFoundError:
    from recommended_prices import OZON_HEADERS, _normalize_offer_id

BASE = "https://api-seller.ozon.ru"
LOG_DIR = "actions history"
LOG_FILE = "auto_actions_log.jsonl"
TIMEOUT = 60


class AutoActionsError(RuntimeError):
    pass


def _post(path: str, payload: Dict[str, Any], attempts: int = 3) -> Dict[str, Any]:
    last = ""
    for attempt in range(attempts):
        try:
            response = requests.post(BASE + path, headers=OZON_HEADERS, json=payload, timeout=TIMEOUT)
        except requests.RequestException as exc:
            last = exc.__class__.__name__
        else:
            if response.status_code == 200:
                return response.json()
            if response.status_code < 500 and response.status_code != 429:
                raise AutoActionsError(f"Ozon отклонил {path}: HTTP {response.status_code} {response.text[:300]}")
            last = f"HTTP {response.status_code}"
        time.sleep(3 * (attempt + 1))
    raise AutoActionsError(f"Ozon временно не отвечает на {path} ({last}). Повторите через несколько минут.")


def _number(value: Any) -> Optional[float]:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if number > 0 else None


def _flag(value: Any) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().upper()
    if text in ("TRUE", "ENABLED"):
        return True
    if text in ("FALSE", "DISABLED"):
        return False
    return None


# ---------------------------------------------------------------- журнал

def log_path(root: Path) -> Path:
    return Path(root) / LOG_DIR / LOG_FILE


def read_log(root: Path) -> List[Dict[str, Any]]:
    path = log_path(root)
    if not path.exists():
        return []
    entries = []
    for line in path.read_text(encoding="utf-8").splitlines():
        try:
            entries.append(json.loads(line))
        except ValueError:
            continue
    return entries


def append_log(root: Path, entries: List[Dict[str, Any]]) -> None:
    path = log_path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        for entry in entries:
            handle.write(json.dumps(entry, ensure_ascii=False) + "\n")


def last_auto_add_change(entries: List[Dict[str, Any]]) -> Dict[int, Dict[str, Any]]:
    """Последнее подтверждённое изменение автодобавления по товару: {product_id: запись журнала}."""
    result: Dict[int, Dict[str, Any]] = {}
    for entry in entries:
        if entry.get("action") in ("auto_add_off", "auto_add_on") and entry.get("confirmed"):
            result[int(entry["product_id"])] = entry
    return result


# ---------------------------------------------------------------- чтение

def load_prices() -> List[Dict[str, Any]]:
    items, cursor = [], ""
    while True:
        payload: Dict[str, Any] = {"filter": {"visibility": "ALL"}, "limit": 1000}
        if cursor:
            payload["cursor"] = cursor
        body = _post("/v5/product/info/prices", payload)
        page = body.get("items") or []
        items += page
        cursor = body.get("cursor") or ""
        if not page or not cursor or len(page) < 1000:
            return items


def load_timers(product_ids: List[int]) -> Dict[int, Dict[str, Any]]:
    result: Dict[int, Dict[str, Any]] = {}
    for start in range(0, len(product_ids), 100):
        body = _post("/v1/product/action/timer/status", {"product_ids": [str(pid) for pid in product_ids[start:start + 100]]})
        for status in body.get("statuses") or []:
            result[int(status["product_id"])] = status
    return result


def build_rows(prices: List[Dict[str, Any]], timers: Dict[int, Dict[str, Any]], log: List[Dict[str, Any]],
               now: Optional[datetime] = None) -> List[Dict[str, Any]]:
    """Строки для таблицы: по товару — автодобавление, автоприменение, минимальная цена, таймер, предупреждения."""
    now = now or datetime.now(timezone.utc)
    changes = last_auto_add_change(log)
    rows = []
    for item in prices:
        price = item.get("price") or {}
        product_id = int(item["product_id"])
        timer = timers.get(product_id) or {}
        expires = timer.get("expired_at")
        expires_at = datetime.fromisoformat(expires.replace("Z", "+00:00")) if expires else None
        min_price_on = bool(timer.get("min_price_for_auto_actions_enabled"))
        auto_add = _flag(price.get("auto_add_to_ozon_actions_list_enabled"))
        change = changes.get(product_id)
        reverted = bool(change and change["action"] == "auto_add_off" and auto_add)
        rows.append({
            "product_id": product_id,
            "offer_id": _normalize_offer_id(str(item.get("offer_id") or "")),
            "price": _number(price.get("price")),
            "old_price": _number(price.get("old_price")),
            "min_price": _number(price.get("min_price")),
            "auto_add": auto_add,
            "auto_action": _flag(price.get("auto_action_enabled")),
            "min_price_on": min_price_on,
            "expires_at": expires_at.isoformat() if expires_at else None,
            "hours_left": round((expires_at - now).total_seconds() / 3600, 1) if expires_at and min_price_on else None,
            "auto_add_changed_at": change["at"] if change else None,
            "reverted_by_ozon": reverted,
        })
    rows.sort(key=lambda r: (len(r["offer_id"]), r["offer_id"]))
    return rows


def load_state(root: Path) -> Dict[str, Any]:
    prices = load_prices()
    timers = load_timers([int(item["product_id"]) for item in prices])
    return {"loaded": datetime.now().isoformat(timespec="seconds"), "rows": build_rows(prices, timers, read_log(root))}


# ---------------------------------------------------------------- изменения (только по подтверждению в интерфейсе)

def _import_prices(rows: List[Dict[str, Any]], extra: Dict[str, Any]) -> Dict[int, str]:
    """
    /v1/product/import/prices с одним изменённым признаком. Цены не трогаем: если Ozon требует их в запросе,
    повторяем для таких товаров с текущими ценами (те же значения — ничего не меняется). Возвращает {product_id: ошибка}.
    """
    def send(batch: List[Dict[str, Any]], with_prices: bool) -> Dict[int, str]:
        prices = []
        for row in batch:
            entry = {"offer_id": row["offer_id"], "product_id": row["product_id"], **extra}
            if with_prices:
                for key in ("price", "old_price", "min_price"):
                    if row.get(key) is not None:
                        entry[key] = f"{row[key]:g}"
                if row.get("old_price") is None:
                    entry["old_price"] = "0"
            prices.append(entry)
        errors: Dict[int, str] = {}
        for start in range(0, len(prices), 1000):
            body = _post("/v1/product/import/prices", {"prices": prices[start:start + 1000]})
            for result in body.get("result") or []:
                if not result.get("updated"):
                    messages = "; ".join(str(e.get("message") or e.get("code")) for e in result.get("errors") or [])
                    errors[int(result.get("product_id") or 0)] = messages or "Ozon не применил изменение"
        return errors

    errors = send(rows, with_prices=False)
    retry = [row for row in rows if row["product_id"] in errors and "price" in errors[row["product_id"]].lower()]
    if retry:
        second = send(retry, with_prices=True)
        for row in retry:
            errors.pop(row["product_id"], None)
        errors.update(second)
    return errors


def apply_change(root: Path, action: str, rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    action: auto_add_off | auto_add_on | min_price_on | timer_update.
    Отправляет изменение, перечитывает состояние в Ozon и пишет журнал. Возвращает итог для интерфейса.
    """
    if not rows:
        return {"ok": True, "done": 0, "expected": 0, "failed": [], "message": "Ничего не выбрано."}
    errors: Dict[int, str] = {}
    if action == "auto_add_off":
        errors = _import_prices(rows, {"auto_add_to_ozon_actions_list_enabled": "DISABLED"})
    elif action == "auto_add_on":
        errors = _import_prices(rows, {"auto_add_to_ozon_actions_list_enabled": "ENABLED"})
    elif action == "min_price_on":
        without_min = [row for row in rows if not row.get("min_price")]
        for row in without_min:
            errors[row["product_id"]] = "не задана минимальная цена — укажите её в «Цены и скидки»"
        ready = [row for row in rows if row.get("min_price")]
        if ready:
            errors.update(_import_prices(ready, {"min_price_for_auto_actions_enabled": True}))
    elif action == "timer_update":
        product_ids = [str(row["product_id"]) for row in rows]
        for start in range(0, len(product_ids), 100):
            _post("/v1/product/action/timer/update", {"product_ids": product_ids[start:start + 100]})
    else:
        raise AutoActionsError(f"Неизвестное действие: {action}")

    time.sleep(2)  # Ozon применяет изменения не мгновенно
    state = {row["product_id"]: row for row in load_state(root)["rows"]}
    now = datetime.now().isoformat(timespec="seconds")
    log, failed, done = [], [], 0
    for row in rows:
        fresh = state.get(row["product_id"], {})
        if action == "auto_add_off":
            confirmed = fresh.get("auto_add") is False
        elif action == "auto_add_on":
            confirmed = fresh.get("auto_add") is True
        elif action == "min_price_on":
            confirmed = bool(fresh.get("min_price_on"))
        else:
            confirmed = bool(fresh.get("expires_at")) and fresh.get("expires_at") != row.get("expires_at")
        reason = errors.get(row["product_id"])
        if confirmed:
            done += 1
        else:
            failed.append({"offer_id": row["offer_id"], "reason": reason or "Ozon принял запрос, но состояние не изменилось"})
        log.append({"at": now, "action": action, "product_id": row["product_id"], "offer_id": row["offer_id"],
                    "confirmed": confirmed, "expires_at": fresh.get("expires_at"), "error": reason})
    append_log(root, log)
    labels = {"auto_add_off": "Автодобавление отключено", "auto_add_on": "Автодобавление включено",
              "min_price_on": "Учёт минимальной цены включён", "timer_update": "Таймер продлён"}
    message = f"{labels[action]}: {done} из {len(rows)} (проверено по данным Ozon)."
    if failed:
        message += f" Не применилось: {len(failed)}."
    return {"ok": not failed, "done": done, "expected": len(rows), "failed": failed, "message": message,
            "state": list(state.values())}
