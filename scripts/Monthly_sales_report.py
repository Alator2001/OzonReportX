import os
import requests
import pandas as pd
from datetime import datetime, timedelta
from dateutil.parser import isoparse
import time
import json
from decimal import Decimal
from dotenv import load_dotenv

# 🔐 Загружаем данные для авторизации из переменных окружения
load_dotenv()
CLIENT_ID = os.getenv('OZON_CLIENT_ID')
API_KEY = os.getenv('OZON_API_KEY')

if not CLIENT_ID or not API_KEY:
    raise RuntimeError("Отсутствуют переменные OZON_CLIENT_ID или OZON_API_KEY. Укажите их в .env или окружении.")

HEADERS = {
    'Client-Id': CLIENT_ID,
    'Api-Key': API_KEY,
    'Content-Type': 'application/json'
}

# 📅 Получаем диапазон дат за прошлый месяц
def get_last_month_date_range():
    today = datetime.now()
    first_day_current_month = today.replace(day=1)
    last_day_last_month = first_day_current_month - timedelta(days=1)
    first_day_last_month = last_day_last_month.replace(day=1)
    
    date_from = first_day_last_month.strftime('%Y-%m-%dT00:00:00Z')
    date_to = last_day_last_month.strftime('%Y-%m-%dT23:59:59Z')
    
    return date_from, date_to

# 📄 Загрузка карты себестоимости из внешнего файла
def load_cost_map():
    script_dir = os.path.dirname(__file__)
    repo_root = os.path.abspath(os.path.join(script_dir, '..'))

    candidates = [
        os.path.join(repo_root, 'costs.xlsx'),
        os.path.join(repo_root, 'costs.csv')
    ]

    for path in candidates:
        if os.path.exists(path):
            try:
                if path.endswith('.xlsx'):
                    df = pd.read_excel(path)
                else:
                    df = pd.read_csv(path)

                # Нормализуем имена столбцов
                lower_cols = {c.lower(): c for c in df.columns}
                # Поддерживаемые варианты названий
                key_col = None
                cost_col = None

                for variant in ['prefix', 'префикс', 'код', 'артикул', 'offer_id']:
                    if variant in lower_cols:
                        key_col = lower_cols[variant]
                        break
                for variant in ['cost', 'себестоимость', 'цена', 'стоимость']:
                    if variant in lower_cols:
                        cost_col = lower_cols[variant]
                        break

                if not key_col or not cost_col:
                    print(f"⚠️ Файл {os.path.basename(path)} найден, но столбцы не распознаны. \nОжидаются столбцы: 'prefix'/'префикс'/'код'/'артикул' и 'cost'/'себестоимость'.")
                    continue

                mapping = {}
                for _, row in df.iterrows():
                    key = str(row.get(key_col, '')).strip()
                    if not key or key.lower() == 'nan':
                        continue
                    try:
                        value = float(row.get(cost_col, 0) or 0)
                    except Exception:
                        continue
                    mapping[key] = value

                print(f"🧾 Загружена карта себестоимости из {os.path.basename(path)}: {len(mapping)} записей")
                return mapping
            except Exception as e:
                print(f"⚠️ Не удалось прочитать {os.path.basename(path)}: {e}")

    print("ℹ️ Файл себестоимости не найден (costs.xlsx или costs.csv в корне репозитория). Будет использовано значение 0.")
    return {}

# 📥 Получаем список заказов FBS (Fulfillment by Seller)
def get_orders():
    # now = datetime.now()
    # date_from = now.replace(day=1).strftime('%Y-%m-%dT00:00:00Z')
    # date_to = now.strftime('%Y-%m-%dT23:59:59Z')
    date_from, date_to = get_last_month_date_range()
    url = 'https://api-seller.ozon.ru/v3/posting/fbs/list'
    result = []
    limit = 100

    # Статусы заказов, которые необходимо получить
    STATUSES = ["awaiting_packaging", "awaiting_deliver", "delivering", "delivered", "cancelled"]

    for status in STATUSES:
        print(f"📥 [FBS] Получаем заказы со статусом: {status}")
        offset = 0
        while True:
            payload = {
                "filter": {
                    "since": date_from,
                    "to": date_to,
                    "status": status
                },
                "limit": limit,
                "offset": offset,
                "with": {
                    "analytics_data": True,
                    "financial_data": True
                }
            }

            response = requests.post(url, headers=HEADERS, json=payload)
            response.raise_for_status()
            data = response.json()

            postings = data.get("result", {}).get("postings", [])
            if not postings:
                break

            for p in postings:
                p["__schema"] = "FBS"  # Добавляем пометку о схеме
            result.extend(postings)
            offset += limit
            time.sleep(0.2)  # Пауза между запросами

    return result

# 📥 Получаем список заказов FBO (Fulfillment by Ozon)
def get_fbo_orders():
    # now = datetime.now()
    # date_from = now.replace(day=1).strftime('%Y-%m-%dT00:00:00Z')
    # date_to = now.strftime('%Y-%m-%dT23:59:59Z')
    date_from, date_to = get_last_month_date_range()

    url = 'https://api-seller.ozon.ru/v2/posting/fbo/list'
    result = []
    limit = 100

    STATUSES = ["awaiting_deliver", "delivering", "delivered", "cancelled"]

    for status in STATUSES:
        print(f"📥 [FBO] Получаем заказы со статусом: {status}")
        offset = 0
        while True:
            payload = {
                "dir": "ASC",
                "filter": {
                    "since": date_from,
                    "to": date_to,
                    "status": status
                },
                "limit": limit,
                "offset": offset,
                "with": {
                    "analytics_data": True,
                    "financial_data": True
                }
            }

            response = requests.post(url, headers=HEADERS, json=payload)
            response.raise_for_status()

            #print("📨 Ответ от API:", response.status_code)
            #print(response.text)

            data = response.json()

            if isinstance(data, list):
                postings = data
            elif isinstance(data, dict) and "result" in data:
                postings = data["result"]
            else:
                print(f"⚠️ Ожидался список заказов, но получено: {data}")
                break

            if not postings:
                break

            for p in postings:
                if isinstance(p, dict):
                    p["__schema"] = "FBO"  # Добавляем пометку о схеме
            result.extend(postings)
            offset += limit
            time.sleep(0.2)

    return result

# 💳 Получаем финансовые транзакции по заказу
def get_transactions(posting_number, date_from, date_to):
    url = "https://api-seller.ozon.ru/v3/finance/transaction/list"

    payload = {
        "filter": {
            "date": {
                "from": date_from,
                "to": date_to
            },
            "posting_number": posting_number
        },
        "page_size": 100,
        "page": 1
    }

    all_operations = []

    while True:
        response = requests.post(url, headers=HEADERS, json=payload)

        if response.status_code != 200:
            print(f"❌ Ошибка при запросе транзакций для {posting_number}: {response.status_code}")
            print(response.text)
            return []

        data = response.json().get("result", {})
        operations = data.get("operations", [])
        all_operations.extend(operations)

        if len(operations) < payload["page_size"]:
            break
        payload["page"] += 1

    return all_operations

# 📊 Преобразуем данные в Excel
def to_excel(postings, output_file=None):
    from datetime import datetime
    import pandas as pd

    date_from, date_to = get_last_month_date_range()
    rows = []
    total_posts = max(len(postings or []), 1)

# Получаем прошлый месяц и год
    now = datetime.now()
    if now.month == 1:
        month = 12
        year = now.year - 1
    else:
        month = now.month - 1
        year = now.year

    # Название месяца на русском в родительном падеже (Сентябрь → сентября)
    months = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]
    month_name = months[month - 1]

    # Формируем путь и имя файла в папке ../reports относительно этого скрипта
    if not output_file:
        script_dir = os.path.dirname(__file__)
        reports_dir = os.path.abspath(os.path.join(script_dir, '..', 'reports'))
        os.makedirs(reports_dir, exist_ok=True)
        output_file = os.path.join(reports_dir, f"{month_name} {year}.xlsx")


    # карта себестоимости: ключ может быть точным offer_id или префиксом
    cost_map = load_cost_map()

    for idx, post in enumerate(postings, start=1):
        posting_number = post.get("posting_number", "")
        status = post.get("status", "")                         # Статус
        date = post.get("shipment_date", "")                    # Дата отгрузки
        items = post.get("products", []) or []

        # Если в заказе нет товаров — пропускаем
        if not items:
            continue

        # Количество (сумма по позициям)
        quantity_total = sum(int(it.get("quantity", 0) or 0) for it in items)

        # Заголовок строки — первая позиция
        head = items[0]
        name = str(head.get("name", ""))
        # Все артикулы заказа (без дублей, в исходном порядке)
        seen = set()
        offer_ids_list = []
        for it in items:
            oid = str(it.get("offer_id", ""))
            if oid and oid not in seen:
                seen.add(oid)
                offer_ids_list.append(oid)
        offer_ids_joined = ", ".join(offer_ids_list)

        # Себестоимость (по всем товарам, со знаком минус)
        cost_price = 0.0
        for it in items:
            oid = str(it.get("offer_id", ""))
            q = int(it.get("quantity", 0) or 0)

            # 1) Пытаемся найти точное совпадение по offer_id
            unit_cost = None
            if oid in cost_map:
                unit_cost = float(cost_map.get(oid, 0))

            unit_cost = unit_cost if unit_cost is not None else 0.0
            cost_price -= unit_cost * q

        # Агрегация транзакций по заказу (без дублей, без эквайринга)
        amount = 0.0
        sale_commission = 0.0
        price = 0.0

        transactions = get_transactions(posting_number, date_from, date_to)
        for trans in transactions or []:
            amount += float(trans.get("amount") or 0)
            sale_commission += float(trans.get("sale_commission") or 0)
            price += float(trans.get("accruals_for_sale") or 0)

        # Формируем значения в зависимости от статуса
        if status == "delivering":
            amount_cell = "-"
            sale_commission_cell = "-"
            delivery_cost_cell = "-"
            profit_cell = "-"
        elif status == "cancelled":
            amount_cell = amount
            sale_commission_cell = "-"
            delivery_cost_cell = "-"
            profit_cell = amount
            cost_price = 0.0   # ← себестоимость обнуляем при отмене
        elif status == "delivered":
            amount_cell = amount
            sale_commission_cell = sale_commission
            delivery_cost_cell = - amount + price + sale_commission
            profit_cell = amount + cost_price
        else:
            amount_cell = "-"
            sale_commission_cell = "-"
            delivery_cost_cell = "-"
            profit_cell = "-"

        rows.append({
            "Статус": status,
            "Номер заказа": posting_number,
            "Название товара": name,
            "Артикул": offer_ids_joined,
            "Количество шт.": quantity_total,
            "Цена продажи": price,
            "Комиссия за продажу Ozon": sale_commission_cell,
            "Логистика (Включает операционные ошибки продавца)": delivery_cost_cell,
            "Сумма начисления": amount_cell,
            "Себестоимость": cost_price,
            "Прибыль": profit_cell,
            "Дата отгрузки": date,
            "Схема": post.get("__schema", "")
        })

        # выводим прогресс каждые 5 записей и на финише
        if idx % 5 == 0 or idx == total_posts:
            percent = int(idx * 100 / total_posts)
            print(f"\r⚙️ Обработка заказов: {percent}%", end="", flush=True)

    df = pd.DataFrame(rows)
    df.to_excel(output_file, index=False)
    print("\r✅ Обработка заказов: 100%")
    print(f"✅ Отчёт сохранён: {output_file}")
    return output_file

from openpyxl import load_workbook

def calc_business_indicators(filename):
    print("💲 Рассчёт бизнес показателей")

    # Открываем Excel-файл
    wb = load_workbook(filename)
    ws = wb.active  # Можно заменить на ws = wb["Имя_листа"], если нужно конкретный лист

    # Считаем Общую выручку
    sales_revenue = 0
    for cell in ws["F"][1:]: 
        if isinstance(cell.value, (int, float)):
            sales_revenue += cell.value

    # Считаем Чистую прибыль
    net_profit = 0
    for cell in ws["K"][1:]: 
        if isinstance(cell.value, (int, float)):
            net_profit += cell.value

    # Считаем Себестоимость
    cost_price = 0
    for cell in ws["J"][1:]: 
        if isinstance(cell.value, (int, float)):
            cost_price += cell.value

    net_profit_margin = (net_profit/sales_revenue)*100
    cogs = sales_revenue + cost_price
    gross_profit_margin =(cogs/sales_revenue)*100
    operating_expenses = cogs - net_profit
    # Записываем результат
    ws["P1"] = "Общаяя выручка"
    ws["Q1"] = sales_revenue
    ws["P2"] = "Чистая прибыль"
    ws["Q2"] = net_profit
    ws["P3"] = "Итоговая себестоимость"
    ws["Q3"] = cost_price
    ws["P4"] = "Рентабельность  по чистой прибыли (Net Profit Margin) %"
    ws["Q4"] = net_profit_margin
    ws["P5"] = "COGS (валовая прибыль)"
    ws["Q5"] = cogs
    ws["P6"] = "Gross Profit Margin Рентабельность по валовой прибыли %"
    ws["Q6"] = gross_profit_margin
    ws["P7"] = "Операционные расходы"
    ws["Q7"] = operating_expenses

    # Сохраняем изменения
    wb.save(filename)
    print(f"✅ Рассчёт бизнес показателей завершён")

# 🚀 Точка входа
def main():
    print("📦 Получаем список заказов за текущий месяц...")

    fbs_orders = get_orders()
    fbo_orders = get_fbo_orders()

    all_orders = fbs_orders + fbo_orders
    print(f"🔢 Найдено заказов: {len(all_orders)}")
    
    # Имя файла формируется внутри to_excel как "<Месяц> <Год>.xlsx"
    output_file = to_excel(all_orders)
    
    calc_business_indicators(output_file)

if __name__ == "__main__":
    main()