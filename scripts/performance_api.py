"""
Модуль для работы с Ozon Performance API (рекламные кампании, затраты на маркетинг).

Документация: https://docs.ozon.ru/api/performance/#tag/Campaign

Основные методы Performance API (по документации):
1. GET /api/client/campaign - получение списка рекламных кампаний
2. POST /api/client/statistics - запрос статистики по кампаниям (возвращает UUID отчёта)
3. GET /api/client/statistics/{UUID} - проверка статуса формирования отчёта
4. GET /api/client/statistics/report - скачивание готового отчёта по UUID

ВАЖНО: 
- С 15 января 2025 хост performance.ozon.ru перестал работать, используется только api-performance.ozon.ru
- Авторизация: JWT токен через OAuth (POST /api/client/token)
- Client ID формат: "XXXXX-XXXXX@advertising.performance.ozon.ru"
- OZON_PERF_API_KEY в .env = client_secret (не готовый токен!)
"""
import os
from typing import List, Dict, Any, Optional, Tuple
import requests
import time

# Кэш для токена (чтобы не запрашивать каждый раз)
_token_cache: Dict[str, tuple] = {}  # {cache_key: (token, expires_at)}


def get_performance_api_credentials() -> Tuple[Optional[str], Optional[str]]:
    """
    Загружает credentials для Performance API из .env
    
    ВАЖНО: 
    - OZON_PERF_CLIENT_ID = client_id (формат: "XXXXX-XXXXX@advertising.performance.ozon.ru")
    - OZON_PERF_API_KEY = client_secret (используется для получения JWT токена через OAuth)
    """
    perf_client_id = os.getenv('OZON_PERF_CLIENT_ID')
    perf_client_secret = os.getenv('OZON_PERF_API_KEY')  # Это client_secret, не готовый токен!
    if perf_client_id and perf_client_secret:
        return perf_client_id, perf_client_secret
    return None, None


def get_performance_token(session: requests.Session, client_id: str, client_secret: str) -> Optional[str]:
    """
    Получает JWT токен для Performance API через OAuth endpoint.
    
    Документация: https://docs.ozon.ru/api/performance/
    
    Метод: POST /api/client/token
    Хост: api-performance.ozon.ru
    
    Параметры:
    - client_id: Client ID (формат: "XXXXX-XXXXX@advertising.performance.ozon.ru")
    - client_secret: Client Secret (из OZON_PERF_API_KEY в .env)
    - grant_type: "client_credentials"
    
    Ответ:
    - access_token: JWT токен
    - expires_in: время жизни токена в секундах (обычно 1800 = 30 минут)
    - token_type: "Bearer"
    
    Токен кэшируется до истечения срока действия.
    """
    if not client_id or not client_secret:
        return None
    
    # Проверяем кэш токена
    cache_key = f"{client_id}:{client_secret[:10]}"
    if cache_key in _token_cache:
        token, expires_at = _token_cache[cache_key]
        if time.time() < expires_at:
            return token
        # Токен истёк, удаляем из кэша
        del _token_cache[cache_key]
    
    # Получаем новый токен через OAuth endpoint
    url = "https://api-performance.ozon.ru/api/client/token"
    
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials"
    }
    
    try:
        resp = session.post(url, headers=headers, json=payload, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        
        access_token = data.get("access_token")
        expires_in = data.get("expires_in", 1800)  # По умолчанию 30 минут
        
        if access_token:
            # Кэшируем токен (сохраняем с запасом - минус 60 секунд до истечения)
            expires_at = time.time() + expires_in - 60
            _token_cache[cache_key] = (access_token, expires_at)
            return access_token
        else:
            print(f"⚠️ Не удалось получить access_token из ответа: {data}")
            return None
            
    except requests.exceptions.HTTPError as e:
        print(f"⚠️ Ошибка при получении токена Performance API: HTTP {e.response.status_code}")
        if e.response.text:
            try:
                error_data = e.response.json()
                print(f"   Ответ: {error_data}")
            except:
                print(f"   Ответ (текст): {e.response.text[:300]}")
        return None
    except Exception as e:
        print(f"⚠️ Ошибка при получении токена Performance API: {str(e)}")
        return None


def list_campaigns(session: requests.Session, perf_client_id: str, perf_token: str, 
                  state: Optional[str] = None, adv_object_type: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Получает список кампаний через Performance API.
    
    Документация: https://docs.ozon.ru/api/performance/#tag/Campaign
    
    Метод: GET /api/client/campaign
    Авторизация: Authorization: Bearer {JWT_TOKEN}
    
    Параметры запроса (опционально):
    - advObjectType: тип рекламного объекта (SKU, SEARCH_PROMO, BANNER и т.д.)
    - state: состояние кампании (CAMPAIGN_STATE_RUNNING, CAMPAIGN_STATE_STOPPED и т.д.)
    - page, pageSize: для пагинации
    """
    # По документации: GET /api/client/campaign
    # Хост: api-performance.ozon.ru
    # Авторизация: Authorization: Bearer {JWT_TOKEN}
    url = "https://api-performance.ozon.ru/api/client/campaign"
    
    headers = {
        "Authorization": f"Bearer {perf_token}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    all_campaigns = []
    page = 1
    page_size = 100
    
    try:
        while True:
            params = {
                "page": page,
                "pageSize": page_size
            }
            
            # Добавляем опциональные параметры
            if adv_object_type:
                params["advObjectType"] = adv_object_type
            if state:
                params["state"] = state
            
            resp = session.get(url, headers=headers, params=params, timeout=10)
            
            if resp.status_code != 200:
                print(f"⚠️ Ошибка при получении списка кампаний:")
                print(f"   URL: {url}")
                print(f"   Статус: {resp.status_code}")
                try:
                    error_data = resp.json()
                    print(f"   Ответ: {error_data}")
                except:
                    print(f"   Ответ (текст): {resp.text[:300]}")
                break
            
            data = resp.json()
            
            # По документации ответ имеет структуру: {"list": [...]}
            campaigns_list = []
            if isinstance(data, dict):
                if "list" in data:
                    # list может быть массивом или одним объектом
                    list_data = data["list"]
                    if isinstance(list_data, list):
                        campaigns_list = list_data
                    elif isinstance(list_data, dict):
                        campaigns_list = [list_data]
                else:
                    # Fallback: пробуем другие возможные поля
                    campaigns_list = data.get("result", data.get("campaigns", data.get("items", [])))
            elif isinstance(data, list):
                campaigns_list = data
            
            if not campaigns_list or not isinstance(campaigns_list, list):
                break
            
            all_campaigns.extend(campaigns_list)
            
            # Если получили меньше запрошенного размера страницы - это последняя страница
            if len(campaigns_list) < page_size:
                break
            
            page += 1
        
        if all_campaigns:
            print(f"✅ Список кампаний получен: найдено {len(all_campaigns)} кампаний")
            return all_campaigns
        else:
            print("ℹ️ Кампании не найдены")
            return []
            
    except requests.exceptions.HTTPError as e:
        print(f"⚠️ HTTP ошибка при получении списка кампаний: {e.response.status_code}")
        if e.response.text:
            try:
                error_data = e.response.json()
                print(f"   Ответ: {error_data}")
            except:
                print(f"   Ответ (текст): {e.response.text[:300]}")
        return []
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Ошибка запроса при получении списка кампаний: {str(e)}")
        return []
    except Exception as e:
        print(f"⚠️ Неожиданная ошибка при получении списка кампаний: {str(e)}")
        return []


def filter_cpc_campaigns(campaigns: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Фильтрует кампании по типу оплаты 'CPC' (оплата за клик).
    
    По документации:
    - advObjectType: "SKU" = Оплата за клик (уже отфильтровано в list_campaigns)
    - paymentType: может быть "CPC", "CPO" и т.д.
    
    Все кампании уже отфильтрованы по advObjectType="SKU" в list_campaigns,
    поэтому просто возвращаем все.
    """
    # Все кампании уже отфильтрованы по advObjectType="SKU" в list_campaigns,
    # поэтому просто возвращаем все
    return campaigns


def request_statistics_report(session: requests.Session, perf_client_id: str, perf_token: str, 
                              campaign_ids: List[int], date_from: str, date_to: str) -> Optional[str]:
    """
    Запрашивает создание отчёта со статистикой по кампаниям.
    Возвращает UUID отчёта для последующей проверки статуса и скачивания.
    
    Метод: POST /api/client/statistics
    Хост: api-performance.ozon.ru (старый performance.ozon.ru больше не работает с 15 января 2025)
    """
    url = "https://api-performance.ozon.ru/api/client/statistics"
    
    headers = {
        "Authorization": f"Bearer {perf_token}",
        "Content-Type": "application/json",
        "Client-Id": perf_client_id
    }
    
    payload = {
        "campaigns": campaign_ids,
        "dateFrom": date_from.split("T")[0],
        "dateTo": date_to.split("T")[0],
        "groupBy": "DATE"  # или "CUMULATIVE" для итоговой суммы
    }
    
    try:
        resp = session.post(url, headers=headers, json=payload, timeout=30)
        if resp.status_code != 200:
            print(f"⚠️ Ошибка при создании отчёта статистики:")
            print(f"   URL: {url}")
            print(f"   Статус: {resp.status_code}")
            try:
                error_data = resp.json()
                print(f"   Ответ: {error_data}")
            except:
                print(f"   Ответ (текст): {resp.text[:300]}")
            return None
        
        data = resp.json()
        # API может возвращать UUID в разных регистрах и полях
        report_uuid = (data.get("UUID") or data.get("uuid") or 
                      data.get("reportId") or data.get("report_id") or 
                      data.get("id") or data.get("Id"))
        if report_uuid:
            return str(report_uuid)
        else:
            print(f"⚠️ Не удалось получить UUID отчёта из ответа: {data}")
            return None
    except requests.exceptions.HTTPError as e:
        print(f"⚠️ HTTP ошибка при создании отчёта статистики: {e.response.status_code}")
        if e.response.text:
            try:
                error_data = e.response.json()
                print(f"   Ответ: {error_data}")
            except:
                print(f"   Ответ (текст): {e.response.text[:300]}")
        return None
    except Exception as e:
        print(f"⚠️ Ошибка при создании отчёта статистики: {str(e)}")
        return None


def get_statistics_report_status(session: requests.Session, perf_client_id: str, perf_token: str, 
                                 report_uuid: str) -> Optional[str]:
    """
    Проверяет статус формирования отчёта.
    Возвращает статус: "pending", "ready", "error" или None если не удалось проверить.
    
    Метод: GET /api/client/statistics/{UUID}
    Хост: api-performance.ozon.ru (старый performance.ozon.ru больше не работает с 15 января 2025)
    """
    url = f"https://api-performance.ozon.ru/api/client/statistics/{report_uuid}"
    
    headers = {
        "Authorization": f"Bearer {perf_token}",
        "Content-Type": "application/json",
        "Client-Id": perf_client_id
    }
    
    try:
        resp = session.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        return data.get("status") or data.get("state")
    except Exception as e:
        print(f"⚠️ Ошибка при проверке статуса отчёта: {str(e)}")
        return None


def download_statistics_report(session: requests.Session, perf_client_id: str, perf_token: str, 
                               report_uuid: str) -> Optional[Dict[str, Any]]:
    """
    Скачивает готовый отчёт по UUID и парсит данные.
    Возвращает словарь с агрегированными показателями или None.
    
    Метод: GET /api/client/statistics/report?uuid={UUID}
    Хост: api-performance.ozon.ru (старый performance.ozon.ru больше не работает с 15 января 2025)
    """
    url = "https://api-performance.ozon.ru/api/client/statistics/report"
    
    headers = {
        "Authorization": f"Bearer {perf_token}",
        "Content-Type": "application/json",
        "Client-Id": perf_client_id
    }
    
    params = {"uuid": report_uuid}
    
    try:
        resp = session.get(url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        # Отчёт может быть в JSON или CSV формате
        content_type = resp.headers.get("Content-Type", "")
        if "application/json" in content_type:
            return resp.json()
        elif "text/csv" in content_type or "application/zip" in content_type:
            # Если CSV/ZIP - нужно парсить отдельно
            print("⚠️ Отчёт в CSV/ZIP формате - требуется дополнительная обработка")
            return None
    except Exception as e:
        print(f"⚠️ Ошибка при скачивании отчёта: {str(e)}")
        return None
    
    return None


def get_campaign_stats_for_month(session: requests.Session, perf_client_id: str, perf_token: str, campaign_ids: List[int], 
                                  date_from: str, date_to: str) -> Dict[str, Any]:
    """
    Получает статистику по кампаниям за указанный период через прямой метод API.
    
    Использует GET /api/client/statistics/campaign/product/json для получения данных сразу,
    без необходимости ждать готовности асинхронного отчёта.
    """
    if not campaign_ids:
        return {"total_cost": 0.0, "total_clicks": 0, "campaigns_count": 0}
    
    # Используем прямой метод для получения статистики
    statistics = get_campaign_statistics_json(session, perf_client_id, perf_token, campaign_ids, date_from, date_to)
    
    if not statistics:
        return {"total_cost": 0.0, "total_clicks": 0, "campaigns_count": 0}
    
    # Парсим и суммируем показатели
    total_cost = 0.0
    total_clicks = 0
    
    for stat in statistics:
        # API использует поля: moneySpent, clicks
        # Значения приходят как строки, могут быть с запятыми вместо точек
        cost_str = stat.get("moneySpent") or stat.get("cost") or stat.get("spent") or 0
        clicks_str = stat.get("clicks") or 0
        
        # Парсим числа (заменяем запятые на точки)
        def parse_number(value):
            if not value:
                return 0
            if isinstance(value, (int, float)):
                return float(value)
            value_str = str(value).replace(",", ".").replace(" ", "")
            try:
                return float(value_str)
            except (ValueError, TypeError):
                return 0
        
        def parse_int(value):
            if not value:
                return 0
            if isinstance(value, int):
                return value
            if isinstance(value, float):
                return int(value)
            value_str = str(value).replace(",", ".").replace(" ", "")
            try:
                return int(float(value_str))
            except (ValueError, TypeError):
                return 0
        
        total_cost += parse_number(cost_str)
        total_clicks += parse_int(clicks_str)
    
    return {
        "total_cost": total_cost,
        "total_clicks": total_clicks,
        "campaigns_count": len(campaign_ids)
    }


def _get_stats_direct(session: requests.Session, perf_client_id: str, perf_token: str, campaign_ids: List[int],
                      date_from: str, date_to: str) -> Dict[str, Any]:
    """Fallback метод: прямой запрос статистики (если асинхронный механизм не работает)
    
    Хост: api-performance.ozon.ru (старый performance.ozon.ru больше не работает с 15 января 2025)
    """
    url = f"https://api-performance.ozon.ru/api/client/{perf_client_id}/adv/v1/statistics/campaign"
    
    headers = {
        "Authorization": f"Bearer {perf_token}",
        "Content-Type": "application/json",
        "Client-Id": perf_client_id
    }
    
    payload = {
        "campaign_ids": campaign_ids,
        "fromDate": date_from.split("T")[0],
        "toDate": date_to.split("T")[0],
        "groupBy": "CUMULATIVE"
    }
    
    try:
        resp = session.post(url, headers=headers, json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        total_cost = 0.0
        total_clicks = 0
        stats_list = data.get("result", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
        for stat in stats_list:
            total_cost += float(stat.get("cost", stat.get("spent", 0)) or 0)
            total_clicks += int(stat.get("clicks", stat.get("clickCount", 0)) or 0)
        return {
            "total_cost": total_cost,
            "total_clicks": total_clicks,
            "campaigns_count": len(campaign_ids)
        }
    except Exception as e:
        print(f"⚠️ Ошибка при прямом запросе статистики: {str(e)}")
        return {"total_cost": 0.0, "total_clicks": 0, "campaigns_count": 0}


def get_active_campaigns_for_month(session: requests.Session, perf_client_id: str, perf_token: str,
                                   date_from: str, date_to: str) -> List[Dict[str, Any]]:
    """
    Получает список активных кампаний за указанный месяц.
    
    Фильтрует кампании по:
    - state = CAMPAIGN_STATE_RUNNING (активные кампании)
    - Период активности кампании пересекается с указанным месяцем
    """
    # Получаем все активные кампании
    all_campaigns = list_campaigns(session, perf_client_id, perf_token, 
                                   state="CAMPAIGN_STATE_RUNNING")
    
    if not all_campaigns:
        return []
    
    # Парсим даты периода
    from datetime import datetime
    try:
        # Приводим все даты к naive datetime (без timezone) для корректного сравнения
        if "T" in date_from:
            period_from = datetime.fromisoformat(date_from.replace("Z", "+00:00"))
            # Убираем timezone, оставляем только дату
            if period_from.tzinfo is not None:
                period_from = period_from.replace(tzinfo=None)
        else:
            period_from = datetime.strptime(date_from.split("T")[0], "%Y-%m-%d")
        
        if "T" in date_to:
            period_to = datetime.fromisoformat(date_to.replace("Z", "+00:00"))
            # Убираем timezone, оставляем только дату
            if period_to.tzinfo is not None:
                period_to = period_to.replace(tzinfo=None)
        else:
            period_to = datetime.strptime(date_to.split("T")[0], "%Y-%m-%d")
    except Exception as e:
        print(f"⚠️ Ошибка парсинга дат: {e}")
        return all_campaigns  # Возвращаем все, если не удалось распарсить
    
    # Фильтруем кампании, которые были активны в указанном периоде
    active_in_period = []
    for campaign in all_campaigns:
        from_date_str = campaign.get("fromDate", "")
        to_date_str = campaign.get("toDate", "")
        
        if not from_date_str:
            # Если нет даты начала, считаем что кампания активна
            active_in_period.append(campaign)
            continue
        
        try:
            # Парсим даты кампании и приводим к naive datetime
            if "T" in from_date_str:
                camp_from = datetime.fromisoformat(from_date_str.replace("Z", "+00:00"))
                # Убираем timezone
                if camp_from.tzinfo is not None:
                    camp_from = camp_from.replace(tzinfo=None)
            else:
                camp_from = datetime.strptime(from_date_str, "%Y-%m-%d")
            
            camp_to = None
            if to_date_str:
                if "T" in to_date_str:
                    camp_to = datetime.fromisoformat(to_date_str.replace("Z", "+00:00"))
                    # Убираем timezone
                    if camp_to.tzinfo is not None:
                        camp_to = camp_to.replace(tzinfo=None)
                else:
                    camp_to = datetime.strptime(to_date_str, "%Y-%m-%d")
            
            # Проверяем пересечение периодов
            # Кампания активна в периоде, если:
            # - дата начала кампании <= конец периода И
            # - (дата окончания кампании >= начало периода ИЛИ дата окончания не указана)
            if camp_from <= period_to and (camp_to is None or camp_to >= period_from):
                active_in_period.append(campaign)
        except Exception as e:
            # Если не удалось распарсить даты кампании, включаем её в список
            print(f"⚠️ Не удалось распарсить даты кампании {campaign.get('id', 'unknown')}: {e}")
            active_in_period.append(campaign)
    
    return active_in_period


def get_campaign_statistics_json(session: requests.Session, perf_client_id: str, perf_token: str,
                                 campaign_ids: List[int], date_from: str, date_to: str) -> Optional[List[Dict[str, Any]]]:
    """
    Получает статистику по кампаниям в формате JSON.
    
    Метод: GET /api/client/statistics/campaign/product/json
    
    Параметры:
    - campaignIds: список ID кампаний (массив строк)
    - dateFrom, dateTo: даты периода в формате ГГГГ-ММ-ДД
    - from, to: даты периода в формате RFC3339 (альтернатива)
    """
    url = "https://api-performance.ozon.ru/api/client/statistics/campaign/product/json"
    
    headers = {
        "Authorization": f"Bearer {perf_token}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    # Форматируем даты
    date_from_formatted = date_from.split("T")[0] if "T" in date_from else date_from
    date_to_formatted = date_to.split("T")[0] if "T" in date_to else date_to
    
    # Формируем параметры запроса
    # По документации campaignIds должен быть массивом строк
    params = {
        "dateFrom": date_from_formatted,
        "dateTo": date_to_formatted
    }
    
    # Добавляем campaignIds если они указаны
    # В GET запросах массивы передаются как повторяющиеся параметры
    if campaign_ids:
        campaign_ids_str = [str(cid) for cid in campaign_ids]
        # requests автоматически обработает список, создав параметры вида campaignIds=id1&campaignIds=id2
        params["campaignIds"] = campaign_ids_str
    
    try:
        resp = session.get(url, headers=headers, params=params, timeout=30)
        
        if resp.status_code != 200:
            print(f"⚠️ HTTP ошибка при получении статистики кампаний: {resp.status_code}")
            print(f"   URL: {url}")
            print(f"   Параметры: {params}")
            try:
                error_data = resp.json()
                print(f"   Ответ: {error_data}")
            except:
                print(f"   Ответ (текст): {resp.text[:500]}")
            return None
        
        data = resp.json()
        
        # По документации ответ может быть в разных форматах
        if isinstance(data, list):
            return data
        elif isinstance(data, dict):
            # Пробуем различные возможные поля (включая 'rows', который использует API)
            result = (data.get("rows") or data.get("result") or data.get("data") or 
                     data.get("list") or data.get("items") or [])
            
            if isinstance(result, list):
                return result
            elif isinstance(result, dict):
                return [result]
        
        return []
    except requests.exceptions.HTTPError as e:
        print(f"⚠️ HTTP ошибка при получении статистики кампаний: {e.response.status_code}")
        if e.response.text:
            try:
                error_data = e.response.json()
                print(f"   Ответ: {error_data}")
            except:
                print(f"   Ответ (текст): {e.response.text[:300]}")
        return None
    except Exception as e:
        print(f"⚠️ Ошибка при получении статистики кампаний: {str(e)}")
        return None


def get_active_campaigns_with_statistics(session: requests.Session, perf_client_id: str, perf_token: str,
                                         date_from: str, date_to: str) -> List[Dict[str, Any]]:
    """
    Получает активные кампании за месяц вместе со статистикой.
    
    Возвращает список словарей, каждый содержит:
    - Данные кампании (id, title, state, budget и т.д.)
    - Статистику за период (расход, показы, клики, заказы и т.д.)
    """
    # Получаем активные кампании за период
    active_campaigns = get_active_campaigns_for_month(session, perf_client_id, perf_token, date_from, date_to)
    
    if not active_campaigns:
        return []
    
    # Извлекаем ID кампаний
    campaign_ids = []
    for camp in active_campaigns:
        camp_id = camp.get("id") or camp.get("campaign_id")
        if camp_id:
            try:
                campaign_ids.append(int(camp_id))
            except (ValueError, TypeError):
                continue
    
    if not campaign_ids:
        print("⚠️ Не удалось извлечь ID кампаний")
        return []
    
    # Получаем статистику по кампаниям
    statistics = get_campaign_statistics_json(session, perf_client_id, perf_token, campaign_ids, date_from, date_to)
    
    if not statistics:
        print("⚠️ Не удалось получить статистику по кампаниям")
        # Возвращаем кампании без статистики
        return active_campaigns
    
    # Создаём словарь статистики по ID кампании для быстрого поиска
    stats_by_campaign_id = {}
    for stat in statistics:
        camp_id = stat.get("campaignId") or stat.get("campaign_id") or stat.get("id")
        if camp_id:
            try:
                stats_by_campaign_id[int(camp_id)] = stat
            except (ValueError, TypeError):
                continue
    
    # Объединяем данные кампаний со статистикой
    result = []
    for campaign in active_campaigns:
        camp_id = campaign.get("id") or campaign.get("campaign_id")
        if not camp_id:
            continue
        
        try:
            camp_id_int = int(camp_id)
        except (ValueError, TypeError):
            continue
        
        # Объединяем данные кампании и статистику
        combined = campaign.copy()
        if camp_id_int in stats_by_campaign_id:
            combined.update(stats_by_campaign_id[camp_id_int])
        
        result.append(combined)
    
    return result


def get_cpc_campaigns_for_month(session: requests.Session, date_from: str, date_to: str) -> Dict[str, Any]:
    """Главная функция: получает список CPC кампаний, которые были активны в указанном месяце"""
    perf_client_id, perf_api_key = get_performance_api_credentials()
    if not perf_client_id or not perf_api_key:
        print("ℹ️ Не настроены переменные для Performance API: OZON_PERF_CLIENT_ID и OZON_PERF_API_KEY. Пропускаем автозагрузку маркетинга.")
        return {"total_cost": 0.0, "total_clicks": 0, "campaigns_count": 0}

    print("📢 Получаем данные о рекламных кампаниях (CPC)...")
    
    perf_token = get_performance_token(session, perf_client_id, perf_api_key)
    if not perf_token:
        print("⚠️ Не удалось получить токен для Performance API")
        return {"total_cost": 0.0, "total_clicks": 0, "campaigns_count": 0}
    
    # Получаем активные кампании за период
    active_campaigns = get_active_campaigns_for_month(session, perf_client_id, perf_token, date_from, date_to)
    if not active_campaigns:
        print("ℹ️ Не найдено активных кампаний за указанный период")
        return {"total_cost": 0.0, "total_clicks": 0, "campaigns_count": 0}
    
    # Фильтруем только CPC кампании (SKU тип)
    cpc_campaigns = [c for c in active_campaigns if c.get("advObjectType") == "SKU" or c.get("paymentType") == "CPC"]
    
    if not cpc_campaigns:
        print("ℹ️ Не найдено CPC кампаний (оплата за клик)")
        return {"total_cost": 0.0, "total_clicks": 0, "campaigns_count": 0}
    
    campaign_ids = []
    for camp in cpc_campaigns:
        camp_id = camp.get("id") or camp.get("campaign_id")
        if camp_id:
            try:
                campaign_ids.append(int(camp_id))
            except (ValueError, TypeError):
                continue
    
    if not campaign_ids:
        print("⚠️ Не удалось извлечь ID кампаний")
        return {"total_cost": 0.0, "total_clicks": 0, "campaigns_count": 0}
    
    stats = get_campaign_stats_for_month(session, perf_client_id, perf_token, campaign_ids, date_from, date_to)
    print(f"✅ Найдено CPC кампаний: {stats['campaigns_count']}, затрат: {stats['total_cost']:.2f} ₽, кликов: {stats['total_clicks']}")
    return stats


def get_campaigns_data_for_excel(session: requests.Session, date_from: str, date_to: str) -> Optional[List[Dict[str, Any]]]:
    """
    Получает данные об активных кампаниях за месяц для вывода в Excel.
    
    Возвращает список словарей с данными для таблицы:
    - ID кампании
    - Название кампании
    - Состояние
    - Тип оплаты
    - Бюджет
    - Дневной бюджет
    - Расход (за период)
    - Показы
    - Клики
    - CTR
    - Средняя цена клика
    - Заказы (шт.)
    - Заказы (руб.)
    - ДРР (доля рекламных расходов)
    """
    perf_client_id, perf_api_key = get_performance_api_credentials()
    if not perf_client_id or not perf_api_key:
        return None
    
    perf_token = get_performance_token(session, perf_client_id, perf_api_key)
    if not perf_token:
        return None
    
    # Получаем активные кампании со статистикой
    campaigns_with_stats = get_active_campaigns_with_statistics(session, perf_client_id, perf_token, date_from, date_to)
    
    if not campaigns_with_stats:
        return []
    
    # Формируем данные для Excel
    excel_data = []
    for camp in campaigns_with_stats:
        # Базовые данные кампании
        camp_id = str(camp.get("id") or camp.get("campaign_id") or "")
        title = str(camp.get("title") or "")
        # API статистики возвращает "status", а список кампаний - "state"
        state = str(camp.get("status") or camp.get("state") or "")
        payment_type = str(camp.get("paymentType") or "")
        # API статистики возвращает "objectType", а список кампаний - "advObjectType"
        adv_object_type = str(camp.get("objectType") or camp.get("advObjectType") or "")
        
        # Конвертируем строки в числа (заменяем запятые на точки для float)
        def parse_number(value):
            """Парсит число из строки, заменяя запятые на точки"""
            if not value:
                return 0
            if isinstance(value, (int, float)):
                return float(value)
            # Заменяем запятые на точки и убираем пробелы
            value_str = str(value).replace(",", ".").replace(" ", "")
            try:
                return float(value_str)
            except (ValueError, TypeError):
                return 0
        
        def parse_int(value):
            """Парсит целое число из строки"""
            if not value:
                return 0
            if isinstance(value, int):
                return value
            if isinstance(value, float):
                return int(value)
            # Заменяем запятые на точки для парсинга
            value_str = str(value).replace(",", ".").replace(" ", "")
            try:
                return int(float(value_str))
            except (ValueError, TypeError):
                return 0
        
        # Бюджеты (в миллионных долях рубля, конвертируем в рубли)
        # API возвращает бюджеты как строки с запятыми (например, "8000,00")
        budget_str = camp.get("budget") or 0
        daily_budget_str = camp.get("dailyBudget") or 0
        weekly_budget_str = camp.get("weeklyBudget") or 0
        
        # Парсим бюджеты (они уже в рублях, не в миллионных долях)
        budget_rub = parse_number(budget_str)
        daily_budget_rub = parse_number(daily_budget_str)
        weekly_budget_rub = parse_number(weekly_budget_str)
        
        # Статистика (из ответа API статистики)
        # Маппинг полей API -> Excel:
        # - moneySpent -> Расход за период (руб.)
        # - views -> Показы
        # - clicks -> Клики
        # - orders -> Заказы (шт.)
        # - ordersMoney -> Заказы (руб.)
        # - ctr -> CTR (%)
        # - clickPrice -> Средняя цена клика (руб.)
        # - drr -> ДРР (%)
        # Значения приходят как строки, могут быть с запятыми вместо точек (например, "81,83")
        
        # Расход - API использует "moneySpent" (приоритет)
        cost_str = camp.get("moneySpent") or camp.get("cost") or camp.get("spent") or camp.get("expenses") or 0
        # Показы - API использует "views" (приоритет)
        impressions_str = camp.get("views") or camp.get("impressions") or 0
        # Клики - API использует "clicks"
        clicks_str = camp.get("clicks") or 0
        # Заказы (количество) - API использует "orders"
        orders_count_str = camp.get("orders") or 0
        # Заказы (сумма) - API использует "ordersMoney" (приоритет)
        orders_sum_str = camp.get("ordersMoney") or camp.get("ordersSum") or camp.get("ordersRevenue") or 0
        # CTR - API использует "ctr" (может быть уже в процентах или как десятичная дробь)
        ctr_str = camp.get("ctr") or 0
        # Средняя цена клика - API использует "clickPrice" (приоритет)
        click_price_str = camp.get("clickPrice") or 0
        
        cost = parse_number(cost_str)
        impressions = parse_int(impressions_str)
        clicks = parse_int(clicks_str)
        orders_count = parse_int(orders_count_str)
        orders_sum = parse_number(orders_sum_str)
        
        # CTR и средняя цена клика из API (если есть)
        ctr_from_api = parse_number(ctr_str)
        avg_cpc_from_api = parse_number(click_price_str)
        
        # Рассчитываем производные показатели
        # Используем CTR из API, если есть, иначе рассчитываем
        if ctr_from_api > 0:
            ctr = ctr_from_api
            # Если CTR меньше 1, возможно это десятичная дробь (0.5 = 0.5%), умножаем на 100
            if ctr < 1:
                ctr = ctr * 100
        else:
            ctr = (clicks / impressions * 100) if impressions > 0 else 0.0
        
        # Используем среднюю цену клика из API, если есть, иначе рассчитываем
        if avg_cpc_from_api > 0:
            avg_cpc = avg_cpc_from_api
        else:
            avg_cpc = (cost / clicks) if clicks > 0 else 0.0
        
        # ДРР (доля рекламных расходов) - может быть в API как "drr"
        drr_str = camp.get("drr") or 0
        if drr_str:
            drr = parse_number(drr_str)
            # Если drr меньше 1, возможно это десятичная дробь, умножаем на 100
            if drr < 1:
                drr = drr * 100
        else:
            drr = (cost / orders_sum * 100) if orders_sum > 0 else 0.0
        
        excel_data.append({
            "ID кампании": camp_id,
            "Название кампании": title,
            "Состояние": state,
            "Тип оплаты": payment_type,
            "Тип объекта": adv_object_type,
            "Бюджет (руб.)": budget_rub,
            "Дневной бюджет (руб.)": daily_budget_rub,
            "Недельный бюджет (руб.)": weekly_budget_rub,
            "Расход за период (руб.)": cost,
            "Показы": impressions,
            "Клики": clicks,
            "CTR (%)": round(ctr, 2),
            "Средняя цена клика (руб.)": round(avg_cpc, 2),
            "Заказы (шт.)": orders_count,
            "Заказы (руб.)": orders_sum,
            "ДРР (%)": round(drr, 2) if drr > 0 else 0.0
        })
    
    return excel_data
