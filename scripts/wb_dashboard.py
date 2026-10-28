"""«Бизнес-сводка» Wildberries — данные и блоки для общего экрана сводки (тот же, что у Ozon, см. app_flet.py).

Экран, шапка («Месяц», «Год», «Обновить», «…») и «Динамика» общие для обоих магазинов; этот модуль отвечает за
то, что у WB своё: чтение сохранённых финансовых отчётов (wb reports/<ГГГГ-ММ>.json), загрузку месяца из WB
(«Обновить»), сверку выплат, показатели месяца и блоки «Статусы заказов и возвраты», «Приёмка, хранение и воронка».

Отдельной кнопки пересчёта нет: показатели считаются заново из сохранённого отчёта и wb_costs.xlsx при каждом
открытии сводки и выборе месяца, а Excel перерисовывается, если себестоимость или сверка выплат новее его.
"""
import asyncio
from datetime import date, datetime
import queue
import threading
from types import SimpleNamespace

import flet as ft
from scripts import wb_acceptance
from scripts import wb_finance as finance
from scripts import wb_funnel
from scripts import wb_monthly_report
from scripts import wb_orders
from scripts import wb_report_design as design
from scripts import wb_returns
from scripts import wb_settlement
from scripts import wb_storage
from scripts.marketplace_settings import read_settings
from scripts.report_paths import wb_report_path

HELP = {
    "WB Выручка": "Сколько заплатили покупатели за проданные товары (уже со скидкой WB) по финансовым отчётам WB за месяц.",
    "WB Чистая прибыль": "«К выплате» от WB минус себестоимость из wb_costs.xlsx. Комиссия, логистика, хранение и прочие "
                         "удержания WB уже вычтены из выплаты. Налоги и внешние расходы не входят.",
    "WB Маржа": "Чистая прибыль / выручка.",
    "WB Средний чек": "Выручка / число заказов с продажей.",
    "WB Скидка": "Средняя скидка WB покупателю (СПП). Её оплачивает WB из своей комиссии — на вашу выплату не влияет.",
    "WB Комиссия": "Сколько WB удержал из ваших цен со скидкой продавца по ставке комиссии (у магазина 38–41,5% "
                   "для досок, ниже для икон и оберегов). Процент справа — средняя ставка от ваших цен, а не доля в выручке.",
    "WB Скидка покупателям": "Скидка WB покупателям (СПП). Покупатель платит меньше вашей цены, а разницу WB оплачивает "
                             "из своей комиссии — поэтому в расчёте она идёт плюсом. Комиссия минус эта скидка — "
                             "сколько WB заработал на самом деле.",
    "WB Логистика": "Доставка к покупателю и возвраты, в том числе отменённых заказов (туда и обратно).",
    "WB Вне заказов": "Расходы, не привязанные к заказам с продажей этого месяца: логистика отмен, хранение, досрочный вывод. "
                      "Подробно — лист «Расходы вне заказов» в отчёте.",
    "WB Начислено": "«Итого к оплате» по финансовым отчётам WB за месяц.",
    "WB Закрыто": "Сколько из начисленного за месяц WB уже вывел вам на счёт. Выведено = всё начисленное по отчётам с начала "
                  "магазина − текущий баланс; деньги закрывают месяцы по порядку, сначала себестоимость.",
    "WB Баланс": "Баланс продавца на WB на момент сверки и сколько из него можно вывести сейчас.",
    "WB Статусы": "Order Feed и отчёт о возвратах WB — снимок за последние 31 день, не привязан к выбранному месяцу.",
    "WB Приёмка": "Отдельные отчёты WB за выбранный месяц. Хранение грузится кусками по 8 дней — это может занять несколько минут.",
}


def _rub(value):
    return float(value) if value is not None else None


class WbDashboard:
    """ui — помощники оформления сводки из app_flet: dpanel, dtext, dash_stat, dash_stats_panel, dash_list,
    dash_section_title, dash_pal, money, percent, number, primary. set_sync(active, text) — заглушка над показателями
    на время загрузки (как у Ozon), on_change() — перерисовать сводку, reveal(path, label) — открыть файл."""

    def __init__(self, *, root, page, months, ui, log, reveal, set_sync, on_change, toast):
        self.root, self.page, self.months, self.ui = root, page, months, ui
        self.log, self.reveal, self.set_sync, self.on_change, self.toast = log, reveal, set_sync, on_change, toast
        self.period = None
        self.data = self.summary = self.rows = self.settlement = None
        self.error = None
        self.busy = False
        self.cancel_event = threading.Event()
        self.orders = SimpleNamespace(busy=False, cancel=threading.Event(), summary=None, returns=None,
                                      by_srid={}, text="Ещё не загружены.")
        self.details = SimpleNamespace(busy=False, cancel=threading.Event(), period=None, acceptance=None,
                                       storage=None, funnel=None, text="Ещё не загружены за этот месяц.")

    # ── данные ───────────────────────────────────────────────────────────

    def write_log(self, message):
        self.log(f"[{datetime.now():%H:%M:%S}] [WB] {message}")

    def title(self, year=None, month=None):
        year, month = (year, month) if year else self.period
        return f"{self.months[month - 1]} {year}"

    def load(self, year, month):
        """Прочитать сохранённый отчёт месяца, пересчитать показатели и при необходимости перерисовать Excel."""
        self.period = (year, month)
        self.data = self.summary = self.rows = None
        self.error = None
        try:
            data = finance.load_report(self.root, year, month)
            costs = finance.read_costs(self.root)
            self.settlement = wb_settlement.settle(self.root, costs)
            if data and data.get("reports"):
                self.data, self.summary = data, finance.summarize(data, costs)
                self.rows = design.orders(data, costs)
                self._refresh_excel_if_stale(data, costs)
        except Exception as exc:
            self.error = str(exc) if isinstance(exc, finance.FinanceError) else \
                "Не удалось прочитать отчёт WB или wb_costs.xlsx. Закройте файлы в Excel и откройте сводку снова."
            self.write_log(f"Ошибка чтения: {self.error}")

    def _refresh_excel_if_stale(self, data, costs):
        path = wb_report_path(self.root, data["year"], data["month"])
        sources = [finance.report_path(self.root, data["year"], data["month"]), self.root / "wb_costs.xlsx",
                   self.root / wb_settlement.SETTLEMENT_FILE]
        newest = max((p.stat().st_mtime for p in sources if p.exists()), default=0)
        if path.exists() and path.stat().st_mtime >= newest:
            return
        try:
            wb_monthly_report.refresh_month(self.root, data, costs, self.settlement)
            self.write_log(f"Excel-отчёт WB обновлён по себестоимости и сверке выплат: {path.name}")
        except Exception as exc:
            self.write_log(f"Не удалось обновить Excel-отчёт WB (закройте его в Excel): {exc}")

    def freshness(self):
        if not self.period:
            return "—"
        if not self.data:
            return f"{self.title()} · отчёта ещё нет"
        fetched = str(self.data.get("fetched_at") or "")
        fresh = f" · обновлено {fetched[8:10]}.{fetched[5:7]} в {fetched[11:16]}" if len(fetched) >= 16 else ""
        return self.title() + fresh

    def report_path(self, year, month):
        return wb_report_path(self.root, year, month)

    def chart_data(self, periods):
        """(подписи, выручка, прибыль, маржа %, время самого свежего источника) по сохранённым месяцам WB."""
        labels, revenue, profit, margin, newest = [], [], [], [], 0.0
        costs = finance.read_costs(self.root)
        for month, year in periods:
            data = finance.load_report(self.root, year, month)
            if not data or not data.get("reports"):
                continue
            summary = finance.summarize(data, costs)
            if summary.get("empty"):
                continue
            newest = max(newest, finance.report_path(self.root, year, month).stat().st_mtime)
            labels.append(f"{self.months[month - 1][:3]} {str(year)[-2:]}")
            revenue.append(float(summary["revenue"] or 0))
            profit.append(float(summary["profit"]) if summary["profit"] is not None else 0.0)
            margin.append(float(summary["margin"]) if summary["margin"] is not None else 0.0)
        return labels, revenue, profit, margin, newest

    def open_costs(self):
        try:
            self.reveal(finance.ensure_costs(self.root), "Себестоимость WB")
        except Exception:
            self.toast("Не удалось открыть wb_costs.xlsx. Закройте файл в Excel и повторите.", True)

    # ── «Обновить»: загрузка месяца из WB ─────────────────────────────────

    def cancel(self):
        if self.busy and not self.cancel_event.is_set():
            self.write_log("Запрошена отмена загрузки WB. Ожидание завершения текущего запроса.")
            self.cancel_event.set()

    async def _run(self, work, progress, on_message):
        task = asyncio.create_task(asyncio.to_thread(work))
        while not task.done():
            self._drain(progress, on_message)
            await asyncio.sleep(0.25)
        self._drain(progress, on_message)
        return await task

    def _drain(self, progress, on_message):
        while True:
            try:
                message = progress.get_nowait()
            except queue.Empty:
                return
            self.write_log(message)
            on_message(message)

    async def download(self, year, month):
        if self.busy:
            self.toast("Загрузка WB уже идёт.", True)
            return
        self.busy = True
        self.cancel_event.clear()
        progress = queue.Queue()
        sync = lambda text: self.set_sync(True, text)  # noqa: E731
        sync(f"Обновляю отчёт: загружаю продажи, расходы и выплаты из WB за {self.title(year, month).lower()}…")
        self.write_log(f"Начало загрузки финансовых отчётов: {self.title(year, month)}.")
        token = read_settings(self.root).get("WB_API_TOKEN") or ""

        def fetch():
            client = finance.FinanceClient(token, self.cancel_event, progress.put)
            try:
                return client.fetch(year, month)
            finally:
                client.session.close()

        def settle():
            client = finance.FinanceClient(token, self.cancel_event, progress.put)
            try:
                progress.put("Сверка выплат WB: итоги отчётов за всё время и баланс…")
                wb_settlement.refresh(self.root, client)
            finally:
                client.session.close()

        try:
            data = await self._run(fetch, progress, sync)
            if self.cancel_event.is_set():
                raise finance.FinanceError("Загрузка отменена; предыдущий отчёт сохранён.")
            if (read_settings(self.root).get("WB_API_TOKEN") or "") != token:
                raise finance.FinanceError("Токен WB изменился во время загрузки. Повторите запрос.")
            finance.summarize(data, finance.read_costs(self.root))
            finance.ensure_costs(self.root, data["details"])
            finance.save_report(self.root, data)
            try:
                await self._run(settle, progress, sync)
            except Exception as exc:  # сбой сверки не отменяет сохранённый месяц
                self.write_log(f"Сверка выплат WB не обновлена: {exc}")
            sync("Формирую Excel-отчёт WB…")
            await asyncio.to_thread(self._export, data)
            await asyncio.to_thread(wb_monthly_report.refresh_saved, self.root)
            self.write_log(f"Загрузка завершена: {len(data['reports'])} отчётов, {len(data['details'])} строк.")
            self.toast(f"Отчёт WB за {self.title(year, month).lower()} обновлён.")
        except Exception as exc:
            reason = str(exc) if isinstance(exc, finance.FinanceError) else \
                "Не удалось загрузить или сохранить WB. Проверьте доступ к файлам и повторите."
            self.write_log(f"{'Отмена' if self.cancel_event.is_set() else 'Ошибка загрузки'}: {reason}")
            self.toast(reason, True)
        finally:
            self.busy = False
            self.set_sync(False, "")
            if self.period == (year, month):
                self.load(year, month)
            self.on_change()

    def _export(self, data):
        same = self.details.period == (data["year"], data["month"])
        costs = finance.read_costs(self.root)
        path = wb_monthly_report.build(self.root, data, costs, self.orders.by_srid,
                                       self.details.acceptance if same else None,
                                       self.details.storage if same else None,
                                       self.details.funnel if same else None)
        self.write_log(f"Excel-отчёт WB обновлён: {path.name}")

    # ── статусы заказов и возвраты (31 день) ──────────────────────────────

    async def load_statuses(self):
        state = self.orders
        if state.busy:
            return
        state.busy = True
        state.cancel.clear()
        progress = queue.Queue()
        token = read_settings(self.root).get("WB_API_TOKEN") or ""

        def message(text):
            state.text = text
            self.on_change()

        message("Подключение к Order Feed WB…")

        def fetch_orders():
            client = wb_orders.OrderFeedClient(token, state.cancel, progress.put)
            try:
                return client.fetch()
            finally:
                client.session.close()

        def fetch_returns():
            client = wb_returns.ReturnsClient(token, state.cancel, progress.put)
            try:
                return client.fetch()
            finally:
                client.session.close()

        try:
            orders_data = await self._run(fetch_orders, progress, message)
            state.summary = wb_orders.summarize(orders_data)
            returns = wb_returns.summarize(await self._run(fetch_returns, progress, message))
            state.returns, state.by_srid = returns, returns["by_srid"]
            state.text = (f"Снимок на {orders_data['fetched_at']} · заказов: {state.summary['total']} · "
                          f"возвратов: {returns['total']}")
            self.write_log(state.text)
            if self.data:
                await asyncio.to_thread(self._export, self.data)
        except Exception as exc:
            known = (wb_orders.OrderFeedError, wb_returns.ReturnsError)
            state.text = str(exc) if isinstance(exc, known) else "Не удалось загрузить статусы заказов или возвраты WB."
            self.write_log(f"Ошибка: {state.text}")
        finally:
            state.busy = False
            self.on_change()

    # ── приёмка, хранение, воронка за месяц ───────────────────────────────

    async def load_details(self):
        state = self.details
        if state.busy or not self.period:
            return
        state.busy = True
        state.cancel.clear()
        year, month = self.period
        progress = queue.Queue()
        token = read_settings(self.root).get("WB_API_TOKEN") or ""
        first, last = (date.fromisoformat(d) for d in finance.period_dates(year, month))

        def message(text):
            state.text = text
            self.on_change()

        message(f"Загружаю приёмку, хранение и воронку за {self.title(year, month).lower()}…")

        def fetcher(client_class):
            def run():
                client = client_class(token, state.cancel, progress.put)
                try:
                    return client.fetch(first, last)
                finally:
                    client.session.close()
            return run

        try:
            acceptance = wb_acceptance.summarize(await self._run(fetcher(wb_acceptance.AcceptanceClient), progress, message))
            storage = wb_storage.summarize(await self._run(fetcher(wb_storage.StorageClient), progress, message))
            funnel = wb_funnel.summarize(await self._run(fetcher(wb_funnel.FunnelClient), progress, message))
            state.acceptance, state.storage, state.funnel, state.period = acceptance, storage, funnel, (year, month)
            state.text = f"{self.title(year, month)} · загружено"
            self.write_log(f"Приёмка, хранение и воронка WB загружены: {self.title(year, month)}.")
            if self.data and self.period == (year, month):
                await asyncio.to_thread(self._export, self.data)
        except Exception as exc:
            known = (wb_acceptance.AcceptanceError, wb_storage.StorageError, wb_funnel.FunnelError)
            state.text = str(exc) if isinstance(exc, known) else "Не удалось загрузить приёмку, хранение или воронку WB."
            self.write_log(f"Ошибка: {state.text}")
        finally:
            state.busy = False
            self.on_change()

    # ── показатели для экрана ─────────────────────────────────────────────

    def controls(self, launch_report):
        ui = self.ui
        pal = ui.dash_pal()
        if self.error:
            return [ui.dpanel(ui.dtext(self.error, 13, pal["bad"]))]
        if not self.data:
            return [ui.dpanel(
                ft.Column([
                    ft.Icon(ft.Icons.INSERT_CHART_OUTLINED, size=44, color=pal["muted"]),
                    ui.dtext(f"Отчёта за {self.title().lower()} ещё нет", 18, weight=ft.FontWeight.W_700),
                    ui.dtext("Сформируйте отчёт — сводка заполнится автоматически.", 13, pal["muted"]),
                    ft.Container(height=6),
                    ft.Button("Сформировать отчёт", icon=ft.Icons.REFRESH, bgcolor=ui.primary, color="white",
                              on_click=lambda _e: launch_report()),
                ], spacing=6, horizontal_alignment=ft.CrossAxisAlignment.CENTER, tight=True),
                padding=48, alignment=ft.Alignment(0, 0),
            )] + self.extra_sections()
        summary, rows = self.summary, self.rows
        money, percent, number = ui.money, ui.percent, ui.number
        revenue = _rub(summary["revenue"]) or 0.0
        profit = _rub(summary["profit"])
        margin = _rub(summary["margin"])
        stats = design.order_stats(self.data, rows)
        lines = dict(design.profit_lines(summary, rows))
        outside = sum(i["amount"] for i in design.charges(self.data))

        def share(value):
            return percent(abs(value) / revenue * 100) if value and revenue else ""

        margin_color = None if margin is None else pal["good"] if margin >= 20 else pal["warn"] if margin >= 5 else pal["bad"]
        orders_rows = [("Продано", number(stats["sales"]), "", None, None)]
        orders_rows += [(f"  {name}", number(count), "", None, None) for name, count in stats["by_scheme"].items()]
        orders_rows += [("Возвраты", number(stats["returns"]), "", pal["bad"] if stats["returns"] else None, None),
                        ("Скидка WB покупателю", percent(stats["spp"] * 100) if stats["spp"] is not None else "—", "",
                         None, "WB Скидка")]
        cost = _rub(summary["costs"])
        economy = [
            ("Себестоимость", money(cost) if cost is not None else "—", share(cost), None, None),
            ("Валовая прибыль", money(_rub(summary["gross"])) if summary["gross"] is not None else "—", "", None, None),
            ("Операционные расходы", money(_rub(summary["operating"])) if summary["operating"] is not None else "—", "", None, None),
            ("Чистая прибыль", money(profit) if profit is not None else "—", "",
             None if profit is None else pal["good"] if profit >= 0 else pal["bad"], "WB Чистая прибыль"),
        ]
        commission_key = design.commission_label(rows)
        rate = design.commission_rate(rows)
        expense_keys = [("Скидка WB", "Скидка WB покупателям (оплатил WB)", "WB Скидка покупателям"),
                        ("Комиссия", commission_key, "WB Комиссия"), ("Эквайринг", "Эквайринг", None),
                        ("Логистика", "Логистика", "WB Логистика"), ("Хранение", "Хранение на складах WB", None),
                        ("Приёмка", "Приёмка поставок", None), ("Штрафы и удержания", None, None),
                        ("Досрочный вывод", "Досрочный вывод денег", None)]
        expenses = []
        for label, key, help_key in expense_keys:
            value = (lines["Штрафы"] + lines["Удержания"]) if key is None else lines[key]
            if abs(value) >= 0.005 or label in ("Комиссия", "Логистика"):
                note = share(value)
                if label == "Комиссия" and rate is not None:
                    note = percent(rate * 100)  # ставка от ваших цен, а не доля в выручке — см. подсказку
                elif label == "Скидка WB":
                    note = "оплатил WB"
                expenses.append((label, money(value), note, pal["good"] if value > 0 else None, help_key))
        expenses.append(("в т. ч. вне заказов", money(outside), "", None, "WB Вне заказов"))
        payouts = self._payout_rows(pal)
        groups = [("Заказы", orders_rows), ("Экономика", economy), ("Расходы WB", expenses), ("Выплаты WB", payouts)]
        rows_total = max(len(group_rows) for _title, group_rows in groups)
        controls = [
            ui.dash_stats_panel([
                self._tip(ui.dash_stat("Выручка", money(revenue), f"{number(stats['sales'])} заказов"), "WB Выручка"),
                self._tip(ui.dash_stat("Чистая прибыль", money(profit) if profit is not None else "—",
                                       "после всех расходов" if profit is not None else "не посчитана — см. ниже",
                                       None if profit is None else pal["good"] if profit >= 0 else pal["bad"]), "WB Чистая прибыль"),
                self._tip(ui.dash_stat("Маржа", percent(margin) if margin is not None else "—", "чистая прибыль / выручка",
                                       margin_color), "WB Маржа"),
                self._tip(ui.dash_stat("Средний чек", money(_rub(summary["average"])), "выручка / заказы"), "WB Средний чек"),
            ]),
            ft.ResponsiveRow([ui.dash_list(title, group_rows, rows_total) for title, group_rows in groups],
                             spacing=16, run_spacing=16),
        ]
        notes = list(summary.get("notes") or [])
        losses = [o for o in rows if o["profit"] is not None and o["profit"] < 0]
        if losses:
            articles = ", ".join(sorted({o["article"] for o in losses}))
            notes.append(f"Убыточных заказов: {len(losses)} на {money(-sum(o['profit'] for o in losses))} "
                         f"(артикулы {articles}) — цена продажи не покрывает комиссию, логистику и себестоимость. "
                         "Подробно — лист «Заказы» в отчёте.")
        if notes:
            controls.append(ui.dpanel(ft.Column([ui.dtext(note, 13, pal["warn"]) for note in notes], spacing=6)))
        return controls + self.extra_sections()

    def _tip(self, control, key):
        control.tooltip = HELP.get(key)
        return control

    def _payout_rows(self, pal):
        money = self.ui.money
        month = (self.settlement or {}).get("months", {}).get(f"{self.period[0]}-{self.period[1]:02d}")
        payout = _rub(self.summary["payout"])
        if not month:
            return [("Начислено", money(payout), "", None, "WB Начислено"),
                    ("Сверка выплат", "ещё не проводилась", "", pal["muted"], "WB Закрыто")]
        due = sum(month["due"].values()) / 100
        closed = sum(month["closed"].values()) / 100
        left = month["left"] / 100
        return [
            ("Начислено", money(due), "", None, "WB Начислено"),
            ("Закрыто выплатами", money(closed), "", pal["good"] if month["done"] else None, "WB Закрыто"),
            ("Осталось получить", money(left), "", pal["good"] if month["done"] else pal["warn"], "WB Закрыто"),
            ("На балансе WB", money(self.settlement["balance"] / 100), "", None, "WB Баланс"),
            ("Можно вывести", money(self.settlement["for_withdraw"] / 100), "", None, "WB Баланс"),
        ]

    def extra_sections(self):
        ui = self.ui
        pal = ui.dash_pal()
        orders, details = self.orders, self.details
        status_rows = []
        if orders.summary:
            for title, bucket in (("Всего", orders.summary), ("FBS (свой склад)", orders.summary["fbs"]),
                                  ("FBO (склад WB)", orders.summary["fbo"])):
                status_rows.append((title, self.ui.number(bucket["total"]),
                                    f"в пути {bucket['in_progress']} · доставлено {bucket['delivered']} · "
                                    f"отменено {bucket['cancelled']}", None, None))
        if orders.returns and orders.returns["total"]:
            reasons = sorted(orders.returns["by_reason"].items(), key=lambda item: item[1], reverse=True)[:3]
            status_rows.append(("Возвраты", self.ui.number(orders.returns["total"]),
                                ", ".join(f"{reason} ({count})" for reason, count in reasons), pal["bad"], None))
        same = details.period == self.period
        detail_rows = []
        if same and details.acceptance is not None:
            detail_rows.append(("Приёмка", ui.money(_rub(details.acceptance["total"])),
                                f"товаров {len(details.acceptance['by_nm'])}", None, None))
        if same and details.storage is not None:
            detail_rows.append(("Хранение", ui.money(_rub(details.storage["total"])),
                                f"товаров {len(details.storage['by_nm'])}", None, None))
        if same and details.funnel is not None:
            f = details.funnel
            detail_rows += [("Просмотры карточек", ui.number(f["views"]), "", None, None),
                            ("В корзину", ui.number(f["cart"]), f"{f['add_to_cart_pct']}%", None, None),
                            ("Заказы", ui.number(f["orders"]), f"{f['cart_to_order_pct']}% из корзины", None, None),
                            ("Выкупы", ui.number(f["buyouts"]), f"{f['buyout_pct']}%", None, None)]

        def block(title, caption, help_key, state, rows, on_click):
            button = ft.Button("Обновить", icon=ft.Icons.REFRESH, bgcolor=ui.primary, color="white",
                               disabled=state.busy or self.busy, on_click=on_click)
            head = ft.Row([ui.dtext(state.text, 13, pal["muted"], expand=True),
                           ft.ProgressRing(width=18, height=18, stroke_width=2.2, visible=state.busy), button],
                          vertical_alignment=ft.CrossAxisAlignment.CENTER)
            body = [head] + ([ui.dash_list("", rows, len(rows), col={"xs": 12})] if rows else [])
            panel = ui.dpanel(ft.Column(body, spacing=10))
            panel.tooltip = HELP.get(help_key)
            return [ui.dash_section_title(title, caption), panel]

        return (block("Статусы заказов и возвраты", "Снимок WB за последние 31 день", "WB Статусы", orders, status_rows,
                      lambda _e: self.page.run_task(self.load_statuses))
                + block("Приёмка, хранение и воронка продаж", "Отдельные отчёты WB за выбранный месяц", "WB Приёмка",
                        details, detail_rows, lambda _e: self.page.run_task(self.load_details)))
