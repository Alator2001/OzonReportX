"""Wildberries business summary backed by financial reports."""
import asyncio
from datetime import date, datetime
from pathlib import Path
import queue
import threading

import flet as ft
from scripts import wb_acceptance
from scripts import wb_finance as finance
from scripts import wb_funnel
from scripts import wb_monthly_report
from scripts import wb_orders
from scripts import wb_returns
from scripts import wb_storage
from scripts.marketplace_settings import read_settings


def build_dashboard(*, months, metric, hero_metric, card, text_color,
                    muted_color, primary, accent, surface, page, root, reveal, theme, log):
    today = date.today()
    month = ft.Dropdown(label="Месяц сводки", value=str(today.month), width=230,
                        options=[ft.dropdown.Option(str(i), title) for i, title in enumerate(months, 1)])
    year = ft.Dropdown(label="Год сводки", value=str(today.year), width=170,
                       options=[ft.dropdown.Option(str(y)) for y in range(today.year, 2024, -1)])
    status = ft.Text("Выберите период и загрузите отчёты WB.", size=12, color=muted_color)
    cards = ft.ResponsiveRow(spacing=14, run_spacing=14)
    details = ft.ResponsiveRow(spacing=12, run_spacing=12)
    notes = ft.Column(spacing=6)
    cancel = threading.Event()
    last_report = {"path": None}
    chart = ft.Image(src="", visible=False, border_radius=18)
    chart_note = ft.Text("График строится по сохранённым месяцам выбранного года.", size=12)
    busy = {"value": False}
    order_cards = ft.ResponsiveRow(spacing=14, run_spacing=14)
    order_status = ft.Text("Статусы заказов и возвраты ещё не загружены.", size=12, color=muted_color)
    order_cancel = threading.Event()
    order_busy = {"value": False}
    returns_cards = ft.ResponsiveRow(spacing=14, run_spacing=14)
    returns_index = {"by_srid": {}}
    current_data = {"value": None}
    detail_index = {"acceptance": None, "storage": None, "funnel": None, "period": None}
    detail_cards = ft.ResponsiveRow(spacing=14, run_spacing=14)
    detail_status = ft.Text("Приёмка, хранение и воронка продаж ещё не загружены.", size=12, color=muted_color)
    detail_cancel = threading.Event()
    detail_busy = {"value": False}

    def write_log(message):
        log(f"[{datetime.now():%H:%M:%S}] [WB] {message}")

    def period_title():
        return f"{months[int(month.value) - 1]} {year.value}"

    def fmt(value, suffix=" ₽"):
        return "—" if value is None else f"{value:,.2f}".replace(",", " ") + suffix

    def render(data=None):
        result = finance.summarize(data, finance.read_costs(root)) if data else {"empty": True}
        cards.controls = [
            hero_metric("Выручка", fmt(result.get("revenue")), "По финансовым отчётам WB", surface, primary, col=4),
            hero_metric("Прибыль по отчёту", fmt(result.get("profit")), "Итог к выплате минус себестоимость", surface, primary, col=4),
            hero_metric("Себестоимость", fmt(result.get("costs")), "Отдельный справочник wb_costs.xlsx", surface, accent, col=4),
            metric("Net Margin", fmt(result.get("margin"), " %"), "Прибыль / выручка × 100%", surface, col=3),
            metric("Средний чек продажи", fmt(result.get("average")), "Сумма продаж / уникальные srid продаж", surface, col=3),
            metric("Продажи", str(result.get("sales", "—")), "Уникальные srid продаж за период", surface, col=3),
            metric("Возвраты", str(result.get("returns", "—")), "По операциям возврата WB", surface, col=3),
        ]
        def lines(items):
            return [ft.Text(f"{label}: {fmt(result.get(key))}", size=13) for label, key in items]
        details.controls = [
            card("Прибыль", "", lines([("Итог WB к выплате", "payout"), ("Валовая прибыль", "gross"),
                                       ("Операционные расходы", "operating")]), col=4),
            card("Доставка и хранение", "", lines([("Логистика", "logistics"), ("Хранение", "storage"),
                                                       ("Приёмка", "acceptance")]), col=4),
            card("Прочие операции WB", "", lines([("Удержания", "deductions"), ("Штрафы", "penalties"),
                                                     ("Доплаты", "additional")]), col=4),
        ]
        notes.controls = [ft.Text(message, color=accent, size=13) for message in result.get("notes", [])]
        current_data["value"] = data
        if data:
            status.value = f"{months[data['month'] - 1]} {data['year']} · Обновлено: {data['fetched_at']} · Отчётов: {len(data['reports'])}"
            if not result.get("empty"):
                export_report(data)
        theme(view)

    def export_report(data):
        same_period = detail_index["period"] == (data["year"], data["month"])
        try:
            report_file = wb_monthly_report.build(
                root, data, finance.read_costs(root), returns_index["by_srid"],
                detail_index["acceptance"] if same_period else None,
                detail_index["storage"] if same_period else None,
                detail_index["funnel"] if same_period else None)
            last_report["path"] = report_file
            open_report_button.disabled = False
            write_log(f"Excel-отчёт WB обновлён: {report_file.name}")
        except Exception as exc:
            write_log(f"Не удалось сформировать Excel-отчёт WB: {exc}")

    def refresh(_e=None):
        write_log(f"Пересчёт сводки из файлов: {period_title()}.")
        try:
            data = finance.load_report(root, int(year.value), int(month.value))
            render(data)
            if not data:
                status.value = "За выбранный месяц нет сохранённых отчётов этого токена WB."
            write_log(status.value)
            for note in notes.controls:
                write_log(f"Внимание: {note.value}")
            chart.visible = False
        except Exception as exc:
            status.value = str(exc) if isinstance(exc, finance.FinanceError) else "Не удалось прочитать отчёт или wb_costs.xlsx. Проверьте файлы."
            write_log(f"Ошибка пересчёта: {status.value}")
            cards.controls = []
            details.controls = []
            notes.controls = []
        page.update()

    async def download():
        if busy["value"]:
            write_log("Повторная загрузка не запущена: предыдущая операция ещё выполняется.")
            return
        busy["value"] = True
        cancel.clear()
        token = ""
        selected_year, selected_month = int(year.value), int(month.value)
        for control in (month, year, download_button, recalculate_button, costs_button, chart_button):
            control.disabled = True
        cancel_button.disabled = False
        progress = queue.Queue()
        status.value = "Подключение к финансовым отчётам WB…"
        write_log(f"Начало загрузки финансовых отчётов: {period_title()}.")
        page.update()
        def fetch():
            client = finance.FinanceClient(token, cancel, progress.put)
            try:
                return client.fetch(selected_year, selected_month)
            finally:
                client.session.close()
        def drain_progress():
            while True:
                try:
                    message = progress.get_nowait()
                except queue.Empty:
                    break
                status.value = message
                write_log(message)

        try:
            token = read_settings(root).get("WB_API_TOKEN") or ""
            task = asyncio.create_task(asyncio.to_thread(fetch))
            while not task.done():
                drain_progress()
                page.update()
                await asyncio.sleep(0.25)
            data = await task
            drain_progress()
            if cancel.is_set():
                raise finance.FinanceError("Загрузка отменена; предыдущий отчёт сохранён.")
            if (read_settings(root).get("WB_API_TOKEN") or "") != token:
                raise finance.FinanceError("Токен WB изменился во время загрузки. Повторите запрос.")
            finance.summarize(data, finance.read_costs(root))
            write_log("Проверка данных завершена. Обновление справочника wb_costs.xlsx.")
            finance.ensure_costs(root, data["details"])
            write_log("Сохранение финансового отчёта WB.")
            finance.save_report(root, data)
            render(data)
            write_log(f"Загрузка завершена: {len(data['reports'])} отчётов, {len(data['details'])} строк. {status.value}")
            for note in notes.controls:
                write_log(f"Внимание: {note.value}")
            chart.visible = False
        except Exception as exc:
            drain_progress()
            status.value = str(exc) if isinstance(exc, finance.FinanceError) else "Не удалось загрузить или сохранить WB. Проверьте доступ к файлам; повторите запрос."
            write_log(f"{'Отмена' if cancel.is_set() else 'Ошибка загрузки'}: {status.value}")
        finally:
            busy["value"] = False
            for control in (month, year, download_button, recalculate_button, costs_button, chart_button):
                control.disabled = False
            cancel_button.disabled = True
            page.update()

    def open_costs(_e):
        write_log("Открытие справочника себестоимости WB.")
        try:
            path = finance.ensure_costs(root)
            write_log("Справочник wb_costs.xlsx подготовлен. Передача команды открытия файла.")
            reveal(path, "Себестоимость WB")
        except Exception:
            status.value = "Не удалось открыть wb_costs.xlsx. Закройте файл в Excel и повторите."
            write_log(f"Ошибка: {status.value}")
            page.update()

    def open_report(_e):
        path = last_report["path"]
        if not path:
            return
        write_log(f"Открытие Excel-отчёта WB: {path.name}.")
        try:
            reveal(path, "Отчёт WB")
        except Exception:
            status.value = "Не удалось открыть Excel-отчёт WB. Закройте файл в Excel и повторите."
            write_log(f"Ошибка: {status.value}")
            page.update()

    def draw_chart(_e):
        write_log(f"Построение графика WB за {year.value} год.")
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        try:
            costs = finance.read_costs(root)
            points = []
            for m in range(1, 13):
                data = finance.load_report(root, int(year.value), m)
                if data:
                    summary = finance.summarize(data, costs)
                    if not summary["empty"]:
                        points.append((m, summary))
            if not points:
                chart.visible = False
                chart_note.value = "Нет сохранённых финансовых отчётов за выбранный год."
                write_log(chart_note.value)
            else:
                figure = Figure(figsize=(10, 3.6), tight_layout=True)
                FigureCanvasAgg(figure)
                axes = figure.subplots()
                axes.plot([p[0] for p in points], [float(p[1]["revenue"]) for p in points], marker="o", label="Выручка")
                axes.plot([p[0] for p in points], [float(p[1]["profit"]) if p[1]["profit"] is not None else float("nan") for p in points], marker="o", label="Прибыль по отчёту")
                axes.set_xticks([p[0] for p in points], [months[p[0]-1][:3] for p in points])
                axes.set_ylabel("Рубли")
                axes.grid(alpha=0.2)
                axes.legend()
                from uuid import uuid4
                target = Path(root) / ".cache" / "wb" / f"trend-{uuid4().hex}.png"
                target.parent.mkdir(parents=True, exist_ok=True)
                figure.savefig(target)
                chart.src = str(target.resolve())
                chart.visible = True
                chart_note.value = "По сохранённым месяцам; себестоимость пересчитана по текущему справочнику WB."
                write_log(f"График построен: {len(points)} месяцев.")
        except Exception:
            chart.visible = False
            chart_note.value = "Не удалось построить график. Проверьте отчёты и себестоимость WB."
            write_log(f"Ошибка: {chart_note.value}")
        page.update()

    def cancel_download(_e):
        if busy["value"] and not cancel.is_set():
            write_log("Запрошена отмена загрузки WB. Ожидание завершения текущего запроса.")
            cancel.set()

    def render_orders(summary=None):
        if not summary:
            order_cards.controls = []
            theme(view)
            return
        def group(title, bucket, col):
            return metric(title, str(bucket["total"]),
                          f"В пути: {bucket['in_progress']} · Доставлено: {bucket['delivered']} · Отменено: {bucket['cancelled']}",
                          surface, col=col)
        order_cards.controls = [
            group("Всего заказов", summary, 4),
            group("FBS (свой склад)", summary["fbs"], 4),
            group("FBO (склад WB)", summary["fbo"], 4),
        ]
        theme(view)

    def render_returns(summary=None):
        if not summary or not summary["total"]:
            returns_cards.controls = []
            theme(view)
            return
        reasons = sorted(summary["by_reason"].items(), key=lambda item: item[1], reverse=True)[:5]
        statuses = sorted(summary["by_status"].items(), key=lambda item: item[1], reverse=True)[:5]
        returns_cards.controls = [
            metric("Возвраты (движение товара)", str(summary["total"]),
                   "Топ причин: " + ", ".join(f"{reason} ({count})" for reason, count in reasons),
                   surface, col=6),
            metric("По статусу логистики", str(summary["total"]),
                   ", ".join(f"{name} ({count})" for name, count in statuses),
                   surface, col=6),
        ]
        theme(view)

    async def refresh_operational_data():
        if order_busy["value"]:
            write_log("Обновление статусов заказов и возвратов уже выполняется.")
            return
        order_busy["value"] = True
        order_cancel.clear()
        orders_button.disabled = True
        progress = queue.Queue()
        order_status.value = "Подключение к Order Feed WB…"
        write_log("Начало загрузки статусов заказов и возвратов WB (последние 31 день).")
        page.update()
        token = read_settings(root).get("WB_API_TOKEN") or ""
        def drain_progress():
            while True:
                try:
                    message = progress.get_nowait()
                except queue.Empty:
                    break
                order_status.value = message
                write_log(message)
        async def run(fetcher):
            task = asyncio.create_task(asyncio.to_thread(fetcher))
            while not task.done():
                drain_progress()
                page.update()
                await asyncio.sleep(0.25)
            drain_progress()
            return await task
        try:
            def fetch_orders_data():
                client = wb_orders.OrderFeedClient(token, order_cancel, progress.put)
                try:
                    return client.fetch()
                finally:
                    client.session.close()
            orders_data = await run(fetch_orders_data)
            order_summary = wb_orders.summarize(orders_data)
            render_orders(order_summary)
            write_log(f"Статусы заказов обновлены: заказов {order_summary['total']} "
                      f"(FBS: {order_summary['fbs']['total']}, FBO: {order_summary['fbo']['total']}).")

            def fetch_returns_data():
                client = wb_returns.ReturnsClient(token, order_cancel, progress.put)
                try:
                    return client.fetch()
                finally:
                    client.session.close()
            returns_data = await run(fetch_returns_data)
            returns_summary = wb_returns.summarize(returns_data)
            returns_index["by_srid"] = returns_summary["by_srid"]
            render_returns(returns_summary)
            write_log(f"Отчёт о возвратах обновлён: {returns_summary['total']} записей за 31 день.")

            order_status.value = (f"Снимок на {orders_data['fetched_at']} · заказов: {order_summary['total']} "
                                   f"· возвратов: {returns_summary['total']}")
            if current_data["value"]:
                write_log("Обновление Excel-отчёта причинами/статусами возвратов.")
                export_report(current_data["value"])
        except Exception as exc:
            drain_progress()
            known = (wb_orders.OrderFeedError, wb_returns.ReturnsError)
            order_status.value = str(exc) if isinstance(exc, known) else "Не удалось загрузить статусы заказов или возвраты WB."
            write_log(f"Ошибка загрузки статусов/возвратов: {order_status.value}")
        finally:
            order_busy["value"] = False
            orders_button.disabled = False
            page.update()

    def render_details():
        acceptance, storage, funnel = detail_index["acceptance"], detail_index["storage"], detail_index["funnel"]
        if not acceptance and not storage and not funnel:
            detail_cards.controls = []
            theme(view)
            return
        controls = []
        if acceptance is not None:
            controls.append(metric("Приёмка по товарам", fmt(acceptance["total"]),
                                   f"Строк отчёта: {acceptance['rows']} · товаров: {len(acceptance['by_nm'])}",
                                   surface, col=4))
        if storage is not None:
            controls.append(metric("Хранение по товарам", fmt(storage["total"]),
                                   f"Строк отчёта: {storage['rows']} · товаров: {len(storage['by_nm'])}",
                                   surface, col=4))
        if funnel is not None:
            controls.append(metric("Воронка продаж (органика)",
                                   f"{funnel['views']} просмотров",
                                   f"В корзину: {funnel['cart']} ({funnel['add_to_cart_pct']}%) · "
                                   f"Заказы: {funnel['orders']} ({funnel['cart_to_order_pct']}%) · "
                                   f"Выкупы: {funnel['buyouts']} ({funnel['buyout_pct']}%)",
                                   surface, col=4))
        detail_cards.controls = controls
        theme(view)

    async def fetch_month_details():
        if detail_busy["value"]:
            write_log("Загрузка приёмки/хранения/воронки уже выполняется.")
            return
        detail_busy["value"] = True
        detail_cancel.clear()
        selected_year, selected_month = int(year.value), int(month.value)
        for control in (month, year, download_button, recalculate_button, details_button):
            control.disabled = True
        detail_cancel_button.disabled = False
        progress = queue.Queue()
        detail_status.value = f"Подключение к отчётам WB о приёмке, хранении и воронке продаж: {period_title()}…"
        write_log(f"Начало загрузки приёмки, хранения и воронки продаж WB: {period_title()}.")
        page.update()
        token = read_settings(root).get("WB_API_TOKEN") or ""
        date_from, date_to = finance.period_dates(selected_year, selected_month)
        from datetime import date as date_cls
        date_from, date_to = date_cls.fromisoformat(date_from), date_cls.fromisoformat(date_to)
        def drain_progress():
            while True:
                try:
                    message = progress.get_nowait()
                except queue.Empty:
                    break
                detail_status.value = message
                write_log(message)
        async def run(fetcher):
            task = asyncio.create_task(asyncio.to_thread(fetcher))
            while not task.done():
                drain_progress()
                page.update()
                await asyncio.sleep(0.25)
            drain_progress()
            return await task
        try:
            def fetch_acceptance():
                client = wb_acceptance.AcceptanceClient(token, detail_cancel, progress.put)
                try:
                    return client.fetch(date_from, date_to)
                finally:
                    client.session.close()
            acceptance_data = await run(fetch_acceptance)
            acceptance_summary = wb_acceptance.summarize(acceptance_data)
            write_log(f"Приёмка WB загружена: {acceptance_summary['rows']} строк, {fmt(acceptance_summary['total'])}.")

            def fetch_storage():
                client = wb_storage.StorageClient(token, detail_cancel, progress.put)
                try:
                    return client.fetch(date_from, date_to)
                finally:
                    client.session.close()
            storage_data = await run(fetch_storage)
            storage_summary = wb_storage.summarize(storage_data)
            write_log(f"Хранение WB загружено: {storage_summary['rows']} строк, {fmt(storage_summary['total'])}.")

            def fetch_funnel():
                client = wb_funnel.FunnelClient(token, detail_cancel, progress.put)
                try:
                    return client.fetch(date_from, date_to)
                finally:
                    client.session.close()
            funnel_data = await run(fetch_funnel)
            funnel_summary = wb_funnel.summarize(funnel_data)
            write_log(f"Воронка продаж WB загружена: {funnel_summary['products']} товаров, "
                      f"просмотров {funnel_summary['views']}, заказов {funnel_summary['orders']}.")

            detail_index["acceptance"] = acceptance_summary
            detail_index["storage"] = storage_summary
            detail_index["funnel"] = funnel_summary
            detail_index["period"] = (selected_year, selected_month)
            render_details()
            detail_status.value = (f"{period_title()} · приёмка: {fmt(acceptance_summary['total'])} · "
                                   f"хранение: {fmt(storage_summary['total'])} · "
                                   f"воронка: {funnel_summary['views']} просмотров, {funnel_summary['orders']} заказов")
            if current_data["value"] and (current_data["value"]["year"], current_data["value"]["month"]) == (selected_year, selected_month):
                write_log("Обновление Excel-отчёта детализацией приёмки/хранения/воронки.")
                export_report(current_data["value"])
        except Exception as exc:
            drain_progress()
            known = (wb_acceptance.AcceptanceError, wb_storage.StorageError, wb_funnel.FunnelError)
            detail_status.value = str(exc) if isinstance(exc, known) else "Не удалось загрузить приёмку/хранение/воронку WB."
            write_log(f"Ошибка загрузки приёмки/хранения/воронки: {detail_status.value}")
        finally:
            detail_busy["value"] = False
            for control in (month, year, download_button, recalculate_button, details_button):
                control.disabled = False
            detail_cancel_button.disabled = True
            page.update()

    def cancel_details(_e):
        if detail_busy["value"] and not detail_cancel.is_set():
            write_log("Запрошена отмена загрузки приёмки/хранения WB.")
            detail_cancel.set()

    download_button = ft.Button("Загрузить из WB", on_click=lambda _e: page.run_task(download))
    recalculate_button = ft.TextButton("Пересчитать из файлов", on_click=refresh)
    costs_button = ft.TextButton("Себестоимость WB", on_click=open_costs)
    open_report_button = ft.TextButton("Открыть Excel-отчёт", disabled=True, on_click=open_report)
    chart_button = ft.TextButton("Построить график за год", on_click=draw_chart)
    cancel_button = ft.TextButton("Отменить загрузку", disabled=True, on_click=cancel_download)
    orders_button = ft.Button("Обновить статусы и возвраты", on_click=lambda _e: page.run_task(refresh_operational_data))
    details_button = ft.Button("Загрузить приёмку, хранение и воронку", on_click=lambda _e: page.run_task(fetch_month_details))
    detail_cancel_button = ft.TextButton("Отменить", disabled=True, on_click=cancel_details)
    month.on_select = refresh
    year.on_select = refresh
    view = ft.Column([
        ft.Text("Бизнес-сводка · Wildberries", size=26, weight=ft.FontWeight.W_800, color=text_color),
        ft.Row([month, year, download_button, cancel_button], wrap=True),
        ft.Row([recalculate_button, costs_button, open_report_button], wrap=True),
        status, notes, cards, details,
        card("Статусы заказов и возвраты (FBS + FBO)",
             "Order Feed и отчёт о возвратах WB — снимок за последние 31 день, не привязан к выбранному месяцу.", [
            ft.Row([orders_button], wrap=True), order_status, order_cards, returns_cards,
        ]),
        card("Приёмка, хранение и воронка продаж по товарам",
             "Отдельные отчёты WB за выбранный месяц (не за 31 день) — можно загрузить любой прошлый месяц. Загрузка хранения идёт кусками по 8 дней и может занять несколько минут.", [
            ft.Row([details_button, detail_cancel_button], wrap=True), detail_status, detail_cards,
        ]),
        card("Особенности расчёта WB", "", [
            ft.Text("Источник — ежедневные финансовые отчёты WB за выбранный период, а не дата создания заказа. Текущий месяц может быть неполным.", size=12),
            ft.Text("Прибыль по отчёту = итог WB к выплате − себестоимость. Удержания, логистика и другие суммы показываются отдельно и повторно не вычитаются. Внешние расходы и налоги не включены.", size=12),
            ft.Text("Продажа с отрицательной прибылью остаётся продажей. Продажа и её возврат в одном периоде имеют нулевую итоговую себестоимость. Для возвратов прошлых периодов расчёт прибыли приостановлен до выбора правила.", size=12),
            ft.Text("Статусы «В пути»/«Доставлено»/«Отменено» и причины возвратов берутся отдельно из Order Feed и отчёта о возвратах WB (последние 31 день) — в финансовом отчёте их нет. В Excel-отчёте причина и статус возврата проставляются только для заказов, которые попали в оба отчёта, и не влияют на суммы. Приёмка, хранение и воронка продаж по товарам — из отдельных отчётов WB за тот же месяц; суммы приёмки/хранения могут немного отличаться от строк «Приёмка»/«Хранение» выше. Воронка продаж — органический трафик (просмотры/корзина/заказы/выкупы), не реклама, ни на что не влияет. Рекламная аналитика пока не подключена: её расходы нельзя повторно вычитать без сверки с удержаниями WB.", size=12),
        ]),
        card("Динамика показателей", "", [chart_button, chart, chart_note]),
    ], spacing=18)
    try:
        render(finance.load_report(root, int(year.value), int(month.value)))
    except Exception:
        render()
        status.value = "Не удалось прочитать сохранённые данные WB. Проверьте отчёт и справочник себестоимости."
    return view
