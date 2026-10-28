"""Раздел «Акции» магазина WB: календарь акций с бустом и порогами участия и баланс «буст ↔ скидка ↔ маржа»
по каждому артикулу (scripts/wb_promotions.py). Только чтение. Оформление — как раздел «Цены» (wb_prices_view.py)."""
import asyncio
from datetime import date, datetime
import time

import flet as ft
from scripts import wb_prices, wb_promotions
from scripts.marketplace_settings import read_settings
from scripts.wb_prices_view import AUTO_SYNC_INTERVAL, GREEN, RED

YELLOW = "#FFEB9C"
PROMO_COLUMNS = [("Акция", 26), ("Тип", 8), ("Даты", 12), ("Когда", 10), ("Ваших в акции", 13), ("Буст", 7),
                 ("До следующего уровня", 16)]
ARTICLE_COLUMNS = [("Артикул", 7), ("Скидка сейчас", 7), ("Цена сейчас", 9), ("Скидка в акции", 7),
                   ("Цена в акции", 9), ("Статус", 10), ("Прибыль сейчас", 9), ("Прибыль в акции", 9),
                   ("Прибыль на минимуме", 9), ("Окупается при росте продаж", 9), ("Что делать", 14),
                   ("Переходов в день: со скидкой акции | с меньшей", 11), ("По факту", 14)]
SERIES_COLUMNS = [("Серия", 8), ("Артикулов", 8), ("Дней: со скидкой акции | с меньшей", 14),
                  ("Переходов в день: со скидкой акции | с меньшей", 16), ("Заказов в день: со скидкой акции | с меньшей", 16),
                  ("Переходов больше в", 10)]
# Цвет совета (wb_promotions.advice): красный — нужно действие, жёлтый — решить, зелёный — всё правильно.
TONE_FILL = {"act": RED, "consider": YELLOW, "ok": GREEN}


def per_day(value, digits=1):
    return "—" if value is None else f"{value:.{digits}f}".replace(".", ",")


def _date(text):
    try:
        return datetime.fromisoformat(str(text).replace("Z", "+00:00"))
    except ValueError:
        return None


class WbPromotionsView:
    """ui — как у WbPricesView: dpanel, dtext, dash_stat, dash_pal, money, percent, cell, primary, title_color, muted_color."""

    def __init__(self, *, root, page, ui, log, toast):
        self.root, self.page, self.ui, self.log, self.toast = root, page, ui, log, toast
        self.promotions = None
        self.rows = None
        self.error = None
        self.busy = False
        self.last_sync = 0.0
        self.show_all = False
        self.status = ft.Text("—", size=13, color=ui.muted_color)
        self.stats_area = ft.Container()
        self.promo_table = ft.Column(spacing=0)
        self.article_table = ft.Column(spacing=0)
        self.series_table = ft.Column(spacing=0)
        self.series_rows = None
        self.sync_label = ft.Text("Синхронизация с WB: календарь акций, цены и скидки…", size=13,
                                  weight=ft.FontWeight.W_600, text_align=ft.TextAlign.CENTER)
        self.overlay = ft.Container(
            left=0, top=0, right=0, bottom=0, visible=False, border_radius=18, alignment=ft.Alignment(0, -1),
            padding=ft.Padding.only(top=80), data="preserve_excel_fill",
            content=ft.Column([ft.ProgressRing(width=36, height=36, stroke_width=3), self.sync_label],
                              spacing=14, tight=True, horizontal_alignment=ft.CrossAxisAlignment.CENTER))
        header = ft.ResponsiveRow([
            ft.Container(col={"xs": 12, "lg": 6}, content=ft.Column(
                [ft.Text("Акции", size=26, weight=ft.FontWeight.W_800, color=ui.title_color), self.status], spacing=2)),
            ft.Container(col={"xs": 12, "lg": 6}, alignment=ft.Alignment(1, 0), content=ft.Row([
                ft.Button("Синхронизировать", icon=ft.Icons.SYNC, bgcolor=ui.primary, color="white",
                          tooltip="Обновит календарь акций, цены и скидки из WB. Запускается сама при входе в раздел",
                          on_click=lambda _e: self.page.run_task(self.sync)),
            ], spacing=10, alignment=ft.MainAxisAlignment.END)),
        ], vertical_alignment=ft.CrossAxisAlignment.CENTER, run_spacing=12)
        self.promo_panel = ui.dpanel(ft.Column([
            ft.Row([ui.dtext("Акции сейчас и скоро", 15, weight=ft.FontWeight.W_700),
                    ui.dtext("Буст получают все участники акции; его уровень зависит от доли ваших товаров с остатком "
                             "в акции. Выводя товар, вы можете снизить буст остальным.", 12, ui.muted_color,
                             expand=True, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS)],
                   spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            self.promo_table,
        ], spacing=8), padding=ft.Padding.only(left=20, right=20, top=12, bottom=8))
        self.article_panel = ui.dpanel(ft.Column([
            ft.Row([ui.dtext("Артикулы: буст или маржа", 15, weight=ft.FontWeight.W_700),
                    ui.dtext("«Окупается при росте продаж» — во сколько раз акция должна поднять продажи, чтобы "
                             "заработать столько же, сколько на минимальной цене. Наведите на строку — подробности.",
                             12, ui.muted_color, expand=True, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS)],
                   spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            self.article_table,
        ], spacing=8), padding=ft.Padding.only(left=20, right=20, top=12, bottom=8))
        self.series_panel = ui.dpanel(ft.Column([
            ft.Row([ui.dtext("Серии: что даёт скидка акции на деле", 15, weight=ft.FontWeight.W_700),
                    ui.dtext("Переходы в карточку и заказы в день — в дни со скидкой акции и с меньшей. У серии данных "
                             "больше, чем у одного артикула. WB отдаёт статистику по дням только за неделю, программа "
                             "копит её при каждой синхронизации.", 12, ui.muted_color, expand=True, max_lines=2,
                             overflow=ft.TextOverflow.ELLIPSIS)],
                   spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            self.series_table,
        ], spacing=8), padding=ft.Padding.only(left=20, right=20, top=12, bottom=8))
        self.view = ft.Column([header, self.stats_area, ft.Stack([ft.Column(
            [self.promo_panel, self.article_panel, self.series_panel], spacing=14), self.overlay])], spacing=14)

    def write_log(self, message):
        self.log(f"[{datetime.now():%H:%M:%S}] [WB] {message}")

    # ── данные ───────────────────────────────────────────────────────────

    def refresh(self, push_update=True):
        self.error = None
        try:
            self.promotions = wb_promotions.load(self.root)
            snapshot = wb_prices.load_snapshot(self.root)
            prices = wb_prices.compute(self.root, snapshot) if snapshot else None
            history, daily = wb_promotions.load_history(self.root), wb_promotions.load_daily(self.root)
            orders = (self.promotions or {}).get("order_discounts", {})
            self.rows = wb_promotions.article_rows(prices["rows"], history, orders, daily=daily) if prices else None
            self.series_rows = wb_promotions.series_rows(self.rows, history, orders, daily) if self.rows else None
        except Exception as exc:
            self.rows = None
            self.error = f"Не удалось посчитать акции WB: {exc}. Закройте wb_costs.xlsx в Excel и повторите."
        self.render(push_update)

    def enter(self):
        self.refresh(push_update=False)
        if not self.busy and time.monotonic() - self.last_sync >= AUTO_SYNC_INTERVAL:
            self.page.run_task(self.sync)

    async def sync(self):
        if self.busy:
            return
        token = read_settings(self.root).get("WB_API_TOKEN") or ""
        if not token:
            self.toast("Добавьте API-токен WB в настройках.", True)
            return
        self.busy = True
        self.last_sync = time.monotonic()
        self.set_sync(True)
        self.write_log("Синхронизация акций WB: календарь, цены и скидки, скидки в заказах.")
        try:
            snapshot = await asyncio.to_thread(wb_prices.sync, self.root, token)
            nm_ids = [int(nm_id) for nm_id in snapshot["prices"]]
            data = await asyncio.to_thread(wb_promotions.sync, self.root, token, None, nm_ids)
            if data.get("order_discounts_error"):
                self.write_log(f"Скидки в заказах не обновлены (взяты прежние): {data['order_discounts_error']}")
            if data.get("daily_error"):
                self.write_log(f"Статистика карточек по дням не обновлена: {data['daily_error']}")
            self.write_log(f"Акции WB обновлены: {len(data['promotions'])} акций.")
        except Exception as exc:
            message = str(exc) if isinstance(exc, (wb_prices.WbPricesError, wb_promotions.WbPromotionsError)) \
                else f"Не удалось загрузить акции WB: {exc}"
            self.write_log(message)
            self.toast(message, True)
        finally:
            self.busy = False
            self.set_sync(False)
            self.refresh()

    def set_sync(self, active):
        dark = self.ui.dash_pal()["dark"]
        self.overlay.visible = active
        self.overlay.bgcolor = ("#CC1E2328" if dark else "#CCF5F7FB") if active else None
        self.overlay.content.controls[1].color = "#F3F4F6" if dark else self.ui.title_color
        if self.page.controls:
            try:
                self.page.update(self.overlay)
            except Exception:
                pass

    # ── отрисовка ─────────────────────────────────────────────────────────

    def render(self, push_update=True):
        ui, pal = self.ui, self.ui.dash_pal()
        if self.error or not self.promotions or self.rows is None:
            self.status.value = "Акции WB ещё не загружены" if not self.error else "Ошибка"
            self.stats_area.content = None
            text = self.error or "Нажмите «Синхронизировать» — программа загрузит календарь акций, цены и скидки из WB."
            self.promo_table.controls = [ui.dtext(text, 13, pal["bad"] if self.error else pal["muted"])]
            self.article_table.controls = []
            self.series_table.controls = []
        else:
            fetched = datetime.fromisoformat(self.promotions["fetched_at"])
            self.status.value = (f"{len(self.promotions['promotions'])} акций · обновлено {fetched:%d.%m в %H:%M} · "
                                 "цена в акции — по наибольшей скидке за 30 дней (история скидок и заказы)")
            self.stats_area.content = self.stats(pal)
            self.promo_table.controls = self.promo_rows(pal)
            self.article_table.controls = self.article_rows(pal)
            self.series_table.controls = self.series_lines(pal)
        if push_update and self.page.controls:
            try:
                self.page.update(self.view)
            except Exception:
                pass

    def stats(self, pal):
        ui = self.ui
        today = date.today()
        running = [p for p in self.promotions["promotions"]
                   if (_date(p.get("startDateTime")) or datetime.max).date() <= today]
        close = [p for p in running if (wb_promotions.tiers(p)["need"] or 99) <= 3]
        left = sum(1 for r in self.rows if r["status"] == "вышел из акции")
        exits = sum(1 for r in self.rows if r["advice"]["tone"] == "act")
        stats = [
            ui.dash_stat("Акций идёт", str(len(running)), f"и ещё {len(self.promotions['promotions']) - len(running)} скоро"),
            ui.dash_stat("Буст рядом", str(len(close)), "акций, где до следующего уровня ≤ 3 товаров",
                         pal["warn"] if close else None),
            ui.dash_stat("Вышли из акций", str(left), "скидка ниже, чем была за 30 дней"),
            ui.dash_stat("Нужно поднять цену", str(exits), "в акции убыток или буст не окупится",
                         pal["bad"] if exits else pal["good"]),
        ]
        return ui.dpanel(ft.ResponsiveRow(stats, spacing=24, run_spacing=12),
                         padding=ft.Padding.symmetric(horizontal=24, vertical=16))

    def header_row(self, columns, pal):
        return ft.Container(padding=ft.Padding.only(left=8, right=8, bottom=8), content=ft.Row(
            [ft.Text(name, expand=weight, size=11, weight=ft.FontWeight.W_600, color=pal["muted"], max_lines=2,
                     overflow=ft.TextOverflow.ELLIPSIS) for name, weight in columns],
            spacing=12, vertical_alignment=ft.CrossAxisAlignment.END))

    def line(self, columns, values, pal, tooltip=None):
        cells = [self.ui.cell(name, {"value": values[name][0], "fill_color": values[name][1]}, weight)
                 for name, weight in columns]
        return ft.Container(padding=ft.Padding.symmetric(horizontal=8, vertical=6),
                            border=ft.Border.only(top=ft.BorderSide(1, pal["separator"])), data="preserve_excel_fill",
                            tooltip=tooltip, content=ft.Row(cells, spacing=12,
                                                            vertical_alignment=ft.CrossAxisAlignment.CENTER))

    def promo_rows(self, pal):
        today = date.today()
        lines = [self.header_row(PROMO_COLUMNS, pal)]
        for promo in self.promotions["promotions"]:
            start, end = _date(promo.get("startDateTime")), _date(promo.get("endDateTime"))
            t = wb_promotions.tiers(promo)
            started = start and start.date() <= today
            when = (f"ещё {(end.date() - today).days} дн." if started and end else
                    f"через {(start.date() - today).days} дн." if start else "—")
            if t["need"] is None:
                next_level = "максимальный уровень"
            else:
                next_level = f"+{t['need']} тов. → {t['next_boost']:g}%"
            boost_fill = None if not t["inside"] else GREEN if t["need"] is None else YELLOW if t["need"] <= 3 else None
            values = {
                "Акция": (promo.get("name") or "", None),
                "Тип": ("авто" if promo.get("type") == "auto" else "обычная", None),
                "Даты": (f"{start:%d.%m} — {end:%d.%m}" if start and end else "—", None),
                "Когда": (when, GREEN if started else None),
                "Ваших в акции": (f"{t['inside']} из {t['total']} ({t['share']:.0f}%)", None),
                "Буст": (f"{t['boost']:g}%" if t["inside"] else "—", None),
                "До следующего уровня": (next_level, boost_fill),
            }
            tip = (promo.get("description") or "").strip()[:600] or None
            if promo.get("exceptionProductsCount"):
                tip = f"Исключено вами до старта: {promo['exceptionProductsCount']} тов.\n" + (tip or "")
            lines.append(self.line(PROMO_COLUMNS, values, pal, tip))
        return lines

    def article_rows(self, pal):
        ui = self.ui
        money = lambda value: ui.money(value) if value is not None else "—"  # noqa: E731
        tone_order = {"act": 0, "consider": 1, None: 2, "ok": 3}
        rows = sorted(self.rows, key=lambda r: (r["verdict"].startswith("конфликта нет"),
                                                tone_order[r["advice"]["tone"]], r["article"]))
        calm = [r for r in rows if r["verdict"].startswith("конфликта нет")]
        shown = rows if self.show_all else [r for r in rows if not r["verdict"].startswith("конфликта нет")]
        lines = [self.header_row(ARTICLE_COLUMNS, pal)]
        for r in shown:
            multiplier = ("не окупится" if r["profit_promo"] is not None and r["profit_promo"] <= 0 else
                          f"×{r['multiplier']:.1f}".replace(".", ",") if r["multiplier"] else "—")
            values = {
                "Артикул": (r["article"], None),
                "Скидка сейчас": (f"{r['discount']} %", None),
                "Цена сейчас": (money(r["discounted_price"]), None),
                "Скидка в акции": (f"{r['promo_discount']} %" if r["promo_discount"] is not None else "—", None),
                "Цена в акции": (money(r["promo_price"]), RED if r["min_price"] and r["promo_price"] < r["min_price"] else None),
                "Статус": (r["status"], YELLOW if r["status"] == "вышел из акции" else None),
                "Прибыль сейчас": (money(r["profit_now"]), RED if (r["profit_now"] or 0) < 0 else None),
                "Прибыль в акции": (money(r["profit_promo"]), RED if (r["profit_promo"] or 0) < 0 else None),
                "Прибыль на минимуме": (money(r["profit_min"]), None),
                "Окупается при росте продаж": (multiplier, None),
                "Что делать": (r["advice"]["text"], TONE_FILL.get(r["advice"]["tone"])),
                "Переходов в день: со скидкой акции | с меньшей": (
                    f"{per_day(r['effect']['promo']['views'])} | {per_day(r['effect']['plain']['views'])}", None),
                "По факту": (r["fact_advice"]["text"], TONE_FILL.get(r["fact_advice"]["tone"])),
            }
            lines.append(self.line(ARTICLE_COLUMNS, values, pal, self.tooltip(r)))
        if calm:
            label = (f"Скрыть товары без конфликта ({len(calm)})" if self.show_all
                     else f"Показать товары без конфликта ({len(calm)}): цена акции у них не ниже минимальной")
            lines.append(ft.Container(padding=ft.Padding.only(top=8), content=ft.TextButton(label, on_click=self.toggle)))
        return lines

    def series_lines(self, pal):
        lines = [self.header_row(SERIES_COLUMNS, pal)]
        for s in self.series_rows or []:
            e = s["effect"]
            uplift = e["views_uplift"]
            values = {
                "Серия": (s["series"], None), "Артикулов": (f"{s['articles']} шт.", None),  # голое число ячейка показала бы как рубли
                "Дней: со скидкой акции | с меньшей": (f"{e['promo']['days']} | {e['plain']['days']}", None),
                "Переходов в день: со скидкой акции | с меньшей": (
                    f"{per_day(e['promo']['views'])} | {per_day(e['plain']['views'])}", None),
                "Заказов в день: со скидкой акции | с меньшей": (
                    f"{per_day(e['promo']['orders'], 2)} | {per_day(e['plain']['orders'], 2)}", None),
                "Переходов больше в": (f"×{per_day(uplift)}" if uplift else "мало данных", None),
            }
            tip = ("Дни считаются в сумме по артикулам серии. Скидка до начала истории скидок взята из самой ранней "
                   "записи — это оценка." if e["estimated"] else None)
            lines.append(self.line(SERIES_COLUMNS, values, pal, tip))
        return lines

    def toggle(self, _e=None):
        self.show_all = not self.show_all
        self.render()

    def tooltip(self, r):
        money = self.ui.money
        if r["notes"]:
            return "Не посчитано: " + "; ".join(r["notes"]) + "."
        lines = [f"Скидка в акции {r['promo_discount']}% — {r['promo_source']}"
                 + (f" ({r['promo_date']})" if r.get("promo_date") else "") + f": цена {money(r['promo_price'])}."]
        if r["min_price"]:
            lines.append(f"Минимальная цена {money(r['min_price'])}: прибыль {money(r['profit_min'])} с продажи.")
        if r["profit_promo"] is not None:
            lines.append(f"В акции прибыль {money(r['profit_promo'])} с продажи.")
        if r["multiplier"]:
            lines.append(f"Акция выгодна, если с бустом продаж станет больше в {r['multiplier']:.1f} раза и более."
                         .replace(".", ",", 1))
        lines.append("Вердикт: до ×1,5 — держать, до ×3 — смотреть по продажам, больше — выходить.")
        effect = r["effect"]
        lines.append(f"По факту: {effect['promo']['days']} дн. со скидкой акции "
                     f"({per_day(effect['promo']['views'])} переходов в день), {effect['plain']['days']} дн. с меньшей "
                     f"({per_day(effect['plain']['views'])}). Для вывода нужно хотя бы {wb_promotions.MIN_DAYS} дня "
                     "с каждой стороны; переходы — замена продаж при одинаковой конверсии."
                     + (" Скидка до начала истории — оценка." if effect["estimated"] else ""))
        return "\n".join(lines)
