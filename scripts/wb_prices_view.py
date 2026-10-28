"""Раздел «Цены» магазина WB — в том же виде, что «Цены и скидки» Ozon (app_flet.py), только для чтения.

Шапка: маржа WB (свои настройки, wb_price_settings.json), «Синхронизировать» (снимок цен из WB), «…».
Ниже — те же показатели и таблица «Товары», что у Ozon. Заявок на скидку у WB нет, а минимальная и желательная
цена считаются сами при каждом показе (scripts/wb_prices.py), поэтому кнопки «Рассчитать цены» тоже нет.
"""
import asyncio
from datetime import datetime
import time

import flet as ft
from scripts import wb_prices
from scripts.marketplace_settings import read_settings

AUTO_SYNC_INTERVAL = 300  # как у Ozon: при входе в раздел — не чаще раза в 5 минут
RED, GREEN, DARK_GREEN = "#FFC7CE", "#C6EFCE", "#A9D8B8"
COLUMNS = [  # (название, вес ширины)
    ("Артикул", 10), ("Себестоимость", 11), ("Минимальная цена", 12), ("Желательная цена", 12),
    ("Цена до скидки", 12), ("Скидка", 8), ("Можно ставить скидку", 11), ("Цена со скидкой", 12), ("Рентабельность", 11),
    ("Прибыль с единицы", 12),
]


def parse_margin(text):
    return float(str(text or "").strip().replace(",", ".").rstrip("%")) if str(text or "").strip() else None


class WbPricesView:
    """ui — помощники из app_flet: dpanel, dtext, dash_stat, dash_pal, money, percent, cell(column, payload, expand),
    primary, accent, title_color, muted_color. reveal(path, label) — открыть файл, toast(text, error) — сообщение."""

    def __init__(self, *, root, page, ui, log, reveal, toast):
        self.root, self.page, self.ui, self.log, self.reveal, self.toast = root, page, ui, log, reveal, toast
        self.result = None
        self.error = None
        self.busy = False
        self.last_sync = 0.0
        settings = wb_prices.load_settings(root)
        self.margin_min = ft.TextField(label="Мин. маржа", value=f"{settings['min_margin']:g}", width=120, dense=True,
                                       tooltip="Доля, например 0.25 = 25%. Минимальная маржа — порог «ниже нельзя», "
                                               "желаемая — цель расчёта. Маржа WB хранится отдельно от Ozon.",
                                       on_submit=self.save_margin)
        self.margin_des = ft.TextField(label="Желаемая", value=f"{settings['desired_margin']:g}", width=120, dense=True,
                                       tooltip=self.margin_min.tooltip, on_submit=self.save_margin)
        self.status = ft.Text("—", size=13, color=ui.muted_color)
        self.stats_area = ft.Container()
        self.table = ft.Column(spacing=0)
        self.sync_label = ft.Text("Синхронизация с WB: обновляю цены и скидки…", size=13, weight=ft.FontWeight.W_600,
                                  text_align=ft.TextAlign.CENTER)
        self.overlay = ft.Container(
            left=0, top=0, right=0, bottom=0, visible=False, border_radius=18, alignment=ft.Alignment(0, -1),
            padding=ft.Padding.only(top=120), data="preserve_excel_fill",
            content=ft.Column([ft.ProgressRing(width=36, height=36, stroke_width=3), self.sync_label],
                              spacing=14, tight=True, horizontal_alignment=ft.CrossAxisAlignment.CENTER))
        more = ft.PopupMenuButton(icon=ft.Icons.MORE_HORIZ, tooltip="Ещё", items=[
            ft.PopupMenuItem(content=ft.Text("Открыть wb_costs.xlsx"), on_click=lambda _e: self.open_costs()),
            ft.PopupMenuItem(content=ft.Text("Перечитать таблицу"), on_click=lambda _e: self.refresh()),
        ])
        header = ft.ResponsiveRow([
            ft.Container(col={"xs": 12, "lg": 4}, content=ft.Column(
                [ft.Text("Цены", size=26, weight=ft.FontWeight.W_800, color=ui.title_color), self.status], spacing=2)),
            ft.Container(col={"xs": 12, "lg": 8}, alignment=ft.Alignment(1, 0), content=ft.Row([
                self.margin_min, self.margin_des,
                ft.IconButton(icon=ft.Icons.SAVE_OUTLINED, tooltip="Сохранить маржу", on_click=self.save_margin),
                ft.Button("Синхронизировать", icon=ft.Icons.SYNC, bgcolor=ui.primary, color="white",
                          tooltip="Обновит текущие цены и скидки из WB. Запускается сама при входе в раздел",
                          on_click=lambda _e: self.page.run_task(self.sync)),
                more,
            ], spacing=10, wrap=True, run_spacing=8, alignment=ft.MainAxisAlignment.END,
                vertical_alignment=ft.CrossAxisAlignment.CENTER)),
        ], vertical_alignment=ft.CrossAxisAlignment.CENTER, run_spacing=12)
        self.table_panel = ui.dpanel(ft.Column([
            ft.Row([ui.dtext("Товары", 15, weight=ft.FontWeight.W_700),
                    ui.dtext("Наведите на строку, чтобы увидеть комиссию, логистику и цену для покупателя. "
                             "Себестоимость — в wb_costs.xlsx (меню «…»).", 12, ui.muted_color, expand=True,
                             max_lines=1, overflow=ft.TextOverflow.ELLIPSIS)],
                   spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Stack([self.table, self.overlay]),
        ], spacing=8), padding=ft.Padding.only(left=20, right=20, top=12, bottom=8))
        self.view = ft.Column([header, self.stats_area, self.table_panel], spacing=14)

    def write_log(self, message):
        self.log(f"[{datetime.now():%H:%M:%S}] [WB] {message}")

    # ── данные ───────────────────────────────────────────────────────────

    def refresh(self, push_update=True):
        """Посчитать таблицу по сохранённому снимку цен, отчётам WB и wb_costs.xlsx — без запросов к WB."""
        self.error = None
        try:
            snapshot = wb_prices.load_snapshot(self.root)
            self.result = wb_prices.compute(self.root, snapshot) if snapshot else None
        except Exception as exc:
            self.result = None
            self.error = f"Не удалось посчитать цены WB: {exc}. Закройте wb_costs.xlsx в Excel и повторите."
        self.render(push_update)

    def enter(self):
        """Вход в раздел: показать сохранённое и синхронизировать с WB, если снимок старше 5 минут."""
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
        self.write_log("Синхронизация цен WB: цены, скидки, предметы товаров.")
        try:
            await asyncio.to_thread(wb_prices.sync, self.root, token)
            self.write_log("Цены WB обновлены.")
        except Exception as exc:
            message = str(exc) if isinstance(exc, wb_prices.WbPricesError) else f"Не удалось загрузить цены WB: {exc}"
            self.write_log(message)
            self.toast(message, True)
        finally:
            self.busy = False
            self.set_sync(False)
            self.refresh()

    def set_sync(self, active):
        self.overlay.visible = active
        dark = self.ui.dash_pal()["dark"]
        self.overlay.bgcolor = ("#CC1E2328" if dark else "#CCF5F7FB") if active else None
        self.overlay.content.controls[1].color = "#F3F4F6" if dark else self.ui.title_color
        self.table.disabled = active
        if self.page.controls:
            try:
                self.page.update(self.overlay, self.table)
            except Exception:
                pass  # раздел сейчас не на экране

    def save_margin(self, _e=None):
        try:
            low, high = parse_margin(self.margin_min.value), parse_margin(self.margin_des.value)
            settings = wb_prices.save_settings(self.root, low, high)
        except (ValueError, wb_prices.WbPricesError) as exc:
            self.toast(str(exc) if isinstance(exc, wb_prices.WbPricesError) else "Маржа WB — число, например 0,25.", True)
            return
        self.toast(f"Маржа WB сохранена: минимальная {settings['min_margin']:.0%}, желаемая {settings['desired_margin']:.0%}.")
        self.refresh()

    def open_costs(self):
        from scripts import wb_finance
        try:
            self.reveal(wb_finance.ensure_costs(self.root), "Себестоимость WB")
        except Exception:
            self.toast("Не удалось открыть wb_costs.xlsx. Закройте файл в Excel и повторите.", True)

    # ── отрисовка ─────────────────────────────────────────────────────────

    def render(self, push_update=True):
        ui = self.ui
        pal = ui.dash_pal()
        if self.error or not self.result:
            self.status.value = "Цены WB ещё не загружены" if not self.error else "Ошибка"
            self.stats_area.content = None
            text = self.error or "Нажмите «Синхронизировать» — программа загрузит цены и скидки из WB."
            self.table.controls = [ui.dtext(text, 13, pal["bad"] if self.error else pal["muted"])]
        else:
            result = self.result
            rows = result["rows"]
            fetched = datetime.fromisoformat(result["fetched_at"])
            self.status.value = (f"{len(rows)} товаров · цены WB от {fetched:%d.%m в %H:%M} · "
                                 f"комиссия и логистика по отчётам за {', '.join(result['months']) or '—'}")
            self.stats_area.content = self.stats(rows, result["settings"])
            self.table.controls = [self.header_row(pal)] + [self.row(row, result, pal) for row in rows]
        if push_update and self.page.controls:
            try:
                self.page.update(self.view)
            except Exception:
                pass

    def stats(self, rows, settings):
        ui, pal = self.ui, self.ui.dash_pal()
        counted = [row for row in rows if row["margin"] is not None]
        below = sum(1 for row in counted if row["discounted_price"] < row["min_price"])
        no_cost = sum(1 for row in rows if row["cost"] is None)
        low, high = settings["min_margin"] * 100, settings["desired_margin"] * 100
        avg = sum(row["margin"] for row in counted) / len(counted) * 100 if counted else None
        profit = sum(row["profit"] for row in counted) / len(counted) if counted else None
        color = None if avg is None else pal["bad"] if avg < low else pal["good"] if avg <= high else pal["accent"]
        stats = [
            ui.dash_stat("Товаров", str(len(rows)), f"без себестоимости: {no_cost}" if no_cost else "все с себестоимостью"),
            ui.dash_stat("Ниже минимальной цены", str(below), "цена со скидкой < минимальной",
                         pal["bad"] if below else pal["good"]),
            ui.dash_stat("Средняя рентабельность", ui.percent(avg) if avg is not None else "—",
                         f"цель {ui.percent(low)} – {ui.percent(high)}", color),
            ui.dash_stat("Прибыль с единицы", ui.money(profit) if profit is not None else "—", "в среднем по товарам",
                         pal["bad"] if profit is not None and profit < 0 else None),
        ]
        return ui.dpanel(ft.ResponsiveRow(stats, spacing=24, run_spacing=12),
                         padding=ft.Padding.symmetric(horizontal=24, vertical=16))

    def header_row(self, pal):
        return ft.Container(padding=ft.Padding.only(left=8, right=8, bottom=8), content=ft.Row(
            [ft.Text(name, expand=weight, size=11, weight=ft.FontWeight.W_600, color=pal["muted"], max_lines=2,
                     overflow=ft.TextOverflow.ELLIPSIS) for name, weight in COLUMNS],
            spacing=12, vertical_alignment=ft.CrossAxisAlignment.END))

    def row(self, row, result, pal):
        ui = self.ui
        money = lambda value: ui.money(value) if value is not None else "—"  # noqa: E731
        settings = result["settings"]
        price, low, high = row["discounted_price"], row["min_price"], row["desired_price"]
        price_fill = None if low is None else RED if price < low else DARK_GREEN if high and price >= high else GREEN
        margin_fill = None
        if row["margin"] is not None:
            margin_fill = RED if row["margin"] < settings["min_margin"] else GREEN if row["margin"] <= settings["desired_margin"] else None
        allowed, discount_fill = self.discount_range(row)
        values = {
            "Артикул": (row["article"], None), "Себестоимость": (money(row["cost"]), None),
            "Минимальная цена": (money(low), None), "Желательная цена": (money(high), None),
            "Цена до скидки": (money(row["price"]), None), "Скидка": (f"{row['discount']} %", discount_fill),
            "Можно ставить скидку": (allowed, None),
            "Цена со скидкой": (money(price), price_fill),
            "Рентабельность": (ui.percent(row["margin"] * 100) if row["margin"] is not None else "—", margin_fill),
            "Прибыль с единицы": (money(row["profit"]), RED if row["profit"] is not None and row["profit"] < 0 else None),
        }
        cells = [ui.cell(name, {"value": text, "fill_color": fill}, weight) for name, weight in COLUMNS
                 for text, fill in [values[name]]]
        return ft.Container(padding=ft.Padding.symmetric(horizontal=8, vertical=6),
                            border=ft.Border.only(top=ft.BorderSide(1, pal["separator"])), data="preserve_excel_fill",
                            tooltip=self.tooltip(row), content=ft.Row(cells, spacing=12,
                                                                      vertical_alignment=ft.CrossAxisAlignment.CENTER))

    @staticmethod
    def discount_range(row):
        """«30–40 %» — скидки продавца при текущей цене до скидки: от желательной цены до минимальной.
        Подсветка текущей скидки: в диапазоне — зелёная, больше допустимой — красная."""
        high, low = row["discount_max"], row["discount_desired"]
        if row["min_price"] is None:
            return "—", None
        if high is None:
            return "поднять цену до скидки", RED
        text = f"до {high} %" if low is None or low == high else f"{low}–{high} %"
        if low is None:
            text += " (желательная цена недостижима)"
        # Скидка не больше «желательной» — цена не ниже желательной (тёмно-зелёная, как у цены со скидкой).
        fill = RED if row["discount"] > high else GREEN if low is None or row["discount"] > low else DARK_GREEN
        return text, fill

    def tooltip(self, row):
        if row["notes"]:
            return f"{row['subject']}. Не посчитано: " + "; ".join(row["notes"]) + "."
        lines = [row["subject"]]
        if row["discount_max"] is not None:
            lines.append(f"Скидка {row['discount_desired'] if row['discount_desired'] is not None else '—'}% — цена "
                         f"не ниже желательной, {row['discount_max']}% — не ниже минимальной (при цене до скидки "
                         f"{self.ui.money(row['price'])}). Вписывайте целое число в «Скидку продавца» в кабинете WB.")
        lines += [f"Комиссия WB {row['commission'] * 100:.1f}% ({row['commission_source']})".replace(".", ","),
                  f"Эквайринг {row['acquiring'] * 100:.1f}%".replace(".", ","),
                  f"Логистика {self.ui.money(row['logistics'])} за проданную штуку ({row['logistics_source']})"]
        if row["buyer_price"] is not None:
            lines.append(f"Покупатель платит ≈ {self.ui.money(row['buyer_price'])} — со скидкой WB, её оплачивает WB")
        return "\n".join(lines)
