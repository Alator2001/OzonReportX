from decimal import Decimal
import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import Mock, patch

from scripts import wb_finance as wb
import test_technical_fixes as technical


def report(**updates):
    row = {"reportId": 1, "currency": "RUB", "dateFrom": "2026-09-01", "dateTo": "2026-09-01",
           "retailAmountSum": "1000", "bankPaymentSum": "600", "deliveryServiceSum": "100",
           "paidStorageSum": "30", "paidAcceptanceSum": "10", "deductionSum": "20",
           "penaltySum": "0", "additionalPaymentSum": "0"}
    return dict(row, **updates)


def detail(row_id=1, operation="Продажа", **updates):
    row = {"reportId": 1, "rrdId": row_id, "currency": "RUB", "sellerOperName": operation,
           "srid": "order", "nmId": 123, "quantity": 1, "retailAmount": "1000"}
    return dict(row, **updates)


def data(rows=None):
    return {"reports": [report()], "details": rows if rows is not None else [detail()],
            "version": 1, "year": 2026, "month": 9, "fetched_at": "2026-09-13", "account": wb.identity("test")}


class WBFinanceTests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def test_financial_expenses_are_not_subtracted_twice(self):
        result = wb.summarize(data(), {"123": Decimal(700)})
        self.assertEqual(result["profit"], -100)
        self.assertEqual(result["margin"], -10)
        self.assertEqual(result["sales"], 1)
        self.assertEqual(result["returns"], 0)

    def test_sale_and_return_have_zero_costs(self):
        payload = data([detail(), detail(2, "Возврат")])
        payload["reports"] = [report(retailAmountSum="0", bankPaymentSum="-100")]
        result = wb.summarize(payload, {})
        self.assertEqual(result["costs"], 0)
        self.assertEqual(result["profit"], -100)
        self.assertEqual(result["returns"], 1)
        self.assertIsNone(result["margin"])

    def test_missing_cost_or_prior_period_return_blocks_profit(self):
        self.assertIsNone(wb.summarize(data(), {})["profit"])
        result = wb.summarize(data([detail(operation="Возврат")]), {"123": Decimal(100)})
        self.assertIsNone(result["profit"])
        self.assertTrue(result["notes"])

    def test_logistics_rows_are_not_sales_even_with_sale_document_type(self):
        result = wb.summarize(data([detail(), detail(2, "Логистика", docTypeName="Продажа", quantity=0)]), {"123": Decimal(100)})
        self.assertEqual(result["costs"], 100)
        self.assertEqual(result["sales"], 1)

    def test_complete_pagination_and_mismatched_reports(self):
        client = wb.FinanceClient("test")
        with patch.object(client, "request", side_effect=[[report()], [detail()], [detail(2, srid="order2")], None]) as fetch:
            result = client.fetch(2026, 9)
            self.assertEqual(len(result["details"]), 2)
            self.assertEqual(fetch.call_args_list[2].args[1]["rrdId"], 1)
            self.assertEqual(fetch.call_args_list[0].args[1]["period"], "daily")
        for bad in (detail(reportId=2), detail(currency="USD"), detail(rrdId=0)):
            with patch.object(client, "request", side_effect=[[report()], [bad]]):
                with self.assertRaises(wb.FinanceError):
                    client.fetch(2026, 9)

    def test_request_errors_do_not_expose_token_or_response_body(self):
        session = Mock()
        session.post.return_value = Mock(status_code=403, text="secret-token")
        client = wb.FinanceClient("secret-token", session=session)
        with patch.object(wb, "_next_request", 0):
            with self.assertRaises(wb.FinanceError) as error:
                client.request("/list", {})
        self.assertNotIn("secret-token", str(error.exception))
        self.assertEqual(session.post.call_args.kwargs["headers"]["Authorization"], "secret-token")

    def test_cancel_and_rate_limit_do_not_retry_early(self):
        cancel = threading.Event()
        session = Mock()
        response = Mock(status_code=429, headers={"X-Ratelimit-Retry": "120"})
        session.post.return_value = response
        client = wb.FinanceClient("test", cancel=cancel, session=session,
                                  progress=lambda message: cancel.set())
        with patch.object(wb, "_next_request", 0):
            with self.assertRaises(wb.FinanceError):
                client.request("/list", {})
        session.post.assert_called_once()

    def test_order_rows_reconcile_with_summary_payout_and_split_multi_item_orders(self):
        rows_payload = [
            detail(1, "Продажа", nmId=123, retailAmount="1000", forPay="900",
                   deliveryService="50", paidStorage="10"),
            detail(2, "Продажа", nmId=456, srid="order", retailAmount="500", forPay="400"),
            detail(3, "Доставка", nmId=0, srid="other-order", quantity=0, retailAmount="0",
                   forPay="0", deliveryService="20"),
        ]
        payload = data(rows_payload)
        payload["reports"] = [report(bankPaymentSum="1220", deliveryServiceSum="70", paidStorageSum="10")]
        summary = wb.summarize(payload, {"123": Decimal(200), "456": Decimal(50)})
        rows, residual = wb.order_rows(payload, {"123": Decimal(200), "456": Decimal(50)})
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["Номер заказа"], "order")
        self.assertEqual(row["Количество шт."], 2)
        self.assertEqual(row["Цена продажи"], 1500)
        self.assertEqual(row["Себестоимость"], -250)
        self.assertAlmostEqual(row["Прибыль"], row["Сумма начисления"] - 250)
        total = sum(r["Сумма начисления"] for r in rows) + float(residual)
        self.assertAlmostEqual(total, float(summary["payout"]))

    def test_order_rows_mark_missing_cost_without_blocking_other_orders(self):
        payload = data([detail(1, "Продажа", nmId=123, forPay="900")])
        rows, _residual = wb.order_rows(payload, {})
        self.assertIsNone(rows[0]["Себестоимость"])
        self.assertIsNone(rows[0]["Прибыль"])

    def test_monthly_report_writes_orders_and_matches_dashboard_totals(self):
        from scripts import wb_monthly_report
        payload = data([detail(1, "Продажа", nmId=123, forPay="900", retailAmount="1000")])
        payload["reports"] = [report(bankPaymentSum="900")]
        costs = {"123": Decimal(200)}
        summary = wb.summarize(payload, costs)
        output_file = wb_monthly_report.build(self.directory, payload, costs)
        self.assertEqual(output_file.name, "WB Сентябрь 2026.xlsx")
        workbook = wb.load_workbook(output_file)
        try:
            sheet = workbook["Заказы"]
            self.assertEqual(sheet.cell(row=2, column=2).value, "order")
            self.assertEqual(sheet.cell(row=2, column=9).value, 700)
            indicators = {sheet.cell(row=r, column=16).value: sheet.cell(row=r, column=17).value
                          for r in range(1, sheet.max_row + 1) if sheet.cell(row=r, column=16).value}
            self.assertEqual(indicators["Прибыль по отчёту"], float(summary["profit"]))
            self.assertEqual(indicators["Итог WB к выплате"], float(summary["payout"]))
        finally:
            workbook.close()

    def test_monthly_report_enriches_matching_orders_with_return_reason(self):
        from scripts import wb_monthly_report
        payload = data([detail(1, "Продажа", nmId=123, srid="order", forPay="900", retailAmount="1000")])
        payload["reports"] = [report(bankPaymentSum="900")]
        costs = {"123": Decimal(200)}
        returns_by_srid = {"order": {"reason": "Цвет", "status": "В пути в пвз"},
                           "unrelated-order": {"reason": "Брак", "status": "Завершён"}}
        output_file = wb_monthly_report.build(self.directory, payload, costs, returns_by_srid)
        workbook = wb.load_workbook(output_file)
        try:
            sheet = workbook["Заказы"]
            self.assertEqual(sheet.cell(row=2, column=2).value, "order")
            self.assertEqual(sheet.cell(row=2, column=11).value, "Цвет")
            self.assertEqual(sheet.cell(row=2, column=12).value, "В пути в пвз")
        finally:
            workbook.close()

    def test_monthly_report_leaves_return_columns_blank_without_a_match(self):
        from scripts import wb_monthly_report
        payload = data([detail(1, "Продажа", nmId=123, srid="order", forPay="900", retailAmount="1000")])
        payload["reports"] = [report(bankPaymentSum="900")]
        output_file = wb_monthly_report.build(self.directory, payload, {"123": Decimal(200)})
        workbook = wb.load_workbook(output_file)
        try:
            sheet = workbook["Заказы"]
            self.assertIsNone(sheet.cell(row=2, column=11).value)
            self.assertIsNone(sheet.cell(row=2, column=12).value)
        finally:
            workbook.close()

    def test_monthly_report_adds_acceptance_and_storage_sheets(self):
        from scripts import wb_monthly_report
        payload = data([detail(1, "Продажа", nmId=123, forPay="900", retailAmount="1000")])
        payload["reports"] = [report(bankPaymentSum="900")]
        acceptance = {"total": Decimal("75"), "rows": 2,
                     "by_nm": {"123": {"count": 3, "total": Decimal("75"), "subject": "Доски"}}}
        storage = {"total": Decimal("12.5"), "rows": 5,
                  "by_nm": {"123": {"total": Decimal("12.5"), "subject": "Доски", "vendor_code": "1114"}}}
        output_file = wb_monthly_report.build(self.directory, payload, {"123": Decimal(200)},
                                              acceptance=acceptance, storage=storage)
        workbook = wb.load_workbook(output_file)
        try:
            self.assertIn("Приёмка", workbook.sheetnames)
            self.assertIn("Хранение", workbook.sheetnames)
            acceptance_sheet = workbook["Приёмка"]
            self.assertEqual(acceptance_sheet.cell(row=2, column=1).value, "123")
            self.assertEqual(acceptance_sheet.cell(row=2, column=3).value, 75.0)
            storage_sheet = workbook["Хранение"]
            self.assertEqual(storage_sheet.cell(row=2, column=3).value, 12.5)
            indicators = {workbook["Заказы"].cell(row=r, column=16).value: workbook["Заказы"].cell(row=r, column=17).value
                          for r in range(1, workbook["Заказы"].max_row + 1) if workbook["Заказы"].cell(row=r, column=16).value}
            self.assertEqual(indicators["Приёмка по товарам (лист «Приёмка»), сумма"], 75.0)
            self.assertEqual(indicators["Хранение по товарам (лист «Хранение»), сумма"], 12.5)
        finally:
            workbook.close()

    def test_monthly_report_adds_sales_funnel_sheet(self):
        from scripts import wb_monthly_report
        payload = data([detail(1, "Продажа", nmId=123, forPay="900", retailAmount="1000")])
        payload["reports"] = [report(bankPaymentSum="900")]
        funnel = {"products": 1, "views": 100, "cart": 10, "orders": 3, "buyouts": 2,
                 "add_to_cart_pct": 10.0, "cart_to_order_pct": 30.0, "buyout_pct": 66.7,
                 "by_nm": {"123": {"nmId": "123", "title": "Доска", "vendor_code": "1114", "subject": "Доски",
                                  "views": 100, "cart": 10, "orders": 3, "order_sum": Decimal(2700),
                                  "buyouts": 2, "buyout_sum": Decimal(1800), "cancels": 1,
                                  "add_to_cart_pct": 10, "cart_to_order_pct": 30, "buyout_pct": 67}}}
        output_file = wb_monthly_report.build(self.directory, payload, {"123": Decimal(200)}, funnel=funnel)
        workbook = wb.load_workbook(output_file)
        try:
            self.assertIn("Воронка продаж", workbook.sheetnames)
            sheet = workbook["Воронка продаж"]
            self.assertEqual(sheet.cell(row=2, column=1).value, "123")
            self.assertEqual(sheet.cell(row=2, column=4).value, 100)
            indicators = {workbook["Заказы"].cell(row=r, column=16).value: workbook["Заказы"].cell(row=r, column=17).value
                          for r in range(1, workbook["Заказы"].max_row + 1) if workbook["Заказы"].cell(row=r, column=16).value}
            self.assertEqual(indicators["Воронка продаж (лист «Воронка продаж») — просмотры"], 100)
        finally:
            workbook.close()

    def test_separate_costs_preserve_values_and_reject_invalid_numbers(self):
        path = wb.ensure_costs(self.directory, [detail(vendorCode="=SECRET()")])
        self.assertEqual(path.name, "wb_costs.xlsx")
        workbook = wb.load_workbook(path)
        self.assertEqual(workbook.active['B2'].data_type, 's')
        workbook.active['D2'] = 123
        workbook.save(path)
        workbook.close()
        wb.ensure_costs(self.directory, [detail(), detail(2, nmId=456)])
        self.assertEqual(wb.read_costs(self.directory), {"123": Decimal(123), "456": None})
        self.assertFalse((self.directory / 'costs.xlsx').exists())
        with self.assertRaises(wb.FinanceError):
            wb.number('NaN')

    def test_cache_is_scoped_to_token_and_atomic(self):
        wb.save_report(self.directory, data())
        with patch.object(wb, "read_settings", return_value={"WB_API_TOKEN": "another"}):
            self.assertIsNone(wb.load_report(self.directory, 2026, 9))
        with patch.object(wb, "read_settings", return_value={"WB_API_TOKEN": "test"}):
            self.assertEqual(wb.load_report(self.directory, 2026, 9)["details"], [detail()])
        path = wb.report_path(self.directory, 2026, 9)
        before = path.read_bytes()
        with patch.object(wb.json, "dumps", side_effect=ValueError):
            with self.assertRaises(ValueError):
                wb.save_report(self.directory, data())
        self.assertEqual(path.read_bytes(), before)

    def test_desktop_summary_uses_actual_card_builders(self):
        import asyncio
        import flet as ft
        from scripts import wb_dashboard
        from scripts.wb_dashboard import build_dashboard
        namespace = dict(ft=ft, SURFACE="white", TEXT="black", MUTED="grey", BORDER="grey")
        helpers = {name: technical.ui_function(name, namespace) for name in ("card", "metric", "hero_metric")}
        logger, page = Mock(), Mock()
        view = build_dashboard(months=[str(n) for n in range(1, 13)], **helpers,
                               text_color="black", muted_color="grey", primary="green",
                               accent="orange", surface="white", page=page, root=self.directory,
                               reveal=Mock(), theme=Mock(), log=logger)
        self.assertEqual(view.controls[0].value, "Бизнес-сводка · Wildberries")
        self.assertEqual(len(view.controls[5].controls), 7)
        view.controls[2].controls[0].on_click(None)
        self.assertTrue(any("Пересчёт" in c.args[0] for c in logger.call_args_list))

        def client_factory(token, cancel, progress):
            client = Mock()
            def fetch(*args):
                progress("Последняя страница получена")
                return data()
            client.fetch.side_effect = fetch
            return client

        view.controls[1].controls[2].on_click(None)
        download = page.run_task.call_args.args[0]
        with patch.object(wb_dashboard, "read_settings", return_value={"WB_API_TOKEN": "hidden-token"}), \
                patch.object(wb, "FinanceClient", side_effect=client_factory), \
                patch.object(wb, "ensure_costs"), patch.object(wb, "save_report"):
            asyncio.run(download())
        messages = "\n".join(c.args[0] for c in logger.call_args_list)
        self.assertIn("Последняя страница получена", messages)
        self.assertIn("Загрузка завершена", messages)
        self.assertIn("Внимание:", messages)
        self.assertNotIn("hidden-token", messages)

        with patch.object(wb_dashboard, "read_settings", return_value={"WB_API_TOKEN": "hidden-token"}), \
                patch.object(wb, "FinanceClient", side_effect=OSError("hidden-token")):
            asyncio.run(download())
        self.assertIn("Ошибка загрузки", logger.call_args.args[0])
        self.assertNotIn("hidden-token", logger.call_args.args[0])


class OzonStatusTests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def test_only_return_status_removes_costs(self):
        monthly = technical.monthly
        for status, amount, expected_profit, expected_cost in (
            ("delivered", 50, -50, -100), ("delivered", 150, 50, -100),
            ("returned", -30, -30, 0), ("returned", 20, 20, 0),
            ("awaiting_deliver", 0, "-", 0),
        ):
            with self.subTest(status=status, amount=amount):
                postings = [{"posting_number": "order", "status": status,
                             "products": [{"offer_id": "sku", "quantity": 1}]}]
                finances = {"order": [{"amount": amount, "sale_commission": -10, "accruals_for_sale": 100,
                                       "has_commission_data": True}]}
                with patch.object(monthly, "get_transactions_by_posting", return_value=finances), \
                        patch.object(monthly, "load_cost_map", return_value={"sku": 100}), \
                        patch.object(monthly, "_safe_save_excel", return_value="unused") as save:
                    monthly.to_excel(postings, "2026-09-01", "2026-09-02", 9, 2026,
                                     output_file="unused", session=Mock())
                row = save.call_args.args[0].iloc[0]
                self.assertEqual(row["Статус"], status)
                self.assertEqual(row["Себестоимость"], expected_cost)
                self.assertEqual(row["Прибыль"], expected_profit)

    def test_zero_accrual_price_falls_back_to_posting_price(self):
        """/v1/finance/accrual/by-day can split a posting into fee-only accruals (commission: null)
        whose sale-with-commission entry landed in a different month; the posting's own per-unit
        price still reflects the sale and must not be reported as a 0 ₽ "Цена продажи"."""
        monthly = technical.monthly
        postings = [{"posting_number": "order", "status": "delivered",
                     "products": [{"offer_id": "sku", "quantity": 2, "price": "150.0000"}]}]
        finances = {"order": [{"amount": -75, "sale_commission": 0, "accruals_for_sale": 0}]}
        with patch.object(monthly, "get_transactions_by_posting", return_value=finances), \
                patch.object(monthly, "load_cost_map", return_value={}), \
                patch.object(monthly, "_safe_save_excel", return_value="unused") as save:
            monthly.to_excel(postings, "2026-09-01", "2026-09-02", 9, 2026,
                             output_file="unused", session=Mock())
        row = save.call_args.args[0].iloc[0]
        self.assertEqual(row["Цена продажи"], 300.0)

    def test_nonzero_accrual_price_is_not_overridden_by_posting_price(self):
        monthly = technical.monthly
        postings = [{"posting_number": "order", "status": "delivered",
                     "products": [{"offer_id": "sku", "quantity": 1, "price": "999.0000"}]}]
        finances = {"order": [{"amount": 90, "sale_commission": -10, "accruals_for_sale": 100,
                              "has_commission_data": True}]}
        with patch.object(monthly, "get_transactions_by_posting", return_value=finances), \
                patch.object(monthly, "load_cost_map", return_value={}), \
                patch.object(monthly, "_safe_save_excel", return_value="unused") as save:
            monthly.to_excel(postings, "2026-09-01", "2026-09-02", 9, 2026,
                             output_file="unused", session=Mock())
        row = save.call_args.args[0].iloc[0]
        self.assertEqual(row["Цена продажи"], 100.0)

    def test_missing_commission_data_is_reported_as_its_own_pending_status(self):
        """Real example: posting 0250581479-0070-1, shipped 2026-08-20, accrual posted 2026-09-03 —
        no sale-with-commission accrual fell inside August. Reporting profit as amount+cost (a real
        loss on paper) would be wrong: the sale happened, Ozon just hasn't booked the commission yet.
        The row must be pulled out of "delivered" into its own status so cost/profit aren't counted
        as real until the accrual actually arrives (in whatever month that turns out to be)."""
        monthly = technical.monthly
        postings = [{"posting_number": "order", "status": "delivered",
                     "products": [{"offer_id": "sku", "quantity": 1, "price": "490.0000"}]}]
        finances = {"order": [{"amount": -121.35, "sale_commission": 0, "accruals_for_sale": 0,
                               "has_commission_data": False}]}
        with patch.object(monthly, "get_transactions_by_posting", return_value=finances), \
                patch.object(monthly, "load_cost_map", return_value={"sku": 175}), \
                patch.object(monthly, "_safe_save_excel", return_value="unused") as save:
            monthly.to_excel(postings, "2026-08-01", "2026-08-31", 8, 2026,
                             output_file="unused", session=Mock())
        row = save.call_args.args[0].iloc[0]
        label = "Ожидает расчёта Ozon (начисление ещё не пришло)"
        self.assertEqual(row["Статус"], "ожидает расчёта")
        self.assertEqual(row["Цена продажи"], 490.0)
        self.assertEqual(row["Комиссия за продажу Ozon"], label)
        self.assertEqual(row["Логистика (Включает операционные ошибки продавца)"], label)
        self.assertEqual(row["Себестоимость"], 0.0)
        self.assertEqual(row["Прибыль"], label)

    def test_pending_cross_border_buyout_reports_profit_as_unknown(self):
        """Real example: posting 0221866132-0045-1 — is_marketplace_buyout (cross-border, Ozon
        settles commission/payout on customs/currency timing) with commission_amount=payout=0 in
        Ozon's own live snapshot too, not just a request-window gap. Cost of goods must not be
        charged before revenue is recognized, and profit must not be reported as a real loss."""
        monthly = technical.monthly
        postings = [{"posting_number": "order", "status": "delivered",
                     "products": [{"offer_id": "sku", "quantity": 1, "price": "490.0000",
                                  "is_marketplace_buyout": True}]}]
        finances = {"order": [{"amount": -121.35, "sale_commission": 0, "accruals_for_sale": 0,
                               "has_commission_data": False}]}
        with patch.object(monthly, "get_transactions_by_posting", return_value=finances), \
                patch.object(monthly, "load_cost_map", return_value={"sku": 175}), \
                patch.object(monthly, "_safe_save_excel", return_value="unused") as save:
            monthly.to_excel(postings, "2026-08-01", "2026-08-31", 8, 2026,
                             output_file="unused", session=Mock())
        row = save.call_args.args[0].iloc[0]
        label = "Ожидает расчёта Ozon (трансграничный заказ)"
        self.assertEqual(row["Статус"], "ожидает расчёта")
        self.assertEqual(row["Цена продажи"], 490.0)
        self.assertEqual(row["Комиссия за продажу Ozon"], label)
        self.assertEqual(row["Логистика (Включает операционные ошибки продавца)"], label)
        self.assertEqual(row["Себестоимость"], 0.0)
        self.assertEqual(row["Прибыль"], label)

    def test_post_delivery_return_is_reported_as_returned_despite_ozon_status(self):
        """Ozon never flips a posting's status away from "delivered" for a return that happens
        after the item was received — real example: posting 40621894-2993-1, sale (+1501/-735)
        and its same-day return (-1501/+735) both booked as separate accruals, status stays
        "delivered". The report must not charge cost of goods for stock that came back, and must
        not fall back to the item's list price (that zero is a real net, not a data gap)."""
        monthly = technical.monthly
        postings = [{"posting_number": "order", "status": "delivered",
                     "products": [{"offer_id": "sku", "quantity": 1, "price": "1501.0000"}]}]
        finances = {"order": [
            {"amount": 668.23, "sale_commission": -735, "accruals_for_sale": 1501,
             "is_return": False, "has_commission_data": True},
            {"amount": -765.51, "sale_commission": 735, "accruals_for_sale": -1501,
             "is_return": True, "has_commission_data": True},
            {"amount": -103, "sale_commission": 0, "accruals_for_sale": 0,
             "is_return": False, "has_commission_data": False},
        ]}
        with patch.object(monthly, "get_transactions_by_posting", return_value=finances), \
                patch.object(monthly, "load_cost_map", return_value={"sku": 100}), \
                patch.object(monthly, "_safe_save_excel", return_value="unused") as save:
            monthly.to_excel(postings, "2026-08-01", "2026-08-31", 8, 2026,
                             output_file="unused", session=Mock())
        row = save.call_args.args[0].iloc[0]
        self.assertEqual(row["Статус"], "returned")
        self.assertEqual(row["Цена продажи"], 0)
        self.assertEqual(row["Себестоимость"], 0.0)
        self.assertAlmostEqual(row["Прибыль"], 668.23 - 765.51 - 103)

    def test_pending_orders_are_excluded_from_profit_and_cost_totals(self):
        monthly = technical.monthly
        postings = [
            {"posting_number": "settled", "status": "delivered",
             "products": [{"offer_id": "sku", "quantity": 1, "price": "1000.0000"}]},
            {"posting_number": "pending", "status": "delivered",
             "products": [{"offer_id": "sku", "quantity": 1, "price": "500.0000"}]},
        ]
        finances = {
            "settled": [{"amount": 600, "sale_commission": -400, "accruals_for_sale": 1000,
                        "has_commission_data": True}],
            "pending": [{"amount": -50, "sale_commission": 0, "accruals_for_sale": 0,
                        "has_commission_data": False}],
        }
        target = self.directory / "report.xlsx"
        with patch.object(monthly, "get_transactions_by_posting", return_value=finances), \
                patch.object(monthly, "load_cost_map", return_value={"sku": 200}):
            monthly.to_excel(postings, "2026-08-01", "2026-08-31", 8, 2026, output_file=str(target))
            monthly.calc_business_indicators(str(target), ozon_promotion_cost_override=0,
                                             external_marketing_cost_override=0)
        workbook = wb.load_workbook(target)
        try:
            sheet = workbook["Заказы"]
            values = {sheet.cell(row=r, column=16).value: sheet.cell(row=r, column=17).value
                     for r in range(1, sheet.max_row + 1) if sheet.cell(row=r, column=16).value}
            self.assertEqual(values["Заказы, ожидающие расчёта Ozon (не учтены в прибыли/себестоимости)"], 1)
            self.assertEqual(values["Чистая прибыль"], 400)
            self.assertEqual(values["Итоговая себестоимость"], -200)
            self.assertEqual(values["Количество доставленных заказов"], 1)
        finally:
            workbook.close()

    def test_payout_reconciliation_block_shows_ozon_balance_and_early_payment_fee(self):
        monthly = technical.monthly
        postings = [{"posting_number": "order", "status": "delivered",
                     "products": [{"offer_id": "sku", "quantity": 1, "price": "1000.0000"}]}]
        finances = {"order": [{"amount": 600, "sale_commission": -400, "accruals_for_sale": 1000,
                               "has_commission_data": True}]}
        balance_data = {
            "total": {
                "opening_balance": {"value": 100.0}, "closing_balance": {"value": 50.0},
                "accrued": {"value": 550.0}, "payments": [{"value": -600.0}],
            },
            "cashflows": {
                "sales": {"amount": {"value": 1000}, "fee": {"value": -400}},
                "returns": {"amount": {"value": 0}, "fee": {"value": 0}},
                "services": [{"name": "early_payment", "amount": {"value": -25.5}},
                            {"name": "acquiring", "amount": {"value": -12}}],
            },
        }
        target = self.directory / "report.xlsx"
        with patch.object(monthly, "get_transactions_by_posting", return_value=finances), \
                patch.object(monthly, "load_cost_map", return_value={"sku": 300}), \
                patch.object(monthly, "get_monthly_balance_reports", return_value=[balance_data]):
            monthly.to_excel(postings, "2026-08-01T00:00:00Z", "2026-08-31T23:59:59Z", 8, 2026, output_file=str(target))
            monthly.calc_business_indicators(str(target), date_from="2026-08-01T00:00:00Z",
                                             date_to="2026-08-31T23:59:59Z",
                                             ozon_promotion_cost_override=0, external_marketing_cost_override=0)
        workbook = wb.load_workbook(target)
        try:
            sheet = workbook["Заказы"]
            values = {sheet.cell(row=r, column=16).value: sheet.cell(row=r, column=17).value
                     for r in range(1, sheet.max_row + 1) if sheet.cell(row=r, column=16).value}
            self.assertEqual(values["Выплачено Ozon за период (реальный перевод)"], -600.0)
            self.assertEqual(values["Входящий баланс на начало периода"], 100.0)
            self.assertEqual(values["Исходящий баланс на конец периода"], 50.0)
            self.assertEqual(values["Начислено Ozon за период (по балансу, календарные дни)"], 550.0)
            self.assertEqual(values["Сумма начисления по нашему расчёту (по датам отгрузки)"], 600)
            self.assertAlmostEqual(values["Расхождение: баланс Ozon минус наш расчёт"], 550.0 - 600)
            self.assertEqual(values["Комиссия за ранний вывод средств"], -25.5)
            self.assertEqual(values["acquiring"], -12)
        finally:
            workbook.close()
