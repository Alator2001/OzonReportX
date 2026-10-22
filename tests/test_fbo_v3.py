"""Offline regressions for the FBO v3 migration."""
import unittest
from unittest.mock import Mock

import test_technical_fixes as technical
from test_technical_fixes import monthly


class FboV3Tests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def session(self, *pages):
        session = Mock()
        session.post.side_effect = [Mock(json=Mock(return_value=page)) for page in pages]
        return session

    def test_cursor_pagination_and_new_money_format(self):
        posting = {
            "posting_number": "1", "status": "awaiting_packaging",
            "products": [{"price": {"amount": "28990", "currency": "RUB"}}],
            "financial_data": {"products": [{"commission": {
                "amount": 100, "percent": 10, "currency": "RUB"}}]},
        }
        session = self.session(
            {"postings": [posting], "has_next": True, "cursor": "opaque-next"},
            {"postings": [{"posting_number": "2"}], "has_next": False, "cursor": ""},
        )
        result = monthly.get_fbo_orders("start", "end", session=session)
        self.assertEqual([p["posting_number"] for p in result], ["1", "2"])
        self.assertEqual(result[0]["products"][0]["price"], "28990")
        self.assertEqual(result[0]["financial_data"]["products"][0]["commission_amount"], 100)
        self.assertTrue(all(p["__schema"] == "FBO" for p in result))
        calls = session.post.call_args_list
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[0].args[0], "https://api-seller.ozon.ru/v3/posting/fbo/list")
        payload = calls[0].kwargs["json"]
        self.assertEqual(payload, {
            "sort_dir": "ASC", "cursor": "", "limit": 100,
            "filter": {"since": "start", "to": "end", "statuses": [
                "awaiting_packaging", "awaiting_deliver", "delivering", "delivered", "cancelled"]},
            "with": {"analytics_data": True, "financial_data": True},
        })
        self.assertEqual(calls[1].kwargs["json"]["cursor"], "opaque-next")

    def test_failed_second_page_raises_instead_of_returning_partial_orders(self):
        session = self.session({"postings": [{}], "has_next": True, "cursor": "next"})
        first = session.post.side_effect
        session.post.side_effect = [next(first), RuntimeError("HTTP 429")]
        with self.assertRaisesRegex(RuntimeError, "429"):
            monthly.get_fbo_orders("start", "end", session=session)

    def test_missing_or_repeated_cursor_stops_pagination(self):
        for cursor in (None, "", "next"):
            with self.subTest(cursor=cursor):
                session = self.session(
                    {"postings": [], "has_next": True, "cursor": "next"},
                    {"postings": [], "has_next": True, "cursor": cursor},
                )
                with self.assertRaisesRegex(RuntimeError, "cursor"):
                    monthly.get_fbo_orders("start", "end", session=session)
                self.assertEqual(session.post.call_count, 2)

    def test_malformed_response_is_not_an_empty_report(self):
        for page in ({"result": []}, {}, {"postings": [], "has_next": "false"}):
            for function in (monthly.get_fbo_orders, monthly.get_orders):
                with self.subTest(page=page, function=function.__name__):
                    with self.assertRaisesRegex(RuntimeError, "Некорректный ответ"):
                        function("start", "end", session=self.session(page))

    def test_fbs_uses_v4_cursor_with_all_statuses_in_one_pass(self):
        posting = {
            "posting_number": "1", "status": "delivered", "shipment_date": "2026-09-20T10:00:00Z",
            "products": [{"offer_id": "7012", "quantity": 1, "price": {"amount": "960", "currency": "RUB"}}],
            "financial_data": {"products": [{"commission": {"amount": -460.8, "percent": 48, "currency": "RUB"}}]},
        }
        session = self.session(
            {"postings": [posting], "has_next": True, "cursor": "opaque-next"},
            {"postings": [{"posting_number": "2"}], "has_next": False, "cursor": ""},
        )
        result = monthly.get_orders("start", "end", session=session)
        self.assertEqual([p["posting_number"] for p in result], ["1", "2"])
        self.assertEqual(result[0]["products"][0]["price"], "960")
        self.assertEqual(result[0]["financial_data"]["products"][0]["commission_amount"], -460.8)
        self.assertTrue(all(p["__schema"] == "FBS" for p in result))
        calls = session.post.call_args_list
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[0].args[0], "https://api-seller.ozon.ru/v4/posting/fbs/list")
        self.assertEqual(calls[0].kwargs["json"], {
            "cursor": "", "limit": 100,
            "filter": {"since": "start", "to": "end", "statuses": [
                "awaiting_packaging", "awaiting_deliver", "delivering", "delivered", "cancelled"]},
            "with": {"analytics_data": True, "financial_data": True},
        })
        self.assertEqual(calls[1].kwargs["json"]["cursor"], "opaque-next")
