import threading
import unittest
from unittest.mock import Mock, patch

from scripts import wb_funnel as funnel
import test_technical_fixes as technical


def product(nm_id=123, views=100, cart=10, orders=3, buyouts=2, cancel=1):
    return {
        "product": {"nmId": nm_id, "title": "Доска", "vendorCode": "1114", "subjectName": "Доски"},
        "statistic": {"selected": {
            "openCount": views, "cartCount": cart, "orderCount": orders, "orderSum": orders * 900,
            "buyoutCount": buyouts, "buyoutSum": buyouts * 900, "cancelCount": cancel,
            "conversions": {"addToCartPercent": 10, "cartToOrderPercent": 30, "buyoutPercent": 67},
        }},
    }


class WBFunnelTests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def test_fetch_paginates_until_short_page(self):
        client = funnel.FunnelClient("test")
        first_page = {"products": [product(nm_id=i) for i in range(funnel.PAGE_LIMIT)], "currency": "RUB"}
        second_page = {"products": [product(nm_id=9999)], "currency": "RUB"}
        with patch.object(client, "request", side_effect=[first_page, second_page]) as requested:
            from datetime import date
            data = client.fetch(date(2026, 8, 1), date(2026, 8, 31))
        self.assertEqual(len(data["products"]), funnel.PAGE_LIMIT + 1)
        first_call = requested.call_args_list[0].args[0]
        self.assertEqual(first_call["selectedPeriod"], {"start": "2026-08-01", "end": "2026-08-31"})
        self.assertEqual(requested.call_args_list[1].args[0]["offset"], funnel.PAGE_LIMIT)

    def test_summarize_aggregates_totals_and_conversions(self):
        data = {"products": [product(nm_id=1, views=100, cart=10, orders=4, buyouts=2),
                             product(nm_id=2, views=50, cart=5, orders=2, buyouts=1)]}
        summary = funnel.summarize(data)
        self.assertEqual(summary["products"], 2)
        self.assertEqual(summary["views"], 150)
        self.assertEqual(summary["cart"], 15)
        self.assertEqual(summary["orders"], 6)
        self.assertEqual(summary["buyouts"], 3)
        self.assertEqual(summary["add_to_cart_pct"], round(15 / 150 * 100, 1))
        self.assertEqual(summary["cart_to_order_pct"], round(6 / 15 * 100, 1))
        self.assertEqual(summary["buyout_pct"], round(3 / 6 * 100, 1))
        self.assertEqual(set(summary["by_nm"]), {"1", "2"})

    def test_summarize_handles_zero_views_without_dividing_by_zero(self):
        summary = funnel.summarize({"products": []})
        self.assertEqual(summary["products"], 0)
        self.assertIsNone(summary["add_to_cart_pct"])
        self.assertIsNone(summary["cart_to_order_pct"])
        self.assertIsNone(summary["buyout_pct"])

    def test_request_errors_do_not_expose_token_or_response_body(self):
        session = Mock()
        session.post.return_value = Mock(status_code=403, text="secret-token")
        client = funnel.FunnelClient("secret-token", session=session)
        with patch.object(funnel, "_next_request", 0):
            with self.assertRaises(funnel.FunnelError) as error:
                client.request({})
        self.assertNotIn("secret-token", str(error.exception))
        self.assertEqual(session.post.call_args.kwargs["headers"]["Authorization"], "secret-token")

    def test_cancel_and_rate_limit_do_not_retry_early(self):
        cancel = threading.Event()
        session = Mock()
        response = Mock(status_code=429, headers={"X-Ratelimit-Retry": "45"})
        session.post.return_value = response
        client = funnel.FunnelClient("test", cancel=cancel, session=session,
                                     progress=lambda message: cancel.set())
        with patch.object(funnel, "_next_request", 0):
            with self.assertRaises(funnel.FunnelError):
                client.request({})
        session.post.assert_called_once()

    def test_malformed_response_is_rejected(self):
        session = Mock()
        session.post.return_value = Mock(status_code=200, json=lambda: {"data": {"products": "nope"}})
        client = funnel.FunnelClient("test", session=session)
        with patch.object(funnel, "_next_request", 0):
            with self.assertRaises(funnel.FunnelError):
                client.request({})


if __name__ == "__main__":
    unittest.main()
