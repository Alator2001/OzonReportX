import threading
import unittest
from unittest.mock import Mock, patch

from scripts import wb_orders as feed
import test_technical_fixes as technical


def order(status="created", is_mp=False, **updates):
    row = {"nmId": 1, "chrtId": 1, "srid": "order", "status": status, "isMp": is_mp,
           "warehouseName": "Склад WB", "sellerPrice": 1000}
    return dict(row, **updates)


class WBOrderFeedTests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def test_fetch_paginates_using_snapshot_time_and_offset(self):
        client = feed.OrderFeedClient("test")
        first_page = {"snapshotTime": "2026-09-14T00:00:00Z",
                      "orders": [order(srid=str(i)) for i in range(feed.PAGE_LIMIT)]}
        second_page = {"snapshotTime": "2026-09-14T00:00:00Z", "orders": [order(srid="last")]}
        with patch.object(client, "request", side_effect=[first_page, second_page]) as requested:
            result = client.fetch(days=10)
        self.assertEqual(len(result["orders"]), feed.PAGE_LIMIT + 1)
        first_call, second_call = requested.call_args_list
        self.assertEqual(first_call.args[0]["pagination"], {"offset": 0, "limit": feed.PAGE_LIMIT})
        self.assertEqual(second_call.args[0]["pagination"],
                         {"offset": feed.PAGE_LIMIT, "limit": feed.PAGE_LIMIT,
                          "snapshotTime": "2026-09-14T00:00:00Z"})

    def test_fetch_caps_lookback_to_31_days(self):
        client = feed.OrderFeedClient("test")
        with patch.object(client, "request", return_value={"snapshotTime": "", "orders": []}) as requested:
            client.fetch(days=90)
        period = requested.call_args.args[0]["selectedPeriod"]
        self.assertIn("T", period["start"])

    def test_summarize_buckets_by_status_and_channel(self):
        data = {"orders": [
            order(status="created", is_mp=True),
            order(status="buyout", is_mp=True),
            order(status="buyout", is_mp=False),
            order(status="cancel", is_mp=False),
            order(status="cancel", is_mp=False),
        ]}
        summary = feed.summarize(data)
        self.assertEqual(summary["total"], 5)
        self.assertEqual(summary["in_progress"], 1)
        self.assertEqual(summary["delivered"], 2)
        self.assertEqual(summary["cancelled"], 2)
        self.assertEqual(summary["fbs"], {"total": 2, "in_progress": 1, "delivered": 1, "cancelled": 0})
        self.assertEqual(summary["fbo"], {"total": 3, "in_progress": 0, "delivered": 1, "cancelled": 2})

    def test_request_errors_do_not_expose_token_or_response_body(self):
        session = Mock()
        session.post.return_value = Mock(status_code=403, text="secret-token")
        client = feed.OrderFeedClient("secret-token", session=session)
        with patch.object(feed, "_next_request", 0):
            with self.assertRaises(feed.OrderFeedError) as error:
                client.request({})
        self.assertNotIn("secret-token", str(error.exception))
        self.assertEqual(session.post.call_args.kwargs["headers"]["Authorization"], "secret-token")

    def test_cancel_and_rate_limit_do_not_retry_early(self):
        cancel = threading.Event()
        session = Mock()
        response = Mock(status_code=429, headers={"X-Ratelimit-Retry": "120"})
        session.post.return_value = response
        client = feed.OrderFeedClient("test", cancel=cancel, session=session,
                                      progress=lambda message: cancel.set())
        with patch.object(feed, "_next_request", 0):
            with self.assertRaises(feed.OrderFeedError):
                client.request({})
        session.post.assert_called_once()

    def test_malformed_response_is_rejected(self):
        session = Mock()
        session.post.return_value = Mock(status_code=200, json=lambda: {"data": {"orders": "not-a-list"}})
        client = feed.OrderFeedClient("test", session=session)
        with patch.object(feed, "_next_request", 0):
            with self.assertRaises(feed.OrderFeedError):
                client.request({})


if __name__ == "__main__":
    unittest.main()
