import threading
import unittest
from unittest.mock import Mock, patch

from scripts import wb_returns as returns
import test_technical_fixes as technical


def record(srid="r1", reason="Цвет", status="В пути в пвз", **updates):
    row = {"nmId": 1, "srid": srid, "reason": reason, "status": status,
           "returnType": "Возврат по инициативе покупателя", "orderId": 1}
    return dict(row, **updates)


class WBReturnsTests(unittest.TestCase):
    setUp = technical.OfflineTests.setUp

    def test_fetch_returns_records_and_caps_period(self):
        session = Mock()
        session.get.return_value = Mock(status_code=200, json=lambda: {"report": [record()]})
        client = returns.ReturnsClient("test", session=session)
        with patch.object(returns, "_next_request", 0):
            data = client.fetch(days=90)
        self.assertEqual(data["period_days"], returns.MAX_LOOKBACK_DAYS)
        self.assertEqual(len(data["records"]), 1)
        params = session.get.call_args.kwargs["params"]
        self.assertIn("dateFrom", params)
        self.assertIn("dateTo", params)

    def test_summarize_groups_by_reason_status_and_srid(self):
        data = {"records": [record(srid="r1", reason="Цвет", status="В пути в пвз"),
                            record(srid="r2", reason="Брак", status="Завершён"),
                            record(srid="r3", reason="Цвет", status="Завершён")]}
        summary = returns.summarize(data)
        self.assertEqual(summary["total"], 3)
        self.assertEqual(summary["by_reason"]["Цвет"], 2)
        self.assertEqual(summary["by_status"]["Завершён"], 2)
        self.assertEqual(set(summary["by_srid"]), {"r1", "r2", "r3"})
        self.assertEqual(summary["by_srid"]["r1"]["reason"], "Цвет")

    def test_request_errors_do_not_expose_token_or_response_body(self):
        session = Mock()
        session.get.return_value = Mock(status_code=403, text="secret-token")
        client = returns.ReturnsClient("secret-token", session=session)
        with patch.object(returns, "_next_request", 0):
            with self.assertRaises(returns.ReturnsError) as error:
                client.fetch()
        self.assertNotIn("secret-token", str(error.exception))
        self.assertEqual(session.get.call_args.kwargs["headers"]["Authorization"], "secret-token")

    def test_cancel_and_rate_limit_do_not_retry_early(self):
        cancel = threading.Event()
        session = Mock()
        response = Mock(status_code=429, headers={"X-Ratelimit-Retry": "120"})
        session.get.return_value = response
        client = returns.ReturnsClient("test", cancel=cancel, session=session,
                                       progress=lambda message: cancel.set())
        with patch.object(returns, "_next_request", 0):
            with self.assertRaises(returns.ReturnsError):
                client.fetch()
        session.get.assert_called_once()

    def test_malformed_response_is_rejected(self):
        session = Mock()
        session.get.return_value = Mock(status_code=200, json=lambda: {"report": "not-a-list"})
        client = returns.ReturnsClient("test", session=session)
        with patch.object(returns, "_next_request", 0):
            with self.assertRaises(returns.ReturnsError):
                client.fetch()


if __name__ == "__main__":
    unittest.main()
