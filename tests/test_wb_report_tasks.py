import threading
import unittest
from unittest.mock import Mock, patch

from scripts import wb_report_tasks as tasks
import test_technical_fixes as technical


def response(status_code=200, json_body=None, headers=None, text=""):
    mock = Mock(status_code=status_code, headers=headers or {}, text=text)
    if json_body is not None:
        mock.json = lambda: json_body
    else:
        mock.json = Mock(side_effect=ValueError("no body"))
    return mock


class WBReportTaskTests(unittest.TestCase):
    def setUp(self):
        technical.OfflineTests.setUp(self)
        tasks._next_request.clear()

    def run_task(self, session, cancel=None):
        return tasks.run_report_task(
            token="test", cancel=cancel or threading.Event(), progress=lambda message: None, session=session,
            task_key="t", create_url="https://x/create", create_params={"a": 1},
            status_url="https://x/status/{task_id}", download_url="https://x/download/{task_id}",
            poll_interval=0, poll_timeout=5, status_min_interval=0,
        )

    def test_full_cycle_succeeds_after_processing_status(self):
        session = Mock()
        session.get.side_effect = [
            response(json_body={"data": {"taskId": "abc"}}),
            response(json_body={"data": {"id": "abc", "status": "processing"}}),
            response(json_body={"data": {"id": "abc", "status": "done"}}),
            response(json_body=[{"total": 5}]),
        ]
        rows = self.run_task(session)
        self.assertEqual(rows, [{"total": 5}])
        urls = [call.args[0] for call in session.get.call_args_list]
        self.assertEqual(urls, ["https://x/create", "https://x/status/abc", "https://x/status/abc", "https://x/download/abc"])

    def test_download_204_means_empty_report(self):
        session = Mock()
        session.get.side_effect = [
            response(json_body={"data": {"taskId": "abc"}}),
            response(json_body={"data": {"status": "done"}}),
            response(status_code=204),
        ]
        self.assertEqual(self.run_task(session), [])

    def test_poll_timeout_raises(self):
        session = Mock()
        session.get.side_effect = [
            response(json_body={"data": {"taskId": "abc"}}),
        ] + [response(json_body={"data": {"status": "processing"}})] * 200
        with self.assertRaises(tasks.ReportTaskError):
            tasks.run_report_task(
                token="test", cancel=threading.Event(), progress=lambda message: None, session=session,
                task_key="t", create_url="https://x/create", create_params={"a": 1},
                status_url="https://x/status/{task_id}", download_url="https://x/download/{task_id}",
                poll_interval=0.01, poll_timeout=0.2, status_min_interval=0,
            )

    def test_missing_task_id_is_rejected(self):
        session = Mock()
        session.get.return_value = response(json_body={"data": {}})
        with self.assertRaises(tasks.ReportTaskError):
            self.run_task(session)

    def test_cancel_stops_between_polls(self):
        cancel = threading.Event()
        session = Mock()
        def create(*_a, **_k):
            return response(json_body={"data": {"taskId": "abc"}})
        def status(*_a, **_k):
            cancel.set()
            return response(json_body={"data": {"status": "processing"}})
        session.get.side_effect = [create(), status()]
        with self.assertRaises(tasks.ReportTaskError):
            self.run_task(session, cancel=cancel)

    def test_request_errors_do_not_expose_token(self):
        session = Mock()
        session.get.return_value = response(status_code=403, text="secret-token")
        with self.assertRaises(tasks.ReportTaskError) as error:
            self.run_task(session)
        self.assertNotIn("secret-token", str(error.exception))

    def test_rate_limit_uses_retry_header_and_does_not_retry_early(self):
        session = Mock()
        session.get.return_value = response(status_code=429, headers={"X-Ratelimit-Retry": "120"})
        cancel = threading.Event()
        with patch.object(threading.Event, "wait", side_effect=lambda *_a, **_k: cancel.set()):
            with self.assertRaises(tasks.ReportTaskError):
                self.run_task(session, cancel=cancel)
        session.get.assert_called_once()


if __name__ == "__main__":
    unittest.main()
