import io
import tempfile
import threading
import unittest
import urllib.error
from datetime import datetime, timezone
from email.utils import format_datetime
from pathlib import Path
from unittest.mock import Mock, patch

from test_forwarder import TARGET, FakeMailbox, rule

import forward_outlook_web as fw
from resilience import (
    Cancelled,
    Stopped,
    Throttled,
    Transport,
    WriteUncertain,
    retry_after,
)
from state import Journal


class Clock:
    def __init__(self):
        self.now = 0
        self.waits = []

    def __call__(self):
        return self.now

    def wait(self, seconds):
        self.waits.append(seconds)
        self.now += seconds


def response(body=b"{}"):
    r = io.BytesIO(body)
    r.url = "https://outlook.live.com/"
    return r


def error(code, retry=None):
    return urllib.error.HTTPError(
        "https://outlook.live.com/",
        code,
        "message",
        {"Retry-After": retry} if retry else {},
        io.BytesIO(),
    )


class RetryTests(unittest.TestCase):
    def setup_transport(self, **kwargs):
        self.clock = Clock()
        self.events = []
        return Transport(
            clock=self.clock,
            waiter=self.clock.wait,
            jitter=lambda: 0,
            reporter=self.events.append,
            **kwargs,
        )

    def test_retry_after_formats(self):
        self.assertEqual(retry_after({"Retry-After": "240"}), 240)
        date = format_datetime(datetime.fromtimestamp(1600, timezone.utc), usegmt=True)
        self.assertEqual(retry_after({"Retry-After": date}, now=1000), 600)
        for text in ("bad", "NaN", "inf", "-2"):
            self.assertEqual(retry_after({"Retry-After": text}), 0)

    def test_two_and_four_minute_backoff_then_success(self):
        transport = self.setup_transport()
        opener = Mock()
        opener.open.side_effect = [error(429), error(429), response()]
        self.assertEqual(transport.fetch(opener, "url", safe=True)[0], b"{}")
        self.assertEqual([e["wait_seconds"] for e in self.events], [120, 240])
        self.assertGreaterEqual(self.clock.now, 360)
        self.assertEqual(opener.open.call_count, 3)

    def test_server_delay_is_not_shortened(self):
        transport = self.setup_transport()
        opener = Mock()
        opener.open.side_effect = [error(429, "600"), response()]
        transport.fetch(opener, "url", safe=True)
        self.assertGreaterEqual(self.clock.now, 600)

    def test_exhaustion_opens_shared_circuit(self):
        transport = self.setup_transport(retries=1)
        opener = Mock()
        opener.open.side_effect = [error(429), error(429)]
        with self.assertRaises(Throttled):
            transport.fetch(opener, "url", safe=True)
        second = Mock()
        with self.assertRaises(Throttled):
            transport.fetch(second, "url", safe=True)
        second.open.assert_not_called()

    def test_excessive_server_wait_stops_instead_of_retrying_early(self):
        transport = self.setup_transport(max_wait=300)
        opener = Mock()
        opener.open.side_effect = error(429, "600")
        with self.assertRaises(Throttled):
            transport.fetch(opener, "url", safe=True)
        self.assertEqual(opener.open.call_count, 1)

    def test_cooldown_applies_to_other_workers(self):
        transport = self.setup_transport()
        transport.pause(120, "http_429", 1)
        observed = []
        opener = Mock()
        opener.open.side_effect = lambda *a, **k: (
            observed.append(self.clock.now) or response()
        )
        transport.fetch(opener, "url", safe=True)
        self.assertGreaterEqual(observed[0], 120)

    def test_safe_read_server_and_body_failures_retry(self):
        transport = self.setup_transport()
        broken = Mock()
        broken.__enter__ = Mock(return_value=broken)
        broken.__exit__ = Mock(return_value=False)
        broken.read.side_effect = ConnectionResetError()
        opener = Mock()
        opener.open.side_effect = [error(503, "10"), broken, response()]
        transport.fetch(opener, "url", safe=True)
        self.assertEqual(opener.open.call_count, 3)

    def test_mutation_transient_failure_is_never_replayed(self):
        for failure in (error(503), TimeoutError(), urllib.error.URLError("reset")):
            transport = self.setup_transport()
            opener = Mock()
            opener.open.side_effect = failure
            with self.assertRaises(WriteUncertain):
                transport.fetch(opener, "url", mutation=True)
            self.assertEqual(opener.open.call_count, 1)

    def test_explicit_429_mutation_rejection_can_retry(self):
        transport = self.setup_transport()
        opener = Mock()
        opener.open.side_effect = [error(429), response()]
        transport.fetch(opener, "url", mutation=True)
        self.assertEqual(opener.open.call_count, 2)
        self.assertGreaterEqual(self.clock.now, 120)

    def test_bad_credentials_and_forbidden_are_not_retried(self):
        for code in (400, 401, 403):
            transport = self.setup_transport()
            opener = Mock()
            opener.open.side_effect = error(code)
            with self.assertRaises(Stopped):
                transport.fetch(opener, "url", safe=True)
            self.assertEqual(opener.open.call_count, 1)

    def test_cancellation_wakes_minutes_long_sleep(self):
        sleeping = threading.Event()
        transport = Transport(reporter=lambda _: None)

        def waiter(seconds):
            sleeping.set()
            return transport.cancelled.wait(seconds)

        transport.waiter = waiter
        transport.pause(120, "http_429", 1)
        outcomes = []

        def worker():
            try:
                transport.gate()
            except Cancelled:
                outcomes.append("cancelled")

        thread = threading.Thread(target=worker)
        thread.start()
        self.assertTrue(sleeping.wait(1))
        transport.cancel()
        thread.join(timeout=1)
        self.assertFalse(thread.is_alive())
        self.assertEqual(outcomes, ["cancelled"])


class JournalTests(unittest.TestCase):
    def test_persistence_and_no_raw_identifiers(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "state.sqlite"
            Journal(path).begin("source@example.com", TARGET)
            self.assertTrue(Journal(path).pending("source@example.com", TARGET))
            self.assertNotIn(b"source@example.com", path.read_bytes())
            self.assertNotIn(TARGET.encode(), path.read_bytes())
            Journal(path).clear("source@example.com", TARGET)
            self.assertFalse(Journal(path).pending("source@example.com", TARGET))

    def test_changing_destination_cannot_hide_an_unresolved_write(self):
        with tempfile.TemporaryDirectory() as folder:
            journal = Journal(Path(folder) / "state.sqlite")
            journal.begin("source@example.com", TARGET)
            box = fw.Mailbox("source@example.com", "fake", journal=journal)
            box.call = Mock()
            result = fw.configure(box, "other@example.com", True)
            self.assertEqual(result["status"], "unresolved_write_to_other_destination")
            box.call.assert_not_called()
            journal.clear("source@example.com", "other@example.com")
            self.assertTrue(journal.pending("source@example.com", TARGET))

    def test_second_instance_cannot_use_same_state(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "state.sqlite"
            with Journal(path):
                with self.assertRaises(Stopped):
                    with Journal(path):
                        pass
            with Journal(path):
                pass

    def test_pending_write_prevents_second_create(self):
        box = FakeMailbox()
        box.pending_write = lambda _: True
        result = fw.configure(box, TARGET, True)
        self.assertEqual(result["status"], "unresolved_previous_write")
        self.assertEqual(box.writes, 0)

    @patch.object(fw.time, "sleep")
    def test_uncertain_write_reconciles_without_replaying(self, _):
        box = FakeMailbox()

        def create(target):
            box.writes += 1
            box.current.append(rule())
            raise WriteUncertain("socket closed")

        box.create_rule = create
        self.assertTrue(fw.configure(box, TARGET, True)["verified"])
        self.assertEqual(box.writes, 1)

    @patch.object(fw.time, "sleep")
    def test_uncertain_write_without_readback_remains_pending(self, _):
        with tempfile.TemporaryDirectory() as folder:
            journal = Journal(Path(folder) / "state.sqlite")
            box = fw.Mailbox("source@example.com", "fake", journal=journal)
            box.read = lambda: {"AddressString": None}
            box.rules = lambda: []
            box._create_rule = Mock(side_effect=WriteUncertain("socket closed"))
            self.assertEqual(
                fw.configure(box, TARGET, True)["status"], "write_outcome_unknown"
            )
            self.assertTrue(journal.pending(box.email, TARGET))
            self.assertEqual(
                fw.configure(box, TARGET, True)["status"], "unresolved_previous_write"
            )
            self.assertEqual(box._create_rule.call_count, 1)

    def test_definite_rejection_clears_intent(self):
        with tempfile.TemporaryDirectory() as folder:
            journal = Journal(Path(folder) / "state.sqlite")
            box = fw.Mailbox("source@example.com", "fake", journal=journal)
            box._create_rule = Mock(side_effect=Stopped("HTTP 403"))
            with self.assertRaises(Stopped):
                box.create_rule(TARGET)
            self.assertFalse(journal.pending(box.email, TARGET))

    def test_bad_success_response_retains_intent(self):
        with tempfile.TemporaryDirectory() as folder:
            journal = Journal(Path(folder) / "state.sqlite")
            transport = Mock()
            transport.fetch.return_value = (b"not json", "url")
            box = fw.Mailbox("source@example.com", "fake", transport, journal)
            with self.assertRaises(WriteUncertain):
                box.create_rule(TARGET)
            self.assertTrue(journal.pending(box.email, TARGET))

    def test_readback_clears_crash_record_without_creating(self):
        with tempfile.TemporaryDirectory() as folder:
            journal = Journal(Path(folder) / "state.sqlite")
            journal.begin("source@example.com", TARGET)
            box = fw.Mailbox("source@example.com", "fake", journal=journal)
            box.read = lambda: {"AddressString": None}
            box.rules = lambda: [rule()]
            box._create_rule = Mock()
            self.assertTrue(fw.configure(box, TARGET, True)["verified"])
            self.assertFalse(journal.pending(box.email, TARGET))
            box._create_rule.assert_not_called()

    def test_incomplete_rule_listing_cannot_trigger_a_write(self):
        box = fw.Mailbox("source@example.com", "fake")
        box.call = Mock(return_value={"WasSuccessful": True})
        with self.assertRaises(Stopped):
            box.rules()

    def test_expired_authorization_refreshes_once(self):
        mailbox = Mock()
        with (
            patch.object(fw, "authenticate", return_value="fake") as auth,
            patch.object(fw, "Mailbox", return_value=mailbox),
            patch.object(
                fw,
                "configure",
                side_effect=[
                    fw.AuthenticationExpired("expired"),
                    {"status": "already_configured", "verified": True},
                ],
            ),
        ):
            result = fw.run_account(
                1,
                ("source@example.com", "password"),
                TARGET,
                False,
                Transport(reporter=lambda _: None),
            )
        self.assertTrue(result["verified"])
        self.assertEqual(auth.call_count, 2)
