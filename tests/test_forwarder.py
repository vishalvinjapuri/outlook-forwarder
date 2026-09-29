import contextlib
import io
import tempfile
import threading
import unittest
import urllib.error
from pathlib import Path
from unittest.mock import patch

import forward_outlook_web as fw
import gmail_filters as gf

TARGET = "destination@example.com"


def rule():
    return {
        "Name": fw.rule_name(TARGET),
        "Enabled": True,
        "ForwardTo": [{"SmtpAddress": TARGET}],
        "StopProcessingRules": False,
    }


class FakeMailbox:
    def __init__(self, rules=None, stale=False, global_address=None):
        self.current = rules or []
        self.stale = stale
        self.writes = 0
        self.global_address = global_address

    def read(self):
        return {"AddressString": self.global_address}

    def rules(self):
        return self.current

    def create_rule(self, target):
        self.writes += 1
        if not self.stale:
            self.current.append(rule())


class RulesTests(unittest.TestCase):
    def test_create_and_rerun_without_duplicate(self):
        box = FakeMailbox()
        self.assertEqual(fw.configure(box, TARGET, True)["status"], "configured")
        self.assertEqual(
            fw.configure(box, TARGET, True)["status"], "already_configured"
        )
        self.assertEqual(box.writes, 1)

    def test_preview_does_not_write(self):
        box = FakeMailbox()
        self.assertEqual(fw.configure(box, TARGET)["status"], "would_configure")
        self.assertEqual(box.writes, 0)

    def test_conflicts_do_not_write(self):
        other = {"Name": "Other rule", "RedirectTo": [{"SmtpAddress": TARGET}]}
        for box in [
            FakeMailbox([other]),
            FakeMailbox([rule(), other]),
            FakeMailbox(global_address=TARGET),
            FakeMailbox([rule(), rule()]),
        ]:
            with self.subTest(box=box):
                self.assertFalse(fw.configure(box, TARGET, True)["verified"])
                self.assertEqual(box.writes, 0)

    def test_modified_managed_rule_requires_review(self):
        for key, value in [
            ("Enabled", False),
            ("InError", True),
            ("SubjectContainsWords", ["something"]),
            ("DeleteMessage", True),
            ("UnknownFutureAction", True),
        ]:
            changed = rule()
            changed[key] = value
            box = FakeMailbox([changed])
            self.assertFalse(fw.configure(box, TARGET, True)["verified"])
            self.assertEqual(box.writes, 0)

    @patch.object(fw.time, "sleep")
    def test_accepted_write_without_readback_fails(self, sleep):
        box = FakeMailbox(stale=True)
        self.assertFalse(fw.configure(box, TARGET, True)["verified"])
        self.assertEqual(box.writes, 1)
        self.assertEqual(sleep.call_count, 2)

    def test_null_server_defaults_are_inactive(self):
        current = rule()
        current.update(WithImportance="NullImportance", From=[], Priority=1)
        self.assertTrue(fw.matches(current, TARGET))


class ConcurrencyTests(unittest.TestCase):
    def test_bounded_parallelism_and_complete_results(self):
        barrier = threading.Barrier(3)
        lock = threading.Lock()
        active = 0
        maximum = 0

        def task(number, *_):
            nonlocal active, maximum
            with lock:
                active += 1
                maximum = max(maximum, active)
            barrier.wait(timeout=5)
            with lock:
                active -= 1
            return {"account_number": number, "status": "configured", "verified": True}

        with patch.object(fw, "run_account", side_effect=task):
            results = list(fw.run_accounts([("a", "b")] * 6, 4, TARGET, True, 3))
        self.assertEqual(maximum, 3)
        self.assertEqual({r["account_number"] for r in results}, set(range(4, 10)))

    def test_throttle_stops_unscheduled_accounts(self):
        def task(number, *_):
            return {"account_number": number, "status": "throttled", "verified": False}

        with patch.object(fw, "run_account", side_effect=task) as worker:
            results = list(fw.run_accounts([("a", "b")] * 8, 1, TARGET, True, 2))
        self.assertEqual(worker.call_count, 2)
        self.assertEqual(
            sum(r["status"] == "not_started_after_throttle" for r in results), 6
        )

    def test_regular_failure_does_not_stop_batch(self):
        result = {"status": "error", "verified": False}
        with patch.object(fw, "run_account", return_value=result) as worker:
            self.assertEqual(
                len(list(fw.run_accounts([("a", "b")] * 5, 1, TARGET, False, 2))), 5
            )
        self.assertEqual(worker.call_count, 5)

    def test_one_worker_serial(self):
        with patch.object(
            fw,
            "run_account",
            side_effect=lambda number, *_: {
                "account_number": number,
                "status": "configured",
                "verified": True,
            },
        ):
            results = list(fw.run_accounts([("a", "b")] * 5, 1, TARGET, True, 1))
        self.assertEqual([r["account_number"] for r in results], [1, 2, 3, 4, 5])


class InputAndSecurityTests(unittest.TestCase):
    def test_password_colons_and_duplicates(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "accounts.txt"
            path.write_text("a@example.com:pass:word\n\n")
            self.assertEqual(fw.accounts(path), [("a@example.com", "pass:word")])
            path.write_text("a@example.com:secret\nA@example.com:other\n")
            with self.assertRaisesRegex(fw.Stopped, "Duplicate account"):
                fw.accounts(path)

    def test_credential_exception_is_not_logged(self):
        with patch.object(
            fw, "authenticate", side_effect=ValueError("secret password and token")
        ):
            result = fw.run_account(1, ("a@example.com", "secret"), TARGET, True)
        self.assertEqual(result["reason"], "ValueError")
        self.assertNotIn("secret", str(result))
        self.assertNotIn("a@example.com", str(result))

    def test_redirect_origins(self):
        for url in [
            "http://login.live.com/",
            "https://evil.example/",
            "https://login.live.com@evil.example/",
            "https://u:p@login.live.com/",
        ]:
            with self.assertRaises(fw.Stopped):
                fw.checked(url)

    def test_429_is_distinct_and_no_retry(self):
        from unittest.mock import Mock

        opener = Mock()
        opener.open.side_effect = urllib.error.HTTPError(
            "https://login.live.com/", 429, "limit", {}, io.BytesIO()
        )
        with self.assertRaises(fw.Throttled):
            fw.Transport(retries=0).fetch(opener, "https://login.live.com/")
        self.assertEqual(opener.open.call_count, 1)

    def test_cli_rejects_invalid_worker_count(self):
        with patch(
            "sys.argv",
            ["forwarder", "--accounts", "none", "--to", TARGET, "--workers", "0"],
        ):
            with (
                contextlib.redirect_stderr(io.StringIO()),
                self.assertRaises(SystemExit) as exc,
            ):
                fw.main()
        self.assertEqual(exc.exception.code, 2)


class GmailTests(unittest.TestCase):
    def test_exact_coverage_and_only_never_spam_action(self):
        addresses = [f"account{i}@example.com" for i in range(103)]
        root = gf.build_filters(addresses).getroot()
        found = []
        for entry in root.findall(f"{{{gf.ATOM}}}entry"):
            props = {
                node.get("name"): node.get("value")
                for node in entry.findall(f"{{{gf.APPS}}}property")
            }
            self.assertEqual(set(props), {"from", "shouldNeverSpam"})
            self.assertEqual(props["shouldNeverSpam"], "true")
            chunk = props["from"].split(" OR ")
            self.assertLessEqual(len(chunk), 10)
            found.extend(chunk)
        self.assertEqual(found, addresses)

    def test_passwords_never_enter_filter_input(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "accounts.txt"
            path.write_text(
                "a@example.com:secret:password\nA@example.com:other\nb@example.com\n"
            )
            self.assertEqual(
                gf.read_addresses(path), ["a@example.com", "b@example.com"]
            )


if __name__ == "__main__":
    unittest.main()
