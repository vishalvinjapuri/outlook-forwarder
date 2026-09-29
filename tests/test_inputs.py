import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import forward_outlook as graph
import forward_outlook_web as web
import gmail_filters
from input_files import read_text_input
from resilience import Stopped


class TextInputTests(unittest.TestCase):
    def test_any_filename_and_supported_encoding(self):
        text = "first@example.com: p:a:ss \r\n\r\nsecond@example.com:pass\r\n"
        with tempfile.TemporaryDirectory() as folder:
            for name in ("accounts.txt", "accounts.list", "accounts", "plain.rtf"):
                for encoding in ("utf-8", "utf-8-sig", "utf-16", "utf-32"):
                    with self.subTest(name=name, encoding=encoding):
                        path = Path(folder) / name
                        path.write_bytes(text.encode(encoding))
                        self.assertEqual(
                            web.accounts(path),
                            [
                                ("first@example.com", " p:a:ss "),
                                ("second@example.com", "pass"),
                            ],
                        )
                        self.assertEqual(
                            graph.read_accounts(path),
                            ["first@example.com", "second@example.com"],
                        )
                        self.assertEqual(
                            gmail_filters.read_addresses(path),
                            ["first@example.com", "second@example.com"],
                        )

    def test_stdin_text_stream(self):
        with patch("sys.stdin", io.StringIO("first@example.com:secret\n")):
            self.assertEqual(web.accounts("-"), [("first@example.com", "secret")])

    def test_stdin_binary_utf16(self):
        stream = Mock()
        stream.buffer = io.BytesIO("first@example.com:secret".encode("utf-16"))
        with patch("sys.stdin", stream):
            self.assertEqual(web.accounts("-"), [("first@example.com", "secret")])

    def test_rtf_without_converter_gives_actionable_error(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "input.rtf"
            path.write_bytes(b"{\\rtf1 data}")
            with patch("input_files.shutil.which", return_value=None):
                with self.assertRaisesRegex(Stopped, "export as plain text"):
                    read_text_input(path)

    def test_rtf_converter_does_not_print_input(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "rich.data"
            path.write_bytes(b"{\\rtf1 secret}")
            with (
                patch("input_files.shutil.which", return_value="/bin/textutil"),
                patch(
                    "input_files.subprocess.run",
                    return_value=Mock(returncode=0, stdout=b"a@example.com:secret"),
                ) as run,
            ):
                self.assertEqual(web.accounts(path), [("a@example.com", "secret")])
                self.assertTrue(run.call_args.kwargs["capture_output"])

    def test_invalid_bytes_fail_safely(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "input"
            path.write_bytes(b"\x80\x81")
            with self.assertRaisesRegex(Stopped, "Unsupported text encoding"):
                read_text_input(path)

    def test_cli_defaults_to_entire_list(self):
        rows = [(f"account{i}@example.com", "password") for i in range(7)]
        with (
            patch(
                "sys.argv",
                ["forwarder", "--accounts", "file", "--to", "destination@example.com"],
            ),
            patch.object(web, "accounts", return_value=rows),
            patch.object(web, "Journal"),
            patch.object(web, "run_accounts", return_value=[]) as run,
        ):
            self.assertEqual(web.main(), 0)
        self.assertEqual(run.call_args.args[0], rows)

    def test_cli_explicit_range_is_preserved(self):
        rows = [(f"account{i}@example.com", "password") for i in range(7)]
        with (
            patch(
                "sys.argv",
                [
                    "forwarder",
                    "--accounts",
                    "file",
                    "--to",
                    "destination@example.com",
                    "--start-at",
                    "2",
                    "--limit",
                    "2",
                ],
            ),
            patch.object(web, "accounts", return_value=rows),
            patch.object(web, "Journal"),
            patch.object(web, "run_accounts", return_value=[]) as run,
        ):
            self.assertEqual(web.main(), 0)
        self.assertEqual(run.call_args.args[0], rows[1:3])

    def test_unknown_action_is_not_treated_as_safe_read(self):
        transport = Mock()
        transport.fetch.return_value = (
            json.dumps({"WasSuccessful": True}).encode(),
            "url",
        )
        box = web.Mailbox("source@example.com", "fake", transport)
        box.call("EnableInboxRule", {})
        self.assertTrue(transport.fetch.call_args.kwargs["mutation"])
        self.assertFalse(transport.fetch.call_args.kwargs["safe"])
