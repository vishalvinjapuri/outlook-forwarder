#!/usr/bin/env python3
"""Headless consumer Outlook forwarding via the web client's private HTTP API.
Bounded concurrency, shared cooldowns, verified writes, and resumable state.
"""

import argparse
import gzip
import http.cookiejar
import json
import re
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from html.parser import HTMLParser
from pathlib import Path

from input_files import read_text_input
from resilience import (
    AuthenticationExpired,
    Cancelled,
    Stopped,
    Throttled,
    Transport,
    WriteUncertain,
)
from state import Journal

HOSTS = {
    "login.live.com",
    "account.live.com",
    "account.microsoft.com",
    "outlook.live.com",
    "login.microsoftonline.com",
    "privacynotice.account.microsoft.com",
}
HEADER = {
    "__type": "JsonRequestHeaders:#Exchange",
    "RequestServerVersion": "V2018_01_08",
}


def checked(url):
    p = urllib.parse.urlsplit(url)
    if p.scheme != "https" or p.hostname not in HOSTS or p.username or p.password:
        raise Stopped("Unexpected sign-in origin")
    return url


class Fields(HTMLParser):
    def __init__(self, html):
        super().__init__()
        self.fields = {}
        self.feed(html)

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "input" and a.get("name"):
            self.fields[a["name"]] = a.get("value", "")


class Capture(urllib.request.HTTPRedirectHandler):
    def __init__(self):
        self.captured = None

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        p = urllib.parse.urlsplit(checked(newurl))
        q = urllib.parse.parse_qs(p.fragment)
        if (
            p.hostname == "outlook.live.com"
            and p.path == "/owa/auth/dt.aspx"
            and q.get("access_token")
        ):
            self.captured = q
            return None
        return super().redirect_request(req, fp, code, msg, headers, newurl)


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args):
        return None


def authenticate(email, password, transport=None):
    transport = transport or Transport(retries=0)
    state = secrets.token_urlsafe(32)
    capture = Capture()
    opener = urllib.request.build_opener(
        capture, urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar())
    )
    opener.addheaders = [
        ("User-Agent", "Mozilla/5.0"),
        ("Accept-Language", "en-US,en;q=0.9"),
    ]
    query = urllib.parse.urlencode(
        dict(
            client_id="292841",
            scope="service::outlook.office.com::SA_20MIN",
            response_type="token",
            redirect_uri="https://outlook.live.com/owa/auth/dt.aspx",
            state=state,
        )
    )
    body, referer = transport.fetch(
        opener, "https://login.live.com/oauth20_authorize.srf?" + query, safe=True
    )
    html = body.decode("utf-8", "replace")
    match = re.search(r"(?:var\s+)?ServerData\s*=\s*", html)
    if not match:
        raise Stopped("Microsoft sign-in page changed or requires interaction")
    config = json.JSONDecoder().raw_decode(html[match.end() :])[0]
    fields = Fields(config.get("sFTTag", "")).fields
    if not fields.get("PPFT") or not config.get("urlPost"):
        raise Stopped("Microsoft did not provide a password form")
    # Credentials are submitted only to Microsoft's actual password endpoint.
    post = checked(config["urlPost"])
    if urllib.parse.urlsplit(post).hostname != "login.live.com":
        raise Stopped("Unexpected password endpoint")
    fields.update(
        login=email,
        loginfmt=email,
        passwd=password,
        type="11",
        LoginOptions="3",
        PPSX="Passp",
        NewUser="1",
        i13="0",
        i19="0",
    )
    req = urllib.request.Request(
        post,
        data=urllib.parse.urlencode(fields).encode(),
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": "https://login.live.com",
            "Referer": referer,
        },
    )
    transport.fetch(opener, req, capture_redirect=True)
    q = capture.captured
    if not q or q.get("state") != [state] or not q.get("access_token"):
        raise Stopped(
            "Sign-in unsuccessful or Microsoft requires MFA, CAPTCHA, consent, or account recovery"
        )
    token = q["access_token"][0]
    if any(c in token for c in '\r\n"'):
        raise Stopped("Unexpected token format")
    return 'MSAuth1.0 usertoken="' + token + '", type="MSACT"'


class Mailbox:
    def __init__(self, email, token, transport=None, journal=None):
        self.email, self.token = email, token
        self.transport = transport or Transport(retries=0)
        self.journal = journal
        self.opener = urllib.request.build_opener(NoRedirect())

    def call(self, action, data):
        req = urllib.request.Request(
            "https://outlook.live.com/owa/0/service.svc?action=" + action + "&app=Mail",
            data=json.dumps(data).encode(),
            headers={
                "Authorization": self.token,
                "Content-Type": "application/json; charset=utf-8",
                "Action": action,
                "X-AnchorMailbox": "SMTP:" + self.email,
            },
        )
        mutation = action not in {"GetMailboxByIdentity", "GetInboxRule"}
        body, _ = self.transport.fetch(
            self.opener, req, safe=not mutation, mutation=mutation
        )
        try:
            if body.startswith(b"\x1f\x8b"):
                body = gzip.decompress(body)
            result = json.loads(body)
            if not isinstance(result, dict) or "WasSuccessful" not in result:
                raise ValueError("Unexpected API response")
        except (ValueError, OSError, EOFError):
            if mutation:
                raise WriteUncertain(
                    "Write returned an unreadable response; readback required"
                ) from None
            raise Stopped("Unexpected Outlook response format") from None
        if not result.get("WasSuccessful"):
            raise Stopped(
                "Outlook rejected operation; code " + str(result.get("ErrorCode"))
            )
        return result

    def read(self):
        return self.call(
            "GetMailboxByIdentity",
            {
                "__type": "IdentityRequest:#Exchange",
                "Header": HEADER,
                "Identity": {"RawIdentity": self.email},
            },
        )["MailboxOptions"]

    def rules(self):
        result = self.call(
            "GetInboxRule",
            {
                "__type": "GetInboxRuleRequest:#Exchange",
                "Header": HEADER,
                "UseServerRulesLoader": True,
            },
        )
        collection = result.get("InboxRuleCollection")
        if not isinstance(collection, dict) or "InboxRules" not in collection:
            raise Stopped("Outlook rule listing is incomplete")
        rules = collection["InboxRules"]
        if rules is None:
            return []
        if not isinstance(rules, list) or not all(
            isinstance(rule, dict) for rule in rules
        ):
            raise Stopped("Unexpected Outlook rule listing")
        return rules

    def pending_write(self, target):
        return bool(self.journal and self.journal.pending(self.email, target))

    def conflicting_pending(self, target):
        return bool(self.journal and self.journal.conflicts(self.email, target))

    def clear_pending(self, target):
        if self.journal:
            self.journal.clear(self.email, target)

    def create_rule(self, target):
        if self.journal:
            self.journal.begin(self.email, target)
        try:
            self._create_rule(target)
        except WriteUncertain:
            raise
        except (Stopped, AuthenticationExpired):
            # Definitive rejection or cancellation before sending: safe to retry later.
            self.clear_pending(target)
            raise

    def _create_rule(self, target):
        self.call(
            "NewInboxRule",
            {
                "__type": "NewInboxRuleRequest:#Exchange",
                "Header": HEADER,
                "InboxRule": {
                    "Name": rule_name(target),
                    "Enabled": True,
                    "Priority": 0,
                    "StopProcessingRules": False,
                    "ForwardTo": [
                        {
                            "DisplayName": target,
                            "SmtpAddress": target,
                            "RoutingType": "SMTP",
                        }
                    ],
                },
                "Force": False,
                "AlwaysDeleteOutlookRulesBlob": False,
            },
        )


def rule_name(target):
    return "HeadlessForward-" + target.lower()


def matches(rule, target):
    addresses = [
        x.get("SmtpAddress", "").casefold() for x in rule.get("ForwardTo") or []
    ]
    metadata = {
        "__type",
        "Description",
        "DescriptionTimeFormat",
        "DescriptionTimeZone",
        "Enabled",
        "InError",
        "ErrorType",
        "SupportedByTask",
        "WarningMessages",
        "RuleProvider",
        "Identity",
        "Name",
        "Priority",
        "ForwardTo",
    }
    inactive = (
        None,
        False,
        0,
        "",
        "NullInboxRuleMessageFlag",
        "NullInboxRuleMessageType",
        "NullImportance",
        "NullSensitivity",
    )
    # Every condition, exception and other action must be inactive. Unknown active
    # properties fail closed, including later manual edits to our named rule.
    extra = any(
        value not in inactive and value != [] and value != {}
        for key, value in rule.items()
        if key not in metadata
    )
    return (
        rule.get("Name") == rule_name(target)
        and rule.get("Enabled") is True
        and not rule.get("InError")
        and not extra
        and addresses == [target.casefold()]
    )


def configure(mailbox, target, apply=False):
    if getattr(mailbox, "conflicting_pending", lambda _: False)(target):
        return {"status": "unresolved_write_to_other_destination", "verified": False}
    # Read mailbox configuration first to detect existing global forwarding.
    before = mailbox.read()
    existing = (before.get("AddressString") or "").removeprefix("smtp:")
    if existing:
        return {
            "status": "existing_mailbox_forwarding_review_required",
            "verified": False,
        }
    rules = mailbox.rules()
    owned = [r for r in rules if r.get("Name", "").startswith("HeadlessForward-")]
    others = [r for r in rules if r not in owned]
    if any(
        r.get(k)
        for r in others
        for k in ("ForwardTo", "RedirectTo", "ForwardAsAttachmentTo")
    ):
        return {"status": "existing_forwarding_rule_review_required", "verified": False}
    if owned:
        if len(owned) == 1 and matches(owned[0], target):
            getattr(mailbox, "clear_pending", lambda _: None)(target)
            return {"status": "already_configured", "verified": True}
        return {"status": "conflict_existing_managed_rule", "verified": False}
    if getattr(mailbox, "pending_write", lambda _: False)(target):
        return {"status": "unresolved_previous_write", "verified": False}
    if not apply:
        return {"status": "would_configure", "verified": False}
    uncertain = False
    try:
        mailbox.create_rule(target)
    except WriteUncertain:
        uncertain = True
    for delay in (0, 5, 15):
        if delay:
            time.sleep(delay)
        found = [r for r in mailbox.rules() if r.get("Name") == rule_name(target)]
        if len(found) == 1 and matches(found[0], target):
            getattr(mailbox, "clear_pending", lambda _: None)(target)
            return {"status": "configured", "verified": True}
    return {
        "status": "write_outcome_unknown"
        if uncertain
        else "save_accepted_but_rule_not_verified",
        "verified": False,
    }


def accounts(path):
    content = read_text_input(path)
    result, seen = [], set()
    for number, line in enumerate(content.splitlines(), 1):
        if not line.strip():
            continue
        email, sep, password = line.partition(":")
        email = email.strip()
        if (
            not sep
            or not password
            or not re.fullmatch(r"[^\s@:]+@[^\s@:]+\.[^\s@:]+", email)
        ):
            raise Stopped("Invalid account format on line " + str(number))
        if email.casefold() in seen:
            raise Stopped("Duplicate account on line " + str(number))
        seen.add(email.casefold())
        result.append((email, password))
    return result


def run_account(number, credentials, target, apply, transport=None, journal=None):
    """Each worker owns its cookie jar and token; no session state is shared."""
    email, password = credentials
    try:
        transport = transport or Transport()
        for refresh in range(2):
            try:
                mailbox = Mailbox(
                    email, authenticate(email, password, transport), transport, journal
                )
                result = configure(mailbox, target, apply)
                break
            except AuthenticationExpired:
                if refresh:
                    raise
                transport.report(
                    {"event": "refresh_authorization", "account_number": number}
                )
    except Cancelled:
        result = {"status": "cancelled", "verified": False}
    except Throttled as exc:
        result = {"status": "throttled", "reason": str(exc), "verified": False}
    except Stopped as exc:
        result = {"status": "stopped", "reason": str(exc), "verified": False}
    except Exception as exc:
        # Exception bodies/URLs can contain credentials or authentication tokens.
        result = {"status": "error", "reason": type(exc).__name__, "verified": False}
    return {"account_number": number, "target": target, **result}


def run_accounts(
    selected, start_at, target, apply, workers, transport=None, journal=None
):
    """Bounded workers sharing one retry policy, cancellation event, and journal."""
    transport = transport or Transport()
    rows = iter(enumerate(selected, start_at))
    executor = ThreadPoolExecutor(max_workers=workers)
    pending = set()
    throttled = False
    try:

        def submit_next():
            row = next(rows, None)
            if row is not None:
                pending.add(
                    executor.submit(
                        run_account, row[0], row[1], target, apply, transport, journal
                    )
                )

        for _ in range(min(workers, len(selected))):
            submit_next()
        while pending:
            done, pending = wait(pending, return_when=FIRST_COMPLETED)
            results = [future.result() for future in done]
            throttled |= any(result["status"] == "throttled" for result in results)
            yield from results
            if not throttled:
                for _ in done:
                    submit_next()
        if throttled:
            for number, _ in rows:
                yield {
                    "account_number": number,
                    "target": target,
                    "status": "not_started_after_throttle",
                    "verified": False,
                }
    finally:
        transport.cancel()
        # In-flight writes cannot be cancelled safely. Drain them before exit.
        executor.shutdown(wait=True, cancel_futures=True)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--accounts",
        type=Path,
        required=True,
        help="email:password text file (any filename), or - for stdin",
    )
    p.add_argument("--to", required=True, help="Forwarding destination")
    p.add_argument("--start-at", type=int, default=1, help="One-based account number")
    p.add_argument("--limit", type=int, help="Maximum accounts (default: all)")
    p.add_argument(
        "--workers", type=int, default=3, help="Concurrent accounts, 1–10 (default: 3)"
    )
    p.add_argument(
        "--retries",
        type=int,
        default=3,
        help="Retries per request after initial attempt (default: 3)",
    )
    p.add_argument(
        "--cooldown",
        type=int,
        default=120,
        help="Minimum rate-limit delay in seconds (default: 120)",
    )
    p.add_argument(
        "--max-wait",
        type=int,
        default=1800,
        help="Maximum individual retry delay in seconds (default: 1800)",
    )
    p.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="Socket timeout in seconds (default: 30)",
    )
    p.add_argument(
        "--state",
        type=Path,
        default=Path(__file__).resolve().parent / "private/state.sqlite",
        help="Durable write-intent journal and run lock",
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument(
        "--apply",
        action="store_true",
        help="Create forwarding rules and verify readback",
    )
    mode.add_argument(
        "--check", action="store_true", help="Read only; exit nonzero unless configured"
    )
    args = p.parse_args()
    if (
        args.start_at < 1
        or (args.limit is not None and args.limit < 1)
        or not 1 <= args.workers <= 10
        or not 0 <= args.retries <= 10
        or args.cooldown <= 0
        or args.max_wait < args.cooldown
        or args.timeout <= 0
        or not re.fullmatch(r"[^\s@:]+@[^\s@:]+\.[^\s@:]+", args.to)
    ):
        p.error("Invalid address, account range, or worker count")
    try:
        rows = accounts(args.accounts)
    except Stopped as exc:
        p.error(str(exc))
    except Exception as exc:
        p.error("Unable to read account file (" + type(exc).__name__ + ")")
    end = args.start_at - 1 + args.limit if args.limit is not None else None
    selected = rows[args.start_at - 1 : end]
    if not selected:
        p.error("No accounts selected")
    failed = False
    transport = Transport(args.retries, args.cooldown, args.max_wait, args.timeout)
    try:
        with Journal(args.state) as journal:
            for result in run_accounts(
                selected,
                args.start_at,
                args.to,
                args.apply,
                args.workers,
                transport,
                journal,
            ):
                print(json.dumps(result), flush=True)
                failed |= not result["verified"] and (
                    args.apply or args.check or result["status"] != "would_configure"
                )
    except Stopped as exc:
        print(
            json.dumps({"status": "stopped", "reason": str(exc), "verified": False}),
            flush=True,
        )
        return 1
    except Exception as exc:
        print(
            json.dumps({"status": "local_error", "reason": type(exc).__name__}),
            flush=True,
        )
        return 1
    return int(failed)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
