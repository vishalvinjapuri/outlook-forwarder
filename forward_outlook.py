#!/usr/bin/env python3
"""Create future-mail forwarding rules for personal Microsoft accounts."""

import argparse
import hashlib
import json
import os
import re
import stat
import sys
import tempfile
import time
from pathlib import Path
from urllib.parse import quote

from input_files import read_text_input

BASE = "https://graph.microsoft.com/v1.0"
RULES = BASE + "/me/mailFolders/inbox/messageRules"
SCOPES = ["User.Read", "MailboxSettings.ReadWrite"]
EMAIL = re.compile(r"[^\s@:]+@[^\s@:]+\.[^\s@:]+")


def read_accounts(path):
    content = read_text_input(path)
    accounts = []
    for number, line in enumerate(content.splitlines(), 1):
        if not line.strip():
            continue
        # Accept email:password input, but never use or store the password.
        address = line.split(":", 1)[0].strip().lower()
        if not EMAIL.fullmatch(address):
            raise ValueError(f"Invalid email on line {number}; line content withheld.")
        if address not in accounts:
            accounts.append(address)
    if not accounts:
        raise ValueError("No accounts found.")
    return accounts


def graph(session, method, url, **kwargs):
    if not (url.startswith(BASE + "/")):
        raise RuntimeError("Unexpected Graph pagination URL.")
    # Only reads are retried. An interrupted POST may already have succeeded.
    for attempt in range(5):
        try:
            response = session.request(
                method, url, timeout=40, allow_redirects=False, **kwargs
            )
        except Exception:
            raise RuntimeError(
                "Network failure; rerun this account to reconcile its rule."
            ) from None
        if (
            method == "GET"
            and response.status_code in (429, 500, 502, 503, 504)
            and attempt < 4
        ):
            delay = response.headers.get("Retry-After", "")
            time.sleep(min(int(delay) if delay.isdigit() else 2**attempt, 60))
            continue
        if not 200 <= response.status_code < 300:
            raise RuntimeError(
                f"Graph HTTP {response.status_code}; request-id="
                f"{response.headers.get('request-id', 'unavailable')}. "
                "Rerun to reconcile before assuming a write failed."
            )
        return response.json() if response.content else {}


def list_rules(session):
    rules, url = [], RULES
    while url:
        page = graph(session, "GET", url)
        rules.extend(page.get("value", []))
        url = page.get("@odata.nextLink")
    return rules


def has_predicates(value):
    return any(bool(v) for v in (value or {}).values())


def matches(rule, destination):
    actions = rule.get("actions") or {}
    recipients = actions.get("forwardTo") or []
    return (
        rule.get("isEnabled") is True
        and not rule.get("hasError", False)
        and not has_predicates(rule.get("conditions"))
        and not has_predicates(rule.get("exceptions"))
        and len(recipients) == 1
        and recipients[0].get("emailAddress", {}).get("address", "").lower()
        == destination
        and not any(bool(v) for k, v in actions.items() if k != "forwardTo")
    )


def ensure_rule(session, destination, check_only=False):
    rules = list_rules(session)
    name = "GraphForwarder-" + hashlib.sha256(destination.encode()).hexdigest()[:16]
    owned = [r for r in rules if r.get("displayName") == name]
    if owned:
        if len(owned) == 1 and matches(owned[0], destination):
            return "verified_existing"
        raise RuntimeError(
            "A rule with this script's name differs from the expected rule. "
            "Inspect it in Outlook before continuing."
        )
    if any(matches(r, destination) for r in rules):
        return "verified_equivalent_existing"
    if check_only:
        return "missing_rule"
    # First in the Inbox rule sequence so later stop-processing rules do not block it.
    # Existing rule contents are preserved; Outlook may shift their sequence numbers.
    body = {
        "displayName": name,
        "sequence": 1,
        "isEnabled": True,
        "conditions": {},
        "exceptions": {},
        "actions": {
            "forwardTo": [{"emailAddress": {"address": destination}}],
            "stopProcessingRules": False,
        },
    }
    created = graph(session, "POST", RULES, json=body)
    saved = graph(session, "GET", RULES + "/" + quote(created["id"], safe=""))
    if not matches(saved, destination):
        raise RuntimeError(
            "Rule write returned, but readback did not match. Inspect Outlook."
        )
    return "created_and_verified"


def save_tokens(path, data):
    path = Path(path).expanduser().resolve()
    fd, temp = tempfile.mkstemp(prefix=".oauth-", dir=path.parent)
    try:
        with os.fdopen(fd, "w") as handle:
            json.dump(data, handle)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)


def load_tokens(path, client_id):
    path = Path(path).expanduser()
    if not path.exists():
        return {"client_id": client_id, "accounts": {}}
    if stat.S_IMODE(path.stat().st_mode) & 0o077:
        raise ValueError("Token file must be private. Run chmod 600 on it.")
    data = json.loads(path.read_text())
    if data.get("client_id") != client_id or not isinstance(data.get("accounts"), dict):
        raise ValueError("Token file format or client ID mismatch.")
    return data


def authenticate(client_id, expected, token_file, authorize=False):
    import requests

    tokens = load_tokens(token_file, client_id)
    if authorize:
        import msal

        app = msal.PublicClientApplication(
            client_id, authority="https://login.microsoftonline.com/consumers"
        )
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(
                "Cannot start device sign-in. Check app registration/public client settings."
            )
        print(
            f"Sign in as {expected}. Use a private browser window to avoid the wrong account."
        )
        print(flow["message"], flush=True)
        result = app.acquire_token_by_device_flow(flow)
    else:
        refresh_token = tokens["accounts"].get(expected)
        if not refresh_token:
            raise RuntimeError(
                "No refresh token for this account. Run --authorize first."
            )
        try:
            response = requests.post(
                "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
                data={
                    "client_id": client_id,
                    "grant_type": "refresh_token",
                    "refresh_token": refresh_token,
                    "scope": "https://graph.microsoft.com/User.Read "
                    "https://graph.microsoft.com/MailboxSettings.ReadWrite offline_access",
                },
                timeout=40,
                allow_redirects=False,
            )
            result = response.json()
        except Exception:
            raise RuntimeError("Token request failed; no rule was changed.") from None
    if "access_token" not in result:
        raise RuntimeError(
            "Authorization failed or consent expired. Run --authorize again. "
            "No rule was changed."
        )
    session = requests.Session()
    session.headers.update({"Authorization": "Bearer " + result["access_token"]})
    try:
        profile = graph(session, "GET", BASE + "/me?$select=mail,userPrincipalName")
        identities = {
            str(profile.get(k) or "").lower() for k in ("mail", "userPrincipalName")
        }
        if expected not in identities:
            raise RuntimeError(
                "Signed-in mailbox does not match the requested address. "
                "Use its primary address if your list contains an alias."
            )
        if result.get("refresh_token"):
            tokens["accounts"][expected] = result["refresh_token"]
            save_tokens(token_file, tokens)
        elif authorize:
            raise RuntimeError(
                "No refresh token returned; offline authorization is required."
            )
        return session
    except Exception:
        session.close()
        raise


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--accounts", required=True, help="RTF on macOS, or one email per text line"
    )
    parser.add_argument("--to", required=True, help="Destination Gmail address")
    parser.add_argument(
        "--client-id", help="Your Microsoft app registration's client ID"
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--apply",
        action="store_true",
        help="Headless HTTP: create rules using saved OAuth tokens",
    )
    mode.add_argument(
        "--check",
        action="store_true",
        help="Headless HTTP: inspect rules without changing them",
    )
    mode.add_argument(
        "--authorize",
        action="store_true",
        help="One-time device sign-in; save tokens without changing rules",
    )
    parser.add_argument(
        "--token-file",
        help="Private JSON refresh-token file (required outside preview)",
    )
    parser.add_argument(
        "--start-at", type=int, default=1, help="Resume at this 1-based account number"
    )
    parser.add_argument("--limit", type=int, help="Process only this many accounts")
    args = parser.parse_args()
    destination = args.to.strip().lower()
    if not EMAIL.fullmatch(destination) or destination.split("@")[1] not in (
        "gmail.com",
        "googlemail.com",
    ):
        parser.error("--to must be a Gmail address")
    accounts = read_accounts(args.accounts)
    if not 1 <= args.start_at <= len(accounts) or (
        args.limit is not None and args.limit < 1
    ):
        parser.error("Invalid --start-at or --limit")
    selected = list(enumerate(accounts, 1))[args.start_at - 1 :]
    if args.limit:
        selected = selected[: args.limit]
    if not (args.apply or args.check or args.authorize):
        print(f"PREVIEW: {len(selected)} of {len(accounts)} accounts -> {destination}")
        print("Future Inbox mail only. No sign-ins or changes performed.")
        print(
            "Run --authorize once, then --apply for headless setup or --check to inspect."
        )
        return 0
    if not args.client_id or not args.token_file:
        parser.error("--client-id and --token-file are required outside preview")
    for number, address in selected:
        print(f"\n[{number}/{len(accounts)}] {address} -> {destination}", flush=True)
        try:
            with authenticate(
                args.client_id, address, args.token_file, args.authorize
            ) as session:
                status = (
                    "authorized"
                    if args.authorize
                    else ensure_rule(session, destination, check_only=args.check)
                )
                print(
                    json.dumps({"account_number": number, "status": status}), flush=True
                )
        except Exception as exc:
            print(f"Stopped at account {number}: {type(exc).__name__}", file=sys.stderr)
            print(
                f"Resume with --start-at {number}. Existing matching rules are reused.",
                file=sys.stderr,
            )
            return 1
    print("Finished. Rule verification confirms configuration, not Gmail delivery.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit("Interrupted. Rerun using --start-at for the current account.")
    except (ValueError, OSError) as exc:
        sys.exit(str(exc))
