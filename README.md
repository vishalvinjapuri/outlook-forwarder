# Outlook Forwarder

Configure future-mail forwarding from personal Outlook.com / Hotmail accounts to a destination mailbox using headless HTTP requests. Includes a Gmail filter generator and an optional Microsoft Graph implementation.

The main script uses Python's standard library. No browser automation, app registration, or package installation is needed for the web API route.

## How it works

For each account, the script signs into Microsoft, obtains Outlook's consumer authorization token, checks current forwarding settings, creates an Inbox rule, and reads the rule back. Outlook runs the rule afterward; the script does not need to stay running.

- Original messages stay in Outlook. Other rules can continue processing.
- Existing matching rules are reused. Conflicting forwarding or edited managed rules require review.
- The new rule is placed first; Outlook may shift existing rule priorities.
- Passwords, cookies, and web API tokens remain in memory and are not logged or saved by the script.
- Each account has an isolated session. Three accounts run concurrently by default.

**This uses an unofficial private Outlook web API.** Microsoft can change it without notice. MFA, CAPTCHA, account recovery, consent, and unsupported sign-in pages stop that account; the script does not bypass them. Use accounts you own or administer.

## Requirements

Python 3.10 or newer on Windows, Linux, or macOS. Plain-text files and piped input need no platform-specific tools. The optional Graph script has separate dependencies and authorization requirements.

## Quick start

Create a private account file, one `email:password` pair per line:

```text
first@example.com:replace-with-password
second@example.com:another-password
```

Use actual personal Microsoft account addresses in your file. Passwords can contain colons and spaces; password whitespace is preserved. Blank lines are ignored; duplicate accounts are rejected. Files can have any name or extension (`.txt`, `.list`, or no extension). UTF-8, including a BOM, and BOM-marked UTF-16/UTF-32 are supported, with Windows or Unix line endings. Keep inputs in the ignored `private/` directory:

```sh
mkdir -p private
# Create private/accounts.txt in your editor, then restrict access on macOS/Linux:
chmod 600 private/accounts.txt
```

Preview every account in the input file. This signs in and reads settings but does not create rules:

```sh
python3 forward_outlook_web.py \
  --accounts private/accounts.txt --to destination@example.com
```

Create and verify forwarding for every account:

```sh
python3 forward_outlook_web.py \
  --accounts private/accounts.txt --to destination@example.com --apply
```

All input accounts are selected by default. Use `--start-at` and `--limit` to select a range, and `--workers` to set concurrency.

You can also pipe plain text from a file or another program:

```sh
cat private/accounts.txt | python3 forward_outlook_web.py \
  --accounts - --to destination@example.com --apply
```

On Windows, use `python` in place of `python3` if that is your installed command. For example, PowerShell can pipe `Get-Content private/accounts.txt` into the same `--accounts -` invocation. Plain text is the recommended input format. Legacy RTF files are additionally supported on macOS when `textutil` is available; other platforms should export RTF as plain text first.

Verify without changing forwarding:

```sh
python3 forward_outlook_web.py \
  --accounts private/accounts.txt --to destination@example.com \
  --check --workers 3
```

## Speed and recovery

`--workers` accepts 1–10; the default is 3. Independent network waits overlap while operations within each account stay sequential. Use `--workers 1` for the original serial behavior. Actual speed depends on Microsoft latency and throttling; no fixed production speedup is promised.

At most the selected number of workers is in flight. All workers share one cooldown: a rate limit pauses new top-level HTTP requests across the batch, including requests for other accounts. Requests already in flight can finish. Recovery traffic is spaced by at least 250 ms to avoid releasing every worker at once.

### Retry policy

| Failure | Behavior |
| --- | --- |
| HTTP 429 | Wait at least 2 minutes, then 4, then 8, with 0–5 seconds of jitter. Honor a longer `Retry-After` value (seconds or HTTP date). |
| Read timeout, dropped connection, HTTP 408/500/502/503/504 | Retry reads with 5/10/20-second backoff plus jitter; honor longer server delays. |
| Expired mailbox authorization (401) | Obtain fresh authorization once, then reconcile current rules. |
| Invalid credentials, MFA, CAPTCHA, consent, account recovery, HTTP 403 | Stop that account; no challenge bypass or repeated password attempts. |
| Write timeout, connection loss, unexpected response, or server error | Do not replay the mutation. Read back the rule; preserve an unresolved write record if the outcome cannot be confirmed. |

The default allows three retries after the initial request. An explicit HTTP 429 is a rejected request and can be retried after cooldown, including during rule creation. Other uncertain writes are never automatically replayed. If rate limits persist through the retry budget, the shared circuit stops further requests and unscheduled accounts are reported as `not_started_after_throttle`.

A requested delay longer than `--max-wait` (30 minutes by default) stops the batch instead of retrying earlier than Microsoft asked. Permanent failures on one account do not otherwise stop the batch.

```sh
python3 forward_outlook_web.py \
  --accounts private/accounts.txt --to destination@example.com \
  --apply --workers 3 \
  --retries 3 --cooldown 120 --max-wait 1800 --timeout 30 \
  > private/results.jsonl 2> private/events.jsonl
```

`--timeout` is the socket timeout, not a whole-run deadline. Cooldowns and verification increase total run time. JSON account results go to stdout; retry/cooldown and authorization-refresh events go to stderr without credentials, tokens, request URLs, or response bodies.

Output is JSON Lines, one result per account, in completion order. `account_number` is the one-based position among nonblank input lines. A zero exit code means all selected accounts meet the requested mode: preview can report `would_configure`, while `--check` and `--apply` require verified rules. Account/runtime failures use code 1; invalid CLI/input uses code 2; Ctrl+C uses code 130.

### Crash-safe write journal

The web CLI uses `private/state.sqlite` beside the script by default. Override it with `--state /path/to/state.sqlite` and keep using that same file for subsequent runs. The SQLite journal contains hashed account/destination identifiers, no passwords or tokens. It records intent durably **before** creating a rule, then clears that record only after matching rule readback or a definitive rejection.

After a crash or uncertain write, rerunning first reads Outlook. If the rule is present and matches, the journal entry is cleared. If absent, it reports `unresolved_previous_write` and does not create a second rule. Changing destinations does not sidestep an unresolved write. For such a result, inspect Outlook's rules and rerun `--check` after allowing time for propagation. Do not delete the state file to bypass this protection. If manual recovery is needed, first establish the mailbox state before clearing a specific journal entry.

A separate SQLite run lock prevents overlapping CLI instances using the same state file. The lock releases automatically on process exit, including a crash. This is not a distributed lock: use one host and the same local state path per account list. Separate state files or machines can still race. Avoid network filesystems for the journal.

After rate-limit exhaustion, wait and rerun the affected range. Ctrl+C interrupts cooldown waits, stops scheduling new work, and drains in-flight socket operations (subject to their timeout). A request already sent cannot be undone. Local state can be updated during `--check` to reconcile earlier writes, but mailbox settings are not changed.

## Spam and delivery limits

**Inbox rules do not forward messages delivered directly to Outlook Junk.** Existing messages are not forwarded. Add trusted senders to Outlook's Safe Senders list if appropriate, then send a fresh test message.

A verified rule proves the saved configuration, not delivery to the destination. Check the destination Inbox and Spam folder using a new test email.

To generate Gmail “Never send it to Spam” filters:

```sh
python3 gmail_filters.py \
  --accounts private/accounts.txt --output private/gmail-filters.xml
```

This utility also accepts a plain address-only list. It creates groups of up to ten From addresses, never writes passwords, and makes no network requests. In Gmail, use **Settings → See all settings → Filters and Blocked Addresses → Import filters**, upload the XML, select the filters, and create them. Mark existing wanted spam as **Not spam** separately.

Check the forwarded message's actual From address: if it preserves the original sender, filters matching the forwarding accounts will not match it.

## Remove forwarding

In each account, open **Outlook Settings → Mail → Rules** and disable or delete `HeadlessForward-<destination>`. The script does not automatically delete rules or change destinations.

## Optional Microsoft Graph route

`forward_outlook.py` uses Microsoft's documented Graph API. It requires an app registration and one-time delegated authorization for each account; subsequent runs use refresh tokens. See [GRAPH.md](GRAPH.md). This route is separate from the concurrent web API script and has not been live-validated as part of the web API rollout.

## Development

Run the offline tests without dependencies or account credentials:

```sh
python3 -m unittest discover -s tests -v
```

Tests cover concurrency, shared cooldowns, Retry-After dates, retry exhaustion, cancellation, authorization refresh, crash recovery, run locking, unknown write outcomes, conflicting rules, credential-safe errors, and Gmail XML generation. Test clocks simulate minutes of backoff without sleeping for minutes. No tests sign in or send mail. CI runs the same suite on Python 3.10 and 3.14 on Linux and Windows.

The resilience policies are fault-tested offline; they are not evidence of a production SLA or proof that Microsoft's private API will remain compatible. The optional Graph script does not use this web CLI's retry controller or write journal.

`private/` is excluded from Git. Keep credentials, generated filters, and run reports there. No actual account lists or credentials belong in commits.

## References

- [Microsoft: throttling and Retry-After guidance](https://learn.microsoft.com/en-us/graph/throttling) (general retry guidance; the web API is separate)
- [Microsoft: Outlook rules](https://support.microsoft.com/en-us/outlook/mail/manage-email-messages-by-using-rules-in-outlook)
- [Microsoft: safe senders](https://support.microsoft.com/en-us/outlook/safe-senders-in-outlook-com)
- [Google: Gmail filter import/export](https://support.google.com/mail/answer/6579)
- [outlook-cli](https://github.com/adyavanapalli/outlook-cli) and [Outlook web modules](https://github.com/ladifire-opensource/outlook.live.com_modules), useful background on the web API

The working web request shapes were checked against Outlook's own JavaScript and live rule readback. This project is not affiliated with Microsoft or Google.
