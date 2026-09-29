# Forward future Hotmail / Outlook mail to Gmail

This script sets an enabled Inbox rule in each personal Microsoft mailbox using Microsoft Graph. Microsoft runs the rule after setup; you do not need to keep this script running. Existing messages are not forwarded. The rule does not delete the original or stop other rules.

Plain text with any filename or `--accounts -` for stdin is supported across platforms. UTF-8 and BOM-marked UTF-16/UTF-32 are accepted; RTF conversion is an optional macOS feature. Only the email before each colon is used. Passwords are never used, printed, or saved. Refresh tokens are saved in your explicitly supplied private token file (mode 0600); access tokens stay in memory. Keep that file outside shared folders and source control. No credentials are included in this download.

## One-time Microsoft setup

1. In the [Microsoft Entra admin center](https://entra.microsoft.com/), open **App registrations → New registration**. Choose a supported account type that includes **personal Microsoft accounts**.
2. Copy the **Application (client) ID**.
3. Under **Authentication**, enable **Allow public client flows** (device-code sign-in). No client secret is needed.
4. Under **API permissions**, add Microsoft Graph **delegated** permissions `User.Read` and `MailboxSettings.ReadWrite`.

You need a Microsoft Entra directory where you can register an app. Consumer Hotmail accounts must each sign in and grant consent; having their passwords does not enable unattended Graph setup. Microsoft's password grant does not support personal accounts.

## Install and preview

Open a terminal in the folder containing this README and run:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python forward_outlook.py --accounts private/accounts.txt --to destination@example.com
```

Preview only parses the list; it performs no network requests and changes nothing.

## One-time authorization (requires Microsoft sign-in)

```sh
python forward_outlook.py --accounts private/accounts.txt --to destination@example.com --client-id YOUR_APP_ID --token-file private/oauth-tokens.json --authorize
```

This saves OAuth refresh tokens and makes no forwarding changes. Each account must complete Microsoft sign-in once. You can authorize just one first with `--limit 1`, then continue with `--start-at 2`. Expired or revoked consent requires authorization again.

If you already have refresh tokens issued to your own public-client application with these permissions, provide a private JSON file in this shape (placeholder values shown):

```json
{"client_id":"YOUR_APP_ID","accounts":{"example@hotmail.com":"OAUTH_REFRESH_TOKEN"}}
```

Run `chmod 600 private/oauth-tokens.json` on an imported token file. These are OAuth tokens, not passwords, and must belong to the specified client ID.

## Headless HTTP: configure one mailbox first

```sh
python forward_outlook.py --accounts private/accounts.txt --to destination@example.com --client-id YOUR_APP_ID --token-file private/oauth-tokens.json --apply --limit 1
```

The `--apply` and `--check` modes use HTTP requests only, with no browser or input prompts. They exchange refresh tokens for access tokens, verify mailbox identity, and call Graph. Rotated refresh tokens are saved atomically. If a listed email is only an alias, use the primary mailbox address instead.

Send a new test email from another address to the Hotmail mailbox and confirm receipt in Gmail, including Spam. The script verifies the saved rule, but that does not prove message delivery. Inbox rules do not cover messages diverted to Junk before rule processing.

## Configure the remaining accounts

```sh
python forward_outlook.py --accounts private/accounts.txt --to destination@example.com --client-id YOUR_APP_ID --token-file private/oauth-tokens.json --apply --start-at 2
```

This run is unattended after authorization. On failure the script stops and prints the account number needed for `--start-at`. On rerun it reads current rules and reuses an enabled equivalent rule. Interrupted writes are not automatically retried: rerunning reconciles current state first. Run one instance at a time to avoid concurrent rule creation.

The new rule is placed first so later rules that stop processing do not block it. Existing rule contents remain intact; their sequence numbers may shift. Existing conditional forwarding rules may still cause additional forwards.

Use `--check` instead of `--apply` for headless verification without changing forwarding rules. OAuth token rotation can still update the local token file. There is no need for a recurring job: forwarding continues server-side after setup.

## Disable or change destination

In each mailbox, open Outlook settings → Mail → Rules. Disable/delete the rule named `GraphForwarder-…`. For an existing equivalent rule, inspect its forwarding destination to locate it. To change Gmail destinations, remove the old forwarding rule before running for the new destination; the script never deletes old rules automatically.

## References

- [Graph: create an Inbox rule and supported personal-account permission](https://learn.microsoft.com/en-us/graph/api/mailfolder-post-messagerules?view=graph-rest-1.0)
- [Microsoft device authorization flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code)
- [Password grant limitations for personal accounts](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth-ropc)

Local verification covers parsing and simulated rule behavior. Live Microsoft sign-in and Gmail delivery must be verified when you run setup.
