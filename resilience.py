"""Bounded retries and a process-wide cooldown; never replay uncertain writes."""

import http.client
import json
import math
import random
import ssl
import sys
import threading
import time
import urllib.error
from email.utils import parsedate_to_datetime


class Stopped(Exception):
    """A sanitized failure safe to include in logs."""


class Throttled(Stopped):
    """The shared rate-limit retry budget is exhausted."""


class Cancelled(Stopped):
    pass


class WriteUncertain(Stopped):
    """A mutation may have succeeded. Read back; never blindly repeat it."""


class AuthenticationExpired(Stopped):
    pass


def retry_after(headers, now=None):
    """Accept Retry-After seconds or HTTP dates; malformed values use backoff."""
    value = headers.get("Retry-After") if headers else None
    if not value:
        return 0.0
    try:
        seconds = float(value)
    except (ValueError, TypeError):
        try:
            stamp = parsedate_to_datetime(value)
            if stamp.tzinfo is None:
                return 0.0
            seconds = stamp.timestamp() - (time.time() if now is None else now)
        except (ValueError, TypeError, OverflowError):
            return 0.0
    return max(0.0, seconds) if math.isfinite(seconds) else 0.0


def log_event(event):
    print(json.dumps(event), file=sys.stderr, flush=True)


class Transport:
    def __init__(
        self,
        retries=3,
        cooldown=120,
        max_wait=1800,
        timeout=30,
        clock=time.monotonic,
        waiter=None,
        jitter=None,
        reporter=log_event,
    ):
        if retries < 0 or cooldown <= 0 or max_wait < cooldown or timeout <= 0:
            raise ValueError("Invalid retry policy")
        self.retries, self.cooldown, self.max_wait, self.timeout = (
            retries,
            cooldown,
            max_wait,
            timeout,
        )
        self.clock, self.jitter, self.report = (
            clock,
            jitter or (lambda: random.uniform(0, 5)),
            reporter,
        )
        self.cancelled = threading.Event()
        self.waiter = waiter or self.cancelled.wait
        self.lock = threading.Lock()
        self.until = 0.0
        self.next_slot = 0.0
        self.circuit_open = False

    def cancel(self):
        self.cancelled.set()

    def gate(self):
        """Recheck shared deadlines after waking; space recovery traffic 250ms."""
        while True:
            with self.lock:
                if self.cancelled.is_set():
                    raise Cancelled("Cancelled")
                if self.circuit_open:
                    raise Throttled("Rate-limit budget exhausted; resume later")
                now = self.clock()
                delay = max(self.until, self.next_slot) - now
                if delay <= 0:
                    if self.until:
                        self.next_slot = now + 0.25
                    return
            self.waiter(min(delay, 30))

    def pause(self, seconds, reason, attempt):
        if seconds > self.max_wait:
            with self.lock:
                self.circuit_open = True
            raise Throttled("Server wait exceeds configured maximum; resume later")
        with self.lock:
            self.until = max(self.until, self.clock() + seconds)
        self.report(
            {
                "event": "retry_wait",
                "reason": reason,
                "attempt": attempt,
                "wait_seconds": round(seconds, 1),
            }
        )

    def fetch(
        self, opener, request, *, safe=False, mutation=False, capture_redirect=False
    ):
        """Retry explicit 429s; retry transient failures only for read operations.

        HTTP 429 means the request was rejected. Other mutation failures, even
        after receiving response headers, have unknown outcome and are not retried.
        """
        for attempt in range(self.retries + 1):
            self.gate()
            retry_seconds = 0.0
            reason = None
            try:
                with opener.open(request, timeout=self.timeout) as response:
                    return response.read(), response.url
            except urllib.error.HTTPError as exc:
                code = exc.code
                retry_seconds = retry_after(exc.headers)
                url = exc.url
                exc.close()
                if code == 302 and capture_redirect:
                    return b"", url
                if code == 429:
                    if attempt == self.retries:
                        with self.lock:
                            self.circuit_open = True
                        raise Throttled(
                            "Rate limit persisted after bounded retries"
                        ) from None
                    reason = "http_429"
                    retry_seconds = (
                        max(retry_seconds, min(self.cooldown * 2**attempt, 900))
                        + self.jitter()
                    )
                elif mutation:
                    if code == 401:
                        raise AuthenticationExpired(
                            "Mailbox authorization expired"
                        ) from None
                    if code in (400, 403, 404):
                        raise Stopped(
                            "Outlook rejected request; HTTP " + str(code)
                        ) from None
                    raise WriteUncertain(
                        "Write response uncertain; readback required"
                    ) from None
                elif code == 401:
                    raise AuthenticationExpired(
                        "Mailbox authorization expired"
                    ) from None
                elif (
                    safe
                    and code in (408, 500, 502, 503, 504)
                    and attempt < self.retries
                ):
                    reason = "http_" + str(code)
                    retry_seconds = (
                        max(retry_seconds, min(5 * 2**attempt, 60)) + self.jitter()
                    )
                else:
                    raise Stopped("Microsoft HTTP " + str(code)) from None
            except (urllib.error.URLError, OSError, http.client.HTTPException) as exc:
                if mutation:
                    raise WriteUncertain(
                        "Write connection failed; readback required"
                    ) from None
                cause = getattr(exc, "reason", exc)
                if (
                    isinstance(cause, ssl.SSLError)
                    or not safe
                    or attempt == self.retries
                ):
                    raise Stopped(
                        "Connection failed (" + type(cause).__name__ + ")"
                    ) from None
                reason = "read_connection_failure"
                retry_seconds = min(5 * 2**attempt, 60) + self.jitter()
            self.pause(retry_seconds, reason, attempt + 1)
