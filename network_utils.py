"""
network_utils.py
Small retry helper for the app's Google Sheets/Drive/translation calls.
Older or restricted networks (the kind this app tends to run on) often hit
transient DNS/connection hiccups - e.g. "unable to find the server at
oauth2.googleapis.com" - that succeed a second later. Retrying a couple of
times with a short backoff turns those into a non-event instead of a hard
failure, while a genuinely offline machine still fails with a clear message.
"""

import time


def with_retry(func, *args, attempts=3, base_delay=1.5, **kwargs):
    """Calls func(*args, **kwargs), retrying on any exception up to
    `attempts` times with an increasing delay between tries. Raises a
    RuntimeError with a clear, actionable message if every attempt fails."""
    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            last_exc = e
            if attempt < attempts:
                time.sleep(base_delay * attempt)
    raise RuntimeError(
        f"Network error after {attempts} attempts: {last_exc}. "
        "Check the internet connection on this computer, any firewall/antivirus "
        "that might be blocking it, and that the system date/time is correct "
        "(a wrong clock breaks Google's secure login)."
    )
