"""Read portable account text files or stdin; RTF conversion is optional."""

import shutil
import subprocess
import sys
from pathlib import Path

from resilience import Stopped


def read_text_input(path):
    """Accept any filename, UTF-8/16/32 BOMs, or '-' for piped standard input."""
    if str(path) == "-":
        stream = getattr(sys.stdin, "buffer", sys.stdin)
        raw = stream.read()
        if isinstance(raw, str):
            return raw.removeprefix("\ufeff")
    else:
        raw = Path(path).expanduser().read_bytes()
    if raw.lstrip().startswith(b"{\\rtf"):
        converter = shutil.which("textutil")
        if converter is None:
            raise Stopped(
                "RTF needs macOS textutil; export as plain text on this platform"
            )
        result = subprocess.run(
            [converter, "-convert", "txt", "-stdin", "-stdout"],
            input=raw,
            capture_output=True,
            check=False,
        )
        if result.returncode:
            raise Stopped("Could not convert RTF; export as plain text")
        raw = result.stdout
    if raw.startswith((b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff")):
        encoding = "utf-32"
    elif raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        encoding = "utf-16"
    else:
        encoding = "utf-8-sig"
    try:
        return raw.decode(encoding)
    except UnicodeError:
        raise Stopped(
            "Unsupported text encoding; use UTF-8 or BOM-marked UTF-16/UTF-32"
        ) from None
