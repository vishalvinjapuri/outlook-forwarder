#!/usr/bin/env python3
"""Generate Gmail Never-spam filters from an address or email:password list."""

import argparse
import re
import xml.etree.ElementTree as ET
from pathlib import Path

from input_files import read_text_input

ATOM = "http://www.w3.org/2005/Atom"
APPS = "http://schemas.google.com/apps/2006"


def read_addresses(path):
    content = read_text_input(path)
    addresses = []
    seen = set()
    for number, line in enumerate(content.splitlines(), 1):
        if not line.strip():
            continue
        address = line.strip().partition(":")[0]
        if not re.fullmatch(r"[^\s@:]+@[^\s@:]+\.[^\s@:]+", address):
            raise ValueError(f"Invalid address on line {number}")
        if address.casefold() not in seen:
            seen.add(address.casefold())
            addresses.append(address)
    if not addresses:
        raise ValueError("No addresses found")
    return addresses


def build_filters(addresses, group_size=10):
    if not 1 <= group_size <= 10:
        raise ValueError("Group size must be between 1 and 10")
    ET.register_namespace("", ATOM)
    ET.register_namespace("apps", APPS)
    feed = ET.Element(f"{{{ATOM}}}feed")
    ET.SubElement(feed, f"{{{ATOM}}}title").text = "Outlook forwarding filters"
    for offset in range(0, len(addresses), group_size):
        entry = ET.SubElement(feed, f"{{{ATOM}}}entry")
        ET.SubElement(entry, f"{{{ATOM}}}category", {"term": "filter"})
        ET.SubElement(entry, f"{{{ATOM}}}title").text = "Mail Filter"
        ET.SubElement(
            entry,
            f"{{{APPS}}}property",
            {
                "name": "from",
                "value": " OR ".join(addresses[offset : offset + group_size]),
            },
        )
        ET.SubElement(
            entry, f"{{{APPS}}}property", {"name": "shouldNeverSpam", "value": "true"}
        )
    ET.indent(feed)
    return ET.ElementTree(feed)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--accounts", type=Path, required=True)
    parser.add_argument(
        "--output", type=Path, required=True, help="Prefer a path inside private/"
    )
    args = parser.parse_args()
    try:
        addresses = read_addresses(args.accounts)
        build_filters(addresses).write(
            args.output.expanduser(), encoding="utf-8", xml_declaration=True
        )
    except Exception as exc:
        parser.error("Could not generate filters (" + type(exc).__name__ + ")")
    print(f"Generated filters for {len(addresses)} addresses. No passwords included.")


if __name__ == "__main__":
    main()
