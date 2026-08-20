#!/usr/bin/env python3
"""Deterministic acceptance checks for a chapter candidate."""

import argparse
import hashlib
import re
import sys
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("file", type=Path)
    parser.add_argument("--title")
    parser.add_argument("--opening-line")
    parser.add_argument("--marker")
    parser.add_argument("--marker-max-ratio", type=float)
    parser.add_argument("--min-chars", type=int)
    parser.add_argument("--max-chars", type=int)
    parser.add_argument("--forbid", action="append", default=[])
    return parser.parse_args()


def main():
    args = parse_args()
    data = args.file.read_bytes()
    text = data.decode("utf-8")
    nonspace = len(re.sub(r"\s+", "", text))
    digest = hashlib.sha256(data).hexdigest()
    failures = []

    if args.title and not text.startswith(f"# {args.title}\n"):
        failures.append(f"title mismatch: expected '# {args.title}'")

    if args.opening_line:
        lines = text.splitlines()
        actual = lines[2] if len(lines) > 2 else ""
        expected = f"> {args.opening_line}"
        if actual != expected:
            failures.append("fixed opening line mismatch")

    if args.min_chars is not None and nonspace < args.min_chars:
        failures.append(f"nonspace chars {nonspace} < {args.min_chars}")
    if args.max_chars is not None and nonspace > args.max_chars:
        failures.append(f"nonspace chars {nonspace} > {args.max_chars}")

    if args.marker:
        marker_index = text.find(args.marker)
        if marker_index < 0:
            failures.append(f"marker not found: {args.marker}")
        else:
            marker_ratio = len(re.sub(r"\s+", "", text[:marker_index])) / max(nonspace, 1)
            print(f"MARKER_RATIO={marker_ratio:.4f}")
            if args.marker_max_ratio is not None and marker_ratio > args.marker_max_ratio:
                failures.append(
                    f"marker ratio {marker_ratio:.4f} > {args.marker_max_ratio:.4f}"
                )

    for phrase in args.forbid:
        if phrase in text:
            failures.append(f"forbidden phrase found: {phrase}")

    placeholders = re.findall(r"\[(?:TODO|TBD)\]|【(?:待补|待写|TODO|TBD)】", text, re.I)
    if placeholders:
        failures.append(f"unresolved placeholders: {len(placeholders)}")

    print(f"FILE={args.file}")
    print(f"SHA256={digest}")
    print(f"NONSPACE_CHARS={nonspace}")
    if failures:
        for failure in failures:
            print(f"FAIL={failure}")
        return 1
    print("VERDICT=PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
