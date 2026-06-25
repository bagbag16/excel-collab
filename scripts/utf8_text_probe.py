#!/usr/bin/env python3
"""Read UTF-8 text and emit ASCII-safe JSON for Windows shell checks."""

from __future__ import annotations

import argparse
import codecs
import hashlib
import json
from pathlib import Path


def decode_escape(value: str | None) -> str | None:
    if value is None:
        return None
    return codecs.decode(value, "unicode_escape")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path)
    parser.add_argument("--encoding", default="utf-8-sig")
    parser.add_argument("--contains")
    parser.add_argument("--contains-escape")
    parser.add_argument("--context", type=int, default=0)
    parser.add_argument("--max-matches", type=int, default=50)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    raw = args.path.read_bytes()
    text = raw.decode(args.encoding)
    lines = text.splitlines()
    needle = args.contains if args.contains is not None else decode_escape(args.contains_escape)

    matches = []
    if needle is not None:
        for index, line in enumerate(lines, start=1):
            if needle in line:
                start = max(1, index - args.context)
                end = min(len(lines), index + args.context)
                matches.append(
                    {
                        "line": index,
                        "context_start": start,
                        "context_end": end,
                        "text": line,
                        "context": lines[start - 1 : end],
                    }
                )
                if len(matches) >= args.max_matches:
                    break

    result = {
        "ok": True,
        "path": str(args.path),
        "encoding": args.encoding,
        "sha256": hashlib.sha256(raw).hexdigest(),
        "line_count": len(lines),
        "contains": needle,
        "match_count": len(matches),
        "matches": matches,
    }
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
