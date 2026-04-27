#!/usr/bin/env python3
import argparse
import difflib
import json
import os
import re
import sys
from typing import Dict, List, Optional, Tuple


def normalize(value: str) -> str:
    text = value.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def find_best_line_index(lines: List[str], block_text: str) -> Optional[int]:
    target = normalize(block_text).lower()
    if not target:
        return None

    for idx, line in enumerate(lines):
        if target and target in normalize(line).lower():
            return idx

    best_idx: Optional[int] = None
    best_score = 0.0
    for idx, line in enumerate(lines):
        score = difflib.SequenceMatcher(
            None,
            normalize(line).lower(),
            target,
        ).ratio()
        if score > best_score:
            best_score = score
            best_idx = idx

    if best_idx is None or best_score < 0.45:
        return None
    return best_idx


def build_suggestion(comment: Dict[str, str]) -> str:
    author = normalize(str(comment.get("author", "")))
    timestamp = normalize(str(comment.get("timestamp", "")))
    text = normalize(str(comment.get("comment", "")))
    suffix_parts = [part for part in (author, timestamp) if part]
    suffix = " | ".join(suffix_parts)
    if suffix:
        return f"[Coda Comment] {text} ({suffix})"
    return f"[Coda Comment] {text}"


def apply_comment_with_approval(markdown_path: str, comments: List[Dict[str, str]]) -> int:
    if not os.path.exists(markdown_path):
        raise RuntimeError(f"Markdown file not found: {markdown_path}")

    with open(markdown_path, "r", encoding="utf-8") as handle:
        content = handle.read()
    lines = content.splitlines()

    for idx, entry in enumerate(comments, start=1):
        block_text = normalize(str(entry.get("blockText", "")))
        comment_text = normalize(str(entry.get("comment", "")))
        author = normalize(str(entry.get("author", "")))
        timestamp = normalize(str(entry.get("timestamp", "")))
        if not comment_text:
            continue

        line_index = find_best_line_index(lines, block_text)
        suggestion = build_suggestion(entry)

        print()
        print(f"Comment {idx}/{len(comments)}")
        print(f"Author: {author or '(unknown)'}")
        print(f"Time: {timestamp or '(unknown)'}")
        print(f"Comment: {comment_text}")
        print(f"Matched line: {line_index + 1 if line_index is not None else 'not found'}")
        print("Proposed edit:")
        print(suggestion)

        answer = input("Apply this edit? [y]es / [s]kip / [q]uit: ").strip().lower()
        if answer == "q":
            break
        if answer != "y":
            continue

        if line_index is None:
            lines.append("")
            lines.append(suggestion)
        else:
            insertion_at = line_index + 1
            lines.insert(insertion_at, suggestion)

        with open(markdown_path, "w", encoding="utf-8") as handle:
            handle.write("\n".join(lines) + "\n")
        print("Applied.")

    return 0


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Review coda_comments.json and apply edits to markdown one-by-one with approval."
    )
    parser.add_argument(
        "--markdown",
        default="Typeface Page Draft.md",
        help="Path to markdown file to update.",
    )
    parser.add_argument(
        "--comments",
        default="coda_comments.json",
        help="Path to extracted comments JSON file.",
    )
    return parser.parse_args(argv)


def main(argv: List[str]) -> int:
    args = parse_args(argv)
    if not os.path.exists(args.comments):
        raise RuntimeError(f"Comments file not found: {args.comments}")

    with open(args.comments, "r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, list):
        raise RuntimeError("Comments JSON must be a list")

    comments: List[Dict[str, str]] = []
    for item in data:
        if isinstance(item, dict):
            comments.append(item)

    return apply_comment_with_approval(args.markdown, comments)


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        print("Interrupted.")
        sys.exit(130)
