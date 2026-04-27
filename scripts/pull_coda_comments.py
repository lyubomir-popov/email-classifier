#!/usr/bin/env python3
import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass, asdict
from typing import Iterable, List, Optional, Sequence, Tuple

from playwright.sync_api import (
    BrowserContext,
    Locator,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)


THREAD_SELECTORS: Sequence[str] = (
    "[data-testid*='comment-thread']",
    "[data-testid*='commentThread']",
    "[data-testid*='thread']",
    "[role='listitem'] [data-testid*='comment']",
)
COMMENT_PANEL_SELECTORS: Sequence[str] = (
    "[data-testid*='comments-panel']",
    "[data-testid*='commentPanel']",
    "[data-testid*='commentsSidebar']",
    "[aria-label*='Comments'][role='complementary']",
)
OPEN_COMMENTS_BUTTON_SELECTORS: Sequence[str] = (
    "button[aria-label='Comments']",
    "button[aria-label*='comment' i]",
    "[data-testid*='open-comments']",
    "[data-testid*='commentsButton']",
    "[data-testid*='commentIcon']",
)
COMMENT_TEXT_SELECTORS: Sequence[str] = (
    "[data-testid*='comment-body']",
    "[data-testid*='commentBody']",
    "[data-testid*='message-body']",
    "[data-testid*='messageBody']",
    "[role='article']",
    "p",
)
COMMENT_AUTHOR_SELECTORS: Sequence[str] = (
    "[data-testid*='author']",
    "[data-testid*='user-name']",
    "[data-testid*='display-name']",
    ".author",
    "strong",
)
COMMENT_TIMESTAMP_SELECTORS: Sequence[str] = (
    "time",
    "[data-testid*='timestamp']",
    "[data-testid*='time']",
)
ACTIVE_BLOCK_SELECTORS: Sequence[str] = (
    "[aria-selected='true'] [contenteditable='true']",
    "[data-selected='true'] [contenteditable='true']",
    "[class*='selected'] [contenteditable='true']",
    "[data-testid*='canvas'] [contenteditable='true']",
)
THREAD_LINE_SELECTORS: Sequence[str] = (
    "[data-testid*='comment-item']",
    "[data-testid*='message']",
    "[role='article']",
)


@dataclass
class CodaComment:
    blockText: str
    comment: str
    author: str
    timestamp: str


def normalize_text(value: str) -> str:
    text = value.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def locator_visible(locator: Locator) -> bool:
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except Exception:
        return False


def first_visible_locator(page: Page, selectors: Iterable[str]) -> Optional[Locator]:
    for selector in selectors:
        locator = page.locator(selector)
        if locator_visible(locator):
            return locator.first
    return None


def safe_inner_text(locator: Locator) -> str:
    try:
        return normalize_text(locator.inner_text())
    except Exception:
        return ""


def first_text_in(locator: Locator, selectors: Sequence[str]) -> str:
    for selector in selectors:
        candidate = locator.locator(selector)
        try:
            if candidate.count() <= 0:
                continue
        except Exception:
            continue
        for i in range(min(candidate.count(), 5)):
            text = safe_inner_text(candidate.nth(i))
            if text:
                return text
    return ""


def first_timestamp_in(locator: Locator) -> str:
    for selector in COMMENT_TIMESTAMP_SELECTORS:
        candidate = locator.locator(selector)
        try:
            if candidate.count() <= 0:
                continue
        except Exception:
            continue
        element = candidate.first
        for attr in ("datetime", "title", "aria-label"):
            try:
                value = element.get_attribute(attr)
            except Exception:
                value = None
            if value:
                return normalize_text(value)
        text = safe_inner_text(element)
        if text:
            return text
    return ""


def panel_has_threads(page: Page) -> bool:
    for selector in THREAD_SELECTORS:
        locator = page.locator(selector)
        try:
            if locator.count() > 0:
                return True
        except Exception:
            continue
    return False


def find_thread_locator(page: Page) -> Optional[Locator]:
    for selector in THREAD_SELECTORS:
        locator = page.locator(selector)
        try:
            if locator.count() > 0:
                return locator
        except Exception:
            continue
    return None


def find_comment_panel(page: Page) -> Optional[Locator]:
    panel = first_visible_locator(page, COMMENT_PANEL_SELECTORS)
    if panel is not None:
        return panel
    thread_locator = find_thread_locator(page)
    if thread_locator is None:
        return None
    try:
        return thread_locator.first.locator("xpath=ancestor::*[self::aside or self::section or self::div][1]")
    except Exception:
        return None


def open_comment_panel(page: Page) -> None:
    if panel_has_threads(page):
        return

    button = first_visible_locator(page, OPEN_COMMENTS_BUTTON_SELECTORS)
    if button is not None:
        button.click(timeout=5000)
        page.wait_for_timeout(1000)
        if panel_has_threads(page):
            return

    for label in ("Comments", "Comment"):
        menu = page.get_by_role("button", name=re.compile(label, re.IGNORECASE))
        try:
            if menu.count() > 0 and menu.first.is_visible():
                menu.first.click(timeout=5000)
                page.wait_for_timeout(1000)
                if panel_has_threads(page):
                    return
        except Exception:
            continue


def maybe_wait_for_login(page: Page, login_wait_seconds: int) -> None:
    title = (page.title() or "").lower()
    url = (page.url or "").lower()
    login_indicators = (
        "accounts.google.com",
        "login",
        "sign in",
        "signin",
    )
    if not any(x in title or x in url for x in login_indicators):
        return

    print("Login appears required. Complete login in the opened browser window.")
    print(f"Waiting up to {login_wait_seconds} seconds for you to finish login...")
    deadline = time.time() + login_wait_seconds
    while time.time() < deadline:
        page.wait_for_timeout(1500)
        current_url = (page.url or "").lower()
        if "coda.io" in current_url and not any(x in current_url for x in ("login", "signin")):
            break


def scroll_comment_panel_to_load_all(page: Page, max_rounds: int = 30) -> None:
    panel = find_comment_panel(page)
    if panel is None:
        return

    last_height = -1
    for _ in range(max_rounds):
        try:
            current_height = panel.evaluate("el => el.scrollHeight")
            panel.evaluate("el => { el.scrollTop = el.scrollHeight; }")
        except Exception:
            return
        page.wait_for_timeout(350)
        if current_height == last_height:
            return
        last_height = current_height


def active_block_text(page: Page) -> str:
    for selector in ACTIVE_BLOCK_SELECTORS:
        locator = page.locator(selector)
        try:
            count = locator.count()
        except Exception:
            continue
        if count <= 0:
            continue
        for i in range(min(count, 5)):
            text = safe_inner_text(locator.nth(i))
            if text and len(text) > 2 and "comment" not in text.lower():
                return text
    return ""


def extract_comment_rows_from_thread(thread: Locator, page: Page) -> List[CodaComment]:
    comments: List[CodaComment] = []
    try:
        thread.click(timeout=5000)
        page.wait_for_timeout(300)
    except Exception:
        pass

    block_text = active_block_text(page)
    rows = []
    for selector in THREAD_LINE_SELECTORS:
        candidate = thread.locator(selector)
        try:
            if candidate.count() > 0:
                rows = [candidate.nth(i) for i in range(candidate.count())]
                break
        except Exception:
            continue

    if not rows:
        rows = [thread]

    for row in rows:
        comment_text = first_text_in(row, COMMENT_TEXT_SELECTORS) or safe_inner_text(row)
        author = first_text_in(row, COMMENT_AUTHOR_SELECTORS)
        timestamp = first_timestamp_in(row)
        if not comment_text:
            continue
        comments.append(
            CodaComment(
                blockText=block_text,
                comment=comment_text,
                author=author,
                timestamp=timestamp,
            )
        )

    return comments


def dedupe_comments(items: List[CodaComment]) -> List[CodaComment]:
    seen = set()
    output: List[CodaComment] = []
    for item in items:
        key = (
            item.blockText.strip(),
            item.comment.strip(),
            item.author.strip(),
            item.timestamp.strip(),
        )
        if key in seen:
            continue
        seen.add(key)
        output.append(item)
    return output


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract comment threads from a Coda canvas page into coda_comments.json."
    )
    parser.add_argument("--url", required=True, help="Coda canvas page URL")
    parser.add_argument("--output", default="coda_comments.json", help="Output JSON path")
    parser.add_argument(
        "--user-data-dir",
        default=os.path.join(".playwright", "coda_profile"),
        help="Persistent Chromium profile directory",
    )
    parser.add_argument(
        "--login-wait-seconds",
        type=int,
        default=180,
        help="Max seconds to wait for first-time manual login",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run browser headless (default is visible mode).",
    )
    return parser.parse_args(argv)


def run_extraction(args: argparse.Namespace) -> int:
    os.makedirs(args.user_data_dir, exist_ok=True)
    out_dir = os.path.dirname(os.path.abspath(args.output))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    with sync_playwright() as pw:
        context: BrowserContext = pw.chromium.launch_persistent_context(
            user_data_dir=args.user_data_dir,
            headless=args.headless,
            viewport={"width": 1600, "height": 1000},
        )
        page: Page
        if context.pages:
            page = context.pages[0]
        else:
            page = context.new_page()

        try:
            page.goto(args.url, wait_until="domcontentloaded", timeout=120_000)
        except PlaywrightTimeoutError:
            page.goto(args.url, wait_until="load", timeout=120_000)

        page.wait_for_timeout(2000)
        maybe_wait_for_login(page, login_wait_seconds=args.login_wait_seconds)
        page.wait_for_timeout(1000)
        open_comment_panel(page)
        scroll_comment_panel_to_load_all(page)

        thread_locator = find_thread_locator(page)
        if thread_locator is None:
            print("No comment threads found. Writing empty output.")
            with open(args.output, "w", encoding="utf-8") as handle:
                json.dump([], handle, indent=2, ensure_ascii=False)
            context.close()
            return 0

        rows: List[CodaComment] = []
        thread_count = thread_locator.count()
        for i in range(thread_count):
            thread = thread_locator.nth(i)
            rows.extend(extract_comment_rows_from_thread(thread, page))

        rows = dedupe_comments(rows)
        payload = [asdict(row) for row in rows]
        with open(args.output, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=False)
        print(f"Extracted {len(payload)} comments to {args.output}")
        context.close()
    return 0


def main(argv: List[str]) -> int:
    args = parse_args(argv)
    try:
        return run_extraction(args)
    except KeyboardInterrupt:
        print("Interrupted.")
        return 130


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
