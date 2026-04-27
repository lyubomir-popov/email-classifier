#!/usr/bin/env python3
import argparse
import csv
import datetime as dt
import email
from email import policy
from email.header import decode_header, make_header
from email.utils import parseaddr
import html
import imaplib
import json
import getpass
import os
import re
import sys
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

import requests


ACTIONS = {"KEEP", "UNSUBSCRIBE", "DELETE", "REVIEW"}
LABELS_BY_ACTION = {
    "KEEP": "AI/KEEP",
    "UNSUBSCRIBE": "AI/UNSUBSCRIBE",
    "DELETE": "AI/DELETE",
    "REVIEW": "AI/REVIEW",
}
DEFAULT_PROCESSED_LABEL = "AI/PROCESSED"
DEFAULT_GMAIL_PROMOTIONS_QUERY = "category:promotions"
PERSONAL_MAILBOX_DOMAINS = {
    "gmail.com",
    "outlook.com",
    "hotmail.com",
    "live.com",
    "icloud.com",
    "me.com",
    "mac.com",
    "proton.me",
    "protonmail.com",
    "yahoo.com",
    "gmx.com",
    "aol.com",
    "btinternet.com",
    "sky.com",
}
BRAND_LIKE_NAME_HINTS = {
    "team",
    "support",
    "newsletter",
    "digest",
    "alerts",
    "marketing",
    "sales",
    "deals",
    "offers",
    "no-reply",
    "noreply",
    "info",
    "updates",
    "mail",
    "admin",
    "bot",
    "service",
}
FORCE_KEEP_PATTERNS: Sequence[Tuple[str, re.Pattern[str]]] = [
    (
        "otp_verification",
        re.compile(
            r"\b(otp|verification code|one-time code|one time code|2fa|mfa|two[- ]factor)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "security_access",
        re.compile(
            r"\b(password reset|reset your password|security alert|new sign[- ]in|new device|"
            r"suspicious activity|account access)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "billing_tax",
        re.compile(
            r"\b(invoice|receipt|payment|statement|renewal|subscription receipt|tax|vat|"
            r"billing|utility bill|credit card bill|phone bill)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "order_confirmation_tracking",
        re.compile(
            r"\b(order confirmation|confirmation number|tracking number)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "bank_government",
        re.compile(
            r"\b(bank|credit card|mortgage|hmrc|nhs|dvla|gov\.uk)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "legal_medical",
        re.compile(
            r"\b(legal notice|court|medical|doctor|hospital|insurance claim)\b",
            re.IGNORECASE,
        ),
    ),
]
PROMOTIONS_WEAK_KEEP_SIGNALS = {"billing_tax", "bank_government", "legal_medical"}
NEWSLETTER_PATTERNS: Sequence[Tuple[str, re.Pattern[str]]] = [
    (
        "newsletter_digest",
        re.compile(
            r"\b(daily|morning briefing|evening briefing|your daily|top stories|roundup|"
            r"newsletter|digest|edition|bulletin|daily update|weekly update)\b",
            re.IGNORECASE,
        ),
    ),
]
TRANSACTIONAL_CASE_PRIMARY_PATTERN = re.compile(
    r"\b(remortgage|mortgage application|conveyancing|property (purchase|sale)|"
    r"loan application|case update|application update)\b",
    re.IGNORECASE,
)
TRANSACTIONAL_CASE_CONTEXT_PATTERNS: Sequence[Tuple[str, re.Pattern[str]]] = [
    (
        "case_reference",
        re.compile(
            r"\b(case reference|instruction reference|reference number|application reference)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "id_pattern",
        re.compile(
            r"\b(?:ref(?:erence)?|case|instruction)\s*(?:no\.?|number|ref)?\s*[:#]?\s*"
            r"[A-Z0-9][A-Z0-9\/\-]{5,}\b",
            re.IGNORECASE,
        ),
    ),
    (
        "workflow",
        re.compile(
            r"\b(milestone|offer( issued)?|verified your identity|case tracking|"
            r"exchange|completion)\b",
            re.IGNORECASE,
        ),
    ),
    (
        "uk_postcode",
        re.compile(r"\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b", re.IGNORECASE),
    ),
    (
        "property_address",
        re.compile(
            r"\b\d{1,5}\s+[A-Za-z][A-Za-z\s]{1,40}\s+"
            r"(avenue|road|street|close|lane|drive|court|way)\b",
            re.IGNORECASE,
        ),
    ),
]
DELIVERY_PATTERNS: Sequence[Tuple[str, re.Pattern[str]]] = [
    (
        "delivery_notice",
        re.compile(
            r"\b(order delivered|delivered today|package was delivered|has been delivered|"
            r"delivery update|delivered)\b",
            re.IGNORECASE,
        ),
    ),
]
DELIVERY_RETURNS_PATTERNS: Sequence[Tuple[str, re.Pattern[str]]] = [
    (
        "returns_related",
        re.compile(
            r"\b(return window|return eligible|return by|start a return|return label|"
            r"refund policy)\b",
            re.IGNORECASE,
        ),
    ),
]
SUSPICIOUS_HUMAN_PATTERNS: Sequence[Tuple[str, re.Pattern[str]]] = [
    ("urgent_wire", re.compile(r"\b(urgent wire|wire transfer)\b", re.IGNORECASE)),
    ("gift_card", re.compile(r"\bgift card\b", re.IGNORECASE)),
    ("crypto_wallet", re.compile(r"\b(crypto|bitcoin|wallet)\b", re.IGNORECASE)),
    ("verify_login", re.compile(r"\b(verify your account|confirm login)\b", re.IGNORECASE)),
]
UNSUBSCRIBE_WORD_PATTERN = re.compile(r"\bunsubscribe\b", re.IGNORECASE)
DENYLIST_PATTERNS: Sequence[Tuple[str, re.Pattern[str]]] = [
    ("eventbrite", re.compile(r"eventbrite", re.IGNORECASE)),
    ("guardian", re.compile(r"(guardian\.co\.uk|theguardian\.com)", re.IGNORECASE)),
    ("mailchimp", re.compile(r"mailchimp", re.IGNORECASE)),
    ("sendgrid", re.compile(r"sendgrid", re.IGNORECASE)),
    ("constantcontact", re.compile(r"constantcontact", re.IGNORECASE)),
    ("campaign", re.compile(r"campaign-", re.IGNORECASE)),
    ("marketing_domain", re.compile(r"(^e\..*\.marketing$|\.marketing$)", re.IGNORECASE)),
]
AGGRESSIVE_SYSTEM_PROMPT = """You classify email using a strict non-lenient policy.
Return only strict JSON with exactly this schema:
{"action":"KEEP|UNSUBSCRIBE|DELETE|REVIEW","confidence":0-1,"reason":"<=120 chars"}

Policy:
- KEEP is extremely narrow:
  security/account access, billing/receipts/tax, legal/medical/bank/government, personal human mail, work-critical.
- For newsletters/marketing/promotions: prefer UNSUBSCRIBE, otherwise DELETE.
- If uncertain, choose REVIEW, never optimistic KEEP.
- confidence must be 0..1.
- reason <=120 chars.
Do not include markdown, preface, or extra keys.
"""
NORMAL_SYSTEM_PROMPT = """You classify promotional emails into actions.
Return only strict JSON with exactly this schema:
{"action":"KEEP|UNSUBSCRIBE|DELETE|REVIEW","confidence":0-1,"reason":"<=120 chars"}
"""

ENV_FILE_PATH = ".env.local"
ENV_BACKEND_KEY = "EMAIL_CLASSIFIER_BACKEND"
ENV_OPENAI_MODEL_KEY = "OPENAI_MODEL"
ENV_OLLAMA_URL_KEY = "OLLAMA_URL"
ENV_OLLAMA_MODEL_KEY = "OLLAMA_MODEL"


def load_env_file(path: str) -> None:
    if not path or not os.path.exists(path):
        return

    with open(path, "r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("export "):
                line = line[len("export ") :].strip()
            if "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            if not key or key in os.environ:
                continue

            value = value.strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
                value = value[1:-1]
            os.environ[key] = value


def quote_env_value(value: str) -> str:
    if not value:
        return value
    if any(char.isspace() for char in value) or "#" in value:
        escaped = value.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'
    return value


def upsert_env_file(path: str, updates: Dict[str, str]) -> None:
    existing_lines: List[str] = []
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as handle:
            existing_lines = handle.read().splitlines()

    written_keys: Set[str] = set()
    output_lines: List[str] = []
    for line in existing_lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#") and "=" in line:
            key = line.split("=", 1)[0].strip()
            if key in updates:
                output_lines.append(f"{key}={quote_env_value(updates[key])}")
                written_keys.add(key)
                continue
        output_lines.append(line)

    if not existing_lines:
        output_lines.extend(
            [
                "# Local-only settings for this machine.",
                "# This file is gitignored.",
                "",
            ]
        )

    for key, value in updates.items():
        if key in written_keys:
            continue
        output_lines.append(f"{key}={quote_env_value(value)}")

    with open(path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(output_lines).rstrip() + "\n")


def can_prompt_interactively() -> bool:
    return sys.stdin.isatty() and sys.stdout.isatty()


def prompt_text(
    label: str,
    *,
    default: Optional[str] = None,
    secret: bool = False,
    allow_blank: bool = False,
) -> str:
    while True:
        suffix = f" [{default}]" if default else ""
        prompt = f"{label}{suffix}: "
        value = getpass.getpass(prompt) if secret else input(prompt)
        value = value.strip()
        if not value and default is not None:
            value = default
        if value or allow_blank:
            return value
        print("A value is required.")


def prompt_choice(label: str, options: Dict[str, str], default: str) -> str:
    while True:
        print(label)
        for key, description in options.items():
            default_suffix = " (default)" if key == default else ""
            print(f"  {key}. {description}{default_suffix}")
        value = input("Choose an option: ").strip().lower() or default
        if value in options:
            return value
        print("Please choose one of the listed options.")


def normalize_app_password(value: str) -> str:
    return re.sub(r"\s+", "", value).strip()


def should_run_setup_wizard(args: argparse.Namespace, backend_explicit: bool) -> bool:
    if args.setup:
        return True
    if not can_prompt_interactively():
        return False
    if args.email and args.app_password:
        if args.cleanup_ai_labels:
            return False
        if args.backend == "openai" and os.getenv("OPENAI_API_KEY", "").strip():
            return False
        if args.backend == "ollama":
            return False
    if backend_explicit:
        return True
    if os.getenv(ENV_BACKEND_KEY, "").strip():
        return True
    return not os.path.exists(ENV_FILE_PATH)


def run_setup_wizard(args: argparse.Namespace, *, backend_explicit: bool) -> None:
    print()
    print("Email Classifier setup")
    print("I will create or update .env.local so this machine can run the tool without extra flags.")
    print()

    if not args.cleanup_ai_labels and not backend_explicit and not os.getenv(ENV_BACKEND_KEY, "").strip():
        chosen_backend = prompt_choice(
            "Choose a classifier backend:",
            {
                "1": "OpenAI API (recommended for most people)",
                "2": "Ollama local model (advanced, requires Ollama running locally)",
            },
            default="1",
        )
        args.backend = "openai" if chosen_backend == "1" else "ollama"

    default_email = args.email or os.getenv("GMAIL_EMAIL", "").strip() or None
    args.email = prompt_text("Gmail address", default=default_email)

    current_password = args.app_password or os.getenv("GMAIL_APP_PASSWORD", "").strip()
    password_default = "leave unchanged" if current_password else None
    password_value = prompt_text(
        "Gmail App Password",
        default=password_default,
        secret=True,
    )
    if password_value == "leave unchanged" and current_password:
        args.app_password = current_password
    else:
        args.app_password = normalize_app_password(password_value)

    env_updates = {
        "GMAIL_EMAIL": args.email,
        "GMAIL_APP_PASSWORD": args.app_password,
    }

    if not args.cleanup_ai_labels:
        env_updates[ENV_BACKEND_KEY] = args.backend
        if args.backend == "openai":
            current_api_key = os.getenv("OPENAI_API_KEY", "").strip()
            api_default = "leave unchanged" if current_api_key else None
            api_key = prompt_text(
                "OpenAI API key",
                default=api_default,
                secret=True,
            )
            if api_key == "leave unchanged" and current_api_key:
                api_key = current_api_key
            os.environ["OPENAI_API_KEY"] = api_key.strip()
            env_updates["OPENAI_API_KEY"] = api_key.strip()
        else:
            default_ollama_url = args.ollama_url or os.getenv(ENV_OLLAMA_URL_KEY, "").strip() or "http://127.0.0.1:11434"
            default_ollama_model = args.ollama_model or os.getenv(ENV_OLLAMA_MODEL_KEY, "").strip() or "llama3:latest"
            args.ollama_url = prompt_text("Ollama URL", default=default_ollama_url)
            args.ollama_model = prompt_text("Ollama model", default=default_ollama_model)
            env_updates[ENV_OLLAMA_URL_KEY] = args.ollama_url
            env_updates[ENV_OLLAMA_MODEL_KEY] = args.ollama_model

    upsert_env_file(ENV_FILE_PATH, env_updates)
    os.environ["GMAIL_EMAIL"] = args.email
    os.environ["GMAIL_APP_PASSWORD"] = args.app_password
    os.environ[ENV_BACKEND_KEY] = args.backend
    print()
    print("Saved local settings to .env.local")
    if args.backend == "openai":
        print("Future runs will default to OpenAI unless you override --backend.")
    else:
        print("Future runs will default to Ollama unless you override --backend.")
    print()


@dataclass
class EmailRecord:
    uid: str
    from_header: str
    subject: str
    date: str
    message_id: str
    list_unsubscribe: str
    snippet: str


@dataclass
class Classification:
    action: str
    confidence: float
    reason: str
    raw_response: str


@dataclass
class PolicyContext:
    mode: str
    folder: str
    gmail_query: Optional[str]
    is_promotions_scan: bool
    folder_default_policy: str
    allowlist_emails: Set[str]
    allowlist_domains: Set[str]


@dataclass
class DecisionResult:
    final_action: str
    model_action: str
    model_confidence: Optional[float]
    model_reason: str
    model_raw: str
    decision_source: str
    rule_match: str
    human_sender: bool
    allowlist_match: bool


class ClassifierError(RuntimeError):
    pass


class BaseClassifier:
    backend_name = "base"
    model_name = ""

    def classify(self, record: EmailRecord) -> Classification:
        raise NotImplementedError


class OpenAIResponsesClassifier(BaseClassifier):
    backend_name = "openai"

    def __init__(
        self,
        api_key: str,
        model: str,
        timeout_seconds: int,
        system_prompt: str,
    ) -> None:
        self.api_key = api_key
        self.model_name = model
        self.timeout_seconds = timeout_seconds
        self.system_prompt = system_prompt
        self.url = "https://api.openai.com/v1/responses"

    def classify(self, record: EmailRecord) -> Classification:
        payload = {
            "model": self.model_name,
            "temperature": 0,
            "input": [
                {
                    "role": "system",
                    "content": [{"type": "input_text", "text": self.system_prompt}],
                },
                {
                    "role": "user",
                    "content": [{"type": "input_text", "text": build_user_prompt(record)}],
                },
            ],
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        response = requests.post(
            self.url,
            headers=headers,
            json=payload,
            timeout=self.timeout_seconds,
        )
        if response.status_code >= 400:
            raise ClassifierError(
                f"OpenAI request failed ({response.status_code}): {response.text[:500]}"
            )
        data = response.json()
        text = extract_openai_output_text(data)
        return parse_classification(text)


class OllamaClassifier(BaseClassifier):
    backend_name = "ollama"

    def __init__(
        self,
        base_url: str,
        model: str,
        timeout_seconds: int,
        system_prompt: str,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.model_name = model
        self.timeout_seconds = timeout_seconds
        self.system_prompt = system_prompt
        self.use_openai_compat = self.base_url.endswith("/v1")

    def classify(self, record: EmailRecord) -> Classification:
        if self.use_openai_compat:
            return self._classify_openai_compat(record)
        return self._classify_native(record)

    def _classify_native(self, record: EmailRecord) -> Classification:
        endpoint = self.base_url
        if not endpoint.endswith("/api/generate"):
            endpoint = f"{endpoint}/api/generate"
        payload = {
            "model": self.model_name,
            "prompt": f"{self.system_prompt}\n\n{build_user_prompt(record)}",
            "stream": False,
            "format": "json",
            "options": {"temperature": 0},
        }
        response = requests.post(endpoint, json=payload, timeout=self.timeout_seconds)
        if response.status_code >= 400:
            raise ClassifierError(
                f"Ollama native request failed ({response.status_code}): {response.text[:500]}"
            )
        data = response.json()
        text = str(data.get("response", "")).strip()
        if not text:
            raise ClassifierError("Ollama native response had no text")
        return parse_classification(text)

    def _classify_openai_compat(self, record: EmailRecord) -> Classification:
        endpoint = f"{self.base_url}/chat/completions"
        payload = {
            "model": self.model_name,
            "temperature": 0,
            "messages": [
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": build_user_prompt(record)},
            ],
        }
        response = requests.post(endpoint, json=payload, timeout=self.timeout_seconds)
        if response.status_code >= 400:
            raise ClassifierError(
                f"Ollama compat request failed ({response.status_code}): {response.text[:500]}"
            )
        data = response.json()
        try:
            text = data["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError) as exc:
            raise ClassifierError(f"Ollama compat response parse error: {exc}") from exc
        return parse_classification(str(text))

def decode_header_value(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def html_to_text(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    value = html.unescape(value)
    return normalize_text(value)


def decode_part_payload(part: email.message.Message) -> str:
    payload = part.get_payload(decode=True)
    charset = part.get_content_charset() or "utf-8"
    if payload is None:
        value = part.get_payload()
        return value if isinstance(value, str) else ""
    try:
        return payload.decode(charset, errors="replace")
    except LookupError:
        return payload.decode("utf-8", errors="replace")


def extract_plain_text_snippet(msg: email.message.Message, max_chars: int = 2000) -> str:
    plain_parts: List[str] = []
    html_parts: List[str] = []

    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            disposition = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disposition:
                continue
            content_type = part.get_content_type().lower()
            if content_type not in {"text/plain", "text/html"}:
                continue
            decoded = decode_part_payload(part).strip()
            if not decoded:
                continue
            if content_type == "text/plain":
                plain_parts.append(decoded)
            else:
                html_parts.append(decoded)
    else:
        content_type = msg.get_content_type().lower()
        decoded = decode_part_payload(msg).strip()
        if content_type == "text/plain":
            plain_parts.append(decoded)
        elif content_type == "text/html":
            html_parts.append(decoded)
        elif decoded:
            plain_parts.append(decoded)

    body = "\n".join(plain_parts).strip()
    if not body and html_parts:
        body = html_to_text("\n".join(html_parts))
    body = normalize_text(body)
    return body[:max_chars]


def parse_rfc822_message(uid: str, rfc822_bytes: bytes) -> EmailRecord:
    msg = email.message_from_bytes(rfc822_bytes, policy=policy.default)
    return EmailRecord(
        uid=uid,
        from_header=decode_header_value(msg.get("From")),
        subject=decode_header_value(msg.get("Subject")),
        date=decode_header_value(msg.get("Date")),
        message_id=decode_header_value(msg.get("Message-ID")),
        list_unsubscribe=decode_header_value(msg.get("List-Unsubscribe")),
        snippet=extract_plain_text_snippet(msg, max_chars=2000),
    )


def build_user_prompt(record: EmailRecord) -> str:
    return (
        "Classify this email under the provided policy.\n"
        f"From: {record.from_header}\n"
        f"Subject: {record.subject}\n"
        f"Date: {record.date}\n"
        f"List-Unsubscribe: {record.list_unsubscribe}\n"
        "Body snippet:\n"
        f"{record.snippet}\n"
        'Return only JSON: {"action":"KEEP|UNSUBSCRIBE|DELETE|REVIEW","confidence":0-1,"reason":"<=120 chars"}'
    )


def extract_openai_output_text(response_json: Dict[str, object]) -> str:
    output_text = response_json.get("output_text")
    if isinstance(output_text, str) and output_text.strip():
        return output_text

    output = response_json.get("output")
    if isinstance(output, list):
        chunks: List[str] = []
        for item in output:
            if not isinstance(item, dict):
                continue
            content = item.get("content")
            if not isinstance(content, list):
                continue
            for block in content:
                if not isinstance(block, dict):
                    continue
                if block.get("type") in {"output_text", "text"}:
                    text = block.get("text")
                    if isinstance(text, str) and text.strip():
                        chunks.append(text)
        if chunks:
            return "\n".join(chunks)

    raise ClassifierError("OpenAI response did not contain model text output")


def parse_classification(raw: str) -> Classification:
    candidate = raw.strip()
    parsed: Optional[Dict[str, object]] = None
    try:
        value = json.loads(candidate)
        if isinstance(value, dict):
            parsed = value
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", candidate, flags=re.DOTALL)
        if match:
            value = json.loads(match.group(0))
            if isinstance(value, dict):
                parsed = value
    if parsed is None:
        raise ClassifierError(f"Classifier returned invalid JSON: {candidate[:300]}")

    action = str(parsed.get("action", "")).upper().strip()
    if action not in ACTIONS:
        action = "REVIEW"
    confidence_raw = parsed.get("confidence", 0)
    try:
        confidence = float(confidence_raw)
    except (TypeError, ValueError):
        confidence = 0.0
    confidence = max(0.0, min(1.0, confidence))
    reason = str(parsed.get("reason", "")).strip() or "No reason provided"
    reason = reason[:120]

    normalized = {"action": action, "confidence": confidence, "reason": reason}
    return Classification(
        action=action,
        confidence=confidence,
        reason=reason,
        raw_response=json.dumps(normalized, separators=(",", ":"), ensure_ascii=True),
    )


def quote_imap_string(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def parse_mailbox_name(list_line: str) -> str:
    match = re.search(r' "((?:[^"\\]|\\.)*)"$', list_line)
    if match:
        return match.group(1).replace('\\"', '"').replace("\\\\", "\\")
    parts = list_line.split(" ")
    return parts[-1].strip('"')


def list_mailboxes(conn: imaplib.IMAP4_SSL) -> List[str]:
    status, data = conn.list()
    if status != "OK" or data is None:
        return []
    names: List[str] = []
    for line_bytes in data:
        if not line_bytes:
            continue
        line = line_bytes.decode("utf-8", errors="replace")
        names.append(parse_mailbox_name(line))
    return names


def has_gmail_ext(conn: imaplib.IMAP4_SSL) -> bool:
    capabilities = getattr(conn, "capabilities", ())
    for capability in capabilities:
        if isinstance(capability, bytes):
            capability = capability.decode("ascii", errors="ignore")
        if str(capability).upper() == "X-GM-EXT-1":
            return True
    return False


def escape_gmail_query_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def compose_gmail_query(
    base_query: Optional[str],
    exclude_label: Optional[str] = None,
) -> Optional[str]:
    query = (base_query or "").strip()
    excluded = ""
    if exclude_label:
        label = exclude_label.strip()
        if label:
            excluded = f'-label:"{escape_gmail_query_value(label)}"'

    if query and excluded:
        return f"({query}) {excluded}"
    if query:
        return query
    if excluded:
        return excluded
    return None


def combine_gmail_queries(primary: Optional[str], extra: Optional[str]) -> Optional[str]:
    first = (primary or "").strip()
    second = (extra or "").strip()
    if first and second:
        return f"({first}) ({second})"
    if first:
        return first
    if second:
        return second
    return None


def build_any_label_query(labels: Iterable[str]) -> Optional[str]:
    clauses: List[str] = []
    for label in labels:
        value = str(label or "").strip()
        if not value:
            continue
        clauses.append(f'label:"{escape_gmail_query_value(value)}"')
    if not clauses:
        return None
    if len(clauses) == 1:
        return clauses[0]
    return f"({' OR '.join(clauses)})"


def resolve_folder(
    conn: imaplib.IMAP4_SSL,
    explicit_folder: Optional[str],
) -> Tuple[str, Optional[str], str]:
    if explicit_folder:
        return explicit_folder, None, "explicit_folder"

    mailboxes = list_mailboxes(conn)
    lookup = {name.lower(): name for name in mailboxes}
    if "[gmail]/promotions" in lookup:
        return lookup["[gmail]/promotions"], None, "promotions_folder"
    if "promotions" in lookup:
        return lookup["promotions"], None, "promotions_folder"

    if has_gmail_ext(conn):
        if "[gmail]/all mail" in lookup:
            return (
                lookup["[gmail]/all mail"],
                DEFAULT_GMAIL_PROMOTIONS_QUERY,
                "promotions_category_query",
            )
        if "inbox" in lookup:
            return (
                lookup["inbox"],
                DEFAULT_GMAIL_PROMOTIONS_QUERY,
                "promotions_category_query",
            )

    if "[gmail]/all mail" in lookup:
        return lookup["[gmail]/all mail"], None, "all_mail_fallback"
    if "inbox" in lookup:
        return lookup["inbox"], None, "inbox_fallback"
    return "INBOX", None, "default_inbox"


def select_mailbox(conn: imaplib.IMAP4_SSL, folder: str, readonly: bool) -> None:
    status, _ = conn.select(quote_imap_string(folder), readonly=readonly)
    if status != "OK":
        available = ", ".join(list_mailboxes(conn))
        raise RuntimeError(
            f"Could not select folder '{folder}'. Available folders: {available or '(none)'}"
        )


def fetch_uids(
    conn: imaplib.IMAP4_SSL,
    limit: Optional[int],
    gmail_query: Optional[str] = None,
) -> List[str]:
    if gmail_query:
        status, data = conn.uid("SEARCH", None, "X-GM-RAW", quote_imap_string(gmail_query))
    else:
        status, data = conn.uid("SEARCH", None, "ALL")
    if status != "OK" or not data:
        if gmail_query:
            raise RuntimeError(
                f"Unable to search messages with Gmail query '{gmail_query}' in selected mailbox"
            )
        raise RuntimeError("Unable to search messages in selected mailbox")
    raw = data[0]
    if not raw:
        return []
    uids = [x.decode("ascii", errors="ignore") for x in raw.split()]
    uids.reverse()
    if limit is not None and limit > 0:
        return uids[:limit]
    return uids


def fetch_message_bytes(conn: imaplib.IMAP4_SSL, uid: str) -> Optional[bytes]:
    status, data = conn.uid("FETCH", uid, "(RFC822)")
    if status != "OK" or not data:
        return None
    for item in data:
        if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
            return item[1]
    return None


def ensure_labels(conn: imaplib.IMAP4_SSL, labels: Iterable[str]) -> None:
    for label in labels:
        status, _ = conn.create(quote_imap_string(label))
        if status not in {"OK", "NO"}:
            raise RuntimeError(f"Unable to create label '{label}' (status={status})")


def response_contains_label(data: object, label: str) -> bool:
    target = label.lower()
    if not data:
        return False
    if not isinstance(data, list):
        data = [data]
    for item in data:
        text = ""
        if isinstance(item, tuple) and item:
            first = item[0]
            if isinstance(first, bytes):
                text = first.decode("utf-8", errors="ignore")
            else:
                text = str(first)
        elif isinstance(item, bytes):
            text = item.decode("utf-8", errors="ignore")
        elif item is not None:
            text = str(item)
        if "X-GM-LABELS" in text.upper() and target in text.lower():
            return True
    return False


def add_label(conn: imaplib.IMAP4_SSL, uid: str, label: str) -> bool:
    label_args = [f"({quote_imap_string(label)})", label]
    for label_arg in label_args:
        status, data = conn.uid("STORE", uid, "+X-GM-LABELS", label_arg)
        if status != "OK":
            continue
        if response_contains_label(data, label):
            return True

        fetch_status, fetch_data = conn.uid("FETCH", uid, "(X-GM-LABELS)")
        if fetch_status == "OK" and response_contains_label(fetch_data, label):
            return True
    return False


def remove_label(conn: imaplib.IMAP4_SSL, uid: str, label: str) -> None:
    label_args = [f"({quote_imap_string(label)})", label]
    for label_arg in label_args:
        conn.uid("STORE", uid, "-X-GM-LABELS", label_arg)


def extract_present_ai_labels(data: object) -> Set[str]:
    if not data:
        return set()
    if not isinstance(data, list):
        data = [data]
    labels: Set[str] = set()
    pattern = re.compile(r"\bAI\/[A-Z]+\b")
    for item in data:
        text = ""
        if isinstance(item, tuple) and item:
            first = item[0]
            if isinstance(first, bytes):
                text = first.decode("utf-8", errors="ignore")
            else:
                text = str(first)
        elif isinstance(item, bytes):
            text = item.decode("utf-8", errors="ignore")
        elif item is not None:
            text = str(item)
        for match in pattern.finditer(text.upper()):
            labels.add(match.group(0))
    return labels


def set_ai_label(conn: imaplib.IMAP4_SSL, uid: str, target_label: str) -> bool:
    for label in LABELS_BY_ACTION.values():
        if label != target_label:
            remove_label(conn, uid, label)

    if not add_label(conn, uid, target_label):
        return False

    fetch_status, fetch_data = conn.uid("FETCH", uid, "(X-GM-LABELS)")
    if fetch_status != "OK":
        return False
    present = extract_present_ai_labels(fetch_data)
    target_upper = target_label.upper()
    known_labels = {label.upper() for label in LABELS_BY_ACTION.values()}
    return target_upper in present and not ((present & known_labels) - {target_upper})


def extract_sender(from_header: str) -> Tuple[str, str, str]:
    display_name, address = parseaddr(from_header or "")
    display_name = decode_header_value(display_name).strip()
    address = address.strip().lower()
    domain = ""
    if "@" in address:
        domain = address.rsplit("@", 1)[1]
    return display_name, address, domain


def normalize_match_text(record: EmailRecord) -> str:
    return "\n".join(
        [
            record.from_header,
            record.subject,
            record.date,
            record.message_id,
            record.list_unsubscribe,
            record.snippet,
        ]
    )


def has_list_unsubscribe(record: EmailRecord) -> bool:
    return bool(record.list_unsubscribe.strip())


def match_named_pattern(
    patterns: Sequence[Tuple[str, re.Pattern[str]]],
    text: str,
) -> Optional[str]:
    for name, pattern in patterns:
        if pattern.search(text):
            return name
    return None


def is_person_like_display_name(name: str) -> bool:
    cleaned = name.strip()
    if not cleaned:
        return False
    lowered = cleaned.lower()
    if re.search(r"\d", lowered):
        return False
    if any(hint in lowered for hint in BRAND_LIKE_NAME_HINTS):
        return False
    tokens = re.findall(r"[A-Za-z][A-Za-z'\-]+", cleaned)
    if len(tokens) < 2:
        return False
    if len("".join(tokens)) < 5:
        return False
    return True


def sender_matches_allowlist(
    sender_email: str,
    sender_domain: str,
    allowlist_emails: Set[str],
    allowlist_domains: Set[str],
) -> Optional[str]:
    if sender_email and sender_email in allowlist_emails:
        return f"email:{sender_email}"
    for domain in allowlist_domains:
        if not domain:
            continue
        if sender_domain == domain or sender_domain.endswith(f".{domain}"):
            return f"domain:{domain}"
    return None


def sender_matches_denylist(sender_domain: str) -> Optional[str]:
    if not sender_domain:
        return None
    for name, pattern in DENYLIST_PATTERNS:
        if pattern.search(sender_domain):
            return name
    return None


def contains_force_keep_signal(record: EmailRecord) -> Optional[str]:
    return match_named_pattern(FORCE_KEEP_PATTERNS, normalize_match_text(record))


def contains_newsletter_signal(record: EmailRecord) -> Optional[str]:
    return match_named_pattern(NEWSLETTER_PATTERNS, normalize_match_text(record))


def contains_delivery_signal(record: EmailRecord) -> Optional[str]:
    return match_named_pattern(DELIVERY_PATTERNS, normalize_match_text(record))


def contains_delivery_returns_signal(record: EmailRecord) -> Optional[str]:
    return match_named_pattern(DELIVERY_RETURNS_PATTERNS, normalize_match_text(record))


def contains_billing_signal(record: EmailRecord) -> bool:
    return bool(
        re.search(
            r"\b(invoice|receipt|payment|statement|renewal|subscription receipt|tax|vat|billing)\b",
            normalize_match_text(record),
            re.IGNORECASE,
        )
    )


def contains_human_suspicious_signal(record: EmailRecord) -> Optional[str]:
    return match_named_pattern(SUSPICIOUS_HUMAN_PATTERNS, normalize_match_text(record))


def contains_transactional_case_signal(record: EmailRecord) -> Optional[str]:
    text = normalize_match_text(record)
    primary = bool(TRANSACTIONAL_CASE_PRIMARY_PATTERN.search(text))
    context_hits: List[str] = []
    for name, pattern in TRANSACTIONAL_CASE_CONTEXT_PATTERNS:
        if pattern.search(text):
            context_hits.append(name)

    if primary and context_hits:
        return f"{context_hits[0]}"
    if primary and re.search(r"\b(case|application|offer|property|mortgage)\b", text, re.IGNORECASE):
        return "transactional_keywords"
    if len(context_hits) >= 2 and re.search(
        r"\b(case|application|offer|property|mortgage|conveyancing)\b",
        text,
        re.IGNORECASE,
    ):
        return f"{context_hits[0]}+{context_hits[1]}"
    return None


def load_allowlist(path: str) -> Tuple[Set[str], Set[str]]:
    emails: Set[str] = set()
    domains: Set[str] = set()
    if not path or not os.path.exists(path):
        return emails, domains

    with open(path, "r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.split("#", 1)[0].strip().lower()
            if not line:
                continue
            if line.startswith("@"):
                line = line[1:]
            if "@" in line:
                emails.add(line)
            else:
                domains.add(line)
    return emails, domains


def is_promotions_scan(folder: str, gmail_query: Optional[str]) -> bool:
    folder_is_promotions = "promotions" in (folder or "").lower()
    query_is_promotions = DEFAULT_GMAIL_PROMOTIONS_QUERY in (gmail_query or "").lower()
    return folder_is_promotions or query_is_promotions


def evaluate_pre_llm_rules(record: EmailRecord, context: PolicyContext) -> Optional[DecisionResult]:
    display_name, sender_email, sender_domain = extract_sender(record.from_header)
    allowlist_hit = sender_matches_allowlist(
        sender_email=sender_email,
        sender_domain=sender_domain,
        allowlist_emails=context.allowlist_emails,
        allowlist_domains=context.allowlist_domains,
    )
    if allowlist_hit:
        return DecisionResult(
            final_action="KEEP",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match=f"FORCE_KEEP_ALLOWLIST:{allowlist_hit}",
            human_sender=False,
            allowlist_match=True,
        )

    personal_domain = sender_domain in PERSONAL_MAILBOX_DOMAINS
    person_like_name = is_person_like_display_name(display_name)
    no_list_unsubscribe = not has_list_unsubscribe(record)
    if personal_domain and person_like_name and no_list_unsubscribe:
        suspicious = contains_human_suspicious_signal(record)
        if suspicious:
            return DecisionResult(
                final_action="REVIEW",
                model_action="",
                model_confidence=None,
                model_reason="",
                model_raw="",
                decision_source="PRE_RULE",
                rule_match=f"FORCE_REVIEW_HUMAN_SUSPICIOUS:{suspicious}",
                human_sender=True,
                allowlist_match=False,
            )
        return DecisionResult(
            final_action="KEEP",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match="FORCE_KEEP_HUMAN_SENDER",
            human_sender=True,
            allowlist_match=False,
        )

    has_unsubscribe_header = has_list_unsubscribe(record)
    recurring_signal = contains_newsletter_signal(record)
    recurring_or_unsub = bool(recurring_signal) or bool(
        UNSUBSCRIBE_WORD_PATTERN.search(normalize_match_text(record))
    )
    denylist_signal = sender_matches_denylist(sender_domain)
    promotions_delete_mode = (
        context.mode == "aggressive"
        and context.folder_default_policy == "promotions-delete"
        and context.is_promotions_scan
    )
    transactional_case_signal = contains_transactional_case_signal(record)
    if transactional_case_signal:
        return DecisionResult(
            final_action="KEEP",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match=f"FORCE_KEEP_TRANSACTIONAL_CASE:{transactional_case_signal}",
            human_sender=False,
            allowlist_match=False,
        )

    keep_signal = contains_force_keep_signal(record)
    if keep_signal:
        weak_keep = keep_signal in PROMOTIONS_WEAK_KEEP_SIGNALS
        if weak_keep and (recurring_signal or denylist_signal):
            keep_signal = None
        if weak_keep and promotions_delete_mode and has_unsubscribe_header:
            keep_signal = None

    if keep_signal:
        return DecisionResult(
            final_action="KEEP",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match=f"FORCE_KEEP_SIGNAL:{keep_signal}",
            human_sender=False,
            allowlist_match=False,
        )

    delivery_signal = contains_delivery_signal(record)
    if delivery_signal:
        if contains_billing_signal(record):
            return DecisionResult(
                final_action="KEEP",
                model_action="",
                model_confidence=None,
                model_reason="",
                model_raw="",
                decision_source="PRE_RULE",
                rule_match=f"FORCE_KEEP_DELIVERY_WITH_BILLING:{delivery_signal}",
                human_sender=False,
                allowlist_match=False,
            )
        returns_signal = contains_delivery_returns_signal(record)
        if returns_signal:
            return DecisionResult(
                final_action="REVIEW",
                model_action="",
                model_confidence=None,
                model_reason="",
                model_raw="",
                decision_source="PRE_RULE",
                rule_match=f"FORCE_REVIEW_DELIVERY_RETURNS:{returns_signal}",
                human_sender=False,
                allowlist_match=False,
            )
        return DecisionResult(
            final_action="DELETE",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match=f"FORCE_DELETE_DELIVERY_NOTIFICATION:{delivery_signal}",
            human_sender=False,
            allowlist_match=False,
        )

    if has_unsubscribe_header and recurring_or_unsub:
        if promotions_delete_mode:
            return DecisionResult(
                final_action="DELETE",
                model_action="",
                model_confidence=None,
                model_reason="",
                model_raw="",
                decision_source="PRE_RULE",
                rule_match="FORCE_DELETE_RECURRING_LIST_UNSUBSCRIBE_PROMOTIONS",
                human_sender=False,
                allowlist_match=False,
            )
        return DecisionResult(
            final_action="UNSUBSCRIBE",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match="FORCE_UNSUBSCRIBE_RECURRING_LIST_UNSUBSCRIBE",
            human_sender=False,
            allowlist_match=False,
        )

    if denylist_signal:
        if promotions_delete_mode:
            return DecisionResult(
                final_action="DELETE",
                model_action="",
                model_confidence=None,
                model_reason="",
                model_raw="",
                decision_source="PRE_RULE",
                rule_match=f"FORCE_DELETE_DENYLIST_PROMOTIONS:{denylist_signal}",
                human_sender=False,
                allowlist_match=False,
            )
        if has_unsubscribe_header:
            return DecisionResult(
                final_action="UNSUBSCRIBE",
                model_action="",
                model_confidence=None,
                model_reason="",
                model_raw="",
                decision_source="PRE_RULE",
                rule_match=f"FORCE_UNSUBSCRIBE_DENYLIST:{denylist_signal}",
                human_sender=False,
                allowlist_match=False,
            )
        return DecisionResult(
            final_action="DELETE",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match=f"FORCE_DELETE_DENYLIST:{denylist_signal}",
            human_sender=False,
            allowlist_match=False,
        )

    if recurring_signal:
        if promotions_delete_mode:
            return DecisionResult(
                final_action="DELETE",
                model_action="",
                model_confidence=None,
                model_reason="",
                model_raw="",
                decision_source="PRE_RULE",
                rule_match=f"FORCE_DELETE_NEWSLETTER_PROMOTIONS:{recurring_signal}",
                human_sender=False,
                allowlist_match=False,
            )
        if has_unsubscribe_header:
            return DecisionResult(
                final_action="UNSUBSCRIBE",
                model_action="",
                model_confidence=None,
                model_reason="",
                model_raw="",
                decision_source="PRE_RULE",
                rule_match=f"FORCE_UNSUBSCRIBE_NEWSLETTER:{recurring_signal}",
                human_sender=False,
                allowlist_match=False,
            )
        return DecisionResult(
            final_action="DELETE",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match=f"FORCE_DELETE_NEWSLETTER:{recurring_signal}",
            human_sender=False,
            allowlist_match=False,
        )

    if promotions_delete_mode:
        return DecisionResult(
            final_action="DELETE",
            model_action="",
            model_confidence=None,
            model_reason="",
            model_raw="",
            decision_source="PRE_RULE",
            rule_match="FORCE_DELETE_PROMOTIONS_DEFAULT",
            human_sender=False,
            allowlist_match=False,
        )

    return None


def apply_post_llm_overrides(
    record: EmailRecord,
    context: PolicyContext,
    classification: Classification,
) -> Tuple[str, str]:
    final_action = classification.action
    overrides: List[str] = []
    promotions_delete_mode = (
        context.mode == "aggressive"
        and context.folder_default_policy == "promotions-delete"
        and context.is_promotions_scan
    )
    transactional_case_signal = contains_transactional_case_signal(record)
    if transactional_case_signal and final_action in {"DELETE", "UNSUBSCRIBE"}:
        final_action = "KEEP"
        overrides.append(f"transactional_case_keep:{transactional_case_signal}")

    confidence_threshold = 0.85 if context.mode == "aggressive" else 0.75
    if classification.confidence < confidence_threshold and final_action != "REVIEW":
        final_action = "REVIEW"
        overrides.append(f"low_confidence<{confidence_threshold:.2f}")

    keep_signal = contains_force_keep_signal(record)
    if keep_signal and final_action not in {"KEEP", "REVIEW"}:
        final_action = "REVIEW"
        overrides.append(f"safety_keep_signal:{keep_signal}")

    if context.mode == "aggressive":
        newsletter_signal = contains_newsletter_signal(record)
        if final_action == "KEEP" and (has_list_unsubscribe(record) or newsletter_signal):
            if promotions_delete_mode:
                final_action = "DELETE"
                overrides.append("keep_demoted_to_delete_promotions")
            else:
                final_action = "UNSUBSCRIBE"
                overrides.append("keep_demoted_to_unsubscribe")

        if promotions_delete_mode:
            if final_action == "KEEP":
                final_action = "DELETE"
                overrides.append("promotions_default_delete")
            if final_action == "UNSUBSCRIBE":
                final_action = "DELETE"
                overrides.append("promotions_unsubscribe_to_delete")

    if final_action in {"DELETE", "UNSUBSCRIBE"}:
        if contains_delivery_signal(record):
            if contains_billing_signal(record):
                final_action = "KEEP"
                overrides.append("delivery_with_billing_keep")
            elif contains_delivery_returns_signal(record):
                final_action = "REVIEW"
                overrides.append("delivery_returns_review")
            else:
                final_action = "DELETE"
                overrides.append("delivery_notification_delete")

    return final_action, ",".join(overrides)


def decide_message_action(
    record: EmailRecord,
    classifier: BaseClassifier,
    context: PolicyContext,
) -> DecisionResult:
    pre_rule = evaluate_pre_llm_rules(record, context)
    if pre_rule is not None:
        return pre_rule

    classification = classifier.classify(record)
    final_action, post_overrides = apply_post_llm_overrides(
        record=record,
        context=context,
        classification=classification,
    )
    rule_match = "LLM_DECISION"
    if post_overrides:
        rule_match = f"{rule_match}|POST_OVERRIDE:{post_overrides}"
    return DecisionResult(
        final_action=final_action,
        model_action=classification.action,
        model_confidence=classification.confidence,
        model_reason=classification.reason,
        model_raw=classification.raw_response,
        decision_source="LLM",
        rule_match=rule_match,
        human_sender=False,
        allowlist_match=False,
    )


def build_classifier(args: argparse.Namespace) -> BaseClassifier:
    system_prompt = AGGRESSIVE_SYSTEM_PROMPT if args.mode == "aggressive" else NORMAL_SYSTEM_PROMPT
    if args.backend == "openai":
        api_key = os.getenv("OPENAI_API_KEY", "").strip()
        if not api_key:
            raise RuntimeError(
                "Missing OpenAI API key. Run 'python main.py --setup' or set OPENAI_API_KEY in .env.local."
            )
        return OpenAIResponsesClassifier(
            api_key=api_key,
            model=args.openai_model,
            timeout_seconds=args.request_timeout,
            system_prompt=system_prompt,
        )

    return OllamaClassifier(
        base_url=args.ollama_url,
        model=args.ollama_model,
        timeout_seconds=args.request_timeout,
        system_prompt=system_prompt,
    )

def write_audit_header(writer: csv.DictWriter) -> None:
    writer.writeheader()


def cleanup_ai_labels(
    conn: imaplib.IMAP4_SSL,
    folder: str,
    gmail_query: Optional[str],
    uids: List[str],
    dry_run: bool,
    audit_log_path: str,
    processed_label: str,
    include_processed: bool,
) -> None:
    fields = [
        "ts_utc",
        "folder",
        "gmail_query",
        "uid",
        "removed_labels",
        "cleanup_applied",
        "dry_run",
        "error",
    ]
    labels_to_remove = list(LABELS_BY_ACTION.values())
    if include_processed:
        labels_to_remove.append(processed_label)

    with open(audit_log_path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        write_audit_header(writer)

        total = len(uids)
        for index, uid in enumerate(uids, start=1):
            now = dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
            log_row = {
                "ts_utc": now,
                "folder": folder,
                "gmail_query": gmail_query or "",
                "uid": uid,
                "removed_labels": ",".join(labels_to_remove),
                "cleanup_applied": False,
                "dry_run": dry_run,
                "error": "",
            }
            try:
                if not dry_run:
                    for label in labels_to_remove:
                        remove_label(conn, uid, label)
                    log_row["cleanup_applied"] = True
            except Exception as exc:
                log_row["error"] = str(exc)[:400]

            writer.writerow(log_row)
            print(
                f"[{index}/{total}] UID {uid} cleanup "
                f"(removed={log_row['removed_labels']}, applied={log_row['cleanup_applied']})"
            )


def process_mailbox(
    conn: imaplib.IMAP4_SSL,
    classifier: BaseClassifier,
    context: PolicyContext,
    uids: List[str],
    dry_run: bool,
    audit_log_path: str,
    processed_label: str,
) -> None:
    fields = [
        "ts_utc",
        "folder",
        "gmail_query",
        "uid",
        "from",
        "subject",
        "date",
        "message_id",
        "list_unsubscribe",
        "snippet",
        "backend",
        "mode",
        "model_name",
        "decision_source",
        "rule_match",
        "human_sender",
        "allowlist_match",
        "model_action",
        "model_confidence",
        "model_reason",
        "model_raw",
        "final_action",
        "label",
        "label_applied",
        "processed_label",
        "processed_marked",
        "dry_run",
        "error",
    ]

    with open(audit_log_path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        write_audit_header(writer)

        if not dry_run:
            labels_to_ensure = list(LABELS_BY_ACTION.values())
            if processed_label not in labels_to_ensure:
                labels_to_ensure.append(processed_label)
            ensure_labels(conn, labels_to_ensure)

        total = len(uids)
        for index, uid in enumerate(uids, start=1):
            now = dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
            log_row = {
                "ts_utc": now,
                "folder": context.folder,
                "gmail_query": context.gmail_query or "",
                "uid": uid,
                "from": "",
                "subject": "",
                "date": "",
                "message_id": "",
                "list_unsubscribe": "",
                "snippet": "",
                "backend": classifier.backend_name,
                "mode": context.mode,
                "model_name": classifier.model_name,
                "decision_source": "",
                "rule_match": "",
                "human_sender": False,
                "allowlist_match": False,
                "model_action": "",
                "model_confidence": "",
                "model_reason": "",
                "model_raw": "",
                "final_action": "REVIEW",
                "label": LABELS_BY_ACTION["REVIEW"],
                "label_applied": False,
                "processed_label": processed_label,
                "processed_marked": False,
                "dry_run": dry_run,
                "error": "",
            }

            try:
                message_bytes = fetch_message_bytes(conn, uid)
                if not message_bytes:
                    raise RuntimeError("Could not fetch RFC822 payload")

                record = parse_rfc822_message(uid, message_bytes)
                log_row["from"] = record.from_header
                log_row["subject"] = record.subject
                log_row["date"] = record.date
                log_row["message_id"] = record.message_id
                log_row["list_unsubscribe"] = record.list_unsubscribe
                log_row["snippet"] = record.snippet

                decision = decide_message_action(
                    record=record,
                    classifier=classifier,
                    context=context,
                )
                label = LABELS_BY_ACTION[decision.final_action]

                log_row["decision_source"] = decision.decision_source
                log_row["rule_match"] = decision.rule_match
                log_row["human_sender"] = decision.human_sender
                log_row["allowlist_match"] = decision.allowlist_match
                log_row["model_action"] = decision.model_action
                if decision.model_confidence is not None:
                    log_row["model_confidence"] = f"{decision.model_confidence:.4f}"
                log_row["model_reason"] = decision.model_reason
                log_row["model_raw"] = decision.model_raw
                log_row["final_action"] = decision.final_action
                log_row["label"] = label

                if not dry_run:
                    log_row["label_applied"] = set_ai_label(conn, uid, label)
                    if not log_row["label_applied"]:
                        log_row["error"] = "Label write not confirmed by server"
                    else:
                        log_row["processed_marked"] = add_label(conn, uid, processed_label)
                        if not log_row["processed_marked"]:
                            log_row["error"] = "Processed label write not confirmed by server"
            except Exception as exc:
                log_row["error"] = str(exc)[:400]
                log_row["final_action"] = "REVIEW"
                log_row["label"] = LABELS_BY_ACTION["REVIEW"]

            writer.writerow(log_row)
            print(
                f"[{index}/{total}] UID {uid} -> {log_row['final_action']} "
                f"(source={log_row['rule_match'] or 'error'}, label={log_row['label']}, "
                f"applied={log_row['label_applied']}, processed={log_row['processed_marked']})"
            )


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify Gmail folder messages and apply AI labels."
    )
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Interactive first-run setup. Prompts for local settings and writes .env.local.",
    )
    parser.add_argument("--email", default=os.getenv("GMAIL_EMAIL", "").strip(), help="Gmail address")
    parser.add_argument(
        "--app-password",
        default=os.getenv("GMAIL_APP_PASSWORD", "").strip(),
        help="Gmail App Password",
    )
    parser.add_argument("--imap-host", default="imap.gmail.com")
    parser.add_argument("--folder", default=None, help='Mailbox, e.g. "Promotions"')
    parser.add_argument(
        "--gmail-query",
        default=None,
        help='Gmail X-GM-RAW query, e.g. "category:promotions"',
    )
    parser.add_argument("--limit", type=int, default=200, help="Max messages to process")
    parser.add_argument("--dry-run", action="store_true", help="Do not create/apply labels")
    parser.add_argument(
        "--cleanup-ai-labels",
        action="store_true",
        help="Remove AI action labels from matched messages and exit.",
    )
    parser.add_argument(
        "--cleanup-include-processed",
        action="store_true",
        help="With --cleanup-ai-labels, also remove the processed label.",
    )
    parser.add_argument(
        "--processed-label",
        default=DEFAULT_PROCESSED_LABEL,
        help="Label used to mark messages already handled.",
    )
    parser.add_argument(
        "--skip-processed",
        dest="skip_processed",
        action="store_true",
        default=True,
        help="Skip messages that already have --processed-label (default: on).",
    )
    parser.add_argument(
        "--no-skip-processed",
        dest="skip_processed",
        action="store_false",
        help="Reprocess messages even if they already have --processed-label.",
    )
    parser.add_argument(
        "--backend",
        choices=["openai", "ollama"],
        default=os.getenv(ENV_BACKEND_KEY, "ollama").strip().lower() or "ollama",
    )
    parser.add_argument(
        "--openai-model",
        default=os.getenv(ENV_OPENAI_MODEL_KEY, "gpt-4.1-mini").strip() or "gpt-4.1-mini",
    )
    parser.add_argument(
        "--ollama-model",
        default=os.getenv(ENV_OLLAMA_MODEL_KEY, "llama3:latest").strip() or "llama3:latest",
    )
    parser.add_argument(
        "--ollama-url",
        default=os.getenv(ENV_OLLAMA_URL_KEY, "http://127.0.0.1:11434").strip() or "http://127.0.0.1:11434",
    )
    parser.add_argument("--request-timeout", type=int, default=90)
    parser.add_argument("--mode", choices=["normal", "aggressive"], default="aggressive")
    parser.add_argument(
        "--folder-default-policy",
        choices=["none", "promotions-delete"],
        default="promotions-delete",
        help="How to handle Promotions context in aggressive mode.",
    )
    parser.add_argument(
        "--allowlist-path",
        default="allowlist.txt",
        help="Path to allowlist of addresses/domains that must always be kept.",
    )
    parser.add_argument(
        "--audit-log",
        default=f"audit_{dt.datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.csv",
    )
    return parser.parse_args(argv)


def validate_args(args: argparse.Namespace) -> None:
    if not args.email:
        raise RuntimeError(
            "Missing Gmail address. Run 'python main.py --setup' or set GMAIL_EMAIL in .env.local or with --email."
        )
    if not args.app_password:
        raise RuntimeError(
            "Missing Gmail App Password. Run 'python main.py --setup' or set GMAIL_APP_PASSWORD in .env.local or with --app-password."
        )
    if args.limit is not None and args.limit <= 0:
        raise RuntimeError("--limit must be > 0")
    args.processed_label = str(args.processed_label or "").strip()
    if not args.processed_label:
        raise RuntimeError("--processed-label must not be empty")
    if args.processed_label.upper() in {label.upper() for label in LABELS_BY_ACTION.values()}:
        raise RuntimeError("--processed-label must be different from action labels")


def connect_imap(email_address: str, app_password: str, host: str) -> imaplib.IMAP4_SSL:
    conn = imaplib.IMAP4_SSL(host)
    status, _ = conn.login(email_address, app_password)
    if status != "OK":
        raise RuntimeError("Gmail IMAP login failed")
    return conn


def main(argv: List[str]) -> int:
    load_env_file(ENV_FILE_PATH)
    args = parse_args(argv)
    backend_explicit = "--backend" in argv
    if args.setup and not can_prompt_interactively():
        raise RuntimeError("--setup requires an interactive terminal.")
    if should_run_setup_wizard(args, backend_explicit):
        run_setup_wizard(args, backend_explicit=backend_explicit)
        if args.setup:
            print("Setup complete. Run your chosen command next.")
            return 0
    validate_args(args)
    allowlist_emails, allowlist_domains = load_allowlist(args.allowlist_path)
    classifier: Optional[BaseClassifier] = None
    if not args.cleanup_ai_labels:
        classifier = build_classifier(args)

    conn: Optional[imaplib.IMAP4_SSL] = None
    folder = ""
    gmail_query: Optional[str] = None
    search_query: Optional[str] = None
    try:
        conn = connect_imap(args.email, args.app_password, args.imap_host)
        folder, auto_query, resolution = resolve_folder(conn, args.folder)
        gmail_query = args.gmail_query if args.gmail_query else auto_query
        search_query = gmail_query

        if resolution == "promotions_category_query" and not args.gmail_query:
            print(
                "Promotions is not exposed as IMAP folder; using "
                f"folder='{folder}' + gmail_query='{gmail_query}'."
            )

        gmail_ext = has_gmail_ext(conn)
        if args.cleanup_ai_labels:
            cleanup_labels = list(LABELS_BY_ACTION.values())
            if args.cleanup_include_processed:
                cleanup_labels.append(args.processed_label)
            cleanup_label_query = build_any_label_query(cleanup_labels)
            if gmail_ext:
                search_query = combine_gmail_queries(gmail_query, cleanup_label_query)
            elif not gmail_query:
                raise RuntimeError(
                    "--cleanup-ai-labels requires Gmail X-GM-EXT-1 when --gmail-query is not set."
                )
        elif args.skip_processed:
            if gmail_ext:
                search_query = compose_gmail_query(gmail_query, exclude_label=args.processed_label)
            else:
                print(
                    "X-GM-EXT-1 capability missing; --skip-processed cannot be applied for this run."
                )

        select_mailbox(conn, folder, readonly=args.dry_run)
        uids = fetch_uids(conn, args.limit, gmail_query=search_query)

        if args.cleanup_ai_labels:
            print(
                f"Connected. Folder='{folder}', gmail_query='{search_query or 'ALL'}', "
                f"messages={len(uids)}, cleanup_ai_labels=True, "
                f"cleanup_include_processed={args.cleanup_include_processed}, dry_run={args.dry_run}"
            )
            cleanup_ai_labels(
                conn=conn,
                folder=folder,
                gmail_query=search_query,
                uids=uids,
                dry_run=args.dry_run,
                audit_log_path=args.audit_log,
                processed_label=args.processed_label,
                include_processed=args.cleanup_include_processed,
            )
            print(f"Audit log written to: {args.audit_log}")
            return 0

        if classifier is None:
            raise RuntimeError("Internal error: classifier not initialized")

        context = PolicyContext(
            mode=args.mode,
            folder=folder,
            gmail_query=search_query,
            is_promotions_scan=is_promotions_scan(folder, search_query),
            folder_default_policy=args.folder_default_policy,
            allowlist_emails=allowlist_emails,
            allowlist_domains=allowlist_domains,
        )

        print(
            f"Connected. Folder='{folder}', gmail_query='{search_query or 'ALL'}', "
            f"promotions_scan={context.is_promotions_scan}, messages={len(uids)}, "
            f"backend={args.backend}, mode={args.mode}, dry_run={args.dry_run}, "
            f"skip_processed={args.skip_processed}, processed_label='{args.processed_label}'"
        )
        process_mailbox(
            conn=conn,
            classifier=classifier,
            context=context,
            uids=uids,
            dry_run=args.dry_run,
            audit_log_path=args.audit_log,
            processed_label=args.processed_label,
        )
        print(f"Audit log written to: {args.audit_log}")
        return 0
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass
            try:
                conn.logout()
            except Exception:
                pass


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
