"""Standardized event log with correlation IDs for tracing order processing."""

from __future__ import annotations

import json
import logging
import time
import uuid
from pathlib import Path

logger = logging.getLogger(__name__)

_EVENT_LOG_PATH = Path("output/events.jsonl")
_DEAD_LETTER_PATH = Path("output/dead_letter.jsonl")


def new_correlation_id() -> str:
    """Generate a short unique correlation ID."""
    return uuid.uuid4().hex[:12]


def log_event(
    correlation_id: str,
    event: str,
    order_number: str | None = None,
    source_file: str | None = None,
    status: str = "info",
    details: dict | None = None,
) -> None:
    """Append a structured event to the event log.

    Events are written to output/events.jsonl as one JSON object per line.
    """
    record = {
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "cid": correlation_id,
        "event": event,
        "status": status,
    }
    if order_number:
        record["order"] = order_number
    if source_file:
        record["file"] = source_file
    if details:
        record["details"] = details

    _append_jsonl(_EVENT_LOG_PATH, record)

    # Also log to standard logger for console output
    log_fn = logger.error if status == "error" else logger.info
    log_fn("[%s] %s: %s %s", correlation_id, event, status,
           order_number or source_file or "")


def log_dead_letter(
    correlation_id: str,
    source_file: str,
    error: str,
    stage: str,
    order_number: str | None = None,
    email_subject: str | None = None,
    email_sender: str | None = None,
    retry_count: int = 0,
) -> None:
    """Write a failed order to the dead-letter log for retry.

    Dead-letter entries contain enough info to replay the order.
    """
    record = {
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "cid": correlation_id,
        "source_file": source_file,
        "order_number": order_number,
        "stage": stage,
        "error": error,
        "retry_count": retry_count,
        "resolved": False,
    }
    if email_subject:
        record["email_subject"] = email_subject
    if email_sender:
        record["email_sender"] = email_sender

    _append_jsonl(_DEAD_LETTER_PATH, record)
    logger.error(
        "[%s] DEAD LETTER: %s feilet i '%s': %s",
        correlation_id, source_file, stage, error,
    )


def list_dead_letters(unresolved_only: bool = True) -> list[dict]:
    """Read all dead-letter entries."""
    if not _DEAD_LETTER_PATH.exists():
        return []

    entries = []
    for line in _DEAD_LETTER_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        entry = json.loads(line)
        if unresolved_only and entry.get("resolved"):
            continue
        entries.append(entry)
    return entries


def resolve_dead_letter(correlation_id: str) -> bool:
    """Mark a dead-letter entry as resolved by rewriting the file."""
    if not _DEAD_LETTER_PATH.exists():
        return False

    lines = _DEAD_LETTER_PATH.read_text(encoding="utf-8").splitlines()
    found = False
    new_lines = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        entry = json.loads(line)
        if entry.get("cid") == correlation_id:
            entry["resolved"] = True
            found = True
        new_lines.append(json.dumps(entry, ensure_ascii=False))

    if found:
        _DEAD_LETTER_PATH.write_text(
            "\n".join(new_lines) + "\n", encoding="utf-8"
        )
    return found


def list_events(
    correlation_id: str | None = None,
    order_number: str | None = None,
    last_n: int = 50,
) -> list[dict]:
    """Read events, optionally filtered by correlation ID or order number."""
    if not _EVENT_LOG_PATH.exists():
        return []

    entries = []
    for line in _EVENT_LOG_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        entry = json.loads(line)
        if correlation_id and entry.get("cid") != correlation_id:
            continue
        if order_number and entry.get("order") != order_number:
            continue
        entries.append(entry)

    return entries[-last_n:]


def _append_jsonl(path: Path, record: dict) -> None:
    """Append a single JSON record to a JSONL file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")
