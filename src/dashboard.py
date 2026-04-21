"""FastAPI dashboard for Ortopartner order automation."""

from __future__ import annotations

import json
import logging
import re
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from .event_log import list_dead_letters, list_events, resolve_dead_letter

logger = logging.getLogger(__name__)

app = FastAPI(title="Ortopartner Ordreflyt", version="1.0")

# ---------------------------------------------------------------------------
# Poll state tracking (visible to dashboard JS)
# ---------------------------------------------------------------------------

_poll_state: dict = {
    "running": False,
    "last_run": None,        # ISO timestamp of last completed run
    "last_result": None,     # number of orders processed
    "last_error": None,      # error message if last run failed
    "next_scheduled": None,  # ISO timestamp of next scheduled run
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_PROCESSED_FILE = Path("output/processed_emails.json")


def _count_processed() -> int:
    if _PROCESSED_FILE.exists():
        return len(json.loads(_PROCESSED_FILE.read_text(encoding="utf-8")))
    return 0


def _recent_events(n: int = 50) -> list[dict]:
    return list_events(last_n=n)


def _order_stats() -> dict:
    """Compute stats from event log."""
    events = list_events(last_n=1000)
    total = 0
    success = 0
    skipped = 0
    errors = 0
    review = 0

    for e in events:
        if e["event"] == "odoo_push":
            total += 1
            if e["status"] == "success":
                success += 1
            elif e["status"] == "skipped":
                skipped += 1
            elif e["status"] == "error":
                errors += 1
            details = e.get("details", {})
            if details.get("review"):
                review += 1
        if e["event"] == "failed":
            total += 1
            errors += 1

    return {
        "total": total,
        "success": success,
        "skipped": skipped,
        "errors": errors,
        "review": review,
        "dead_letters": len(list_dead_letters()),
        "emails_processed": _count_processed(),
    }


# --- Warning classification helpers ---

def _classify_warning(msg: str) -> str:
    """Return a CSS class / severity tag for a warning message."""
    lower = msg.lower()
    if "enhetspris utledet" in lower:
        return "info"
    if "enhetspris hentet fra produktkatalog" in lower:
        return "info"
    if "enhetspris-avvik" in lower or "rabatt-avvik" in lower:
        return "warn"
    if "kontraktsrabatt" in lower or "fylte inn" in lower:
        return "warn"
    if "ny kunde opprettet" in lower:
        return "warn"
    if "finnes ikke i odoo" in lower:
        return "warn"
    if "mangler" in lower:
        return "warn"
    return "info"


def _aggregate_orders(events: list[dict]) -> list[dict]:
    """Group events by order_number and build a summary per order.

    Returns a list of order-state dicts (most recent first) with:
      order_number, customer, status, so_name, confidence, review,
      warnings, total, line_count, last_ts, cid, events.
    """
    # First pass: build cid → received_at from email_received events.
    # These events have no "order" field so they'd be skipped below,
    # but they carry the actual inbox arrival time via their cid.
    cid_received: dict[str, str] = {}
    for ev in events:
        if ev.get("event") == "email_received":
            ra = (ev.get("details") or {}).get("received_at")
            if ra:
                cid_received[ev["cid"]] = ra

    by_order: dict[str, dict] = {}

    for ev in events:
        order_num = ev.get("order")
        if not order_num:
            continue

        state = by_order.setdefault(order_num, {
            "order_number": order_num,
            "customer": None,
            "status": None,
            "so_name": None,
            "confidence": None,
            "review": False,
            "warnings": [],
            "total_amount": None,
            "currency": "NOK",
            "line_count": None,
            "message": None,
            "last_ts": ev["ts"],
            "email_received_at": cid_received.get(ev["cid"]),
            "cid": ev["cid"],
            "archived": False,
            "events": [],
        })

        state["events"].append(ev)
        if ev["ts"] >= state["last_ts"]:
            state["last_ts"] = ev["ts"]
            state["cid"] = ev["cid"]

        details = ev.get("details") or {}

        if ev["event"] == "pdf_parsed":
            if "confidence" in details and state["confidence"] is None:
                state["confidence"] = details["confidence"]

        if ev["event"] == "odoo_push":
            state["status"] = ev["status"]
            if details.get("so_name"):
                state["so_name"] = details["so_name"]
            if details.get("customer"):
                state["customer"] = details["customer"]
            if "confidence" in details:
                state["confidence"] = details["confidence"]
            if "review" in details:
                state["review"] = bool(details["review"])
            if "line_count" in details:
                state["line_count"] = details["line_count"]
            if "total_amount" in details:
                state["total_amount"] = details["total_amount"]
            if "currency" in details:
                state["currency"] = details["currency"] or "NOK"
            if details.get("warnings"):
                state["warnings"] = list(details["warnings"])
            if details.get("message"):
                state["message"] = details["message"]

        if ev["event"] == "sharepoint_archived":
            state["archived"] = True

        if ev["event"] == "failed":
            state["status"] = "error"
            if details.get("error"):
                state["message"] = details["error"]

    orders = list(by_order.values())
    orders.sort(key=lambda o: o["email_received_at"] or o["last_ts"], reverse=True)
    return orders


# Matches warnings emitted by OdooOrderService._make_order_line for unknown products.
# Handles both the "finnes ikke i Odoo — bruker fallback..." and
# "finnes ikke i Odoo og ingen fallback..." variants, but NOT the transport
# heuristic warning (which contains "gjenkjent som frakt").
_UNKNOWN_PRODUCT_RE = re.compile(
    r"Produkt '([^']+)' finnes ikke i Odoo", re.IGNORECASE
)


def _aggregate_unknown_products(orders: list[dict]) -> list[dict]:
    """Scan aggregated orders for unknown-product warnings and count them.

    Returns a list of unknown-product summaries (most frequent first):
        {sku, count, last_customer, last_order, last_ts, resolution}

    `resolution` is "fallback" if the warning mentions it was mapped to the
    generic fallback product, else "unmapped" (line has no product_id).
    Transport auto-mapping is excluded.
    """
    by_sku: dict[str, dict] = {}

    for o in orders:
        for w in o.get("warnings") or []:
            lower = w.lower()
            # Skip transport lines — those are handled cleanly.
            if "gjenkjent som frakt" in lower:
                continue
            m = _UNKNOWN_PRODUCT_RE.search(w)
            if not m:
                continue
            sku = m.group(1).strip()
            entry = by_sku.setdefault(sku, {
                "sku": sku,
                "count": 0,
                "last_customer": None,
                "last_order": None,
                "last_ts": None,
                "resolution": "unmapped",
            })
            entry["count"] += 1
            if entry["last_ts"] is None or (o.get("last_ts") and o["last_ts"] > entry["last_ts"]):
                entry["last_ts"] = o.get("last_ts")
                entry["last_customer"] = o.get("customer")
                entry["last_order"] = o.get("order_number")
            if "bruker fallback-produkt" in lower:
                entry["resolution"] = "fallback"

    result = list(by_sku.values())
    result.sort(key=lambda e: (-e["count"], e["sku"]))
    return result


# ---------------------------------------------------------------------------
# HTML template (inline to keep it simple)
# ---------------------------------------------------------------------------

def _format_amount(amount, currency: str = "NOK") -> str:
    if amount is None:
        return "-"
    try:
        val = float(amount)
    except (TypeError, ValueError):
        return str(amount)
    formatted = f"{val:,.2f}".replace(",", " ").replace(".", ",")
    return f"{formatted} {currency}"


def _render_order_warnings_block(warnings: list[str]) -> str:
    """Render warnings as a grouped, color-coded list (like Odoo message log)."""
    if not warnings:
        return '<div class="no-warnings">Ingen advarsler</div>'

    info = [w for w in warnings if _classify_warning(w) == "info"]
    warn = [w for w in warnings if _classify_warning(w) == "warn"]

    parts = []
    if warn:
        parts.append('<div class="warning-group warn">')
        parts.append('<div class="warning-group-title">Krever oppmerksomhet</div>')
        parts.append('<ul>')
        for w in warn:
            parts.append(f'<li>{_escape(w)}</li>')
        parts.append('</ul>')
        parts.append('</div>')
    if info:
        parts.append('<div class="warning-group info">')
        parts.append('<div class="warning-group-title">Informasjon</div>')
        parts.append('<ul>')
        for w in info:
            parts.append(f'<li>{_escape(w)}</li>')
        parts.append('</ul>')
        parts.append('</div>')
    return "".join(parts)


def _escape(s: str) -> str:
    """Minimal HTML escape."""
    if s is None:
        return ""
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def _render_dashboard(stats: dict, events: list[dict], dead_letters: list[dict]) -> str:
    from .config import is_test_mode, get_test_prefix
    test_mode_banner = ""
    if is_test_mode():
        test_mode_banner = (
            f'<div class="test-banner">TEST-MODUS AKTIV '
            f'&mdash; ordrenr prefikses med <code>{get_test_prefix()}</code></div>'
        )

    # Aggregate events by order for the "Siste ordrer" section.
    # We also use a broader window for the unknown-products summary
    # so the list reflects patterns over time, not just the last 50 events.
    all_orders = _aggregate_orders(list_events(last_n=500))
    unknown_products = _aggregate_unknown_products(all_orders)

    orders = all_orders[:50]

    order_rows = ""
    for o in orders:
        status = o["status"] or "pending"
        status_label = {
            "success": "OK",
            "skipped": "Duplikat",
            "error": "Feil",
            "pending": "Venter",
        }.get(status, status)
        status_class = {
            "success": "badge-success",
            "skipped": "badge-skipped",
            "error": "badge-error",
        }.get(status, "badge-pending")

        # Determine filter category: ok, review, error, skipped
        filter_status = status  # success, skipped, error, pending
        if o["review"] and status == "success":
            filter_status = "review"

        review_badge = ""
        if o["review"]:
            review_badge = '<span class="badge badge-review">REVIEW</span>'

        archived_badge = ""
        if o["archived"]:
            archived_badge = '<span class="badge badge-archived" title="Arkivert i SharePoint">SP</span>'

        confidence = ""
        if o["confidence"] is not None:
            confidence = f'{o["confidence"]*100:.0f}%'

        warn_count = len(o["warnings"])
        warn_info_count = sum(1 for w in o["warnings"] if _classify_warning(w) == "info")
        warn_attention_count = warn_count - warn_info_count
        warn_summary = ""
        if warn_attention_count:
            warn_summary += f'<span class="warn-count warn">{warn_attention_count} advarsel</span> '
        if warn_info_count:
            warn_summary += f'<span class="warn-count info">{warn_info_count} info</span>'
        if not warn_summary:
            warn_summary = '<span class="warn-count none">&mdash;</span>'

        warnings_block = _render_order_warnings_block(o["warnings"])

        customer = _escape(o["customer"] or "-")
        so = _escape(o["so_name"] or "-")
        order_num = _escape(o["order_number"])
        total = _format_amount(o["total_amount"], o["currency"] or "NOK")
        line_count = o["line_count"] if o["line_count"] is not None else "-"
        message = _escape(o["message"] or "")
        # Email received timestamp — prefer actual inbox arrival time over processing time
        email_ra = o.get("email_received_at")  # e.g. "2026-04-21T07:35:29Z"
        if email_ra:
            # Normalize: strip trailing Z, replace T with space
            email_ra_norm = email_ra.rstrip("Z").replace("T", " ")
            sort_ts = email_ra_norm          # ISO-sortable for data-date
            ts_date = email_ra_norm[:10]     # YYYY-MM-DD for time filter
            ts_display = f'{email_ra_norm[8:10]}.{email_ra_norm[5:7]} {email_ra_norm[11:16]}'
        else:
            last = o.get("last_ts", "")
            ts_date = last[:10]
            sort_ts = last
            ts_display = last[11:16] if len(last) > 16 else last[11:19]

        is_test = "true" if o["order_number"] and o["order_number"].upper().startswith("TEST-") else "false"

        order_rows += f'''
        <details class="order-row" data-status="{filter_status}" data-date="{ts_date}" data-is-test="{is_test}" data-customer="{_escape(o['customer'] or '')}">
            <summary>
                <span class="ordre-col">
                    <code>{order_num}</code>
                    <span class="badge {status_class}">{status_label}</span>
                    {review_badge}
                    {archived_badge}
                </span>
                <span class="customer-col">{customer}</span>
                <span class="so-col">{so}</span>
                <span class="conf-col">{confidence}</span>
                <span class="lines-col">{line_count}</span>
                <span class="total-col">{total}</span>
                <span class="warn-col">{warn_summary}</span>
                <span class="ts-col" title="{email_ra or o.get('last_ts','')}">{ts_display}</span>
            </summary>
            <div class="order-detail">
                <div class="order-meta">
                    <div><strong>Ordrenr:</strong> <code>{order_num}</code></div>
                    <div><strong>Kunde:</strong> {customer}</div>
                    <div><strong>SO:</strong> {so}</div>
                    <div><strong>Konfidensverdi:</strong> {confidence or "-"}</div>
                    <div><strong>Linjer:</strong> {line_count}</div>
                    <div><strong>Totalbeløp:</strong> {total}</div>
                    <div><strong>Sist oppdatert:</strong> {o["last_ts"]}</div>
                    <div><strong>CID:</strong> <code>{o["cid"]}</code></div>
                </div>
                {"<div class='order-message'>" + message + "</div>" if message else ""}
                {warnings_block}
            </div>
        </details>'''

    # Render unknown-product rows
    unknown_rows = ""
    for u in unknown_products:
        resolution_label = {
            "fallback": '<span class="badge badge-skipped">Fallback</span>',
            "unmapped": '<span class="badge badge-error">Ukoblet</span>',
        }.get(u["resolution"], "")
        last_seen = u["last_ts"][:10] if u.get("last_ts") else "-"
        last_customer = _escape(u.get("last_customer") or "-")
        last_order = _escape(u.get("last_order") or "-")
        unknown_rows += f"""<tr>
            <td><code>{_escape(u['sku'])}</code></td>
            <td class="uk-count">{u['count']}</td>
            <td>{resolution_label}</td>
            <td>{last_customer}</td>
            <td><code>{last_order}</code></td>
            <td class="uk-ts">{last_seen}</td>
        </tr>"""

    dl_rows = ""
    for dl in dead_letters:
        dl_rows += f"""<tr>
            <td><code>{dl['cid']}</code></td>
            <td>{dl['ts']}</td>
            <td>{dl['source_file']}</td>
            <td>{dl.get('order_number', '-')}</td>
            <td>{dl['stage']}</td>
            <td class="err">{dl['error'][:80]}</td>
            <td><form method="post" action="/api/replay/{dl['source_file']}">
                <button type="submit" class="btn btn-sm">Re-kjor</button>
            </form></td>
        </tr>"""

    event_rows = ""
    for ev in reversed(events[-30:]):
        order = ev.get("order", "")
        status_class = {"error": "err", "ok": "ok"}.get(ev["status"], "")
        details = ""
        if ev.get("details"):
            d = ev["details"]
            if "confidence" in d:
                details = f"{d['confidence']:.0%}"
            elif "so_name" in d:
                details = d["so_name"] or ""
            elif "error" in d:
                details = d["error"][:60]
            elif "files" in d:
                details = f"{d['files']} filer"
        event_rows += f"""<tr>
            <td>{ev['ts'][:10]} {ev['ts'][11:19]}</td>
            <td><code>{ev['cid']}</code></td>
            <td>{ev['event']}</td>
            <td class="{status_class}">{ev['status']}</td>
            <td>{order}</td>
            <td>{details}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html lang="no">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Ortopartner Ordreflyt</title>
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #e2e8f0;
        min-height: 100vh;
    }}

    /* --- Top navigation bar --- */
    .topbar {{
        background: rgba(15, 23, 42, 0.8);
        backdrop-filter: blur(12px);
        border-bottom: 1px solid rgba(148, 163, 184, 0.1);
        padding: 16px 32px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        position: sticky;
        top: 0;
        z-index: 100;
    }}
    .topbar-left {{
        display: flex;
        align-items: center;
        gap: 12px;
    }}
    .topbar-logo {{
        width: 36px;
        height: 36px;
        background: linear-gradient(135deg, #3b82f6, #8b5cf6);
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 18px;
        font-weight: 700;
        color: #fff;
    }}
    .topbar h1 {{
        font-size: 18px;
        font-weight: 700;
        color: #f1f5f9;
        letter-spacing: -0.3px;
    }}
    .topbar .subtitle {{
        font-size: 12px;
        color: #64748b;
        margin-top: 1px;
    }}
    .topbar-actions {{
        display: flex;
        gap: 8px;
        align-items: center;
    }}
    .topbar-actions .live-dot {{
        width: 8px;
        height: 8px;
        background: #22c55e;
        border-radius: 50%;
        animation: pulse 2s ease-in-out infinite;
        margin-right: 4px;
    }}
    @keyframes pulse {{
        0%, 100% {{ opacity: 1; }}
        50% {{ opacity: 0.4; }}
    }}
    @keyframes spin {{
        to {{ transform: rotate(360deg); }}
    }}
    .topbar-actions span.live-label {{
        font-size: 12px;
        color: #22c55e;
        font-weight: 500;
        margin-right: 16px;
    }}
    .poll-status {{
        display: flex;
        align-items: center;
        gap: 8px;
        margin-right: 12px;
    }}
    .poll-status .spinner {{
        width: 16px;
        height: 16px;
        border: 2px solid rgba(96, 165, 250, 0.3);
        border-top-color: #60a5fa;
        border-radius: 50%;
        animation: spin 0.8s linear infinite;
        display: none;
    }}
    .poll-status.polling .spinner {{ display: block; }}
    .poll-status .poll-text {{
        font-size: 11px;
        color: #64748b;
        max-width: 200px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }}
    .poll-status.polling .poll-text {{ color: #60a5fa; }}
    .poll-btn {{
        background: linear-gradient(135deg, #3b82f6, #6366f1);
        color: #fff;
        border: none;
        padding: 10px 20px;
        border-radius: 10px;
        cursor: pointer;
        font-size: 13px;
        font-weight: 600;
        transition: all 0.15s;
        letter-spacing: 0.2px;
    }}
    .poll-btn:hover {{
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.4);
    }}
    .poll-btn:disabled {{
        opacity: 0.5;
        cursor: not-allowed;
        transform: none;
        box-shadow: none;
    }}
    .schedule-info {{
        font-size: 10px;
        color: #475569;
        margin-top: 2px;
    }}

    .container {{
        max-width: 1400px;
        margin: 0 auto;
        padding: 24px 32px;
    }}

    /* --- Test banner --- */
    .test-banner {{
        background: linear-gradient(90deg, rgba(234, 179, 8, 0.15), rgba(234, 179, 8, 0.05));
        color: #fbbf24;
        padding: 12px 20px;
        border-radius: 12px;
        margin-bottom: 20px;
        font-weight: 600;
        border: 1px solid rgba(234, 179, 8, 0.3);
        font-size: 14px;
    }}
    .test-banner code {{
        background: rgba(234, 179, 8, 0.2);
        color: #fbbf24;
        padding: 2px 8px;
        border-radius: 4px;
    }}

    /* --- Stat cards --- */
    .cards {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
        gap: 16px;
        margin-bottom: 28px;
    }}
    .card {{
        background: rgba(30, 41, 59, 0.7);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(148, 163, 184, 0.1);
        border-radius: 16px;
        padding: 20px;
        text-align: center;
        transition: transform 0.15s, border-color 0.15s;
        position: relative;
        overflow: hidden;
    }}
    .card::before {{
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 3px;
        border-radius: 16px 16px 0 0;
    }}
    .card:hover {{
        transform: translateY(-2px);
        border-color: rgba(148, 163, 184, 0.25);
    }}
    .card .num {{
        font-size: 36px;
        font-weight: 700;
        line-height: 1;
        margin-bottom: 6px;
    }}
    .card .label {{
        font-size: 12px;
        color: #94a3b8;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    .card.blue .num {{ color: #60a5fa; }}
    .card.blue::before {{ background: linear-gradient(90deg, #3b82f6, #60a5fa); }}
    .card.green .num {{ color: #4ade80; }}
    .card.green::before {{ background: linear-gradient(90deg, #22c55e, #4ade80); }}
    .card.yellow .num {{ color: #fbbf24; }}
    .card.yellow::before {{ background: linear-gradient(90deg, #eab308, #fbbf24); }}
    .card.red .num {{ color: #f87171; }}
    .card.red::before {{ background: linear-gradient(90deg, #ef4444, #f87171); }}

    /* --- Buttons --- */
    .btn {{
        background: linear-gradient(135deg, #3b82f6, #6366f1);
        color: #fff;
        border: none;
        padding: 10px 20px;
        border-radius: 10px;
        cursor: pointer;
        font-size: 13px;
        font-weight: 600;
        transition: all 0.15s;
        letter-spacing: 0.2px;
    }}
    .btn:hover {{
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.4);
    }}
    .btn-sm {{
        padding: 5px 12px;
        font-size: 11px;
        border-radius: 6px;
    }}

    /* --- Filter bar --- */
    .filter-bar {{
        display: flex;
        gap: 12px;
        margin-bottom: 16px;
        flex-wrap: wrap;
        align-items: center;
    }}
    .filter-group {{
        display: flex;
        gap: 2px;
        background: rgba(30, 41, 59, 0.5);
        border-radius: 10px;
        padding: 3px;
        border: 1px solid rgba(148, 163, 184, 0.1);
    }}
    .filter-group label {{
        font-size: 11px;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 6px 10px;
        margin-right: 4px;
        align-self: center;
    }}
    .filter-btn {{
        padding: 6px 14px;
        border: none;
        background: transparent;
        color: #94a3b8;
        font-size: 12px;
        font-weight: 500;
        cursor: pointer;
        border-radius: 8px;
        transition: all 0.15s;
        white-space: nowrap;
    }}
    .filter-btn:hover {{
        color: #e2e8f0;
        background: rgba(148, 163, 184, 0.1);
    }}
    .filter-btn.active {{
        background: linear-gradient(135deg, #3b82f6, #6366f1);
        color: #fff;
        font-weight: 600;
    }}

    /* --- Sections --- */
    section {{
        background: rgba(30, 41, 59, 0.6);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(148, 163, 184, 0.08);
        border-radius: 16px;
        padding: 20px;
        margin-bottom: 20px;
    }}
    section h2 {{
        font-size: 15px;
        font-weight: 600;
        color: #f1f5f9;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }}
    section h2 .section-icon {{
        width: 24px;
        height: 24px;
        border-radius: 6px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 13px;
    }}

    /* --- Tables --- */
    table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }}
    th {{
        text-align: left;
        padding: 8px 10px;
        border-bottom: 1px solid rgba(148, 163, 184, 0.1);
        color: #64748b;
        font-weight: 600;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    td {{
        padding: 8px 10px;
        border-bottom: 1px solid rgba(148, 163, 184, 0.05);
        color: #cbd5e1;
    }}
    tr:hover td {{
        background: rgba(148, 163, 184, 0.03);
    }}
    code {{
        background: rgba(148, 163, 184, 0.1);
        padding: 2px 7px;
        border-radius: 4px;
        font-size: 11px;
        color: #93c5fd;
        font-family: 'JetBrains Mono', 'Fira Code', monospace;
    }}
    .err {{ color: #f87171; }}
    .ok {{ color: #4ade80; }}
    .empty {{ color: #64748b; padding: 30px; text-align: center; }}

    /* --- Order list section --- */
    .orders-header, .order-row summary {{
        display: grid;
        grid-template-columns: 2.4fr 2fr 1fr 0.7fr 0.5fr 1.3fr 1.3fr 0.8fr;
        gap: 8px;
        padding: 12px 16px;
        align-items: center;
        font-size: 13px;
    }}
    .orders-header {{
        background: rgba(148, 163, 184, 0.05);
        color: #64748b;
        font-weight: 600;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        border-radius: 12px 12px 0 0;
        border-bottom: 1px solid rgba(148, 163, 184, 0.08);
    }}
    .order-row {{
        border-bottom: 1px solid rgba(148, 163, 184, 0.05);
    }}
    .order-row:last-child {{ border-bottom: none; }}
    .order-row summary {{
        cursor: pointer;
        list-style: none;
        transition: background 0.15s;
        border-radius: 4px;
    }}
    .order-row summary::-webkit-details-marker {{ display: none; }}
    .order-row summary:hover {{ background: rgba(148, 163, 184, 0.05); }}
    .order-row[open] summary {{ background: rgba(59, 130, 246, 0.08); }}
    .order-row.hidden-by-filter {{ display: none; }}

    .ordre-col {{ display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }}
    .customer-col {{ color: #cbd5e1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    .so-col {{ font-family: 'JetBrains Mono', monospace; color: #60a5fa; font-size: 12px; }}
    .conf-col, .lines-col {{ text-align: center; color: #94a3b8; }}
    .total-col {{ text-align: right; font-variant-numeric: tabular-nums; color: #f1f5f9; font-weight: 600; }}
    .warn-col {{ text-align: left; }}
    .ts-col {{ text-align: right; color: #64748b; font-family: 'JetBrains Mono', monospace; font-size: 11px; }}

    /* --- Badges --- */
    .badge {{
        display: inline-block;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 10px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }}
    .badge-success {{ background: rgba(34, 197, 94, 0.15); color: #4ade80; border: 1px solid rgba(34, 197, 94, 0.3); }}
    .badge-skipped {{ background: rgba(234, 179, 8, 0.15); color: #fbbf24; border: 1px solid rgba(234, 179, 8, 0.3); }}
    .badge-error {{ background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }}
    .badge-pending {{ background: rgba(148, 163, 184, 0.15); color: #94a3b8; border: 1px solid rgba(148, 163, 184, 0.3); }}
    .badge-review {{ background: rgba(251, 146, 60, 0.15); color: #fb923c; border: 1px solid rgba(251, 146, 60, 0.3); }}
    .badge-archived {{ background: rgba(96, 165, 250, 0.15); color: #60a5fa; border: 1px solid rgba(96, 165, 250, 0.3); }}

    .warn-count {{
        display: inline-block;
        padding: 2px 8px;
        border-radius: 6px;
        font-size: 11px;
        font-weight: 600;
        margin-right: 3px;
    }}
    .warn-count.warn {{ background: rgba(251, 146, 60, 0.15); color: #fb923c; }}
    .warn-count.info {{ background: rgba(96, 165, 250, 0.15); color: #60a5fa; }}
    .warn-count.none {{ color: #475569; }}

    /* --- Order detail panel --- */
    .order-detail {{
        padding: 20px 24px;
        background: rgba(15, 23, 42, 0.5);
        border-top: 1px solid rgba(148, 163, 184, 0.08);
        border-radius: 0 0 8px 8px;
    }}
    .order-meta {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 8px 20px;
        margin-bottom: 16px;
        font-size: 12px;
        color: #94a3b8;
    }}
    .order-meta strong {{ color: #cbd5e1; margin-right: 4px; }}
    .order-message {{
        background: rgba(59, 130, 246, 0.08);
        border-left: 3px solid #3b82f6;
        padding: 10px 14px;
        margin-bottom: 14px;
        font-size: 12px;
        color: #94a3b8;
        border-radius: 0 8px 8px 0;
    }}

    .warning-group {{
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
        border-left: 4px solid #475569;
    }}
    .warning-group.warn {{
        border-left-color: #f59e0b;
        background: rgba(245, 158, 11, 0.06);
    }}
    .warning-group.info {{
        border-left-color: #3b82f6;
        background: rgba(59, 130, 246, 0.06);
    }}
    .warning-group-title {{
        font-weight: 700;
        font-size: 11px;
        text-transform: uppercase;
        color: #64748b;
        margin-bottom: 8px;
        letter-spacing: 0.5px;
    }}
    .warning-group.warn .warning-group-title {{ color: #fbbf24; }}
    .warning-group.info .warning-group-title {{ color: #60a5fa; }}
    .warning-group ul {{
        margin: 0;
        padding-left: 18px;
        font-size: 12px;
        line-height: 1.7;
    }}
    .warning-group li {{ color: #94a3b8; }}
    .no-warnings {{
        color: #475569;
        font-size: 12px;
        font-style: italic;
        padding: 6px 0;
    }}

    .orders-section {{ padding: 0; }}
    .orders-section h2 {{ padding: 20px 20px 14px; }}
    .orders-empty {{ padding: 40px; text-align: center; color: #475569; font-size: 14px; }}

    /* --- Unknown products table --- */
    .unknown-products table {{ margin-top: 4px; }}
    .unknown-products td.uk-count {{
        font-weight: 700;
        color: #fb923c;
        text-align: center;
        font-variant-numeric: tabular-nums;
    }}
    .unknown-products td.uk-ts {{
        color: #64748b;
        font-family: 'JetBrains Mono', monospace;
        font-size: 11px;
    }}
    .unknown-products .uk-hint {{
        margin-top: 14px;
        font-size: 11px;
        color: #64748b;
        font-style: italic;
        border-top: 1px solid rgba(148, 163, 184, 0.1);
        padding-top: 10px;
    }}
    .unknown-products .uk-hint strong {{ color: #94a3b8; }}

    /* --- Collapsible event log --- */
    .events-toggle {{
        cursor: pointer;
        user-select: none;
    }}
    .events-toggle .chevron {{
        display: inline-block;
        transition: transform 0.2s;
        font-size: 12px;
        color: #64748b;
    }}
    details[open] .events-toggle .chevron {{
        transform: rotate(90deg);
    }}

    /* --- Filter result counter --- */
    .filter-result-count {{
        font-size: 12px;
        color: #64748b;
        margin-left: auto;
        font-weight: 500;
    }}

    /* --- Filter banner --- */
    .filter-banner {{
        background: rgba(59, 130, 246, 0.08);
        border: 1px solid rgba(59, 130, 246, 0.2);
        border-radius: 10px;
        padding: 10px 16px;
        margin-bottom: 20px;
        font-size: 12px;
        color: #93c5fd;
        display: flex;
        align-items: center;
        gap: 8px;
    }}

    /* --- View toggle buttons --- */
    .view-btns {{
        display: flex;
        gap: 2px;
        background: rgba(30, 41, 59, 0.5);
        border-radius: 10px;
        padding: 3px;
        border: 1px solid rgba(148, 163, 184, 0.1);
        margin-left: 4px;
    }}
    .view-btn {{
        padding: 6px 12px;
        border: none;
        background: transparent;
        color: #64748b;
        font-size: 15px;
        cursor: pointer;
        border-radius: 8px;
        transition: all 0.15s;
    }}
    .view-btn:hover {{ color: #e2e8f0; background: rgba(148, 163, 184, 0.1); }}
    .view-btn.active {{ background: linear-gradient(135deg, #3b82f6, #6366f1); color: #fff; }}

    /* --- Insights view --- */
    .insights-grid {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 14px;
        margin-bottom: 24px;
    }}
    .insight-card {{
        background: rgba(15, 23, 42, 0.5);
        border: 1px solid rgba(148, 163, 184, 0.08);
        border-radius: 14px;
        padding: 18px 20px;
    }}
    .insight-title {{
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #64748b;
        margin-bottom: 10px;
    }}
    .insight-big {{
        font-size: 34px;
        font-weight: 700;
        color: #f1f5f9;
        line-height: 1;
        margin-bottom: 6px;
    }}
    .insight-sub {{ font-size: 11px; color: #475569; }}

    .insight-section-title {{
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #64748b;
        margin-bottom: 12px;
    }}

    /* --- Three-col insights layout --- */
    .insights-two-col {{
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 24px;
    }}
    @media (max-width: 1100px) {{
        .insights-two-col {{ grid-template-columns: 1fr 1fr; }}
    }}
    @media (max-width: 700px) {{
        .insights-two-col {{ grid-template-columns: 1fr; }}
    }}

    /* --- Customer bars --- */
    .customer-bars {{ display: flex; flex-direction: column; gap: 8px; }}
    .cbar {{
        display: grid;
        grid-template-columns: 150px 1fr 48px;
        align-items: center;
        gap: 10px;
        font-size: 12px;
    }}
    .cbar-name {{ color: #cbd5e1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    .cbar-track {{ background: rgba(148, 163, 184, 0.08); border-radius: 4px; height: 8px; overflow: hidden; }}
    .cbar-fill {{ height: 100%; border-radius: 4px; background: linear-gradient(90deg, #3b82f6, #6366f1); transition: width 0.4s; }}
    .cbar-fill.conf-high  {{ background: linear-gradient(90deg, #22c55e, #4ade80); }}
    .cbar-fill.conf-mid   {{ background: linear-gradient(90deg, #eab308, #fbbf24); }}
    .cbar-fill.conf-low   {{ background: linear-gradient(90deg, #ef4444, #f87171); }}
    .cbar-count {{ text-align: right; color: #64748b; font-variant-numeric: tabular-nums; }}
</style>
</head>
<body>

<div class="topbar">
    <div class="topbar-left">
        <div class="topbar-logo">O</div>
        <div>
            <h1>Ortopartner Ordreflyt</h1>
            <div class="subtitle">E-post &rarr; PDF &rarr; Odoo &rarr; SharePoint</div>
        </div>
    </div>
    <div class="topbar-actions">
        <div class="poll-status" id="pollStatus">
            <div class="spinner"></div>
            <div>
                <div class="poll-text" id="pollText"></div>
                <div class="schedule-info" id="scheduleInfo"></div>
            </div>
        </div>
        <span class="live-dot"></span>
        <span class="live-label">Live</span>
        <button class="poll-btn" id="pollBtn" onclick="triggerPoll()">Sjekk e-post nå</button>
    </div>
</div>

<div class="container">

{test_mode_banner}

<div class="cards">
    <div class="card blue"><div class="num" id="sc-emails">{stats['emails_processed']}</div><div class="label">E-poster</div></div>
    <div class="card green"><div class="num" id="sc-ok" data-server="{stats['success']}">{stats['success']}</div><div class="label">Ordrer OK</div></div>
    <div class="card yellow"><div class="num" id="sc-review" data-server="{stats['review']}">{stats['review']}</div><div class="label">Til review</div></div>
    <div class="card yellow"><div class="num" id="sc-dup" data-server="{stats['skipped']}">{stats['skipped']}</div><div class="label">Duplikater</div></div>
    <div class="card red"><div class="num" id="sc-err" data-server="{stats['dead_letters']}">{stats['dead_letters']}</div><div class="label">Feilede</div></div>
</div>
<div id="filterBanner" class="filter-banner" style="display:none;"></div>

{"" if not dead_letters else f'''<section>
<h2><span class="section-icon" style="background:rgba(239,68,68,0.15);">!</span> Feilede ordrer</h2>
<table>
<tr><th>CID</th><th>Tid</th><th>Fil</th><th>Ordre</th><th>Steg</th><th>Feil</th><th></th></tr>
{dl_rows}
</table>
</section>'''}

{"" if not unknown_products else f'''<section class="unknown-products">
<h2><span class="section-icon" style="background:rgba(251,146,60,0.15);">?</span> Ukjente produkter <span style="font-size:11px; color:#64748b; font-weight:400; margin-left:4px;">({len(unknown_products)} SKU-er mangler i Odoo)</span></h2>
<table>
<tr><th>Artikkelnr.</th><th>Antall</th><th>Handling</th><th>Sist sett hos</th><th>Siste ordre</th><th>Dato</th></tr>
{unknown_rows}
</table>
<p class="uk-hint">Opprett disse produktene i Odoo for automatisk behandling. Linjer merket <strong>Ukoblet</strong> mangler produktkobling i SO-en.</p>
</section>'''}

<section class="orders-section">
<h2><span class="section-icon" style="background:rgba(59,130,246,0.15);">&#9776;</span> Ordrer <span style="font-size:11px; color:#64748b; font-weight:400; margin-left:4px;">(klikk for detaljer)</span></h2>
<div style="padding: 0 20px 14px;">
    <div class="filter-bar">
        <div class="filter-group">
            <label>Periode</label>
            <button class="filter-btn active" onclick="setFilter('time', 'all', this)">Alle</button>
            <button class="filter-btn" onclick="setFilter('time', 'today', this)">I dag</button>
            <button class="filter-btn" onclick="setFilter('time', 'week', this)">Denne uken</button>
            <button class="filter-btn" onclick="setFilter('time', 'month', this)">Denne mnd</button>
        </div>
        <div class="filter-group">
            <label>Status</label>
            <button class="filter-btn active" onclick="setFilter('status', 'all', this)">Alle</button>
            <button class="filter-btn" onclick="setFilter('status', 'success', this)">OK</button>
            <button class="filter-btn" onclick="setFilter('status', 'review', this)">Review</button>
            <button class="filter-btn" onclick="setFilter('status', 'error', this)">Feil</button>
            <button class="filter-btn" onclick="setFilter('status', 'skipped', this)">Duplikat</button>
        </div>
        <div class="filter-group">
            <label>Vis</label>
            <button class="filter-btn active" onclick="setFilter('scope', 'all', this)">Alle</button>
            <button class="filter-btn" onclick="setFilter('scope', 'prod', this)">Produksjon</button>
            <button class="filter-btn" onclick="setFilter('scope', 'test', this)">Test</button>
        </div>
        <div class="view-btns">
            <button class="view-btn active" onclick="setView('list', this)" title="Liste">&#9776;</button>
            <button class="view-btn" onclick="setView('insights', this)" title="Innsikt">&#9783;</button>
        </div>
        <span class="filter-result-count" id="filterCount"></span>
    </div>
</div>
{f'<div class="orders-empty">Ingen ordrer behandlet enda.</div>' if not orders else f'''
<div class="orders-header">
    <span>Ordre / Status</span>
    <span>Kunde</span>
    <span>SO</span>
    <span>Konfid.</span>
    <span>Linjer</span>
    <span style="text-align:right;">Total</span>
    <span>Advarsler</span>
    <span style="text-align:right;">Mottatt</span>
</div>
<div id="orderList">
{order_rows}
</div>
<div id="insightsView" style="display:none; padding: 0 20px 20px;">
    <div class="insights-grid">
        <div class="insight-card">
            <div class="insight-title">Suksessrate</div>
            <div class="insight-big" id="ins-rate">—</div>
            <div class="insight-sub" id="ins-rate-sub"></div>
        </div>
        <div class="insight-card">
            <div class="insight-title">Snitt konfidensverdi</div>
            <div class="insight-big" id="ins-conf">—</div>
            <div class="insight-sub">Per behandlet ordre</div>
        </div>
        <div class="insight-card">
            <div class="insight-title">Ordrer til review</div>
            <div class="insight-big" id="ins-review">—</div>
            <div class="insight-sub">Krever manuell sjekk</div>
        </div>
        <div class="insight-card">
            <div class="insight-title">Test vs produksjon</div>
            <div class="insight-big" id="ins-testprod">—</div>
            <div class="insight-sub" id="ins-testprod-sub"></div>
        </div>
    </div>
    <div class="insights-two-col">
        <div>
            <div class="insight-section-title">Ordrer per kunde <span id="ins-scope-label"></span></div>
            <div id="ins-customers" class="customer-bars"></div>
        </div>
        <div>
            <div class="insight-section-title">Konfidens per kunde</div>
            <div id="ins-confidence" class="customer-bars"></div>
        </div>
        <div>
            <div class="insight-section-title">Til review per kunde</div>
            <div id="ins-review-customers" class="customer-bars"></div>
        </div>
    </div>
</div>
'''}
</section>

<section>
<details>
    <summary class="events-toggle">
        <h2 style="display:inline-flex; align-items:center; gap:8px; cursor:pointer;">
            <span class="section-icon" style="background:rgba(148,163,184,0.1);">&#8862;</span>
            Hendelseslogg
            <span class="chevron">&#9654;</span>
            <span style="font-size:11px; color:#64748b; font-weight:400;">({len(events)} hendelser)</span>
        </h2>
    </summary>
    {"<p class='empty'>Ingen hendelser enda.</p>" if not events else f'''<table style="margin-top:12px;">
<tr><th>Tid</th><th>CID</th><th>Hendelse</th><th>Status</th><th>Ordre</th><th>Detaljer</th></tr>
{event_rows}
</table>'''}
</details>
</section>

</div><!-- container -->

<script>
// ── Filter + view state ───────────────────────────────────────
const filters = {{ time: 'all', status: 'all', scope: 'all' }};
let currentView = 'list';

function setFilter(type, value, btn) {{
    filters[type] = value;
    const group = btn.parentElement;
    group.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    applyFilters();
}}

function setView(view, btn) {{
    currentView = view;
    document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('orderList').style.display = view === 'list' ? '' : 'none';
    document.getElementById('insightsView').style.display = view === 'insights' ? '' : 'none';
    if (view === 'insights') updateInsights();
}}

function getDateStr(d) {{
    return d.getFullYear() + '-' +
        String(d.getMonth() + 1).padStart(2, '0') + '-' +
        String(d.getDate()).padStart(2, '0');
}}

function getWeekStart(d) {{
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    return new Date(d.getFullYear(), d.getMonth(), diff);
}}

function rowPassesTimeAndScope(row) {{
    const today = getDateStr(new Date());
    const weekStart = getDateStr(getWeekStart(new Date()));
    const monthStart = today.substring(0, 8) + '01';
    const rowDate = row.dataset.date;
    const isTest = row.dataset.isTest === 'true';

    if (filters.scope === 'prod' && isTest) return false;
    if (filters.scope === 'test' && !isTest) return false;
    if (filters.time === 'today' && rowDate !== today) return false;
    if (filters.time === 'week' && rowDate < weekStart) return false;
    if (filters.time === 'month' && rowDate < monthStart) return false;
    return true;
}}

function applyFilters() {{
    const rows = document.querySelectorAll('.order-row');
    let visible = 0;

    rows.forEach(row => {{
        const rowStatus = row.dataset.status;
        let show = rowPassesTimeAndScope(row);
        if (show && filters.status !== 'all' && rowStatus !== filters.status) show = false;
        row.classList.toggle('hidden-by-filter', !show);
        if (show) visible++;
    }});

    const total = rows.length;
    const countEl = document.getElementById('filterCount');
    const isFiltered = filters.time !== 'all' || filters.status !== 'all' || filters.scope !== 'all';
    countEl.textContent = isFiltered ? visible + ' av ' + total + ' ordrer' : total + ' ordrer';

    // Update stat cards from visible rows
    updateStatCards(rows);

    // Update filter banner
    updateFilterBanner();

    if (currentView === 'insights') updateInsights();
}}

function updateStatCards(rows) {{
    // Only recount if scope filter is active (otherwise server stats are accurate)
    if (filters.scope === 'all' && filters.time === 'all') {{
        // Restore server-rendered values (stored as data attrs on the card elements)
        ['sc-ok','sc-review','sc-dup','sc-err'].forEach(id => {{
            const el = document.getElementById(id);
            if (el && el.dataset.server) el.textContent = el.dataset.server;
        }});
        return;
    }}
    // Count from visible rows that pass time+scope (ignore status filter for stats)
    let ok = 0, review = 0, dup = 0, err = 0;
    rows.forEach(row => {{
        if (!rowPassesTimeAndScope(row)) return;
        const s = row.dataset.status;
        if (s === 'success') ok++;
        else if (s === 'review') review++;
        else if (s === 'skipped') dup++;
        else if (s === 'error') err++;
    }});
    setText('sc-ok', ok);
    setText('sc-review', review);
    setText('sc-dup', dup);
    setText('sc-err', err);
}}

function setText(id, val) {{
    const el = document.getElementById(id);
    if (el) el.textContent = val;
}}

function updateFilterBanner() {{
    const banner = document.getElementById('filterBanner');
    const parts = [];
    if (filters.scope === 'prod') parts.push('kun produksjonsordrer');
    if (filters.scope === 'test') parts.push('kun testordrer');
    if (filters.time === 'today') parts.push('i dag');
    else if (filters.time === 'week') parts.push('denne uken');
    else if (filters.time === 'month') parts.push('denne måneden');
    if (filters.status !== 'all') parts.push('status: ' + filters.status);

    if (parts.length) {{
        banner.textContent = '⚡ Filter aktivt: ' + parts.join(' · ') + ' — statistikk oppdatert tilsvarende';
        banner.style.display = 'flex';
    }} else {{
        banner.style.display = 'none';
    }}
}}

function updateInsights() {{
    const rows = document.querySelectorAll('.order-row');
    let ok = 0, review = 0, dup = 0, err = 0, pending = 0;
    let confTotal = 0, confCount = 0;
    let testCount = 0, prodCount = 0;
    const customerMap = {{}};
    const customerConfMap = {{}};  // customer -> {{ total, count }}
    const customerReviewMap = {{}};  // customer -> review count

    rows.forEach(row => {{
        if (!rowPassesTimeAndScope(row)) return;
        const s = row.dataset.status;
        const isTest = row.dataset.isTest === 'true';
        const customer = row.dataset.customer;

        if (s === 'success') ok++;
        else if (s === 'review') review++;
        else if (s === 'skipped') dup++;
        else if (s === 'error') err++;
        else pending++;

        if (isTest) testCount++; else prodCount++;

        // Confidence from the conf-col span text
        const confEl = row.querySelector('.conf-col');
        const pct = confEl ? parseFloat(confEl.textContent) : NaN;
        if (!isNaN(pct)) {{ confTotal += pct; confCount++; }}

        // Customer breakdown
        if (customer) {{
            customerMap[customer] = (customerMap[customer] || 0) + 1;
            if (!isNaN(pct)) {{
                if (!customerConfMap[customer]) customerConfMap[customer] = {{ total: 0, count: 0 }};
                customerConfMap[customer].total += pct;
                customerConfMap[customer].count++;
            }}
            if (s === 'review') {{
                customerReviewMap[customer] = (customerReviewMap[customer] || 0) + 1;
            }}
        }}
    }});

    const total = ok + review + dup + err + pending;
    const processed = ok + review + dup + err;

    // Suksessrate (ok + review av alt behandlet, ikke pending)
    const rate = processed > 0 ? Math.round((ok + review) / processed * 100) : 0;
    document.getElementById('ins-rate').textContent = rate + '%';
    document.getElementById('ins-rate-sub').textContent = ok + ' OK, ' + review + ' til review av ' + processed;

    // Konfidensverdi
    const avgConf = confCount > 0 ? Math.round(confTotal / confCount) : 0;
    document.getElementById('ins-conf').textContent = confCount > 0 ? avgConf + '%' : '—';

    // Review count
    document.getElementById('ins-review').textContent = review;

    // Test vs prod
    document.getElementById('ins-testprod').textContent = prodCount + ' / ' + testCount;
    document.getElementById('ins-testprod-sub').textContent = 'Produksjon / test av ' + total + ' totalt';

    // Scope label
    const scopeLabels = {{ all: '', prod: '(produksjon)', test: '(test)' }};
    document.getElementById('ins-scope-label').textContent = scopeLabels[filters.scope] || '';

    // Ordrer per kunde
    const sorted = Object.entries(customerMap).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const maxOrders = sorted[0] ? sorted[0][1] : 1;
    const barsEl = document.getElementById('ins-customers');
    barsEl.innerHTML = sorted.length ? sorted.map(([name, count]) => `
        <div class="cbar">
            <div class="cbar-name" title="${{name}}">${{name}}</div>
            <div class="cbar-track"><div class="cbar-fill" style="width:${{Math.round(count/maxOrders*100)}}%"></div></div>
            <div class="cbar-count">${{count}}</div>
        </div>`).join('') : '<p style="color:#475569;font-size:13px;">Ingen data</p>';

    // Konfidens per kunde
    const confEntries = Object.entries(customerConfMap)
        .map(([name, d]) => [name, Math.round(d.total / d.count)])
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
    const confBarsEl = document.getElementById('ins-confidence');
    confBarsEl.innerHTML = confEntries.length ? confEntries.map(([name, avg]) => {{
        const cls = avg >= 90 ? 'conf-high' : avg >= 75 ? 'conf-mid' : 'conf-low';
        const color = avg >= 90 ? '#4ade80' : avg >= 75 ? '#fbbf24' : '#f87171';
        return `<div class="cbar">
            <div class="cbar-name" title="${{name}}">${{name}}</div>
            <div class="cbar-track"><div class="cbar-fill ${{cls}}" style="width:${{avg}}%"></div></div>
            <div class="cbar-count" style="color:${{color}}">${{avg}}%</div>
        </div>`;
    }}).join('') : '<p style="color:#475569;font-size:13px;">Ingen data</p>';

    // Til review per kunde
    const reviewEntries = Object.entries(customerReviewMap)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
    const maxReview = reviewEntries[0] ? reviewEntries[0][1] : 1;
    const reviewBarsEl = document.getElementById('ins-review-customers');
    reviewBarsEl.innerHTML = reviewEntries.length ? reviewEntries.map(([name, count]) => `
        <div class="cbar">
            <div class="cbar-name" title="${{name}}">${{name}}</div>
            <div class="cbar-track"><div class="cbar-fill conf-mid" style="width:${{Math.round(count/maxReview*100)}}%"></div></div>
            <div class="cbar-count" style="color:#fbbf24">${{count}}</div>
        </div>`).join('')
        : '<p style="color:#475569;font-size:13px;">Ingen til review</p>';
}}

// ── Poll trigger + status ─────────────────────────────────────
let pollCheckInterval = null;

function triggerPoll() {{
    const btn = document.getElementById('pollBtn');
    btn.disabled = true;
    btn.textContent = 'Sjekker...';

    fetch('/api/poll', {{
        method: 'POST',
        headers: {{ 'Accept': 'application/json' }}
    }}).then(() => {{
        startPollStatusCheck();
    }}).catch(err => {{
        btn.disabled = false;
        btn.textContent = 'Sjekk e-post nå';
        console.error('Poll trigger failed:', err);
    }});
}}

function startPollStatusCheck() {{
    if (pollCheckInterval) clearInterval(pollCheckInterval);
    updatePollUI(true, 'Sjekker e-post...');
    pollCheckInterval = setInterval(checkPollStatus, 2000);
}}

function checkPollStatus() {{
    fetch('/api/poll/status')
        .then(r => r.json())
        .then(state => {{
            if (state.running) {{
                updatePollUI(true, 'Sjekker e-post...');
            }} else if (pollCheckInterval !== null) {{
                // A poll *we triggered this session* just finished
                clearInterval(pollCheckInterval);
                pollCheckInterval = null;

                const btn = document.getElementById('pollBtn');
                btn.disabled = false;
                btn.textContent = 'Sjekk e-post nå';

                if (state.last_error) {{
                    updatePollUI(false, 'Feil: ' + state.last_error);
                }} else {{
                    updatePollUI(false, (state.last_result || 0) + ' ordre(r) behandlet');
                    if (state.last_result > 0) {{
                        setTimeout(() => window.location.reload(), 1500);
                    }}
                }}
            }}

            updateScheduleInfo(state);
        }})
        .catch(() => {{}});
}}

function updatePollUI(isPolling, text) {{
    const container = document.getElementById('pollStatus');
    const textEl = document.getElementById('pollText');
    container.classList.toggle('polling', isPolling);
    textEl.textContent = text;
}}

function updateScheduleInfo(state) {{
    const el = document.getElementById('scheduleInfo');
    if (!el) return;
    const parts = [];
    if (state.last_run) {{
        const lr = new Date(state.last_run);
        parts.push('Sist: ' + lr.toLocaleTimeString('nb-NO', {{ hour: '2-digit', minute: '2-digit' }}));
    }}
    if (state.next_scheduled) {{
        const ns = new Date(state.next_scheduled);
        const now = new Date();
        if (ns.toDateString() === now.toDateString()) {{
            parts.push('Neste: i dag kl. ' + ns.toLocaleTimeString('nb-NO', {{ hour: '2-digit', minute: '2-digit' }}));
        }} else {{
            parts.push('Neste: i morgen kl. ' + ns.toLocaleTimeString('nb-NO', {{ hour: '2-digit', minute: '2-digit' }}));
        }}
    }}
    el.textContent = parts.join(' · ');
}}

// ── Init ──────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {{
    applyFilters();
    // Check poll status on load (in case a poll is running)
    checkPollStatus();
    // Periodically refresh schedule info (every 60s)
    setInterval(checkPollStatus, 60000);
}});
</script>

</body>
</html>"""


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def dashboard():
    stats = _order_stats()
    events = _recent_events(500)
    dead_letters = list_dead_letters()
    return _render_dashboard(stats, events, dead_letters)


@app.get("/api/stats")
async def api_stats():
    return _order_stats()


@app.get("/api/events")
async def api_events(order: str | None = None, cid: str | None = None, limit: int = 50):
    return list_events(correlation_id=cid, order_number=order, last_n=limit)


@app.get("/api/dead-letters")
async def api_dead_letters():
    return list_dead_letters()


@app.post("/api/poll")
async def api_poll(background_tasks: BackgroundTasks, request: Request):
    """Trigger email polling in the background."""
    background_tasks.add_task(_run_poll)
    # AJAX calls get JSON, form submits get redirect
    accept = request.headers.get("accept", "")
    if "application/json" in accept:
        return JSONResponse({"started": True})
    return RedirectResponse("/", status_code=303)


@app.get("/api/poll/status")
async def api_poll_status():
    """Return current poll state for dashboard live updates."""
    return JSONResponse(_poll_state)


@app.post("/api/replay/{filename:path}")
async def api_replay(filename: str, background_tasks: BackgroundTasks):
    """Replay a failed order from downloads/."""
    background_tasks.add_task(_run_replay, filename)
    return RedirectResponse("/", status_code=303)


# ---------------------------------------------------------------------------
# Background tasks
# ---------------------------------------------------------------------------

_poll_lock = threading.Lock()


def _run_poll():
    """Run email poll (with Odoo push) in background."""
    if not _poll_lock.acquire(blocking=False):
        logger.info("E-postsjekk kjorer allerede, hopper over")
        return
    _poll_state["running"] = True
    _poll_state["last_error"] = None
    try:
        from .config import require_graph_config, require_odoo_config
        from .graph_client import GraphClient
        from .email_monitor import EmailMonitor
        from .sharepoint_archiver import SharePointArchiver
        from .odoo_client import OdooClient
        from .odoo_mapper import OdooMapper
        from .odoo_order import OdooOrderService

        # Graph
        gcfg = require_graph_config()
        graph = GraphClient(gcfg["MS_TENANT_ID"], gcfg["MS_CLIENT_ID"], gcfg["MS_CLIENT_SECRET"])

        # Odoo
        ocfg = require_odoo_config()
        client = OdooClient(ocfg["ODOO_URL"], ocfg["ODOO_DB"], ocfg["ODOO_USERNAME"], ocfg["ODOO_PASSWORD"])
        client.authenticate()
        mapper = OdooMapper(client)
        fallback_id = ocfg.get("ODOO_FALLBACK_PRODUCT_ID")
        transport_id = ocfg.get("ODOO_TRANSPORT_PRODUCT_ID")
        odoo_service = OdooOrderService(
            client, mapper,
            fallback_product_id=int(fallback_id) if fallback_id else None,
            transport_product_id=int(transport_id) if transport_id else None,
        )

        # SharePoint
        archiver = SharePointArchiver(graph)

        monitor = EmailMonitor(graph_client=graph, odoo_service=odoo_service,
                                archiver=archiver)
        results = monitor.poll()
        _poll_state["last_result"] = len(results)
        logger.info("E-postsjekk ferdig: %d resultater", len(results))
    except Exception as e:
        logger.exception("E-postsjekk feilet: %s", e)
        _poll_state["last_error"] = str(e)
    finally:
        _poll_state["running"] = False
        _poll_state["last_run"] = datetime.now().isoformat(timespec="seconds")
        _poll_lock.release()


# ---------------------------------------------------------------------------
# Daily scheduler
# ---------------------------------------------------------------------------

_scheduler_thread: threading.Thread | None = None
_scheduler_stop = threading.Event()

# Default: run at 07:00 every day. Override with POLL_SCHEDULE_HOUR in .env.
_SCHEDULE_HOUR = 7


def _get_schedule_hour() -> int:
    from .config import load_config
    cfg = load_config()
    try:
        return int(cfg.get("POLL_SCHEDULE_HOUR", str(_SCHEDULE_HOUR)))
    except (ValueError, TypeError):
        return _SCHEDULE_HOUR


def _next_run_time() -> datetime:
    """Calculate the next scheduled run time."""
    hour = _get_schedule_hour()
    now = datetime.now()
    target = now.replace(hour=hour, minute=0, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)
    return target


def _scheduler_loop():
    """Background thread that triggers _run_poll() once daily."""
    logger.info("Daglig e-postsjekk scheduler startet (kl. %02d:00)", _get_schedule_hour())
    while not _scheduler_stop.is_set():
        next_run = _next_run_time()
        _poll_state["next_scheduled"] = next_run.isoformat(timespec="seconds")

        # Sleep in short intervals so we can stop cleanly
        while datetime.now() < next_run:
            if _scheduler_stop.wait(timeout=30):
                logger.info("Scheduler stoppet")
                return

        logger.info("Daglig planlagt e-postsjekk starter")
        _run_poll()


def start_scheduler():
    """Start the daily poll scheduler (idempotent)."""
    global _scheduler_thread
    if _scheduler_thread and _scheduler_thread.is_alive():
        return
    _scheduler_stop.clear()
    _scheduler_thread = threading.Thread(target=_scheduler_loop, daemon=True, name="poll-scheduler")
    _scheduler_thread.start()


def stop_scheduler():
    """Stop the daily poll scheduler."""
    _scheduler_stop.set()
    if _scheduler_thread:
        _scheduler_thread.join(timeout=5)


# Auto-start scheduler when the dashboard app starts
@app.on_event("startup")
async def _on_startup():
    start_scheduler()
    logger.info("Dashboard startet med daglig e-postsjekk kl. %02d:00", _get_schedule_hour())


def _run_replay(filename: str):
    """Replay a single PDF."""
    try:
        from .event_log import log_event, new_correlation_id
        from .order_parser import parse_order_pdf
        from .validator import needs_manual_review, validate_order
        from .config import require_odoo_config
        from .odoo_client import OdooClient
        from .odoo_mapper import OdooMapper
        from .odoo_order import OdooOrderService

        cid = new_correlation_id()
        pdf_path = Path("downloads") / filename
        if not pdf_path.exists():
            logger.error("Replay: fant ikke %s", pdf_path)
            return

        log_event(cid, "replay_started", source_file=filename)

        order = parse_order_pdf(pdf_path)
        validate_order(order)
        review = needs_manual_review(order)

        ocfg = require_odoo_config()
        client = OdooClient(ocfg["ODOO_URL"], ocfg["ODOO_DB"], ocfg["ODOO_USERNAME"], ocfg["ODOO_PASSWORD"])
        client.authenticate()
        mapper = OdooMapper(client)
        fallback_id = ocfg.get("ODOO_FALLBACK_PRODUCT_ID")
        transport_id = ocfg.get("ODOO_TRANSPORT_PRODUCT_ID")
        service = OdooOrderService(
            client, mapper,
            fallback_product_id=int(fallback_id) if fallback_id else None,
            transport_product_id=int(transport_id) if transport_id else None,
        )

        result = service.push_order(order, needs_review=review)
        log_event(cid, "replay_completed", order_number=order.order_number,
                   status=result.status,
                   details={"so_name": result.so_name, "message": result.message})

        # Resolve dead letter if successful
        if result.status == "success":
            from .event_log import resolve_dead_letter as _resolve
            # Find matching dead letter by filename
            for dl in list_dead_letters():
                if dl["source_file"] == filename:
                    _resolve(dl["cid"])

        logger.info("Replay ferdig: %s -> %s", filename, result.status)
    except Exception as e:
        logger.exception("Replay feilet for %s: %s", filename, e)
