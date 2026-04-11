"""FastAPI dashboard for Ortopartner order automation."""

from __future__ import annotations

import json
import logging
import threading
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from .event_log import list_dead_letters, list_events, resolve_dead_letter

logger = logging.getLogger(__name__)

app = FastAPI(title="Ortopartner Ordreflyt", version="1.0")

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
    orders.sort(key=lambda o: o["last_ts"], reverse=True)
    return orders


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

    # Aggregate events by order for the "Siste ordrer" section
    orders = _aggregate_orders(events)[:25]

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
            warn_summary = '<span class="warn-count none">—</span>'

        warnings_block = _render_order_warnings_block(o["warnings"])

        customer = _escape(o["customer"] or "-")
        so = _escape(o["so_name"] or "-")
        order_num = _escape(o["order_number"])
        total = _format_amount(o["total_amount"], o["currency"] or "NOK")
        line_count = o["line_count"] if o["line_count"] is not None else "-"
        message = _escape(o["message"] or "")

        order_rows += f'''
        <details class="order-row">
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
                <span class="ts-col">{o["last_ts"][11:19]}</span>
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
                <button type="submit" class="btn btn-sm">Re-kjør</button>
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
            <td>{ev['ts'][11:]}</td>
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
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
           background: #f5f5f5; color: #333; padding: 20px; }}
    h1 {{ color: #1a1a2e; margin-bottom: 8px; font-size: 24px; }}
    .subtitle {{ color: #666; margin-bottom: 24px; font-size: 14px; }}
    .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
              gap: 12px; margin-bottom: 24px; }}
    .card {{ background: #fff; border-radius: 8px; padding: 16px; text-align: center;
             box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
    .card .num {{ font-size: 32px; font-weight: 700; color: #1a1a2e; }}
    .card .label {{ font-size: 12px; color: #888; margin-top: 4px; }}
    .card.green .num {{ color: #2d6a4f; }}
    .card.red .num {{ color: #c1121f; }}
    .card.yellow .num {{ color: #e09f3e; }}
    .card.blue .num {{ color: #457b9d; }}
    .actions {{ margin-bottom: 24px; display: flex; gap: 8px; }}
    .btn {{ background: #1a1a2e; color: #fff; border: none; padding: 8px 16px;
            border-radius: 6px; cursor: pointer; font-size: 13px; }}
    .btn:hover {{ background: #2a2a4e; }}
    .btn-sm {{ padding: 4px 10px; font-size: 12px; }}
    section {{ background: #fff; border-radius: 8px; padding: 16px; margin-bottom: 16px;
               box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
    section h2 {{ font-size: 16px; margin-bottom: 12px; color: #1a1a2e; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    th {{ text-align: left; padding: 6px 8px; border-bottom: 2px solid #eee;
          color: #888; font-weight: 600; font-size: 11px; text-transform: uppercase; }}
    td {{ padding: 6px 8px; border-bottom: 1px solid #f0f0f0; }}
    code {{ background: #f0f0f0; padding: 2px 5px; border-radius: 3px; font-size: 11px; }}
    .err {{ color: #c1121f; }}
    .ok {{ color: #2d6a4f; }}
    .empty {{ color: #aaa; padding: 20px; text-align: center; }}
    .test-banner {{ background: #ffe066; color: #1a1a2e; padding: 10px 16px;
                    border-radius: 6px; margin-bottom: 16px; font-weight: 600;
                    border-left: 4px solid #e09f3e; font-size: 14px; }}
    .test-banner code {{ background: #1a1a2e; color: #ffe066; }}

    /* --- Siste ordrer section --- */
    .orders-header, .order-row summary {{
        display: grid;
        grid-template-columns: 2.4fr 2fr 1fr 0.7fr 0.5fr 1.3fr 1.3fr 0.8fr;
        gap: 8px;
        padding: 10px 12px;
        align-items: center;
        font-size: 13px;
    }}
    .orders-header {{
        background: #f0f3f8;
        color: #555;
        font-weight: 600;
        font-size: 11px;
        text-transform: uppercase;
        border-radius: 6px 6px 0 0;
    }}
    .order-row {{
        background: #fff;
        border-bottom: 1px solid #eef0f4;
    }}
    .order-row:last-child {{ border-bottom: none; }}
    .order-row summary {{
        cursor: pointer;
        list-style: none;
        transition: background 0.1s;
    }}
    .order-row summary::-webkit-details-marker {{ display: none; }}
    .order-row summary:hover {{ background: #f8fafc; }}
    .order-row[open] summary {{ background: #f1f5f9; }}

    .ordre-col {{ display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }}
    .customer-col {{ color: #334; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    .so-col {{ font-family: monospace; color: #457b9d; }}
    .conf-col, .lines-col {{ text-align: center; color: #666; }}
    .total-col {{ text-align: right; font-variant-numeric: tabular-nums; color: #1a1a2e; font-weight: 600; }}
    .warn-col {{ text-align: left; }}
    .ts-col {{ text-align: right; color: #999; font-family: monospace; font-size: 11px; }}

    .badge {{
        display: inline-block;
        padding: 2px 7px;
        border-radius: 10px;
        font-size: 10px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.3px;
    }}
    .badge-success {{ background: #d4edda; color: #155724; }}
    .badge-skipped {{ background: #fff3cd; color: #856404; }}
    .badge-error {{ background: #f8d7da; color: #721c24; }}
    .badge-pending {{ background: #e2e3e5; color: #383d41; }}
    .badge-review {{ background: #fce5cd; color: #9c4a00; }}
    .badge-archived {{ background: #cfe2ff; color: #084298; }}

    .warn-count {{
        display: inline-block;
        padding: 1px 6px;
        border-radius: 8px;
        font-size: 11px;
        font-weight: 600;
        margin-right: 3px;
    }}
    .warn-count.warn {{ background: #fce5cd; color: #9c4a00; }}
    .warn-count.info {{ background: #e3f2fd; color: #0d47a1; }}
    .warn-count.none {{ color: #bbb; }}

    .order-detail {{
        padding: 16px 20px;
        background: #fafbfc;
        border-top: 1px solid #e8eaf0;
    }}
    .order-meta {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 8px 16px;
        margin-bottom: 12px;
        font-size: 12px;
        color: #555;
    }}
    .order-meta strong {{ color: #333; margin-right: 4px; }}
    .order-message {{
        background: #fff;
        border-left: 3px solid #457b9d;
        padding: 8px 12px;
        margin-bottom: 12px;
        font-size: 12px;
        color: #444;
        border-radius: 0 4px 4px 0;
    }}

    .warning-group {{
        background: #fff;
        border-radius: 5px;
        padding: 10px 14px;
        margin-bottom: 8px;
        border-left: 4px solid #ccc;
    }}
    .warning-group.warn {{ border-left-color: #e09f3e; background: #fffbf2; }}
    .warning-group.info {{ border-left-color: #457b9d; background: #f4f8fc; }}
    .warning-group-title {{
        font-weight: 700;
        font-size: 11px;
        text-transform: uppercase;
        color: #666;
        margin-bottom: 6px;
        letter-spacing: 0.3px;
    }}
    .warning-group.warn .warning-group-title {{ color: #9c4a00; }}
    .warning-group.info .warning-group-title {{ color: #0d47a1; }}
    .warning-group ul {{
        margin: 0;
        padding-left: 18px;
        font-size: 12px;
        line-height: 1.6;
    }}
    .warning-group li {{ color: #444; }}
    .no-warnings {{
        color: #9aa;
        font-size: 12px;
        font-style: italic;
        padding: 6px 0;
    }}

    .orders-section {{ padding: 0; }}
    .orders-section h2 {{ padding: 16px 16px 12px; }}
    .orders-empty {{ padding: 30px; text-align: center; color: #aaa; }}
</style>
</head>
<body>
<h1>Ortopartner Ordreflyt</h1>
<p class="subtitle">Automatisk ordrebehandling: E-post &rarr; PDF &rarr; Odoo &rarr; SharePoint</p>

{test_mode_banner}

<div class="cards">
    <div class="card blue"><div class="num">{stats['emails_processed']}</div><div class="label">E-poster behandlet</div></div>
    <div class="card green"><div class="num">{stats['success']}</div><div class="label">Ordrer OK</div></div>
    <div class="card yellow"><div class="num">{stats['skipped']}</div><div class="label">Duplikater</div></div>
    <div class="card yellow"><div class="num">{stats['review']}</div><div class="label">Til review</div></div>
    <div class="card red"><div class="num">{stats['dead_letters']}</div><div class="label">Feilede</div></div>
</div>

<div class="actions">
    <form method="post" action="/api/poll"><button type="submit" class="btn">Sjekk e-post nå</button></form>
</div>

{"" if not dead_letters else f'''<section>
<h2>Feilede ordrer (dead-letter)</h2>
<table>
<tr><th>CID</th><th>Tid</th><th>Fil</th><th>Ordre</th><th>Steg</th><th>Feil</th><th></th></tr>
{dl_rows}
</table>
</section>'''}

<section class="orders-section">
<h2>Siste ordrer <span style="font-size:11px; color:#888; font-weight:400;">(klikk for detaljer)</span></h2>
{f'<div class="orders-empty">Ingen ordrer behandlet ennå.</div>' if not orders else f'''
<div class="orders-header">
    <span>Ordre / Status</span>
    <span>Kunde</span>
    <span>SO</span>
    <span>Konfid.</span>
    <span>Linjer</span>
    <span style="text-align:right;">Totalbeløp</span>
    <span>Advarsler</span>
    <span style="text-align:right;">Tid</span>
</div>
{order_rows}
'''}
</section>

<section>
<h2>Siste hendelser (råformat)</h2>
{"<p class='empty'>Ingen hendelser ennå.</p>" if not events else f'''<table>
<tr><th>Tid</th><th>CID</th><th>Hendelse</th><th>Status</th><th>Ordre</th><th>Detaljer</th></tr>
{event_rows}
</table>'''}
</section>

</body>
</html>"""


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def dashboard():
    stats = _order_stats()
    events = _recent_events(50)
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
async def api_poll(background_tasks: BackgroundTasks):
    """Trigger email polling in the background."""
    background_tasks.add_task(_run_poll)
    return RedirectResponse("/", status_code=303)


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
        logger.info("E-postsjekk ferdig: %d resultater", len(results))
    except Exception as e:
        logger.exception("E-postsjekk feilet: %s", e)
    finally:
        _poll_lock.release()


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
