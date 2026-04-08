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
            if details.get("warnings"):
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


# ---------------------------------------------------------------------------
# HTML template (inline to keep it simple)
# ---------------------------------------------------------------------------

def _render_dashboard(stats: dict, events: list[dict], dead_letters: list[dict]) -> str:
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
</style>
</head>
<body>
<h1>Ortopartner Ordreflyt</h1>
<p class="subtitle">Automatisk ordrebehandling: E-post &rarr; PDF &rarr; Odoo &rarr; SharePoint</p>

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

<section>
<h2>Siste hendelser</h2>
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
        odoo_service = OdooOrderService(client, mapper,
                                         fallback_product_id=int(fallback_id) if fallback_id else None)

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
        service = OdooOrderService(client, mapper,
                                    fallback_product_id=int(fallback_id) if fallback_id else None)

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
