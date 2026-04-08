"""CLI runner for order parsing pipeline."""

from __future__ import annotations

import json
import logging
import sys
import time
from pathlib import Path

from .models import OdooResult, ParsedOrder, TrackingUpdate
from .order_parser import parse_order_pdf
from .validator import needs_manual_review, validate_order

logger = logging.getLogger(__name__)


def _init_odoo_service():
    """Initialize Odoo client, mapper, and order service."""
    from .config import require_odoo_config
    from .odoo_client import OdooClient
    from .odoo_mapper import OdooMapper
    from .odoo_order import OdooOrderService

    cfg = require_odoo_config()
    client = OdooClient(
        cfg["ODOO_URL"], cfg["ODOO_DB"],
        cfg["ODOO_USERNAME"], cfg["ODOO_PASSWORD"],
    )
    client.authenticate()

    mapper = OdooMapper(client)

    fallback_id = cfg.get("ODOO_FALLBACK_PRODUCT_ID")
    fallback_product_id = int(fallback_id) if fallback_id else None

    service = OdooOrderService(
        client, mapper,
        fallback_product_id=fallback_product_id,
    )
    return service


def _push_to_odoo(order: ParsedOrder, service) -> OdooResult | None:
    """Push order to Odoo as a draft Quotation.

    Flagged orders get a REVIEW tag and warning message in Odoo.
    """
    review = needs_manual_review(order)

    if review:
        print(
            f"  [3/3] Sender til Odoo som utkast (flagget, "
            f"konfidensverdi: {order.confidence:.0%})...",
            end=" ", flush=True,
        )
    else:
        print("  [3/3] Sender til Odoo som utkast...", end=" ", flush=True)

    result = service.push_order(order, needs_review=review)

    if result.status == "success":
        print(f"OK — {result.message}")
    elif result.status == "skipped":
        print(f"HOPPET OVER — {result.message}")
    else:
        print(f"FEIL — {result.message}")

    if result.warnings:
        for w in result.warnings:
            print(f"    ! {w}")

    if result.purchase_order_ids:
        print(f"  Innkjøpsordre:  {result.purchase_order_ids}")

    return result


def process_single(
    pdf_path: str,
    output_dir: str | None = None,
    push_to_odoo: bool = False,
    odoo_service=None,
) -> ParsedOrder:
    """Process a single PDF and optionally push to Odoo."""
    path = Path(pdf_path)
    total_steps = 3 if push_to_odoo else 2

    print(f"\n{'='*60}")
    print(f"  Prosesserer: {path.name}")
    print(f"{'='*60}")

    start = time.time()

    # Step 1: Parse
    print(f"  [1/{total_steps}] Parser PDF med AI...", end=" ", flush=True)
    order = parse_order_pdf(path)
    print(f"OK ({time.time() - start:.1f}s)")

    # Step 2: Validate
    print(f"  [2/{total_steps}] Validerer...", end=" ", flush=True)
    order = validate_order(order)
    print("OK")

    # Status
    review = needs_manual_review(order)
    status = "FLAGGET" if review else "OK"
    print(f"\n  Status:         {status}")
    print(f"  Konfidensverdi: {order.confidence:.0%}")
    print(f"  Bestillingsnr:  {order.order_number}")
    print(f"  Kunde:          {order.customer_name}")
    print(f"  Antall linjer:  {len(order.line_items)}")
    if order.total_amount:
        print(f"  Totalbeløp:     {order.total_amount:,.2f} {order.currency}")
    if order.warnings:
        print(f"  Advarsler:")
        for w in order.warnings:
            print(f"    - {w}")

    # Step 3: Odoo push
    odoo_result = None
    if push_to_odoo and odoo_service:
        odoo_result = _push_to_odoo(order, odoo_service)

    # Step 4: Alerting
    from .alerting import AlertService
    odoo_for_alerts = None
    if push_to_odoo and odoo_service:
        try:
            from .config import require_odoo_config
            from .odoo_client import OdooClient
            cfg = require_odoo_config()
            odoo_for_alerts = OdooClient(
                cfg["ODOO_URL"], cfg["ODOO_DB"],
                cfg["ODOO_USERNAME"], cfg["ODOO_PASSWORD"],
            )
            odoo_for_alerts.authenticate()
        except Exception:
            pass  # Alerts still log locally even without Odoo

    alert_service = AlertService(odoo_for_alerts)
    alerts = alert_service.check_order(order, odoo_result)
    if alerts:
        print(f"  Varsler sendt:  {len(alerts)}")
        for a in alerts:
            print(f"    ! {a}")

    # Save JSON output
    if output_dir:
        out_path = Path(output_dir)
        out_path.mkdir(parents=True, exist_ok=True)

        json_file = out_path / f"{path.stem}.json"
        output_data = order.model_dump()
        if odoo_result:
            output_data["odoo_result"] = odoo_result.model_dump()
        json_file.write_text(
            json.dumps(output_data, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        print(f"  Lagret:         {json_file}")

    return order


def process_batch(
    pdf_dir: str,
    output_dir: str = "output",
    push_to_odoo: bool = False,
) -> None:
    """Process all PDFs in a directory."""
    pdf_dir_path = Path(pdf_dir)
    pdfs = sorted(pdf_dir_path.glob("*.pdf"))

    if not pdfs:
        print(f"Ingen PDF-filer funnet i {pdf_dir}")
        return

    print(f"\nFant {len(pdfs)} PDF-filer i {pdf_dir}")
    print(f"Output-mappe: {output_dir}")
    if push_to_odoo:
        print("Odoo-push: AKTIVERT")
    print()

    # Initialize Odoo service once for the batch
    odoo_service = None
    if push_to_odoo:
        try:
            odoo_service = _init_odoo_service()
            print("Odoo-tilkobling: OK\n")
        except Exception as e:
            print(f"Odoo-tilkobling FEILET: {e}")
            print("Fortsetter uten Odoo-push.\n")
            push_to_odoo = False

    results: list[dict] = []
    ok_count = 0
    flagged_count = 0
    error_count = 0

    for pdf in pdfs:
        try:
            order = process_single(
                str(pdf), output_dir,
                push_to_odoo=push_to_odoo,
                odoo_service=odoo_service,
            )
            review = needs_manual_review(order)
            results.append(
                {
                    "file": pdf.name,
                    "order_number": order.order_number,
                    "customer": order.customer_name,
                    "status": "FLAGGET" if review else "OK",
                    "confidence": order.confidence,
                    "lines": len(order.line_items),
                    "total": order.total_amount,
                    "warnings": order.warnings,
                }
            )
            if review:
                flagged_count += 1
            else:
                ok_count += 1
        except Exception as e:
            print(f"\n  FEIL ved prosessering av {pdf.name}: {e}")
            results.append(
                {
                    "file": pdf.name,
                    "status": "FEIL",
                    "error": str(e),
                }
            )
            error_count += 1

    # Summary
    print(f"\n{'='*60}")
    print(f"  OPPSUMMERING")
    print(f"{'='*60}")
    print(f"  Totalt:    {len(pdfs)}")
    print(f"  OK:        {ok_count}")
    print(f"  Flagget:   {flagged_count}")
    print(f"  Feil:      {error_count}")

    if flagged_count > 0:
        print(f"\n  Flaggede bestillinger (trenger manuell kontroll):")
        for r in results:
            if r.get("status") == "FLAGGET":
                print(f"    - {r['file']}: {r['order_number']} ({r['confidence']:.0%})")
                for w in r.get("warnings", []):
                    print(f"        {w}")

    # Save summary
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    summary_file = out / "_summary.json"
    summary_file.write_text(
        json.dumps(results, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"\n  Oppsummering lagret: {summary_file}")


def _init_dhl_tracker():
    """Initialize DHL client, Odoo client, and tracker."""
    from .config import require_odoo_config
    from .dhl_client import DhlClient
    from .dhl_tracker import DhlTracker
    from .odoo_client import OdooClient

    cfg = require_odoo_config()

    dhl_key = cfg.get("DHL_API_KEY")
    dhl_secret = cfg.get("DHL_API_SECRET")
    if not dhl_key or not dhl_secret:
        raise ValueError("Mangler DHL_API_KEY/DHL_API_SECRET i .env")

    dhl_client = DhlClient(dhl_key, dhl_secret, cfg.get("DHL_BASE_URL"))

    odoo_client = OdooClient(
        cfg["ODOO_URL"], cfg["ODOO_DB"],
        cfg["ODOO_USERNAME"], cfg["ODOO_PASSWORD"],
    )
    odoo_client.authenticate()

    return DhlTracker(dhl_client, odoo_client)


def cmd_set_tracking(so_name: str, tracking_number: str) -> None:
    """Set a tracking number on a sale order."""
    tracker = _init_dhl_tracker()
    result = tracker.set_tracking_number(so_name, tracking_number)

    print(f"\n  SO:             {result.so_name}")
    print(f"  Trackingnummer: {result.tracking_number}")
    print(f"  Status:         {result.status}")
    print(f"  Melding:        {result.message}")


def cmd_track(so_name: str) -> None:
    """Track a single sale order."""
    tracker = _init_dhl_tracker()
    result = tracker.sync_tracking(so_name)

    print(f"\n  SO:             {result.so_name}")
    print(f"  Trackingnummer: {result.tracking_number}")
    print(f"  Status:         {result.status}")
    if result.dhl_status:
        print(f"  DHL-status:     {result.dhl_status}")
    print(f"  Melding:        {result.message}")

    if result.events:
        print(f"\n  Siste hendelser:")
        for ev in result.events[:5]:
            loc = f" ({ev.location_city})" if ev.location_city else ""
            print(f"    {ev.timestamp}  {ev.status_message}{loc}")

    # Check for alerts
    from .alerting import AlertService
    alert_service = AlertService()
    alerts = alert_service.check_tracking(result)
    if alerts:
        print(f"\n  Varsler sendt:  {len(alerts)}")
        for a in alerts:
            print(f"    ! {a}")


def cmd_track_all() -> None:
    """Track all open sale orders with tracking numbers."""
    tracker = _init_dhl_tracker()
    results = tracker.sync_all_open()

    if not results:
        print("\n  Ingen åpne leveranser med trackingnummer funnet.")
        return

    from .alerting import AlertService
    alert_service = AlertService()
    total_alerts = 0

    print(f"\n  Sporet {len(results)} leveranser:")
    for r in results:
        dhl = f" [{r.dhl_status}]" if r.dhl_status else ""
        print(f"    {r.so_name}: {r.tracking_number} - {r.status}{dhl}")
        alerts = alert_service.check_tracking(r)
        total_alerts += len(alerts)

    if total_alerts:
        print(f"\n  Totalt {total_alerts} varsler sendt (se output/alerts.jsonl)")


def cmd_replay(pdf_path: str) -> None:
    """Re-parse a PDF and push to Odoo. Used to retry failed orders."""
    from .event_log import log_event, new_correlation_id

    cid = new_correlation_id()
    print(f"Replay: {pdf_path} (cid={cid})")
    log_event(cid, "replay_started", source_file=pdf_path)

    try:
        order = parse_order_pdf(pdf_path)
        validate_order(order)
        review = needs_manual_review(order)

        print(f"  Ordre: {order.order_number} (konfidensverdi: {order.confidence:.0%})")
        log_event(cid, "pdf_parsed", order_number=order.order_number,
                   source_file=pdf_path, details={"confidence": order.confidence})

        service = _init_odoo_service()
        result = service.push_order(order, needs_review=review)

        print(f"  Status: {result.status}")
        print(f"  Melding: {result.message}")
        if result.so_name:
            print(f"  SO: {result.so_name}")
        if result.warnings:
            for w in result.warnings:
                print(f"  ! {w}")

        log_event(cid, "replay_completed", order_number=order.order_number,
                   status=result.status,
                   details={"so_name": result.so_name, "message": result.message})

    except Exception as e:
        print(f"  FEIL: {e}")
        log_event(cid, "replay_failed", source_file=pdf_path, status="error",
                   details={"error": str(e)})


def cmd_dead_letters() -> None:
    """Show unresolved dead-letter entries."""
    from .event_log import list_dead_letters

    entries = list_dead_letters(unresolved_only=True)
    if not entries:
        print("Ingen feilede ordrer i dead-letter-koen.")
        return

    print(f"Feilede ordrer ({len(entries)} stk):\n")
    for e in entries:
        print(f"  [{e['cid']}] {e['ts']} | {e['source_file']}")
        print(f"    Steg: {e['stage']} | Feil: {e['error']}")
        if e.get("order_number"):
            print(f"    Ordre: {e['order_number']}")
        print(f"    Re-kjør: python -m src --replay downloads/{e['source_file']}")
        print()


def cmd_events(query: str | None = None) -> None:
    """Show event log, optionally filtered by order number or correlation ID."""
    from .event_log import list_events

    if query and len(query) == 12 and query.isalnum():
        events = list_events(correlation_id=query)
    elif query:
        events = list_events(order_number=query)
    else:
        events = list_events(last_n=30)

    if not events:
        print("Ingen hendelser funnet.")
        return

    print(f"Hendelser ({len(events)} stk):\n")
    for e in events:
        order = e.get("order", "")
        details = ""
        if e.get("details"):
            details = f" | {json.dumps(e['details'], ensure_ascii=False)}"
        print(f"  {e['ts']} [{e['cid']}] {e['event']}: {e['status']} {order}{details}")


def cmd_rollback(so_name: str) -> None:
    """Cancel a sale order in Odoo. Reverts to draft if confirmed."""
    service = _init_odoo_service()
    client = service._client

    # Find the SO
    so_data = client.search_read(
        "sale.order", [["name", "=", so_name]], ["id", "name", "state"], limit=1,
    )
    if not so_data:
        print(f"Fant ikke SO '{so_name}' i Odoo.")
        return

    so = so_data[0]
    print(f"SO: {so['name']} (id={so['id']}, state={so['state']})")

    if so["state"] == "cancel":
        print("  Allerede kansellert.")
        return

    try:
        # If confirmed (sale), cancel it
        if so["state"] in ("sale", "done"):
            client.call("sale.order", "action_cancel", [[so["id"]]])
            print(f"  Kansellert (var bekreftet)")
        elif so["state"] == "draft":
            client.call("sale.order", "action_cancel", [[so["id"]]])
            print(f"  Kansellert (var utkast)")
        else:
            client.call("sale.order", "action_cancel", [[so["id"]]])
            print(f"  Kansellert (var {so['state']})")

        # Log rollback event
        from .event_log import log_event, new_correlation_id
        cid = new_correlation_id()
        log_event(cid, "rollback", order_number=so_name, status="ok",
                   details={"previous_state": so["state"]})

    except Exception as e:
        print(f"  FEIL ved kansellering: {e}")


def _init_email_monitor(push_to_odoo: bool = False):
    """Initialize Graph client, EmailMonitor, and SharePoint archiver."""
    from .config import require_graph_config
    from .graph_client import GraphClient
    from .email_monitor import EmailMonitor
    from .sharepoint_archiver import SharePointArchiver

    cfg = require_graph_config()
    graph = GraphClient(
        cfg["MS_TENANT_ID"], cfg["MS_CLIENT_ID"], cfg["MS_CLIENT_SECRET"],
    )

    odoo_service = None
    if push_to_odoo:
        odoo_service = _init_odoo_service()

    archiver = SharePointArchiver(graph)

    return EmailMonitor(
        graph_client=graph,
        odoo_service=odoo_service,
        archiver=archiver,
        download_dir="downloads",
        output_dir="output",
    )


def cmd_poll_email(push_to_odoo: bool = False) -> None:
    """Poll ordre@ortopartner.no for new orders."""
    print("Sjekker e-post (ordre@ortopartner.no)...")
    monitor = _init_email_monitor(push_to_odoo)
    results = monitor.poll()

    if not results:
        print("  Ingen nye ordrer funnet.")
        return

    print(f"\n  Behandlet {len(results)} PDF-vedlegg:\n")
    for r in results:
        status_icon = {"success": "+", "parsed": "~", "skipped": "=", "error": "!"}.get(
            r["status"], "?"
        )
        so = f" -> {r['so_name']}" if r.get("so_name") else ""
        print(f"  [{status_icon}] {r['filename']}: {r['message']}{so}")


def main():
    """CLI entry point."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    args = sys.argv[1:]

    if not args:
        print("Bruk:")
        print("  python -m src <pdf-fil>                          # Parser en fil")
        print("  python -m src <pdf-fil> --push                   # Parser + push til Odoo (utkast)")
        print("  python -m src --batch <mappe>                    # Parser alle PDFs i mappe")
        print("  python -m src --batch <mappe> --push             # Batch + Odoo-push (utkast)")
        print("  python -m src --batch <mappe> -o <output>        # Med output-mappe")
        print()
        print("  python -m src --poll-email                       # Hent nye ordrer fra e-post")
        print("  python -m src --poll-email --push                # Hent + push til Odoo")
        print()
        print("  python -m src --set-tracking <SO> <trackingnr>   # Sett DHL trackingnr")
        print("  python -m src --track <SO>                       # Spor en leveranse")
        print("  python -m src --track-all                        # Spor alle åpne")
        print()
        print("  python -m src --replay <pdf-fil>                 # Re-kjør en PDF mot Odoo")
        print("  python -m src --dead-letters                     # Vis feilede ordrer")
        print("  python -m src --events <ordrenr|cid>             # Vis hendelseslogg")
        print("  python -m src --rollback <SO-navn>               # Kanseller en SO i Odoo")
        sys.exit(1)

    # --- Ops commands ---
    if args[0] == "--replay":
        if len(args) < 2:
            print("Bruk: python -m src --replay <pdf-fil>")
            sys.exit(1)
        cmd_replay(args[1])
        return

    if args[0] == "--dead-letters":
        cmd_dead_letters()
        return

    if args[0] == "--events":
        query = args[1] if len(args) > 1 else None
        cmd_events(query)
        return

    if args[0] == "--rollback":
        if len(args) < 2:
            print("Bruk: python -m src --rollback <SO-navn>")
            sys.exit(1)
        cmd_rollback(args[1])
        return

    # --- Email monitor commands ---
    if args[0] == "--poll-email":
        push = "--push" in args
        cmd_poll_email(push_to_odoo=push)
        return

    # --- DHL tracking commands ---
    if args[0] == "--set-tracking":
        if len(args) < 3:
            print("Bruk: python -m src --set-tracking <SO-navn> <trackingnummer>")
            sys.exit(1)
        cmd_set_tracking(args[1], args[2])
        return

    if args[0] == "--track":
        if len(args) < 2:
            print("Bruk: python -m src --track <SO-navn>")
            sys.exit(1)
        cmd_track(args[1])
        return

    if args[0] == "--track-all":
        cmd_track_all()
        return

    # --- Order processing commands ---
    push_to_odoo = "--push" in args

    # Remove flags from args
    args = [a for a in args if a != "--push"]

    # Initialize Odoo service for single file mode
    odoo_service = None
    if push_to_odoo and args[0] != "--batch":
        try:
            odoo_service = _init_odoo_service()
            print("Odoo-tilkobling: OK")
        except Exception as e:
            print(f"Odoo-tilkobling FEILET: {e}")
            sys.exit(1)

    if args[0] == "--batch":
        pdf_dir = args[1] if len(args) > 1 else "testeposter"
        output_dir = "output"
        if "-o" in args:
            idx = args.index("-o")
            output_dir = args[idx + 1]
        process_batch(pdf_dir, output_dir, push_to_odoo)
    else:
        process_single(
            args[0], "output",
            push_to_odoo=push_to_odoo,
            odoo_service=odoo_service,
        )


if __name__ == "__main__":
    main()
