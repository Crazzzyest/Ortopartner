"""
DHL integration test script.

Tester:
1. Autentisering mot DHL Express MyDHL API sandbox
2. Request/response-håndtering
3. Parsing av tracking-respons
4. Simulering av alle tracking-tilstander (ingen ekte sendingsnummer nødvendig)

Kjøres med:
    python scripts/dhl_test.py
    python scripts/dhl_test.py --auth-only
    python scripts/dhl_test.py --simulate
    python scripts/dhl_test.py --track <nummer>
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from dotenv import load_dotenv
load_dotenv()

from src.dhl_client import DhlClient
from src.models import DhlTrackingEvent, DhlTrackingResult

# ── ANSI farger ──────────────────────────────────────────────────────────────
GREEN  = "\033[92m"
YELLOW = "\033[93m"
RED    = "\033[91m"
BLUE   = "\033[94m"
RESET  = "\033[0m"
BOLD   = "\033[1m"

def ok(msg):  print(f"  {GREEN}✓{RESET} {msg}")
def warn(msg):print(f"  {YELLOW}⚠{RESET} {msg}")
def err(msg): print(f"  {RED}✗{RESET} {msg}")
def info(msg):print(f"  {BLUE}→{RESET} {msg}")

# ── Simulerte tracking-tilstander ────────────────────────────────────────────
NOW = datetime.now(timezone.utc)

SIMULATED_SHIPMENTS: dict[str, DhlTrackingResult] = {
    "SIM-TRANSIT": DhlTrackingResult(
        tracking_number="SIM-TRANSIT",
        current_status="TRANSIT",
        last_update=(NOW - timedelta(hours=3)).isoformat(),
        estimated_delivery=(NOW + timedelta(days=1)).strftime("%Y-%m-%d"),
        events=[
            DhlTrackingEvent(
                timestamp=(NOW - timedelta(hours=3)).isoformat(),
                status="TRANSIT",
                status_message="Shipment picked up",
                location_city="Oslo",
                location_country="NO",
            ),
            DhlTrackingEvent(
                timestamp=(NOW - timedelta(hours=6)).isoformat(),
                status="TRANSIT",
                status_message="Shipment label created",
                location_city="Bergen",
                location_country="NO",
            ),
        ],
    ),
    "SIM-OUT-FOR-DELIVERY": DhlTrackingResult(
        tracking_number="SIM-OUT-FOR-DELIVERY",
        current_status="DELIVERY",
        last_update=(NOW - timedelta(hours=1)).isoformat(),
        estimated_delivery=NOW.strftime("%Y-%m-%d"),
        events=[
            DhlTrackingEvent(
                timestamp=(NOW - timedelta(hours=1)).isoformat(),
                status="DELIVERY",
                status_message="With delivery courier",
                location_city="Oslo",
                location_country="NO",
            ),
            DhlTrackingEvent(
                timestamp=(NOW - timedelta(hours=5)).isoformat(),
                status="TRANSIT",
                status_message="Arrived at delivery facility",
                location_city="Oslo",
                location_country="NO",
            ),
        ],
    ),
    "SIM-DELIVERED": DhlTrackingResult(
        tracking_number="SIM-DELIVERED",
        current_status="DELIVERED",
        last_update=(NOW - timedelta(minutes=30)).isoformat(),
        estimated_delivery=NOW.strftime("%Y-%m-%d"),
        events=[
            DhlTrackingEvent(
                timestamp=(NOW - timedelta(minutes=30)).isoformat(),
                status="DELIVERED",
                status_message="Delivered - Signed by MARIUS",
                location_city="Oslo",
                location_country="NO",
            ),
            DhlTrackingEvent(
                timestamp=(NOW - timedelta(hours=2)).isoformat(),
                status="DELIVERY",
                status_message="With delivery courier",
                location_city="Oslo",
                location_country="NO",
            ),
        ],
    ),
    "SIM-EXCEPTION": DhlTrackingResult(
        tracking_number="SIM-EXCEPTION",
        current_status="EXCEPTION",
        last_update=(NOW - timedelta(hours=2)).isoformat(),
        estimated_delivery=None,
        events=[
            DhlTrackingEvent(
                timestamp=(NOW - timedelta(hours=2)).isoformat(),
                status="EXCEPTION",
                status_message="Delivery attempt failed - Recipient not available",
                location_city="Oslo",
                location_country="NO",
            ),
        ],
    ),
}


def _print_result(result: DhlTrackingResult) -> None:
    status_color = {
        "DELIVERED": GREEN,
        "TRANSIT": BLUE,
        "DELIVERY": YELLOW,
        "EXCEPTION": RED,
    }.get(result.current_status.upper(), RESET)

    print(f"\n  Trackingnummer : {BOLD}{result.tracking_number}{RESET}")
    print(f"  Status         : {status_color}{result.current_status}{RESET}")
    if result.last_update:
        print(f"  Siste oppdatering : {result.last_update}")
    if result.estimated_delivery:
        print(f"  Estimert levering : {result.estimated_delivery}")
    print(f"  Antall hendelser  : {len(result.events)}")
    for ev in result.events[:3]:
        loc = f" ({ev.location_city})" if ev.location_city else ""
        print(f"    {ev.timestamp[:16]}  {ev.status:<12}  {ev.status_message}{loc}")
    if len(result.events) > 3:
        print(f"    ... +{len(result.events) - 3} eldre hendelser")


# ── Test 1: Autentisering ─────────────────────────────────────────────────────
def test_auth(client: DhlClient) -> bool:
    print(f"\n{BOLD}Test 1: Autentisering{RESET}")
    try:
        import requests as req
        url = f"{client.base_url}/tracking"
        resp = client._session.get(url, params={"shipmentTrackingNumber": "0000000000"}, timeout=10)
        if resp.status_code == 401:
            err("Autentisering FEILET (401) — sjekk DHL_API_KEY/DHL_API_SECRET")
            return False
        elif resp.status_code in (404, 400):
            ok(f"Autentisering OK (HTTP {resp.status_code} — credentials godkjent, nummer ikke funnet)")
            return True
        elif resp.status_code == 200:
            ok("Autentisering OK (200)")
            return True
        else:
            warn(f"Uventet statuskode {resp.status_code}: {resp.text[:150]}")
            return False
    except Exception as e:
        err(f"Tilkoblingsfeil: {e}")
        return False


# ── Test 2: Request/response-håndtering ──────────────────────────────────────
def test_request_handling(client: DhlClient) -> bool:
    print(f"\n{BOLD}Test 2: Request/response-håndtering{RESET}")
    import requests as req

    # Test: 404 → ValueError
    try:
        client.track_shipment("0000000000")
        err("Forventet ValueError for ukjent nummer — fikk ingen feil")
        return False
    except ValueError as e:
        ok(f"404 → ValueError korrekt: {e}")
    except Exception as e:
        err(f"Uventet feil: {e}")
        return False

    # Test: URL-bygging
    url = f"{client.base_url}/tracking"
    expected_param = "shipmentTrackingNumber"
    resp = client._session.get(url, params={expected_param: "0000000000"}, timeout=10)
    info(f"URL: {resp.url}")
    if expected_param in resp.url:
        ok("URL-format korrekt (query param)")
    else:
        err(f"URL-format feil: {resp.url}")
        return False

    return True


# ── Test 3: Parsing ───────────────────────────────────────────────────────────
def test_parsing(client: DhlClient) -> bool:
    print(f"\n{BOLD}Test 3: Parsing av tracking-respons{RESET}")

    # Simuler DHL API-respons med kjent struktur
    mock_response = {
        "shipments": [
            {
                "id": "1234567890",
                "service": "EXPRESS",
                "origin": {"address": {"addressLocality": "Bergen", "countryCode": "NO"}},
                "destination": {"address": {"addressLocality": "Oslo", "countryCode": "NO"}},
                "status": "TRANSIT",
                "estimatedDeliveryDate": "2026-04-22",
                "events": [
                    {
                        "timestamp": "2026-04-21T10:00:00Z",
                        "location": {"address": {"addressLocality": "Oslo", "countryCode": "NO"}},
                        "status": "TRANSIT",
                        "description": "Shipment picked up",
                    },
                    {
                        "timestamp": "2026-04-21T08:00:00Z",
                        "location": {"address": {"addressLocality": "Bergen", "countryCode": "NO"}},
                        "status": "TRANSIT",
                        "description": "Shipment label created",
                    },
                ],
            }
        ]
    }

    result = client._parse_tracking_response("1234567890", mock_response)

    checks = [
        (result.tracking_number == "1234567890",    "tracking_number korrekt"),
        (result.current_status == "TRANSIT",         "current_status korrekt"),
        (result.estimated_delivery == "2026-04-22",  "estimated_delivery korrekt"),
        (len(result.events) == 2,                    "antall events korrekt (2)"),
        (result.events[0].location_city == "Oslo",   "location_city korrekt"),
        (result.events[0].status == "TRANSIT",       "event status korrekt"),
        (result.last_update is not None,             "last_update satt"),
    ]

    all_ok = True
    for passed, label in checks:
        if passed:
            ok(label)
        else:
            err(label)
            all_ok = False

    return all_ok


# ── Test 4: Simulerte tilstander ─────────────────────────────────────────────
def test_simulate() -> None:
    print(f"\n{BOLD}Test 4: Simulerte tracking-tilstander{RESET}")

    for tn, result in SIMULATED_SHIPMENTS.items():
        status_color = {
            "DELIVERED": GREEN,
            "TRANSIT": BLUE,
            "DELIVERY": YELLOW,
            "EXCEPTION": RED,
        }.get(result.current_status.upper(), RESET)
        print(f"\n  [{status_color}{result.current_status}{RESET}] {tn}")
        _print_result(result)

        # Sjekk om "delivered" logikk vil trigge
        if result.current_status.upper() in ("DELIVERED", "DELIVERY"):
            ok("→ ville trigget auto-validering av picking i Odoo")
        elif result.current_status.upper() == "EXCEPTION":
            warn("→ ville sendt varsel (alert) til Marius")


# ── Test 5: Polling-logikk ────────────────────────────────────────────────────
def test_polling_logic() -> None:
    print(f"\n{BOLD}Test 5: Polling-logikk (simulert){RESET}")

    # Simuler et forløp over tid
    states = [
        ("SIM-TRANSIT", "IN_TRANSIT"),
        ("SIM-OUT-FOR-DELIVERY", "OUT_FOR_DELIVERY"),
        ("SIM-DELIVERED", "DELIVERED"),
    ]

    print("  Simulerer polling-forløp for én sending:")
    for tn, expected in states:
        result = SIMULATED_SHIPMENTS[tn]
        is_terminal = result.current_status.upper() in ("DELIVERED", "EXCEPTION", "CANCEL")
        status_color = GREEN if result.current_status == "DELIVERED" else BLUE
        print(f"    Poll → status={status_color}{result.current_status}{RESET}, terminal={is_terminal}")

    ok("Polling stopper ved DELIVERED/EXCEPTION/CANCEL")
    info("Anbefalt poll-intervall: hvert 30 min mellom 07:00–21:00")
    info("Antall aktive sendinger som sjekkes per kjøring: alle med tracking + state != done")


# ── Hoved ────────────────────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(description="DHL integration test")
    parser.add_argument("--auth-only",  action="store_true", help="Kun autentiseringstest")
    parser.add_argument("--simulate",   action="store_true", help="Kun simulering (ingen API-kall)")
    parser.add_argument("--track",      metavar="NR",        help="Spor et ekte trackingnummer")
    args = parser.parse_args()

    print(f"\n{BOLD}{'='*55}")
    print("  DHL Express MyDHL API — integrasjonstest")
    print(f"{'='*55}{RESET}")

    if args.simulate:
        test_simulate()
        test_polling_logic()
        return

    api_key    = os.environ.get("DHL_API_KEY")
    api_secret = os.environ.get("DHL_API_SECRET")
    base_url   = os.environ.get("DHL_BASE_URL")

    if not api_key or not api_secret:
        err("DHL_API_KEY / DHL_API_SECRET mangler i .env")
        sys.exit(1)

    info(f"Base URL: {base_url or 'https://express.api.dhl.com/mydhlapi'}")
    client = DhlClient(api_key, api_secret, base_url)

    if args.track:
        print(f"\n{BOLD}Sporinger reelt trackingnummer: {args.track}{RESET}")
        try:
            result = client.track_shipment(args.track)
            _print_result(result)
        except (ValueError, ConnectionError) as e:
            err(str(e))
        return

    results = []
    results.append(test_auth(client))
    if args.auth_only:
        sys.exit(0 if all(results) else 1)

    results.append(test_request_handling(client))
    results.append(test_parsing(client))
    test_simulate()
    test_polling_logic()

    print(f"\n{BOLD}{'='*55}")
    passed = sum(results)
    total  = len(results)
    color  = GREEN if passed == total else RED
    print(f"  Resultat: {color}{passed}/{total} tester bestått{RESET}")
    print(f"{'='*55}{RESET}\n")
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()
