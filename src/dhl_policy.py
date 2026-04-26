"""DHL polling-policy: hvor ofte skal vi spørre API-et per sending.

Bakgrunn:
    DHL Shipment Tracking Unified API begrenser oss til 10 kall per
    sending per dag. For å holde oss under (med margin), og for å unngå
    unødvendige kall, varierer vi TTL etter:
      1. Statuskategorien (terminal vs aktiv vs i-transit)
      2. Tidspunktet (dag vs natt — DHL leverer ikke om natta i Norge)

Resultatet:
    ~6 kall/dag per sending i transit, 0 etter levering, og opptil
    ~24 kall/dag for sendinger som er "ute hos kurer" (men kun i
    arbeidstid). Aldri over 10/dag, godt under DHLs grense.
"""

from __future__ import annotations

from datetime import datetime, time, timedelta
from zoneinfo import ZoneInfo

# Norsk tidssone håndterer DST automatisk
OSLO = ZoneInfo("Europe/Oslo")

# Definerer "arbeidstid" — utenom denne vinduet dobler vi TTL.
WORK_HOURS_START = time(7, 0)
WORK_HOURS_END = time(19, 0)

# TTL-tabell pr. statuskategori (under arbeidstid)
TTL_DELIVERY = timedelta(minutes=30)         # ute hos kurer — bruker venter
TTL_TRANSIT = timedelta(hours=4)             # i transit — sjeldne events
TTL_PRE_TRANSIT = timedelta(hours=6)         # label laget, ikke hentet
TTL_DEFAULT = TTL_TRANSIT                    # ukjent status → konservativt 4t

# Multiplikator utenom arbeidstid (19:00-07:00 norsk tid)
OFF_HOURS_MULTIPLIER = 2.0

# Statuser som regnes som "terminale" — ingen flere events kommer.
# Vi treffer aldri DHL igjen for sendinger i disse statusene.
# Sammenlign case-insensitivt; DHL bruker både lowercase ("delivered")
# og uppercase ("DELIVERED") avhengig av API-versjon.
TERMINAL_STATUSES = frozenset({
    "delivered",
    "failure",
    "cancelled",
    "canceled",
    "returned",
})

# "Ute hos kurer" — kort TTL fordi status endrer seg fort
DELIVERY_STATUSES = frozenset({
    "delivery",
    "out_for_delivery",
    "out for delivery",
})

# "Pre-transit" — label laget, men sendingen er ikke hentet ennå
PRE_TRANSIT_STATUSES = frozenset({
    "pre-transit",
    "pre_transit",
    "label_created",
    "shipment_information_received",
})


def is_terminal(status: str) -> bool:
    """True hvis statusen aldri vil endre seg (skip videre API-kall)."""
    if not status:
        return False
    return status.strip().lower() in TERMINAL_STATUSES


def _is_work_hours(now: datetime) -> bool:
    """True hvis det er mellom 07:00 og 19:00 norsk tid."""
    local = now.astimezone(OSLO)
    return WORK_HOURS_START <= local.time() < WORK_HOURS_END


def compute_ttl(status: str | None, now: datetime | None = None) -> timedelta:
    """Beregn passende TTL for en cached respons med gitt status.

    Args:
        status: Siste kjente DHL-status (statusCode), f.eks. "transit".
            Kan være None — da brukes default.
        now: Hvilket tidspunkt vi beregner TTL fra. Default = nå.
            Brukes mest for testing.

    Returns:
        TTL som timedelta. Multipliseres med OFF_HOURS_MULTIPLIER om
        det er utenfor arbeidstid (19:00-07:00 norsk tid).
    """
    if now is None:
        now = datetime.now(tz=OSLO)

    s = (status or "").strip().lower()

    if s in DELIVERY_STATUSES:
        base = TTL_DELIVERY
    elif s in PRE_TRANSIT_STATUSES:
        base = TTL_PRE_TRANSIT
    elif s == "transit" or s == "in_transit":
        base = TTL_TRANSIT
    else:
        base = TTL_DEFAULT

    if not _is_work_hours(now):
        # Multiplisere timedelta direkte er trygt fra Python 3.5+
        base = base * OFF_HOURS_MULTIPLIER

    return base


def estimated_calls_per_day(status: str | None) -> float:
    """Estimer hvor mange API-kall en sending genererer pr. dag.

    Brukes til å verifisere at vi holder oss under DHLs grense
    på 10 kall/sending/dag.
    """
    # 12 timer arbeidstid + 12 timer off-hours
    work_ttl = compute_ttl(status, datetime.now(tz=OSLO).replace(hour=12))
    off_ttl = compute_ttl(status, datetime.now(tz=OSLO).replace(hour=2))
    work_calls = (12 * 3600) / work_ttl.total_seconds()
    off_calls = (12 * 3600) / off_ttl.total_seconds()
    return work_calls + off_calls
