"""DHL Shipment Tracking Unified API client.

Vi bruker DHL Shipment Tracking — Unified-API-et (api-eu.dhl.com/track),
ikke MyDHL API. Forskjeller:
  - Auth: 'DHL-API-Key' header (ikke Basic Auth)
  - URL:  https://api-eu.dhl.com/track/shipments
  - Param: trackingNumber + service=express

Vi følger DHLs ToS:
  - Cache 15 min for å begrense API-kall (separat modul, dhl_cache)
  - Strip persondata (signaturnavn) fra events før de eksponeres
  - Eksponentiell backoff ved 429

Caching gjøres av DhlTracker, ikke her — denne klassen er en ren
HTTP-wrapper. Det gjør den lettere å teste.
"""

from __future__ import annotations

import logging
import re
import time

import requests

from .models import DhlTrackingEvent, DhlTrackingResult

logger = logging.getLogger(__name__)

# DHL Express signaturer som vises i event-description ved levering.
# Vi maskerer disse før de skrives til Odoo-chatter (GDPR).
_SIGNATURE_PATTERNS = (
    re.compile(r"(Signed by\s+)([A-ZÆØÅ][\w\-\.\s]{1,60})", re.IGNORECASE),
    re.compile(r"(Signature\s*:\s*)([A-ZÆØÅ][\w\-\.\s]{1,60})", re.IGNORECASE),
    re.compile(r"(Recipient\s*:\s*)([A-ZÆØÅ][\w\-\.\s]{1,60})", re.IGNORECASE),
)

DEFAULT_BASE_URL = "https://api-eu.dhl.com/track"


def _strip_signature(text: str) -> str:
    """Erstatt persondata-signaturer i tekst med '***'."""
    if not text:
        return text
    for pat in _SIGNATURE_PATTERNS:
        text = pat.sub(r"\1***", text)
    return text


class DhlClient:
    """REST-klient for DHL Shipment Tracking Unified API.

    api_secret tas inn for bakoverkompatibilitet med tidligere kode, men
    brukes ikke — Unified-API-et autentiserer kun med API-key i header.
    """

    def __init__(self, api_key: str, api_secret: str | None = None,
                 base_url: str | None = None):
        self.api_key = api_key
        self.api_secret = api_secret  # ubrukt, beholdes for bakoverkompat
        self.base_url = (base_url or DEFAULT_BASE_URL).rstrip("/")
        self._session = requests.Session()
        self._session.headers.update({
            "DHL-API-Key": api_key,
            "Accept": "application/json",
        })

    def track_shipment(self, tracking_number: str,
                       service: str = "express") -> DhlTrackingResult:
        """Hent sporing for ett trackingnummer.

        Returnerer DhlTrackingResult.
        Kaster ValueError ved 404 (nummer ikke funnet).
        Kaster ConnectionError ved varig API-feil.
        """
        url = f"{self.base_url}/shipments"
        params = {"trackingNumber": tracking_number, "service": service}
        logger.debug("DHL Unified tracking: %s %s", url, params)

        resp = self._request_with_backoff(url, params)

        if resp.status_code == 404:
            raise ValueError(f"Trackingnummer ikke funnet: {tracking_number}")

        if resp.status_code == 401:
            raise ConnectionError(
                "DHL autentisering feilet. Sjekk DHL_API_KEY (Unified Tracking app)."
            )

        if resp.status_code != 200:
            raise ConnectionError(
                f"DHL API feil (HTTP {resp.status_code}): {resp.text[:300]}"
            )

        return self._parse_tracking_response(tracking_number, resp.json())

    def track_multiple(self, tracking_numbers: list[str]) -> list[DhlTrackingResult]:
        """Spor flere sendinger. Feil per nr logges, ikke kastes."""
        results: list[DhlTrackingResult] = []
        for tn in tracking_numbers:
            try:
                results.append(self.track_shipment(tn))
            except (ValueError, ConnectionError) as e:
                logger.warning("Kunne ikke spore %s: %s", tn, e)
                results.append(
                    DhlTrackingResult(
                        tracking_number=tn,
                        current_status="ERROR",
                        events=[],
                    )
                )
        return results

    # ------------------------------------------------------------------
    # Interne metoder
    # ------------------------------------------------------------------

    def _request_with_backoff(self, url: str, params: dict,
                              max_retries: int = 3) -> requests.Response:
        """GET med eksponentiell backoff ved 429 og tilkoblingsfeil.

        DHL Unified Tracking er strengt rate-limited (~5 kall/sek burst,
        og en daglig grense). Vi respekterer Retry-After-headeren når
        den finnes, ellers backoff 2s, 4s, 8s.
        """
        backoff = 2.0
        last_exc: Exception | None = None
        for attempt in range(max_retries + 1):
            try:
                resp = self._session.get(url, params=params, timeout=15)
            except (requests.ConnectionError, requests.Timeout) as e:
                last_exc = e
                if attempt < max_retries:
                    logger.warning(
                        "DHL tilkoblingsfeil (forsøk %d/%d): %s — venter %.0fs",
                        attempt + 1, max_retries + 1, e, backoff,
                    )
                    time.sleep(backoff)
                    backoff *= 2
                    continue
                raise ConnectionError(f"DHL API utilgjengelig: {e}") from e

            if resp.status_code == 429 and attempt < max_retries:
                wait = float(resp.headers.get("Retry-After", str(backoff)))
                logger.warning(
                    "DHL rate limit (429), venter %.0fs (forsøk %d/%d)",
                    wait, attempt + 1, max_retries + 1,
                )
                time.sleep(wait)
                backoff *= 2
                continue

            return resp

        if last_exc:
            raise ConnectionError(f"DHL API utilgjengelig: {last_exc}") from last_exc
        raise ConnectionError("DHL API: utløpt antall retries")

    def _parse_tracking_response(
        self, tracking_number: str, data: dict
    ) -> DhlTrackingResult:
        """Parse Unified Tracking-respons til vår modell.

        Schema (forenklet):
          shipments[0].status.statusCode      → current_status
          shipments[0].status.timestamp       → last_update
          shipments[0].estimatedTimeOfDelivery → estimated_delivery
          shipments[0].events[].timestamp     → events[].timestamp
          shipments[0].events[].statusCode    → events[].status
          shipments[0].events[].description   → events[].status_message
          shipments[0].events[].location.address.addressLocality → location_city
        """
        shipments = data.get("shipments", [])
        if not shipments:
            return DhlTrackingResult(
                tracking_number=tracking_number,
                current_status="UNKNOWN",
                events=[],
            )

        shipment = shipments[0]
        events_raw = shipment.get("events", [])
        events: list[DhlTrackingEvent] = []

        for ev in events_raw:
            loc_addr = (ev.get("location") or {}).get("address") or {}
            city = loc_addr.get("addressLocality")
            # DHL formaterer ofte som "BERGEN - NORWAY" — strip suffix
            if city and " - " in city:
                city = city.split(" - ", 1)[0].title()
            description = _strip_signature(ev.get("description", ""))
            events.append(
                DhlTrackingEvent(
                    timestamp=ev.get("timestamp", ""),
                    status=ev.get("statusCode", ev.get("status", "")),
                    status_message=description,
                    location_city=city,
                    location_country=loc_addr.get("countryCode"),
                )
            )

        # current_status fra status-objektet (Unified har dette som dict)
        status_obj = shipment.get("status") or {}
        if isinstance(status_obj, dict):
            current_status = status_obj.get("statusCode", "")
            last_update = status_obj.get("timestamp")
        else:
            # Fallback hvis API-en endrer schema tilbake
            current_status = str(status_obj)
            last_update = events[0].timestamp if events else None

        if not current_status and events:
            current_status = events[0].status

        estimated_delivery = (
            shipment.get("estimatedTimeOfDelivery")
            or shipment.get("estimatedDeliveryDate")
        )

        logger.info(
            "DHL sporing %s: status=%s, %d hendelser",
            tracking_number, current_status, len(events),
        )

        return DhlTrackingResult(
            tracking_number=tracking_number,
            current_status=current_status,
            last_update=last_update,
            estimated_delivery=estimated_delivery,
            events=events,
        )
