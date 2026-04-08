"""DHL Express REST API client for shipment tracking."""

from __future__ import annotations

import logging
import time

import requests

from .models import DhlTrackingEvent, DhlTrackingResult

logger = logging.getLogger(__name__)


class DhlClient:
    """REST client for DHL Express MyDHL API (tracking)."""

    def __init__(self, api_key: str, api_secret: str, base_url: str | None = None):
        self.api_key = api_key
        self.api_secret = api_secret
        self.base_url = (base_url or "https://express.api.dhl.com/mydhlapi").rstrip("/")
        self._session = requests.Session()
        self._session.auth = (api_key, api_secret)
        self._session.headers.update({"Accept": "application/json"})

    def track_shipment(self, tracking_number: str) -> DhlTrackingResult:
        """Track a single shipment by tracking number.

        Returns DhlTrackingResult with events and current status.
        Raises ValueError if tracking number not found.
        Raises ConnectionError on API failures.
        """
        url = f"{self.base_url}/shipments/{tracking_number}/tracking"
        logger.debug("DHL tracking request: %s", url)

        for attempt in range(2):
            try:
                resp = self._session.get(url, timeout=15)
                break
            except (requests.ConnectionError, requests.Timeout) as e:
                if attempt == 0:
                    logger.warning("DHL API tilkoblingsfeil, prover igjen om 2s: %s", e)
                    time.sleep(2)
                else:
                    raise ConnectionError(f"DHL API utilgjengelig: {e}") from e

        # Handle rate limiting
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", "5"))
            logger.warning("DHL rate limit, venter %ds", retry_after)
            time.sleep(retry_after)
            resp = self._session.get(url, timeout=15)

        if resp.status_code == 404:
            raise ValueError(f"Trackingnummer ikke funnet: {tracking_number}")

        if resp.status_code == 401:
            raise ConnectionError("DHL autentisering feilet. Sjekk DHL_API_KEY/DHL_API_SECRET.")

        if resp.status_code != 200:
            raise ConnectionError(
                f"DHL API feil (HTTP {resp.status_code}): {resp.text[:300]}"
            )

        data = resp.json()
        return self._parse_tracking_response(tracking_number, data)

    def track_multiple(self, tracking_numbers: list[str]) -> list[DhlTrackingResult]:
        """Track multiple shipments. Returns results for each (errors logged, not raised)."""
        results = []
        for tn in tracking_numbers:
            try:
                result = self.track_shipment(tn)
                results.append(result)
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

    def _parse_tracking_response(
        self, tracking_number: str, data: dict
    ) -> DhlTrackingResult:
        """Parse DHL tracking API response into our model."""
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
            location = ev.get("location", {}).get("address", {})
            events.append(
                DhlTrackingEvent(
                    timestamp=ev.get("timestamp", ""),
                    status=ev.get("status", ev.get("statusCode", "")),
                    status_message=ev.get("description", ev.get("statusMessage", "")),
                    location_city=location.get("addressLocality", location.get("city")),
                    location_country=location.get("countryCode"),
                )
            )

        # Current status from the most recent event or top-level status
        current_status = shipment.get("status", "")
        if not current_status and events:
            current_status = events[0].status

        last_update = None
        if events:
            last_update = events[0].timestamp

        estimated_delivery = shipment.get("estimatedDeliveryDate")

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
