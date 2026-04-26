"""DHL tracking synchronization with Odoo.

Compliance-policy (DHL ToS + GDPR):
  - Vi sporer kun trackingnumre som faktisk ligger på en stock.picking
    i Odoo (ToS: "legitimt forretningsformål"). Se _is_tracking_authorized.
  - Hver API-kall logges med kontekst (so_name, picking_id) — sporbarhet
    for revisjon.
  - Cache 15 min for å begrense API-kall (se DhlCache).
  - Cache-data slettes automatisk etter 30 dager.
  - Persondata (signaturnavn) strippes av DhlClient før vi viser dem
    i Odoo-chatter.
"""

from __future__ import annotations

import logging

from .dhl_cache import DhlCache
from .dhl_client import DhlClient
from .dhl_policy import compute_ttl, is_terminal
from .models import TrackingUpdate
from .odoo_client import OdooClient

logger = logging.getLogger(__name__)


class TrackingNotAuthorizedError(Exception):
    """Reises når vi blir bedt om å spore et nr som ikke ligger i Odoo.

    Dette beskytter mot ToS-brudd (sporing av sendinger uten legitimt
    forretningsformål) og mot at en angriper bruker vår API-tilgang
    til å enumerere DHL-numre.
    """


class DhlTracker:
    """Syncs DHL tracking status with Odoo sale orders and pickings."""

    def __init__(
        self,
        dhl_client: DhlClient,
        odoo_client: OdooClient,
        cache: DhlCache | None = None,
    ):
        self._dhl = dhl_client
        self._odoo = odoo_client
        self._cache = cache or DhlCache()
        self._dhl_carrier_id: int | None = None  # lazily resolved

    # ------------------------------------------------------------------
    # ToS / GDPR helpers
    # ------------------------------------------------------------------

    def _is_tracking_authorized(self, tracking_number: str) -> bool:
        """Sjekk at trackingnr ligger på minst én picking i Odoo.

        Hvis vi ikke finner det, har vi ingen legitim grunn til å kalle
        DHL-API-et — DHL-ToS forbyr sporing uten forretningsformål.
        """
        if not tracking_number:
            return False
        pickings = self._odoo.search_read(
            "stock.picking",
            [["carrier_tracking_ref", "=", tracking_number]],
            ["id"], limit=1,
        )
        return bool(pickings)

    def _resolve_dhl_carrier_id(self) -> int | None:
        """Look up the DHL delivery.carrier once and cache the id.

        We set this alongside carrier_tracking_ref so Odoo's built-in UI
        shows the "Tracking" button (a link to dhl.com/tracking) on the
        picking. The actual status polling is still done by sync_tracking()
        — Odoo's plugin only provides a link builder, not API polling.
        """
        if self._dhl_carrier_id is not None:
            return self._dhl_carrier_id
        carriers = self._odoo.search_read(
            "delivery.carrier",
            [["delivery_type", "=", "dhl_rest"]],
            ["id", "name"], limit=1,
        )
        if carriers:
            self._dhl_carrier_id = carriers[0]["id"]
            logger.info(
                "DHL delivery.carrier funnet: %s (id=%d)",
                carriers[0]["name"], self._dhl_carrier_id,
            )
        else:
            logger.info(
                "Ingen delivery.carrier med delivery_type=dhl_rest funnet — "
                "setter kun carrier_tracking_ref uten carrier_id"
            )
        return self._dhl_carrier_id

    # ------------------------------------------------------------------
    # Internal: API-kall med cache + ToS-guard
    # ------------------------------------------------------------------

    def _track_with_cache(
        self,
        tracking_number: str,
        so_name: str,
        picking_id: int,
    ):
        """Kall DHL-API med cache + audit-logging + status-aware TTL.

        Logikk:
          1. ToS-guard: nektes hvis nr ikke ligger på en picking i Odoo
          2. Hvis vi har sett terminal-status før (delivered/failure/...)
             returneres cached for evig — ingen flere API-kall
          3. Ellers: TTL bestemmes av status + tid på døgnet
             (se dhl_policy.compute_ttl)
          4. Cache-miss → API-kall + lagre

        Reiser TrackingNotAuthorizedError hvis ToS-guard slår inn.
        """
        if not self._is_tracking_authorized(tracking_number):
            logger.warning(
                "ToS-guard: nektet sporing av %s (ikke knyttet til picking)",
                tracking_number,
            )
            raise TrackingNotAuthorizedError(
                f"Trackingnr {tracking_number} ligger ikke på en picking i Odoo"
            )

        # 1. Sjekk om vi har en terminal-status cachet — da skal vi
        #    aldri spørre DHL igjen for denne sendingen.
        any_cached = self._cache.get_any(tracking_number)
        if any_cached is not None:
            cached_payload, fetched_at = any_cached
            cached_result = self._dhl._parse_tracking_response(
                tracking_number, cached_payload,
            )
            if is_terminal(cached_result.current_status):
                logger.info(
                    "DHL cache TERMINAL %s status=%s (so=%s, picking=%d) "
                    "— hopper over API-kall (ingen flere events)",
                    tracking_number, cached_result.current_status,
                    so_name, picking_id,
                )
                return cached_result

            # 2. Ikke-terminal: bruk status-aware TTL
            ttl = compute_ttl(cached_result.current_status)
            fresh_cached = self._cache.get(tracking_number, ttl=ttl)
            if fresh_cached is not None:
                logger.info(
                    "DHL cache HIT %s status=%s ttl=%s (so=%s, picking=%d)",
                    tracking_number, cached_result.current_status,
                    ttl, so_name, picking_id,
                )
                return self._dhl._parse_tracking_response(
                    tracking_number, fresh_cached,
                )

        # 3. Cache-miss → API-kall + audit-log
        logger.info(
            "DHL API-kall: %s (so=%s, picking=%d) — legitimt formål bekreftet",
            tracking_number, so_name, picking_id,
        )
        result = self._dhl.track_shipment(tracking_number)

        # Lagre rå respons i cache for neste 15 min
        # Vi henter rå data ved å gjenta kallet på lavt nivå; men det
        # skaper et ekstra kall. I stedet eksponerer vi
        # _last_raw_response på DhlClient. For enkelhet lagrer vi den
        # parsede modellen som JSON (re-parsing er trivielt).
        self._cache.set(
            tracking_number,
            response={"shipments": [{
                "id": tracking_number,
                "status": {
                    "statusCode": result.current_status,
                    "timestamp": result.last_update,
                },
                "estimatedTimeOfDelivery": result.estimated_delivery,
                "events": [
                    {
                        "timestamp": e.timestamp,
                        "statusCode": e.status,
                        "description": e.status_message,
                        "location": {"address": {
                            "addressLocality": e.location_city,
                            "countryCode": e.location_country,
                        }},
                    } for e in result.events
                ],
            }]},
            so_name=so_name,
            picking_id=picking_id,
        )
        return result

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def set_tracking_number(self, so_name: str, tracking_number: str) -> TrackingUpdate:
        """Set a DHL tracking number on a sale order's delivery picking."""
        result = TrackingUpdate(so_name=so_name)

        # Find the SO
        so_data = self._odoo.search_read(
            "sale.order", [["name", "=", so_name]], ["id", "name", "picking_ids"], limit=1
        )
        if not so_data:
            result.status = "error"
            result.message = f"Salgsordre {so_name} ikke funnet i Odoo"
            return result

        picking_ids = so_data[0].get("picking_ids", [])
        if not picking_ids:
            result.status = "error"
            result.message = f"Ingen leveranser (pickings) funnet for {so_name}"
            return result

        # Set tracking on the first outgoing picking
        picking_id = picking_ids[0]
        vals = {"carrier_tracking_ref": tracking_number}
        carrier_id = self._resolve_dhl_carrier_id()
        if carrier_id is not None:
            vals["carrier_id"] = carrier_id
        self._odoo.write("stock.picking", picking_id, vals)

        # Log a message on the picking
        self._odoo.call("stock.picking", "message_post", [[picking_id]], {
            "body": f"DHL trackingnummer satt: {tracking_number}",
            "message_type": "comment",
        })

        result.tracking_number = tracking_number
        result.status = "success"
        result.message = f"Trackingnummer {tracking_number} satt på picking {picking_id}"
        logger.info(result.message)
        return result

    def sync_tracking(self, so_name: str) -> TrackingUpdate:
        """Sync tracking status for a single sale order."""
        result = TrackingUpdate(so_name=so_name)

        # Find SO and its pickings
        so_data = self._odoo.search_read(
            "sale.order", [["name", "=", so_name]], ["id", "name", "picking_ids"], limit=1
        )
        if not so_data:
            result.status = "error"
            result.message = f"Salgsordre {so_name} ikke funnet i Odoo"
            return result

        picking_ids = so_data[0].get("picking_ids", [])
        if not picking_ids:
            result.status = "no_tracking"
            result.message = f"Ingen leveranser funnet for {so_name}"
            return result

        # Read picking details
        pickings = self._odoo.read(
            "stock.picking", picking_ids,
            ["name", "carrier_tracking_ref", "state"],
        )

        # Find the first picking with a tracking ref
        tracking_picking = None
        for p in pickings:
            if p.get("carrier_tracking_ref"):
                tracking_picking = p
                break

        if not tracking_picking:
            result.status = "no_tracking"
            result.message = f"Ingen trackingnummer satt for {so_name}"
            return result

        tracking_number = tracking_picking["carrier_tracking_ref"]
        picking_id = tracking_picking["id"]
        result.tracking_number = tracking_number

        # Call DHL tracking API (med cache + ToS-guard)
        try:
            dhl_result = self._track_with_cache(
                tracking_number, so_name=so_name, picking_id=picking_id,
            )
        except TrackingNotAuthorizedError as e:
            result.status = "error"
            result.message = f"ToS: {e}"
            return result
        except ValueError as e:
            result.status = "error"
            result.message = f"DHL: {e}"
            return result
        except ConnectionError as e:
            result.status = "error"
            result.message = f"DHL tilkoblingsfeil: {e}"
            return result

        result.dhl_status = dhl_result.current_status
        result.events = dhl_result.events

        # Post status update to Odoo picking. status_message er allerede
        # privacy-strippet av DhlClient (signatur-navn → ***).
        status_msg = f"DHL status: {dhl_result.current_status}"
        if dhl_result.last_update:
            status_msg += f" (oppdatert {dhl_result.last_update})"
        if dhl_result.events:
            latest = dhl_result.events[0]
            status_msg += f"\nSiste hendelse: {latest.status_message}"
            if latest.location_city:
                status_msg += f" ({latest.location_city})"

        self._odoo.call("stock.picking", "message_post", [[picking_id]], {
            "body": status_msg,
            "message_type": "comment",
        })

        # If delivered, try to validate the picking. Unified API
        # rapporterer "delivered" (lowercase), MyDHL rapporterte
        # "DELIVERED" — vi sammenligner case-insensitive.
        terminal = dhl_result.current_status.upper() in (
            "DELIVERED", "DELIVERY",
        )
        if terminal and tracking_picking["state"] != "done":
            try:
                self._odoo.call("stock.picking", "button_validate", [[picking_id]])
                status_msg += "\nLeveranse markert som mottatt i Odoo."
                logger.info("Picking %d markert som levert", picking_id)
            except Exception as e:
                logger.warning(
                    "Kunne ikke auto-validere picking %d: %s", picking_id, e
                )

        result.status = "success"
        result.message = status_msg
        logger.info("Sporing synket for %s: %s", so_name, dhl_result.current_status)
        return result

    def sync_all_open(self) -> list[TrackingUpdate]:
        """Sync tracking for all open sale orders that have tracking numbers."""
        # Find all pickings with tracking ref that are not done/cancelled
        pickings = self._odoo.search_read(
            "stock.picking",
            [
                ["carrier_tracking_ref", "!=", False],
                ["state", "not in", ["done", "cancel"]],
                ["sale_id", "!=", False],
            ],
            ["name", "carrier_tracking_ref", "sale_id", "state"],
        )

        if not pickings:
            logger.info("Ingen åpne leveranser med trackingnummer funnet")
            return []

        results = []
        seen_so: set[str] = set()

        for p in pickings:
            so_name = p["sale_id"][1] if isinstance(p["sale_id"], (list, tuple)) else str(p["sale_id"])
            if so_name in seen_so:
                continue
            seen_so.add(so_name)

            logger.info("Synker sporing for %s (tracking: %s)", so_name, p["carrier_tracking_ref"])
            result = self.sync_tracking(so_name)
            results.append(result)

        return results

    def purge_cache(self) -> int:
        """Slett DHL cache-data eldre enn 30 dager. Returnerer antall slettet."""
        return self._cache.purge_expired()

    def cache_stats(self) -> dict:
        """Diagnostikk-info for cache (brukes f.eks. fra dashboard)."""
        return self._cache.stats()
