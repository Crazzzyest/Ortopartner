"""DHL tracking synchronization with Odoo."""

from __future__ import annotations

import logging

from .dhl_client import DhlClient
from .models import TrackingUpdate
from .odoo_client import OdooClient

logger = logging.getLogger(__name__)


class DhlTracker:
    """Syncs DHL tracking status with Odoo sale orders and pickings."""

    def __init__(self, dhl_client: DhlClient, odoo_client: OdooClient):
        self._dhl = dhl_client
        self._odoo = odoo_client

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
        self._odoo.write("stock.picking", picking_id, {
            "carrier_tracking_ref": tracking_number,
        })

        # Log a message on the picking
        self._odoo.call("stock.picking", "message_post", [[picking_id]], {
            "body": f"DHL trackingnummer satt: {tracking_number}",
            "message_type": "comment",
        })

        result.tracking_number = tracking_number
        result.status = "success"
        result.message = f"Trackingnummer {tracking_number} satt pa picking {picking_id}"
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

        # Call DHL tracking API
        try:
            dhl_result = self._dhl.track_shipment(tracking_number)
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

        # Post status update to Odoo picking
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

        # If delivered, try to validate the picking
        if dhl_result.current_status.upper() in ("DELIVERED", "DELIVERY"):
            if tracking_picking["state"] != "done":
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
            logger.info("Ingen apne leveranser med trackingnummer funnet")
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
