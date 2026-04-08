"""Alerting for flagged orders and DHL exceptions.

Posts alerts to Odoo (via message_post on SO) and writes to local alert log.
E-mail alerting will be added when Outlook/M365 Graph access is available.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path

from .models import OdooResult, ParsedOrder, TrackingUpdate
from .odoo_client import OdooClient

logger = logging.getLogger(__name__)

ALERT_LOG = Path(__file__).resolve().parent.parent / "output" / "alerts.jsonl"


class AlertService:
    """Sends alerts for flagged orders and delivery exceptions."""

    def __init__(self, odoo_client: OdooClient | None = None):
        self._odoo = odoo_client

    def check_order(self, order: ParsedOrder, odoo_result: OdooResult | None = None) -> list[str]:
        """Check a parsed order for alert conditions. Returns list of alerts sent."""
        alerts: list[str] = []

        # Alert 1: Low confidence — needs manual review
        if order.confidence < 0.8:
            msg = (
                f"Ordre {order.order_number} fra {order.customer_name} "
                f"har lav konfidensverdi ({order.confidence:.0%}). "
                f"Krever manuell kontroll."
            )
            if order.warnings:
                msg += "\nAdvarsler: " + "; ".join(order.warnings)
            alerts.append(self._send_alert("LAV_KONFIDENSVERDI", msg, order, odoo_result))

        # Alert 2: Many warnings (even if confidence is OK)
        if len(order.warnings) >= 3:
            msg = (
                f"Ordre {order.order_number} har {len(order.warnings)} advarsler: "
                + "; ".join(order.warnings)
            )
            alerts.append(self._send_alert("MANGE_ADVARSLER", msg, order, odoo_result))

        # Alert 3: Odoo push failed
        if odoo_result and odoo_result.status == "error":
            msg = (
                f"Odoo-push feilet for ordre {order.order_number}: "
                f"{odoo_result.message}"
            )
            alerts.append(self._send_alert("ODOO_FEIL", msg, order, odoo_result))

        # Alert 4: Unknown products
        if odoo_result and odoo_result.warnings:
            unknown = [w for w in odoo_result.warnings if "Ukjent produkt" in w]
            if unknown:
                msg = (
                    f"Ordre {order.order_number} har {len(unknown)} ukjente produkter: "
                    + "; ".join(unknown)
                )
                alerts.append(self._send_alert("UKJENT_PRODUKT", msg, order, odoo_result))

        # Alert 5: SO could not be confirmed
        if odoo_result and odoo_result.sale_order_id and not odoo_result.so_confirmed:
            msg = (
                f"Salgsordre {odoo_result.so_name} ble opprettet men kunne ikke bekreftes. "
                f"Sjekk ordren manuelt i Odoo."
            )
            alerts.append(self._send_alert("SO_IKKE_BEKREFTET", msg, order, odoo_result))

        return [a for a in alerts if a]

    def check_tracking(self, update: TrackingUpdate) -> list[str]:
        """Check a tracking update for alert conditions."""
        alerts: list[str] = []

        # Alert: DHL exception
        if update.dhl_status and update.dhl_status.upper() in ("EXCEPTION", "FAILURE"):
            msg = (
                f"DHL-unntak for {update.so_name} "
                f"(tracking {update.tracking_number}): {update.dhl_status}. "
                f"{update.message}"
            )
            alerts.append(self._send_alert(
                "DHL_UNNTAK", msg, tracking_update=update,
            ))

        # Alert: Tracking error (API failure etc)
        if update.status == "error":
            msg = (
                f"Sporingsfeil for {update.so_name} "
                f"(tracking {update.tracking_number}): {update.message}"
            )
            alerts.append(self._send_alert(
                "SPORINGSFEIL", msg, tracking_update=update,
            ))

        # Info: Delivered (positive event, not really an alert)
        if update.dhl_status and update.dhl_status.upper() in ("DELIVERED", "DELIVERY"):
            msg = f"Leveranse for {update.so_name} er levert! (tracking {update.tracking_number})"
            alerts.append(self._send_alert(
                "LEVERT", msg, tracking_update=update, severity="info",
            ))

        return [a for a in alerts if a]

    def _send_alert(
        self,
        alert_type: str,
        message: str,
        order: ParsedOrder | None = None,
        odoo_result: OdooResult | None = None,
        tracking_update: TrackingUpdate | None = None,
        severity: str = "warning",
    ) -> str:
        """Send an alert to all configured channels."""
        logger.warning("ALERT [%s]: %s", alert_type, message)

        # 1. Log to local JSONL file
        self._log_to_file(alert_type, message, severity, order, odoo_result, tracking_update)

        # 2. Post to Odoo (on the SO if we have one)
        if self._odoo and odoo_result and odoo_result.sale_order_id:
            self._post_to_odoo(odoo_result.sale_order_id, alert_type, message, severity)

        # 3. E-mail (TODO: when Outlook/M365 Graph access is available)

        return f"[{alert_type}] {message}"

    def _log_to_file(
        self,
        alert_type: str,
        message: str,
        severity: str,
        order: ParsedOrder | None,
        odoo_result: OdooResult | None,
        tracking_update: TrackingUpdate | None,
    ) -> None:
        """Append alert to local JSONL log file."""
        ALERT_LOG.parent.mkdir(parents=True, exist_ok=True)

        entry = {
            "timestamp": datetime.now().isoformat(),
            "type": alert_type,
            "severity": severity,
            "message": message,
        }
        if order:
            entry["order_number"] = order.order_number
            entry["customer"] = order.customer_name
            entry["confidence"] = order.confidence
        if odoo_result:
            entry["so_name"] = odoo_result.so_name
            entry["odoo_status"] = odoo_result.status
        if tracking_update:
            entry["so_name"] = tracking_update.so_name
            entry["tracking_number"] = tracking_update.tracking_number
            entry["dhl_status"] = tracking_update.dhl_status

        with open(ALERT_LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")

    def _post_to_odoo(
        self, so_id: int, alert_type: str, message: str, severity: str
    ) -> None:
        """Post alert as a message on the sale order in Odoo."""
        try:
            prefix = "ADVARSEL" if severity == "warning" else "INFO"
            body = f"<b>[{prefix}] {alert_type}</b><br/>{message}"
            self._odoo.call("sale.order", "message_post", [[so_id]], {
                "body": body,
                "message_type": "comment",
                "subtype_xmlid": "mail.mt_note",
            })
        except Exception as e:
            logger.error("Kunne ikke poste alert til Odoo SO %d: %s", so_id, e)
