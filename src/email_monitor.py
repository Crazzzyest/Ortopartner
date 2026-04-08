"""Email monitor: polls ordre@ortopartner.no for new orders with PDF attachments."""

from __future__ import annotations

import json
import logging
from pathlib import Path

from .event_log import log_dead_letter, log_event, new_correlation_id
from .graph_client import GraphClient, DEFAULT_MAILBOX
from .models import ParsedOrder
from .order_parser import parse_order_pdf
from .validator import needs_manual_review, validate_order

logger = logging.getLogger(__name__)

_PROCESSED_FILE = Path("output/processed_emails.json")


def _load_processed() -> set[str]:
    """Load set of already-processed message IDs from disk."""
    if _PROCESSED_FILE.exists():
        data = json.loads(_PROCESSED_FILE.read_text(encoding="utf-8"))
        return set(data)
    return set()


def _save_processed(ids: set[str]) -> None:
    """Save processed message IDs to disk."""
    _PROCESSED_FILE.parent.mkdir(parents=True, exist_ok=True)
    _PROCESSED_FILE.write_text(
        json.dumps(sorted(ids), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


class EmailMonitor:
    """Polls a mailbox for new emails with PDF attachments and processes them."""

    def __init__(
        self,
        graph_client: GraphClient,
        odoo_service=None,
        archiver=None,
        download_dir: str | Path = "downloads",
        output_dir: str | Path = "output",
        mailbox: str = DEFAULT_MAILBOX,
    ):
        self._graph = graph_client
        self._odoo = odoo_service
        self._archiver = archiver
        self._download_dir = Path(download_dir)
        self._output_dir = Path(output_dir)
        self._mailbox = mailbox

    def poll(self) -> list[dict]:
        """Check for new emails and process any PDF attachments.

        Uses a local file (output/processed_emails.json) to track which
        messages have already been processed, so we never process the same
        email twice — even without Mail.ReadWrite permission.

        Returns a list of result dicts, one per processed PDF.
        """
        results: list[dict] = []
        processed = _load_processed()

        messages = self._graph.list_messages(
            mailbox=self._mailbox,
            unread_only=False,
            top=50,
        )

        if not messages:
            logger.info("Ingen meldinger i %s", self._mailbox)
            return results

        # Filter out already processed
        new_messages = [m for m in messages if m["id"] not in processed]

        if not new_messages:
            logger.info("Ingen nye meldinger (alle %d allerede behandlet)", len(messages))
            return results

        logger.info("Fant %d nye meldinger (av %d totalt)", len(new_messages), len(messages))

        for msg in new_messages:
            msg_id = msg["id"]
            subject = msg.get("subject", "(uten emne)")
            sender = msg.get("from", {}).get("emailAddress", {}).get("address", "?")

            if not msg.get("hasAttachments"):
                logger.debug("Hopper over melding uten vedlegg: %s", subject)
                processed.add(msg_id)
                continue

            logger.info("Behandler e-post: '%s' fra %s", subject, sender)

            # Get attachments
            attachments = self._graph.list_attachments(msg_id, self._mailbox)
            pdf_attachments = [
                a for a in attachments
                if a.get("name", "").lower().endswith(".pdf")
                and a.get("@odata.type") == "#microsoft.graph.fileAttachment"
            ]

            if not pdf_attachments:
                logger.info("  Ingen PDF-vedlegg, hopper over")
                processed.add(msg_id)
                continue

            logger.info("  Fant %d PDF-vedlegg", len(pdf_attachments))

            for att in pdf_attachments:
                cid = new_correlation_id()
                log_event(cid, "email_received", source_file=att["name"],
                          details={"subject": subject, "sender": sender})
                result = self._process_attachment(msg_id, att, subject, sender, cid)
                result["correlation_id"] = cid
                results.append(result)

            processed.add(msg_id)

        # Save updated processed list
        _save_processed(processed)

        return results

    def _process_attachment(
        self, msg_id: str, attachment: dict, subject: str, sender: str,
        cid: str = "",
    ) -> dict:
        """Download a PDF attachment, parse it, and optionally push to Odoo."""
        att_id = attachment["id"]
        filename = attachment["name"]

        result = {
            "email_subject": subject,
            "email_sender": sender,
            "filename": filename,
            "status": "pending",
            "order_number": None,
            "so_name": None,
            "message": "",
        }

        try:
            # 1. Download PDF
            pdf_path = self._graph.download_attachment(
                msg_id, att_id, self._download_dir, filename, self._mailbox,
            )
            log_event(cid, "pdf_downloaded", source_file=filename)

            # 2. Parse PDF
            order = parse_order_pdf(pdf_path)
            result["order_number"] = order.order_number
            log_event(cid, "pdf_parsed", order_number=order.order_number,
                       source_file=filename, status="ok",
                       details={"confidence": order.confidence})

            # 3. Validate
            validate_order(order)
            review = needs_manual_review(order)

            # 4. Save JSON output
            self._save_output(order)

            # 5. Push to Odoo if service is configured
            if self._odoo:
                odoo_result = self._odoo.push_order(order, needs_review=review)
                result["so_name"] = odoo_result.so_name
                result["status"] = odoo_result.status
                result["message"] = odoo_result.message
                log_event(cid, "odoo_push", order_number=order.order_number,
                           status=odoo_result.status,
                           details={"so_name": odoo_result.so_name,
                                    "warnings": odoo_result.warnings})
            else:
                result["status"] = "parsed"
                result["message"] = (
                    f"Ordre {order.order_number} parset "
                    f"(konfidensverdi: {order.confidence:.0%})"
                )

            # 6. Archive to SharePoint
            if self._archiver:
                year = (order.order_date or "")[:4] or "ukjent"
                json_path = self._output_dir / f"{Path(order.source_file).stem}.json"
                try:
                    urls = self._archiver.archive_order(
                        order_number=order.order_number,
                        year=year,
                        pdf_path=pdf_path,
                        json_path=json_path if json_path.exists() else None,
                        document_type=order.document_type,
                    )
                    if urls:
                        log_event(cid, "sharepoint_archived",
                                   order_number=order.order_number,
                                   details={"files": len(urls)})
                except Exception as e:
                    logger.warning("  SharePoint-arkivering feilet: %s", e)

            log_event(cid, "completed", order_number=order.order_number,
                       status=result["status"])

        except Exception as e:
            logger.exception("  Feil ved behandling av %s", filename)
            result["status"] = "error"
            result["message"] = str(e)
            # Determine which stage failed
            stage = "download" if result["order_number"] is None else "odoo_push"
            log_event(cid, "failed", source_file=filename, status="error",
                       details={"error": str(e), "stage": stage})
            log_dead_letter(
                cid, source_file=filename, error=str(e), stage=stage,
                order_number=result.get("order_number"),
                email_subject=subject, email_sender=sender,
            )

        return result

    def _save_output(self, order: ParsedOrder) -> None:
        """Save parsed order as JSON."""
        self._output_dir.mkdir(parents=True, exist_ok=True)
        stem = Path(order.source_file).stem
        out_path = self._output_dir / f"{stem}.json"
        out_path.write_text(
            json.dumps(order.model_dump(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        logger.debug("Lagret output: %s", out_path)
