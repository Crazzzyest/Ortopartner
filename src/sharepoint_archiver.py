"""SharePoint archiver: uploads order documents to SharePoint with consistent structure."""

from __future__ import annotations

import logging
from pathlib import Path

import requests

from .graph_client import GraphClient

logger = logging.getLogger(__name__)

# Archive folder structure:
#   Documents/Ordrer/{year}/{order_number}/
#     {order_number}_bestilling.pdf
#     {order_number}_konfigurasjonsark.pdf
#     {order_number}_parsed.json

SHAREPOINT_SITE = "ortopartner.sharepoint.com:/sites/AFKI"


class SharePointArchiver:
    """Uploads order documents to SharePoint via Microsoft Graph API."""

    def __init__(self, graph_client: GraphClient, site_id: str | None = None):
        self._graph = graph_client
        self._site_id = site_id
        self._drive_id: str | None = None

    def _ensure_ids(self) -> None:
        """Resolve site ID and drive ID if not already set."""
        if self._drive_id:
            return

        if not self._site_id:
            data = self._graph._get(
                f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE}"
            )
            self._site_id = data["id"]

        drives = self._graph._get(
            f"https://graph.microsoft.com/v1.0/sites/{self._site_id}/drives"
        )
        for d in drives.get("value", []):
            if d["name"] == "Documents":
                self._drive_id = d["id"]
                break

        if not self._drive_id:
            # Fallback to first drive
            self._drive_id = drives["value"][0]["id"]

        logger.info("SharePoint drive: %s", self._drive_id)

    def archive_order(
        self,
        order_number: str,
        year: str,
        pdf_path: Path | str,
        month: str = "00",
        json_path: Path | str | None = None,
        document_type: str = "purchase_order",
    ) -> list[str]:
        """Upload order documents to SharePoint.

        Structure: Ordrer/{year}/{month}/{order_number}/{filename}

        Returns list of uploaded file URLs.
        """
        self._ensure_ids()

        folder_path = f"Ordrer/{year}/{month}/{order_number}"
        uploaded: list[str] = []

        # Determine filename suffix based on document type
        suffix = "konfigurasjonsark" if document_type == "configuration_sheet" else "bestilling"

        # Upload PDF
        pdf_path = Path(pdf_path)
        if pdf_path.exists():
            remote_name = f"{order_number}_{suffix}.pdf"
            url = self._upload_file(folder_path, remote_name, pdf_path)
            if url:
                uploaded.append(url)

        # Upload parsed JSON
        if json_path:
            json_path = Path(json_path)
            if json_path.exists():
                remote_name = f"{order_number}_parsed.json"
                url = self._upload_file(folder_path, remote_name, json_path)
                if url:
                    uploaded.append(url)

        return uploaded

    def _upload_file(self, folder_path: str, filename: str, local_path: Path) -> str | None:
        """Upload a single file to SharePoint. Creates folders as needed.

        Returns the web URL of the uploaded file, or None on failure.
        """
        try:
            content = local_path.read_bytes()
            # Graph API auto-creates intermediate folders with this endpoint
            url = (
                f"https://graph.microsoft.com/v1.0/drives/{self._drive_id}"
                f"/root:/{folder_path}/{filename}:/content"
            )
            resp = requests.put(
                url,
                headers={
                    "Authorization": f"Bearer {self._graph._ensure_token()}",
                    "Content-Type": "application/octet-stream",
                },
                data=content,
                timeout=60,
            )
            resp.raise_for_status()

            web_url = resp.json().get("webUrl", "")
            logger.info("Lastet opp: %s/%s", folder_path, filename)
            return web_url

        except Exception as e:
            logger.error("Feil ved opplasting av %s/%s: %s", folder_path, filename, e)
            return None
