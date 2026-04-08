"""Microsoft Graph API client for reading emails and attachments."""

from __future__ import annotations

import logging
import time
from pathlib import Path

import requests

logger = logging.getLogger(__name__)

# Default mailbox to monitor
DEFAULT_MAILBOX = "ordre@ortopartner.no"


class GraphClient:
    """Thin wrapper around Microsoft Graph API with OAuth2 client credentials."""

    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self._tenant_id = tenant_id
        self._client_id = client_id
        self._client_secret = client_secret
        self._token: str | None = None
        self._token_expires: float = 0

    def _ensure_token(self) -> str:
        """Get a valid access token, refreshing if needed."""
        if self._token and time.time() < self._token_expires - 60:
            return self._token

        resp = requests.post(
            f"https://login.microsoftonline.com/{self._tenant_id}/oauth2/v2.0/token",
            data={
                "grant_type": "client_credentials",
                "client_id": self._client_id,
                "client_secret": self._client_secret,
                "scope": "https://graph.microsoft.com/.default",
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        self._token = data["access_token"]
        self._token_expires = time.time() + data.get("expires_in", 3600)
        logger.debug("Graph token fornyet, utloper om %ds", data.get("expires_in"))
        return self._token

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self._ensure_token()}"}

    def _get(self, url: str, params: dict | None = None) -> dict:
        resp = requests.get(url, headers=self._headers(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def _patch(self, url: str, json_data: dict) -> dict:
        resp = requests.patch(
            url, headers={**self._headers(), "Content-Type": "application/json"},
            json=json_data, timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

    # --- Email operations ---

    def list_messages(
        self,
        mailbox: str = DEFAULT_MAILBOX,
        folder: str = "inbox",
        unread_only: bool = True,
        top: int = 25,
    ) -> list[dict]:
        """List messages in a mailbox folder.

        Returns list of message dicts with id, subject, from, receivedDateTime,
        hasAttachments.
        """
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder}/messages"
        params: dict = {
            "$top": top,
            "$select": "id,subject,from,receivedDateTime,hasAttachments,isRead",
            "$orderby": "receivedDateTime desc",
        }
        if unread_only:
            params["$filter"] = "isRead eq false"

        data = self._get(url, params)
        return data.get("value", [])

    def get_message(self, message_id: str, mailbox: str = DEFAULT_MAILBOX) -> dict:
        """Get a single message with full details."""
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}"
        return self._get(url)

    def list_attachments(self, message_id: str, mailbox: str = DEFAULT_MAILBOX) -> list[dict]:
        """List attachments for a message."""
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}/attachments"
        data = self._get(url)
        return data.get("value", [])

    def download_attachment(
        self,
        message_id: str,
        attachment_id: str,
        dest_dir: Path,
        filename: str,
        mailbox: str = DEFAULT_MAILBOX,
    ) -> Path:
        """Download an attachment to a local file. Returns the file path."""
        url = (
            f"https://graph.microsoft.com/v1.0/users/{mailbox}"
            f"/messages/{message_id}/attachments/{attachment_id}"
        )
        data = self._get(url)

        import base64
        content_bytes = base64.b64decode(data["contentBytes"])

        dest_dir.mkdir(parents=True, exist_ok=True)
        filepath = dest_dir / filename
        filepath.write_bytes(content_bytes)
        logger.info("Lastet ned vedlegg: %s (%d bytes)", filepath, len(content_bytes))
        return filepath

    def mark_as_read(self, message_id: str, mailbox: str = DEFAULT_MAILBOX) -> bool:
        """Mark a message as read. Returns True if successful.

        Requires Mail.ReadWrite permission. Fails gracefully with Mail.Read only.
        """
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}"
        try:
            self._patch(url, {"isRead": True})
            logger.debug("Markert som lest: %s", message_id)
            return True
        except Exception as e:
            logger.warning(
                "Kunne ikke markere som lest (krever Mail.ReadWrite): %s", e
            )
            return False
