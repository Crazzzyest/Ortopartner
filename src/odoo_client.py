"""Low-level XML-RPC client for Odoo V19.0 Enterprise."""

from __future__ import annotations

import logging
import time
import xmlrpc.client
from typing import Any

logger = logging.getLogger(__name__)


class OdooClient:
    """XML-RPC client wrapping Odoo's external API."""

    def __init__(self, url: str, db: str, username: str, password: str):
        self.url = url.rstrip("/")
        self.db = db
        self.username = username
        self.password = password
        self._uid: int | None = None
        self._common = xmlrpc.client.ServerProxy(
            f"{self.url}/xmlrpc/2/common", allow_none=True
        )
        self._object = xmlrpc.client.ServerProxy(
            f"{self.url}/xmlrpc/2/object", allow_none=True
        )

    def authenticate(self) -> int:
        """Authenticate and return uid. Raises ConnectionError on failure."""
        logger.debug("Authenticating to %s as %s", self.url, self.username)
        try:
            uid = self._common.authenticate(
                self.db, self.username, self.password, {}
            )
        except Exception as e:
            raise ConnectionError(
                f"Kunne ikke koble til Odoo ({self.url}): {e}"
            ) from e

        if not uid:
            raise ConnectionError(
                f"Odoo-autentisering feilet. Sjekk ODOO_USERNAME/ODOO_PASSWORD. "
                f"(URL: {self.url}, DB: {self.db}, bruker: {self.username})"
            )

        self._uid = uid
        logger.info("Autentisert mot Odoo (uid=%d)", uid)
        return uid

    @property
    def uid(self) -> int:
        if self._uid is None:
            raise RuntimeError("Ikke autentisert. Kall authenticate() først.")
        return self._uid

    def _execute(self, model: str, method: str, *args, **kwargs) -> Any:
        """Execute an XML-RPC call with retry on transient failures."""
        for attempt in range(2):
            try:
                result = self._object.execute_kw(
                    self.db, self.uid, self.password,
                    model, method, *args, **kwargs
                )
                logger.debug(
                    "Odoo %s.%s(%s) → %s",
                    model, method,
                    str(args)[:200],
                    str(result)[:200],
                )
                return result
            except (ConnectionError, OSError, xmlrpc.client.ProtocolError) as e:
                if attempt == 0:
                    logger.warning("Odoo-tilkoblingsfeil, prøver igjen om 2s: %s", e)
                    time.sleep(2)
                else:
                    raise ConnectionError(
                        f"Odoo-tilkobling feilet etter retry: {e}"
                    ) from e

    def search(self, model: str, domain: list, limit: int = 0) -> list[int]:
        """Search for record IDs matching domain."""
        kwargs = {}
        if limit:
            kwargs["limit"] = limit
        return self._execute(model, "search", [domain], kwargs)

    def search_read(
        self,
        model: str,
        domain: list,
        fields: list[str],
        limit: int = 0,
        order: str | None = None,
    ) -> list[dict]:
        """Search and read records in one call.

        Args:
            model: Odoo model name (e.g. 'sale.order.line').
            domain: Search domain as list of triples.
            fields: Fields to return.
            limit: Max records (0 = no limit).
            order: Sort order (e.g. 'id asc', 'sequence, id').
        """
        kwargs: dict[str, Any] = {"fields": fields}
        if limit:
            kwargs["limit"] = limit
        if order:
            kwargs["order"] = order
        return self._execute(model, "search_read", [domain], kwargs)

    def read(self, model: str, ids: list[int], fields: list[str]) -> list[dict]:
        """Read specific records by ID."""
        return self._execute(model, "read", [ids, fields])

    def create(self, model: str, vals: dict) -> int:
        """Create a record. Returns new record ID."""
        return self._execute(model, "create", [vals])

    def write(self, model: str, record_id: int, vals: dict) -> bool:
        """Update a record."""
        return self._execute(model, "write", [[record_id], vals])

    def call(self, model: str, method: str, args: list, kwargs: dict | None = None) -> Any:
        """Call any method on a model (e.g. action_confirm)."""
        return self._execute(model, method, args, kwargs or {})
