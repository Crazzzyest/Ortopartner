"""SQLite-cache for DHL tracking-respons.

Formål:
  1. Begrense API-kall mot DHL (rate limit + ToS-vennlig)
  2. Sørge for at cached persondata slettes innen 30 dager (GDPR)

Schema:
  tracking_number TEXT PRIMARY KEY
  response_json   TEXT     — rå API-respons som JSON
  fetched_at      TEXT     — ISO 8601, når responsen ble hentet
  so_name         TEXT     — kontekst: hvilken SO bestilte sporing
  picking_id      INTEGER  — kontekst: hvilken picking
  last_used       TEXT     — ISO 8601, sist gang vi leste fra cache

TTL-policy:
  Selve TTL bestemmes av kalleren (typisk DhlTracker via dhl_policy)
  basert på status og tid på døgnet. Default-TTL er 4 timer — som gir
  6 kall/dag per sending, godt under DHLs grense på 10/dag.

  - get(tn, ttl=...) — returnerer cached hvis fersk per gitt TTL
  - get_any(tn)      — returnerer cached uansett alder (for terminal-sjekk)
  - 30 dager retention: hele raden slettes uansett (GDPR)

Vi kjører `purge_expired()` ved hver `get()` og `set()` slik at
slettelogikken ikke krever en egen cron — så lenge man leser eller
skriver fra cachen jevnlig, holdes data ren.
"""

from __future__ import annotations

import json
import logging
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

logger = logging.getLogger(__name__)

DEFAULT_CACHE_PATH = Path("output/dhl_cache.sqlite")

# Default-TTL hvis kalleren ikke spesifiserer noe annet.
# DHL tillater 10 kall/sending/dag → 24/10 = 2.4t minimum.
# 4 timer gir 6 kall/dag per sending med god margin.
FRESH_TTL = timedelta(hours=4)

# Hvor lenge data lagres totalt før de auto-slettes (GDPR-grense)
RETENTION = timedelta(days=30)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


class DhlCache:
    """SQLite-cache for DHL tracking-respons med innebygd retention-policy."""

    def __init__(self, path: str | Path = DEFAULT_CACHE_PATH):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.path)

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS dhl_tracking_cache (
                    tracking_number TEXT PRIMARY KEY,
                    response_json   TEXT NOT NULL,
                    fetched_at      TEXT NOT NULL,
                    so_name         TEXT,
                    picking_id      INTEGER,
                    last_used       TEXT NOT NULL
                )
            """)
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_fetched_at
                ON dhl_tracking_cache (fetched_at)
            """)

    def get(self, tracking_number: str,
            ttl: timedelta | None = None) -> dict | None:
        """Hent fresh cached respons eller None.

        Args:
            tracking_number: trackingnr å slå opp
            ttl: hvor gammel raden får være. Default = FRESH_TTL (4 timer).
                Settes typisk av DhlTracker via dhl_policy.compute_ttl().

        Kjører også purge av rader eldre enn 30 dager som en bivirkning.
        """
        self.purge_expired()

        effective_ttl = ttl if ttl is not None else FRESH_TTL
        cutoff = (datetime.now(timezone.utc) - effective_ttl).isoformat()
        with self._connect() as conn:
            cur = conn.execute(
                "SELECT response_json, fetched_at FROM dhl_tracking_cache "
                "WHERE tracking_number = ? AND fetched_at >= ?",
                (tracking_number, cutoff),
            )
            row = cur.fetchone()
            if not row:
                return None

            # Oppdater last_used
            conn.execute(
                "UPDATE dhl_tracking_cache SET last_used = ? WHERE tracking_number = ?",
                (_now_iso(), tracking_number),
            )
            logger.debug("Cache HIT %s (fetched %s, ttl=%s)",
                         tracking_number, row[1], effective_ttl)
            return json.loads(row[0])

    def get_any(self, tracking_number: str) -> tuple[dict, datetime] | None:
        """Hent cached respons UANSETT alder (innenfor retention).

        Brukes til terminal-status-sjekk: hvis vi en gang har sett
        status=delivered for en sending, skal vi aldri kalle DHL for
        den igjen — uansett om cached respons er gammel.

        Returnerer (response, fetched_at) eller None.
        """
        self.purge_expired()

        with self._connect() as conn:
            cur = conn.execute(
                "SELECT response_json, fetched_at FROM dhl_tracking_cache "
                "WHERE tracking_number = ?",
                (tracking_number,),
            )
            row = cur.fetchone()
            if not row:
                return None

            fetched_at = datetime.fromisoformat(row[1])
            return json.loads(row[0]), fetched_at

    def set(self, tracking_number: str, response: dict,
            so_name: str | None = None,
            picking_id: int | None = None) -> None:
        """Lagre fersk respons. Erstatter eksisterende cache-rad."""
        self.purge_expired()

        now = _now_iso()
        with self._connect() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO dhl_tracking_cache "
                "(tracking_number, response_json, fetched_at, so_name, picking_id, last_used) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (
                    tracking_number,
                    json.dumps(response, ensure_ascii=False),
                    now, so_name, picking_id, now,
                ),
            )
        logger.debug("Cache SET %s (so=%s)", tracking_number, so_name)

    def purge_expired(self) -> int:
        """Slett alle rader eldre enn RETENTION-grensen (30 dager).

        Returnerer antall rader slettet.
        """
        cutoff = (datetime.now(timezone.utc) - RETENTION).isoformat()
        with self._connect() as conn:
            cur = conn.execute(
                "DELETE FROM dhl_tracking_cache WHERE fetched_at < ?",
                (cutoff,),
            )
            deleted = cur.rowcount
        if deleted:
            logger.info("DHL cache: slettet %d rader eldre enn %d dager",
                        deleted, RETENTION.days)
        return deleted

    def stats(self) -> dict:
        """Returnerer enkle stats for diagnose / dashboard."""
        with self._connect() as conn:
            total = conn.execute(
                "SELECT COUNT(*) FROM dhl_tracking_cache"
            ).fetchone()[0]
            cutoff_fresh = (datetime.now(timezone.utc) - FRESH_TTL).isoformat()
            fresh = conn.execute(
                "SELECT COUNT(*) FROM dhl_tracking_cache WHERE fetched_at >= ?",
                (cutoff_fresh,),
            ).fetchone()[0]
            oldest = conn.execute(
                "SELECT MIN(fetched_at) FROM dhl_tracking_cache"
            ).fetchone()[0]
        return {
            "total_entries": total,
            "fresh_entries": fresh,
            "oldest_fetched_at": oldest,
            "retention_days": RETENTION.days,
            "default_ttl_hours": FRESH_TTL.total_seconds() / 3600,
        }

    def clear_all(self) -> int:
        """Slett alle cached rader. Returnerer antall slettet.

        Brukes ved testing eller hvis Ortopartner trekker tilbake samtykke.
        """
        with self._connect() as conn:
            cur = conn.execute("DELETE FROM dhl_tracking_cache")
            deleted = cur.rowcount
        logger.info("DHL cache: tømt totalt (%d rader slettet)", deleted)
        return deleted
