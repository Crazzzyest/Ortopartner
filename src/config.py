"""Centralized configuration loading from .env."""

from __future__ import annotations

import os
from pathlib import Path

import dotenv

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_loaded = False


def load_config() -> dict[str, str]:
    """Load .env and return config dict. Sets os.environ as side-effect."""
    global _loaded
    if not _loaded:
        env_values = dotenv.dotenv_values(_PROJECT_ROOT / ".env")
        for k, v in env_values.items():
            if v and not os.environ.get(k):
                os.environ[k] = v
        _loaded = True

    return {k: v for k, v in os.environ.items()}


def require_odoo_config() -> dict[str, str]:
    """Load config and validate that all required Odoo keys are present."""
    cfg = load_config()

    required = ["ODOO_URL", "ODOO_DB", "ODOO_USERNAME", "ODOO_PASSWORD"]
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        raise ValueError(
            f"Mangler Odoo-konfigurasjon i .env: {', '.join(missing)}"
        )

    return cfg


def require_graph_config() -> dict[str, str]:
    """Load config and validate that all required Microsoft Graph keys are present."""
    cfg = load_config()

    required = ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET"]
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        raise ValueError(
            f"Mangler Microsoft Graph-konfigurasjon i .env: {', '.join(missing)}"
        )

    return cfg


def is_test_mode() -> bool:
    """Check if test mode is active.

    Test mode prefixes order numbers with TEST- so re-runs of real PDFs
    (from testeposter/) don't collide with historical Ortopartner data
    in staging Odoo. Enable by setting TEST_MODE=1 in .env or via the
    --test-mode CLI flag.
    """
    cfg = load_config()
    return cfg.get("TEST_MODE", "").strip().lower() in ("1", "true", "yes", "on")


def get_test_prefix() -> str:
    """Get the prefix to apply to order numbers in test mode.

    Defaults to 'TEST-'. Override via TEST_PREFIX in .env (e.g. 'TEST2-'
    to start a fresh batch without hitting duplicate check on TEST- runs).
    """
    cfg = load_config()
    return cfg.get("TEST_PREFIX", "TEST-") or "TEST-"
