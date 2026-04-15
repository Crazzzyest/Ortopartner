"""Maps ParsedOrder data to Odoo record IDs (partners, products, UoM)."""

from __future__ import annotations

import logging
import re

from .models import ParsedOrder
from .odoo_client import OdooClient

logger = logging.getLogger(__name__)

# Mapping of Norwegian unit strings to Odoo UoM names
_UOM_MAP: dict[str, str] = {
    "stk": "Units",
    "stk.": "Units",
    "pakke": "Units",
    "pk": "Units",
    "plate": "Units",
    "rull": "Units",
    "sett": "Units",
    "set": "Units",
    "crt": "Units",
    "par": "Units",
    "kg": "kg",
    "m": "m",
    "liter": "Liter(s)",
    "l": "Liter(s)",
}

# Words to strip when normalizing partner names for matching
_STRIP_SUFFIXES = re.compile(
    r"\b(AS|A/S|ANS|DA|NUF|SA|avd\.?|avdeling)\b",
    re.IGNORECASE,
)


_NORDIC_MAP = str.maketrans("æøåäöü", "aoaaou")


def _normalize_name(name: str) -> str:
    """Normalize a partner name for fuzzy comparison.

    Strips legal suffixes (AS, A/S), replaces Nordic chars (ø→o, æ→a, å→a),
    removes extra whitespace, and lowercases.
    """
    name = _STRIP_SUFFIXES.sub("", name)
    name = re.sub(r"\([^)]*\)", "", name)  # remove parenthetical like "(Ikke bruk)"
    name = re.sub(r"\s+", " ", name).strip().lower()
    name = name.translate(_NORDIC_MAP)
    return name


def _name_tokens(name: str) -> set[str]:
    """Split normalized name into word tokens for overlap scoring."""
    return set(_normalize_name(name).split())


class OdooMapper:
    """Maps ParsedOrder fields to Odoo record IDs with caching."""

    def __init__(self, client: OdooClient):
        self._client = client
        self._product_cache: dict[str, int | None] = {}
        self._partner_cache: dict[str, int] = {}
        self._uom_cache: dict[str, int] = {}
        self._norway_id: int | None = None
        self._default_uom_id: int | None = None
        self._all_partners: list[dict] | None = None  # lazy-loaded for fuzzy matching

    # --- Partner (customer) ---

    def find_or_create_partner(self, order: ParsedOrder) -> tuple[int, bool]:
        """Find existing partner by name, or create new.

        Returns (partner_id, created) where created is True if a new partner was made.
        """
        name = order.customer_name or "Ukjent kunde"

        if name in self._partner_cache:
            return self._partner_cache[name], False

        partner_id = self._find_partner(name, order)
        created = False

        if partner_id is None:
            partner_id = self._create_partner(order)
            created = True
            logger.info("Opprettet ny partner: %s (id=%d)", name, partner_id)
        else:
            logger.info("Fant eksisterende partner: %s (id=%d)", name, partner_id)

        self._partner_cache[name] = partner_id
        return partner_id, created

    def _find_partner(self, name: str, order: ParsedOrder | None = None) -> int | None:
        """Search for partner using multi-step matching:

        1. Exact name match
        2. ilike (case-insensitive contains), with whitespace normalization
        3. Normalized fuzzy matching (token overlap)
        4. City-based disambiguation for multi-location companies
        """
        # Normalize multiple whitespace to single space for matching
        normalized_name = re.sub(r"\s+", " ", name).strip()

        # Step 1: Exact match
        results = self._client.search_read(
            "res.partner",
            [["name", "=", normalized_name], ["customer_rank", ">", 0]],
            ["id", "name"],
            limit=1,
        )
        if results:
            return results[0]["id"]

        # Step 2: ilike (partial match)
        # Use individual tokens joined by '%' to handle extra whitespace
        # in Odoo partner names (e.g. "Atterås  AS" with double space).
        ilike_pattern = "%".join(normalized_name.split())
        results = self._client.search_read(
            "res.partner",
            [["name", "ilike", ilike_pattern], ["customer_rank", ">", 0]],
            ["id", "name", "city", "customer_rank"],
            limit=20,
        )
        if len(results) == 1:
            logger.info("Partner ilike-match: '%s' -> '%s'", name, results[0]["name"])
            return results[0]["id"]

        # If multiple ilike results, try city disambiguation first
        if len(results) > 1:
            if order and order.delivery_address:
                best = self._disambiguate_by_city(results, order)
                if best:
                    return best

            # Fall back to partner with highest customer_rank (most used)
            results.sort(key=lambda r: r.get("customer_rank", 0), reverse=True)
            logger.info(
                "Partner ilike-match (flere treff, bruker høyest rangert): '%s' -> '%s' (rank=%s)",
                name, results[0]["name"], results[0].get("customer_rank", 0),
            )
            return results[0]["id"]

        # Step 3: Fuzzy token matching against all partners
        match = self._fuzzy_match_partner(name, order)
        if match:
            return match

        return None

    def _disambiguate_by_city(
        self, candidates: list[dict], order: ParsedOrder
    ) -> int | None:
        """Pick the best partner from multiple candidates using delivery city."""
        city = (order.delivery_address.city or "").lower() if order.delivery_address else ""
        if not city:
            return None

        # Check partner name first, then partner city field
        for c in candidates:
            partner_name = c.get("name", "").lower()
            partner_city = c.get("city", "").lower() if c.get("city") else ""
            if city in partner_name or city == partner_city:
                logger.info(
                    "Partner disambiguert via by '%s': '%s'",
                    city, c["name"],
                )
                return c["id"]

        return None

    def _fuzzy_match_partner(
        self, name: str, order: ParsedOrder | None = None
    ) -> int | None:
        """Fuzzy match using normalized token overlap against all Odoo partners.

        Requires at least 50% token overlap and minimum 2 matching tokens.
        """
        if self._all_partners is None:
            self._all_partners = self._client.search_read(
                "res.partner",
                [["customer_rank", ">", 0]],
                ["id", "name", "city"],
            )

        query_tokens = _name_tokens(name)
        if len(query_tokens) < 1:
            return None

        best_score = 0.0
        best_candidates: list[dict] = []

        for partner in self._all_partners:
            partner_tokens = _name_tokens(partner["name"])
            if not partner_tokens:
                continue

            # Jaccard-like overlap: intersection / min(len_a, len_b)
            overlap = query_tokens & partner_tokens
            if not overlap:
                continue

            score = len(overlap) / min(len(query_tokens), len(partner_tokens))

            if score > best_score and len(overlap) >= 2:
                best_score = score
                best_candidates = [partner]
            elif score == best_score and len(overlap) >= 2:
                best_candidates.append(partner)

        if not best_candidates or best_score < 0.5:
            return None

        # If multiple candidates with same score, try city disambiguation
        if len(best_candidates) > 1 and order and order.delivery_address:
            city_match = self._disambiguate_by_city(best_candidates, order)
            if city_match:
                return city_match

        result = best_candidates[0]
        logger.info(
            "Partner fuzzy-match (score=%.0f%%): '%s' -> '%s'",
            best_score * 100, name, result["name"],
        )
        return result["id"]

    def _create_partner(self, order: ParsedOrder) -> int:
        """Create a new res.partner from order data."""
        vals: dict = {
            "name": order.customer_name or "Ukjent kunde",
            "customer_rank": 1,
        }

        if order.customer_reference:
            vals["ref"] = order.customer_reference

        addr = order.delivery_address
        if addr:
            if addr.street:
                vals["street"] = addr.street
            if addr.postal_code:
                vals["zip"] = addr.postal_code
            if addr.city:
                vals["city"] = addr.city
            country_id = self._get_norway_id()
            if country_id:
                vals["country_id"] = country_id

        # Invalidate partner cache so fuzzy matching picks up the new partner
        self._all_partners = None

        return self._client.create("res.partner", vals)

    def _get_norway_id(self) -> int | None:
        """Get res.country ID for Norway."""
        if self._norway_id is None:
            results = self._client.search(
                "res.country", [["code", "=", "NO"]], limit=1
            )
            self._norway_id = results[0] if results else None
        return self._norway_id

    # --- Product ---

    def find_product(self, article_number: str) -> int | None:
        """Look up product by default_code or barcode. Returns product_id or None."""
        if not article_number:
            return None

        key = article_number.strip()
        if key in self._product_cache:
            return self._product_cache[key]

        # Search by default_code (internal reference)
        results = self._client.search_read(
            "product.product",
            [["default_code", "=", key]],
            ["id", "default_code"],
            limit=1,
        )
        if results:
            pid = results[0]["id"]
            self._product_cache[key] = pid
            return pid

        # Try barcode
        results = self._client.search_read(
            "product.product",
            [["barcode", "=", key]],
            ["id", "barcode"],
            limit=1,
        )
        if results:
            pid = results[0]["id"]
            self._product_cache[key] = pid
            return pid

        # Not found
        self._product_cache[key] = None
        logger.warning("Produkt ikke funnet i Odoo: %s", key)
        return None

    def get_fallback_product_id(self, fallback_id: str | None = None) -> int | None:
        """Get or create a fallback product for unknown SKUs."""
        if fallback_id:
            return int(fallback_id)

        # Search for a product named "Ukjent produkt"
        results = self._client.search_read(
            "product.product",
            [["default_code", "=", "UNKNOWN"]],
            ["id"],
            limit=1,
        )
        if results:
            return results[0]["id"]

        return None

    # --- Unit of Measure ---

    def find_uom(self, unit_str: str) -> int:
        """Map unit string to uom.uom ID. Falls back to Units (id=1)."""
        if not unit_str:
            return self._get_default_uom_id()

        key = unit_str.strip().lower()
        if key in self._uom_cache:
            return self._uom_cache[key]

        odoo_name = _UOM_MAP.get(key, "Units")

        results = self._client.search_read(
            "uom.uom",
            [["name", "ilike", odoo_name]],
            ["id", "name"],
            limit=1,
        )

        if results:
            uom_id = results[0]["id"]
        else:
            uom_id = self._get_default_uom_id()

        self._uom_cache[key] = uom_id
        return uom_id

    def _get_default_uom_id(self) -> int:
        """Get the default 'Units' UoM ID."""
        if self._default_uom_id is None:
            results = self._client.search_read(
                "uom.uom",
                [["name", "=", "Units"]],
                ["id"],
                limit=1,
            )
            self._default_uom_id = results[0]["id"] if results else 1
        return self._default_uom_id
