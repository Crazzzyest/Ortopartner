"""Validation rules for parsed orders."""

from __future__ import annotations

import re
from .models import ParsedOrder

# Warning prefixes that represent informational/derived data, not problems.
# These don't count toward the manual-review threshold.
_INFO_WARNING_PREFIXES = (
    "Enhetspris utledet fra totalbeløp",
    "Enhetspris hentet fra produktkatalog",
)


def validate_order(order: ParsedOrder) -> ParsedOrder:
    """Run all validation rules on a parsed order. Mutates warnings and confidence."""
    warnings: list[str] = list(order.warnings)
    confidence = order.confidence

    # Rule 1: Order number must exist and be non-empty
    if not order.order_number or order.order_number.strip() == "":
        warnings.append("Bestillingsnummer mangler")
        confidence -= 0.3

    # Rule 2: Date must be valid ISO format
    if not order.order_date:
        warnings.append("Ordredato mangler")
        confidence -= 0.2
    elif not re.match(r"^\d{4}-\d{2}-\d{2}$", order.order_date):
        warnings.append(f"Ugyldig datoformat: {order.order_date}")
        confidence -= 0.2

    # Rule 2b: Customer name must exist
    if not order.customer_name:
        warnings.append("Kundenavn mangler")
        confidence -= 0.2

    # Rule 3: Customer name should not be Ortopartner
    elif "ortopartner" in order.customer_name.lower():
        warnings.append("Kundenavn er satt til Ortopartner (leverandør) – bør være bestiller")
        confidence -= 0.3

    # Rule 4: Must have at least one line item
    if not order.line_items:
        warnings.append("Ingen ordrelinjer funnet")
        confidence -= 0.4

    # Rule 5: Each line item must have article number and quantity
    for i, item in enumerate(order.line_items):
        if not item.article_number or item.article_number.strip() == "":
            warnings.append(f"Linje {i + 1}: Artikkelnummer mangler")
            confidence -= 0.1
        if item.quantity <= 0:
            warnings.append(f"Linje {i + 1}: Ugyldig antall ({item.quantity})")
            confidence -= 0.1

    # Rule 5b: Derive unit_price from line_total when PDF only had a total
    # column (typical for Ortopartner orders — "Beløp" is line total after
    # discount, no explicit unit price column). We compute the list price
    # before discount so Odoo can combine it with the discount field:
    #   unit_price = line_total / (quantity * (1 - discount/100))
    for i, item in enumerate(order.line_items):
        if (
            item.unit_price is None
            and item.line_total is not None
            and item.quantity > 0
        ):
            disc = (item.discount_percent or 0.0) / 100.0
            divisor = item.quantity * (1.0 - disc)
            if divisor > 0:
                derived = item.line_total / divisor
                item.unit_price = round(derived, 4)
                warnings.append(
                    f"Linje {i + 1}: Enhetspris utledet fra totalbeløp "
                    f"({item.line_total:.2f} / {item.quantity:g}"
                    + (f" / {1 - disc:.2f}" if disc > 0 else "")
                    + f" = {item.unit_price:.2f} NOK)"
                )

    # Rule 6: If total exists, check against sum of line totals
    if order.total_amount is not None:
        line_sum = sum(
            item.line_total for item in order.line_items if item.line_total is not None
        )
        if line_sum > 0:
            diff = abs(order.total_amount - line_sum)
            if diff > 1.0:  # Allow 1 NOK rounding tolerance
                warnings.append(
                    f"Totalbeløp ({order.total_amount}) matcher ikke sum av linjer ({line_sum}). Differanse: {diff}"
                )
                confidence -= 0.15

    # Rule 7: Delivery address should be present
    if order.delivery_address is None:
        warnings.append("Leveringsadresse mangler")
        confidence -= 0.1

    # Rule 8: Configuration sheets get lower default confidence
    if order.document_type == "configuration_sheet":
        warnings.append("Konfigurasjonsark – krever manuell gjennomgang")
        confidence = min(confidence, 0.7)

    # Rule 9: No prices at all
    has_any_price = any(
        item.line_total is not None or item.unit_price is not None
        for item in order.line_items
    )
    if not has_any_price and order.document_type == "purchase_order":
        warnings.append("Ingen priser oppgitt i bestillingen")

    # Rule 10: Strip obsolete "no unit price" warnings that Claude may have
    # emitted — we now derive unit_price from line_total, so these would be
    # misleading. Only keep them if unit_price really is missing on all lines.
    all_lines_priced = order.line_items and all(
        item.unit_price is not None for item in order.line_items
    )
    if all_lines_priced:
        warnings = [
            w for w in warnings
            if "ingen enhetspris" not in w.lower()
            and "mangler enhetspris" not in w.lower()
        ]

    # Clamp confidence
    order.confidence = max(0.0, min(1.0, confidence))
    order.warnings = warnings

    return order


def needs_manual_review(order: ParsedOrder) -> bool:
    """Check if an order needs manual review.

    Informational warnings (e.g. derived unit prices) don't count toward
    the threshold — they're normal for Ortopartner PDFs that only show
    line totals, not every such order should be flagged for review.
    """
    significant_warnings = [
        w for w in order.warnings
        if not w.startswith(_INFO_WARNING_PREFIXES)
        and not any(
            prefix.lower() in w.lower() for prefix in _INFO_WARNING_PREFIXES
        )
    ]
    return order.confidence < 0.8 or len(significant_warnings) > 2
