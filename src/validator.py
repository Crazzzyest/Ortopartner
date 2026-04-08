"""Validation rules for parsed orders."""

from __future__ import annotations

import re
from .models import ParsedOrder


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

    # Clamp confidence
    order.confidence = max(0.0, min(1.0, confidence))
    order.warnings = warnings

    return order


def needs_manual_review(order: ParsedOrder) -> bool:
    """Check if an order needs manual review."""
    return order.confidence < 0.8 or len(order.warnings) > 2
