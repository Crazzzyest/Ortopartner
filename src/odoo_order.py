"""High-level orchestration: ParsedOrder → Sale Order → Confirm → Purchase Order."""

from __future__ import annotations

import logging

from .models import LineItem, OdooResult, ParsedOrder
from .odoo_client import OdooClient
from .odoo_mapper import OdooMapper

logger = logging.getLogger(__name__)


class OdooOrderService:
    """Orchestrates the full Odoo order flow: SO → Confirm → PO."""

    def __init__(
        self,
        client: OdooClient,
        mapper: OdooMapper,
        fallback_product_id: int | None = None,
        transport_product_id: int | None = None,
    ):
        self._client = client
        self._mapper = mapper
        self._fallback_product_id = fallback_product_id
        # Dedicated fallback for transport/freight lines (e.g. F1 "Frakt").
        # Kicks in when the line description looks like transport AND
        # the article number isn't in the catalog.
        self._transport_product_id = transport_product_id

    def push_order(self, order: ParsedOrder, needs_review: bool = False) -> OdooResult:
        """Full pipeline: create SO, confirm, find POs.

        If needs_review is True, the SO is created as a draft Quotation
        (no confirmation) so a human can review and confirm it in Odoo.

        Returns OdooResult with IDs and status.
        """
        result = OdooResult(
            source_file=order.source_file,
            order_number=order.order_number,
        )

        # Carry over parsing/validation warnings (derived unit prices,
        # missing fields, etc.) so they're posted to both the Odoo
        # message log and the event log alongside Odoo-level warnings.
        result.warnings.extend(order.warnings)

        try:
            # 1. Duplicate check
            existing = self._find_existing_so(order.order_number)
            if existing:
                result.status = "skipped"
                result.message = (
                    f"Duplikat: SO med client_order_ref='{order.order_number}' "
                    f"finnes allerede (id={existing})"
                )
                logger.info(result.message)
                return result

            # 2. Resolve partner
            partner_id, partner_created = self._mapper.find_or_create_partner(order)
            result.partner_id = partner_id
            if partner_created:
                result.warnings.append(
                    f"Ny kunde opprettet i Odoo: '{order.customer_name or 'Ukjent kunde'}' "
                    f"(id={partner_id}). Sjekk at dette ikke er en duplikat, og legg til "
                    f"betalingsbetingelser/prisliste manuelt."
                )

            # 3. Create SO with line items
            so_id, warnings = self._create_sale_order(order, partner_id)
            result.sale_order_id = so_id
            result.warnings.extend(warnings)

            # Read back the SO name (e.g. "S00042")
            so_data = self._client.search_read(
                "sale.order", [["id", "=", so_id]], ["name"], limit=1
            )
            if so_data:
                result.so_name = so_data[0]["name"]

            logger.info("SO opprettet: %s (id=%d)", result.so_name, so_id)

            # 4. Tag configuration sheets
            if order.document_type == "configuration_sheet":
                self._tag_config_sheet(so_id, order)

            # 5. If flagged for review, tag and post review message
            if needs_review:
                self._tag_for_review(so_id, order)

            # 6. Post warnings to SO message log
            if result.warnings:
                self._post_warnings(so_id, result.warnings)

            # All orders stay as draft Quotations — confirmation is done
            # manually in Odoo by the user after review.

            result.status = "success"
            if needs_review:
                result.message = (
                    f"Tilbud {result.so_name or so_id} opprettet som utkast "
                    f"(krever manuell kontroll, konfidensverdi: {order.confidence:.0%})"
                )
            else:
                result.message = (
                    f"Tilbud {result.so_name or so_id} opprettet som utkast"
                )

        except Exception as e:
            logger.exception("Feil ved Odoo-push for %s", order.order_number)
            result.status = "error"
            result.message = str(e)

        return result

    def _tag_config_sheet(self, so_id: int, order: ParsedOrder) -> None:
        """Tag a draft SO as a configuration sheet with measurements in the note."""
        tag_id = self._find_or_create_tag("KONFIGURASJONSARK")
        if tag_id:
            try:
                self._client.write("sale.order", so_id, {"tag_ids": [(4, tag_id)]})
            except Exception as e:
                logger.warning("Kunne ikke sette KONFIGURASJONSARK-tag: %s", e)

        # Post config details as a message
        body = f"[KONFIGURASJONSARK] {order.order_number}\n"
        if order.special_instructions:
            body += f"\n{order.special_instructions}"

        try:
            self._client.call("sale.order", "message_post", [[so_id]], {
                "body": body,
                "message_type": "comment",
                "subtype_xmlid": "mail.mt_note",
            })
        except Exception as e:
            logger.warning("Kunne ikke poste config-melding: %s", e)

    def _tag_for_review(self, so_id: int, order: ParsedOrder) -> None:
        """Tag a draft SO for manual review: apply tag + post warning message."""
        # Find or create a "REVIEW" tag
        tag_id = self._find_or_create_tag("REVIEW")
        if tag_id:
            try:
                self._client.write("sale.order", so_id, {"tag_ids": [(4, tag_id)]})
            except Exception as e:
                logger.warning("Kunne ikke sette REVIEW-tag på SO %d: %s", so_id, e)

        # Post review message on the SO with details
        reasons = []
        if order.confidence < 0.8:
            reasons.append(f"Lav konfidensverdi: {order.confidence:.0%}")
        if order.warnings:
            for w in order.warnings:
                reasons.append(w)

        body = f"[REVIEW] Automatisk importert ordre krever manuell kontroll\n"
        body += f"Konfidensverdi: {order.confidence:.0%}\n"
        if reasons:
            body += "\n".join(f"- {r}" for r in reasons) + "\n"
        body += "\nGjennomgå ordrelinjene og bekreft manuelt når alt ser riktig ut."

        try:
            self._client.call("sale.order", "message_post", [[so_id]], {
                "body": body,
                "message_type": "comment",
                "subtype_xmlid": "mail.mt_note",
            })
        except Exception as e:
            logger.warning("Kunne ikke poste review-melding på SO %d: %s", so_id, e)

    def _post_warnings(self, so_id: int, warnings: list[str]) -> None:
        """Post all warnings as a single message on the SO."""
        body = "[ADVARSEL] Automatisk import fant avvik:\n\n"
        body += "\n".join(f"- {w}" for w in warnings)

        try:
            self._client.call("sale.order", "message_post", [[so_id]], {
                "body": body,
                "message_type": "comment",
                "subtype_xmlid": "mail.mt_note",
            })
        except Exception as e:
            logger.warning("Kunne ikke poste advarsler på SO %d: %s", so_id, e)

    def _find_or_create_tag(self, tag_name: str) -> int | None:
        """Find or create a crm.tag for sale orders."""
        try:
            results = self._client.search_read(
                "crm.tag", [["name", "=", tag_name]], ["id"], limit=1
            )
            if results:
                return results[0]["id"]
            return self._client.create("crm.tag", {"name": tag_name})
        except Exception as e:
            logger.warning("Kunne ikke finne/opprette tag '%s': %s", tag_name, e)
            return None

    def _find_existing_so(self, order_number: str) -> int | None:
        """Check if SO with this order_number already exists."""
        results = self._client.search(
            "sale.order",
            [["client_order_ref", "=", order_number]],
            limit=1,
        )
        return results[0] if results else None

    def _create_sale_order(
        self, order: ParsedOrder, partner_id: int
    ) -> tuple[int, list[str]]:
        """Create sale.order with order lines. Returns (SO id, warnings)."""
        warnings: list[str] = []

        order_lines = []
        for i, item in enumerate(order.line_items):
            line_vals, line_warnings = self._make_order_line(item, i + 1)
            order_lines.append((0, 0, line_vals))
            warnings.extend(line_warnings)

        vals: dict = {
            "partner_id": partner_id,
            "client_order_ref": order.order_number,
            "origin": f"PDF:{order.source_file}",
            "order_line": order_lines,
        }

        if order.order_date:
            vals["date_order"] = order.order_date

        if order.special_instructions:
            vals["note"] = order.special_instructions

        so_id = self._client.create("sale.order", vals)

        # After create, Odoo may have applied partner/pricelist defaults
        # (contract discounts, pricelist unit prices) via onchange. Compare
        # what the PDF said vs what Odoo actually saved, and warn on any
        # divergence so Marius can dobbeltsjekke før bekreftelse.
        divergence_warnings = self._check_line_divergence(so_id, order.line_items)
        warnings.extend(divergence_warnings)

        return so_id, warnings

    def _check_line_divergence(
        self, so_id: int, items: list[LineItem]
    ) -> list[str]:
        """Compare PDF values vs Odoo-applied values for price_unit and discount.

        Divergences typically come from Odoo's onchange logic applying
        partner-level contract discounts or pricelist rules that aren't
        mentioned in the customer's PDF. We surface these as warnings
        (not errors) — the values are most likely correct for production,
        but a human should confirm before releasing the order.
        """
        warnings: list[str] = []

        try:
            lines = self._client.search_read(
                "sale.order.line",
                [["order_id", "=", so_id]],
                ["id", "sequence", "name", "price_unit", "discount", "product_uom_qty"],
                order="sequence, id",
            )
        except Exception as e:
            logger.warning("Kunne ikke lese tilbake SO-linjer for %d: %s", so_id, e)
            return warnings

        # Filter out section/note lines that Odoo adds automatically.
        lines = [ln for ln in lines if ln.get("product_uom_qty") is not None]

        if len(lines) != len(items):
            logger.warning(
                "Antall SO-linjer fra Odoo (%d) matcher ikke PDF-linjer (%d) "
                "for SO id=%d — hopper over divergens-sjekk",
                len(lines), len(items), so_id,
            )
            return warnings

        TOL = 0.01  # NOK/percent tolerance for float comparisons

        for idx, (item, line) in enumerate(zip(items, lines), start=1):
            pdf_price = float(item.unit_price or 0.0)
            pdf_discount = float(item.discount_percent or 0.0)
            odoo_price = float(line.get("price_unit") or 0.0)
            odoo_discount = float(line.get("discount") or 0.0)

            label = f"Linje {idx}"
            if item.article_number:
                label = f"Linje {idx} ({item.article_number})"

            # --- Unit price divergence ---
            if pdf_price > TOL and abs(pdf_price - odoo_price) > TOL:
                warnings.append(
                    f"{label}: enhetspris-avvik — PDF sier {pdf_price:.2f} NOK, "
                    f"Odoo anvendte {odoo_price:.2f} NOK (fra prislisten). "
                    f"Dobbeltsjekk før bekreftelse."
                )
            elif pdf_price <= TOL and odoo_price > TOL:
                warnings.append(
                    f"{label}: PDF hadde ingen enhetspris — Odoo fylte inn "
                    f"{odoo_price:.2f} NOK fra prislisten. Verifiser at dette stemmer."
                )
            elif pdf_price > TOL and odoo_price <= TOL:
                warnings.append(
                    f"{label}: PDF sier {pdf_price:.2f} NOK, men Odoo satte "
                    f"enhetspris til 0. Mangler produktkobling/prisliste — "
                    f"fyll inn pris manuelt."
                )

            # --- Discount divergence ---
            if abs(pdf_discount - odoo_discount) > TOL:
                if pdf_discount <= TOL and odoo_discount > TOL:
                    warnings.append(
                        f"{label}: PDF hadde ingen rabatt — Odoo anvendte "
                        f"{odoo_discount:.1f}% (kontraktsrabatt fra kundens "
                        f"prisliste). Bekreft at dette er avtalt."
                    )
                elif pdf_discount > TOL and odoo_discount <= TOL:
                    warnings.append(
                        f"{label}: PDF sier {pdf_discount:.1f}% rabatt, men "
                        f"Odoo anvendte ingen. Legg til manuelt eller sjekk "
                        f"kundens prisliste."
                    )
                else:
                    warnings.append(
                        f"{label}: rabatt-avvik — PDF sier {pdf_discount:.1f}%, "
                        f"Odoo anvendte {odoo_discount:.1f}%. Dobbeltsjekk før "
                        f"bekreftelse."
                    )

        return warnings

    def _make_order_line(
        self, item: LineItem, line_num: int
    ) -> tuple[dict, list[str]]:
        """Build vals dict for a sale.order.line. Returns (vals, warnings)."""
        warnings: list[str] = []

        product_id = self._mapper.find_product(item.article_number)
        if product_id is None and item.article_number:
            # Is this line actually a transport/freight charge from the customer's ERP?
            if self._looks_like_transport(item) and self._transport_product_id:
                product_id = self._transport_product_id
                warnings.append(
                    f"Linje {line_num}: Produkt '{item.article_number}' "
                    f"('{(item.description or '').strip()}') gjenkjent som frakt — "
                    f"koblet til standard fraktprodukt i Odoo."
                )
            elif self._fallback_product_id:
                warnings.append(
                    f"Linje {line_num}: Produkt '{item.article_number}' finnes ikke i Odoo "
                    f"— bruker fallback-produkt. Korriger manuelt for riktig leverandør/innkjøp."
                )
                product_id = self._fallback_product_id
            else:
                warnings.append(
                    f"Linje {line_num}: Produkt '{item.article_number}' finnes ikke i Odoo "
                    f"og ingen fallback er satt. Linjen mangler produktkobling — "
                    f"ingen innkjopsordre (PO) vil bli generert for denne linjen."
                )
                product_id = self._fallback_product_id
        elif product_id is None and not item.article_number:
            warnings.append(
                f"Linje {line_num}: Ingen artikkelnummer oppgitt — "
                f"linjen er ren fritekst uten produktkobling."
            )

        uom_id = self._mapper.find_uom(item.unit)

        description = item.description or ""
        if item.article_number:
            description = f"[{item.article_number}] {description}"

        # Determine price_unit:
        # 1. Use PDF price if set
        # 2. Else fall back to product.list_price (Odoo catalog) if we have a product
        # 3. Else 0.0 (flagged as warning elsewhere)
        # NB: Odoo 19 doesn't fire onchange_product_id on XML-RPC create(),
        #     so a missing price_unit stays as 0 unless we fill it ourselves.
        resolved_price = item.unit_price
        if (resolved_price is None or resolved_price <= 0) and product_id:
            catalog_price = self._fetch_product_list_price(product_id)
            if catalog_price and catalog_price > 0:
                resolved_price = catalog_price
                warnings.append(
                    f"Linje {line_num}: Enhetspris hentet fra produktkatalog "
                    f"(kr {catalog_price:.2f}) — PDF hadde ingen pris. "
                    f"Verifiser mot kundens prisliste før bekreftelse."
                )

        vals: dict = {
            "name": description,
            "product_uom_qty": item.quantity,
            "product_uom_id": uom_id,
            "price_unit": resolved_price or 0.0,
        }

        if product_id:
            vals["product_id"] = product_id

        if item.discount_percent:
            vals["discount"] = item.discount_percent

        return vals, warnings

    _TRANSPORT_KEYWORDS = (
        "transport", "frakt", "shipping", "porto", "fraktkostnad",
        "forsendelse", "leveringsgebyr", "freight",
    )

    def _looks_like_transport(self, item: LineItem) -> bool:
        """Heuristic: does this line look like a freight/transport charge?"""
        haystack = " ".join(
            filter(None, [item.description or "", item.article_number or ""])
        ).lower()
        return any(kw in haystack for kw in self._TRANSPORT_KEYWORDS)

    def _fetch_product_list_price(self, product_id: int) -> float | None:
        """Read list_price from product.product for price-less lines."""
        try:
            rows = self._client.search_read(
                "product.product",
                [["id", "=", product_id]],
                ["list_price"],
                limit=1,
            )
            if rows:
                return float(rows[0].get("list_price") or 0.0)
        except Exception as exc:
            logger.warning("Kunne ikke hente list_price for produkt %s: %s", product_id, exc)
        return None

    def _find_purchase_orders(self, so_name: str) -> list[int]:
        """Find POs generated by procurement after SO confirmation."""
        return self._client.search(
            "purchase.order",
            [["origin", "ilike", so_name]],
        )
