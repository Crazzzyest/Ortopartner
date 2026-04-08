"""Data models for parsed orders."""

from __future__ import annotations

from pydantic import BaseModel, Field


class LineItem(BaseModel):
    article_number: str | None = Field(default=None, description="Artikkelnr / Produktnr / Nr")
    description: str = Field(description="Beskrivelse / Produktnavn")
    quantity: float = Field(description="Antall / Ant. / Rest")
    unit: str = Field(description="Enhet (stk, pakke, plate, rull, Kg, CRT, sett ...)")
    internal_order_ref: str | None = Field(default=None, description="Int. ordre ref.")
    discount_percent: float | None = Field(default=None, description="Rabatt %")
    unit_price: float | None = Field(default=None, description="Enhetspris ekskl. mva")
    line_total: float | None = Field(default=None, description="Beløp / Sum for linjen")


class DeliveryAddress(BaseModel):
    name: str | None = Field(default=None, description="Firmanavn / mottaker")
    street: str | None = Field(default=None, description="Gateadresse")
    postal_code: str | None = Field(default=None, description="Postnummer")
    city: str | None = Field(default=None, description="Poststed")
    country: str | None = Field(default="Norge")


class ParsedOrder(BaseModel):
    order_number: str = Field(description="Bestillingsnr / Ordernr")
    order_date: str | None = Field(default=None, description="Dato (ISO-format YYYY-MM-DD)")
    customer_name: str | None = Field(default=None, description="Bestillende firma/kunde")
    customer_reference: str | None = Field(default=None, description="Vår referanse / kontaktperson")
    internal_order_ref: str | None = Field(default=None, description="Int. ordre ref. (hode-nivå)")
    currency: str = Field(default="NOK")
    line_items: list[LineItem] = Field(description="Ordrelinjer")
    total_amount: float | None = Field(default=None, description="Samlet ordreverdi")
    delivery_address: DeliveryAddress | None = Field(default=None)
    special_instructions: str | None = Field(default=None, description="Viktig-meldinger, faktura-info etc.")
    source_file: str = Field(description="Kilde-PDF filnavn")
    confidence: float = Field(
        default=1.0,
        description="Konfidensverdi 0-1. Under 0.8 = flagges for manuell kontroll",
    )
    warnings: list[str] = Field(default_factory=list, description="Varsler / avvik funnet")
    document_type: str = Field(
        default="purchase_order",
        description="purchase_order | configuration_sheet | unknown",
    )


# ---------------------------------------------------------------------------
# DHL Tracking models
# ---------------------------------------------------------------------------


class DhlTrackingEvent(BaseModel):
    """A single tracking event from DHL."""

    timestamp: str = Field(description="ISO 8601 tidsstempel")
    status: str = Field(description="PICKED_UP | IN_TRANSIT | DELIVERED | EXCEPTION | ...")
    status_message: str = Field(default="")
    location_city: str | None = Field(default=None)
    location_country: str | None = Field(default=None)


class DhlTrackingResult(BaseModel):
    """Result from DHL tracking API for a single shipment."""

    tracking_number: str
    current_status: str = Field(description="Siste DHL-status")
    last_update: str | None = Field(default=None, description="Tidspunkt for siste hendelse")
    estimated_delivery: str | None = Field(default=None)
    events: list[DhlTrackingEvent] = Field(default_factory=list)


class TrackingUpdate(BaseModel):
    """Result of syncing a single SO's tracking with DHL."""

    so_name: str = Field(description="Salgsordre-navn i Odoo (f.eks. S00445)")
    tracking_number: str = Field(default="")
    status: str = Field(default="pending", description="success | error | no_tracking")
    dhl_status: str | None = Field(default=None, description="Siste DHL-status")
    message: str = Field(default="")
    events: list[DhlTrackingEvent] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Odoo result models
# ---------------------------------------------------------------------------


class OdooResult(BaseModel):
    """Result of pushing an order to Odoo."""

    source_file: str = Field(description="Kilde-PDF filnavn")
    order_number: str = Field(description="Bestillingsnr fra PDF")
    status: str = Field(default="pending", description="success | skipped | error")
    message: str = Field(default="")
    partner_id: int | None = Field(default=None, description="res.partner ID i Odoo")
    sale_order_id: int | None = Field(default=None, description="sale.order ID i Odoo")
    so_name: str | None = Field(default=None, description="SO-navn i Odoo (f.eks. S00042)")
    so_confirmed: bool = Field(default=False)
    purchase_order_ids: list[int] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
