"""Generate realistic Norwegian purchase order PDFs for end-to-end testing.

Creates test orders that mimic the Drevelin/Ortopartner format but with
unique order numbers (TEST-YYYYMMDD-NNN) so they don't collide with
historical data in staging Odoo.

Usage:
    python scripts/generate_test_pdfs.py

Output: writes PDFs to test_pdfs/ folder.

Test scenarios covered:
  1. Simple 3-line order, known products, known customer
  2. 4-line order, different customer
  3. Configuration sheet with measurements
  4. Order with unknown article number (triggers fallback warning)
  5. Order from new/unknown customer (triggers partner creation warning)
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "test_pdfs"
TODAY = datetime.now().strftime("%Y%m%d")
TODAY_HUMAN = datetime.now().strftime("%d.%m.%Y")


def _styles():
    ss = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "title", parent=ss["Heading1"], fontSize=16, spaceAfter=12,
            textColor=colors.HexColor("#1a3a6b"),
        ),
        "h2": ParagraphStyle(
            "h2", parent=ss["Heading2"], fontSize=11, spaceAfter=6,
            textColor=colors.HexColor("#1a3a6b"),
        ),
        "normal": ParagraphStyle(
            "normal", parent=ss["Normal"], fontSize=10, leading=13,
        ),
        "small": ParagraphStyle(
            "small", parent=ss["Normal"], fontSize=8, leading=10,
            textColor=colors.HexColor("#555"),
        ),
        "important": ParagraphStyle(
            "important", parent=ss["Normal"], fontSize=10, leading=13,
            textColor=colors.HexColor("#c1121f"), spaceAfter=6,
        ),
    }


def _build_header(order_number, customer_name, order_date, our_ref, story, styles):
    """Build the title + metadata block at the top of the order."""
    story.append(Paragraph(f"{customer_name} - BESTILLING {order_number}", styles["title"]))
    story.append(Spacer(1, 6))

    # Left: supplier info. Right: metadata table.
    supplier_info = [
        Paragraph("<b>ORTOPARTNER AS</b>", styles["normal"]),
        Paragraph("Postboks 123", styles["normal"]),
        Paragraph("7001 Trondheim", styles["normal"]),
        Paragraph("Norge", styles["normal"]),
        Paragraph("post@ortopartner.no", styles["normal"]),
    ]

    metadata = [
        ["Bestillingsnr.", order_number],
        ["Dato", order_date],
        ["Deres ref.", ""],
        ["Vår referanse", our_ref],
        ["Int. ordre ref.", ""],
        ["Valuta", "NOK"],
        ["Side", "1"],
    ]
    meta_table = Table(metadata, colWidths=[35 * mm, 50 * mm])
    meta_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#333")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#888")),
                ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#ccc")),
            ]
        )
    )

    header_layout = Table(
        [[supplier_info, meta_table]],
        colWidths=[85 * mm, 90 * mm],
    )
    header_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header_layout)
    story.append(Spacer(1, 14))


def _build_line_table(lines, story):
    """Build the order lines table."""
    header = [
        ["Artikkelnr.", "Beskrivelse", "Antall", "Enhet", "Rabatt", "Beløp"],
    ]
    rows = []
    total = 0.0
    for item in lines:
        art, desc, qty, unit, discount, amount = item
        rows.append([art, desc, f"{qty:.1f}", unit, f"{discount}%" if discount else "", f"{amount:,.2f}".replace(",", " ")])
        total += amount

    table_data = header + rows + [
        ["", "", "", "", "Samlet ordreverdi (NOK)", f"{total:,.2f}".replace(",", " ")],
    ]

    lines_table = Table(
        table_data,
        colWidths=[28 * mm, 70 * mm, 15 * mm, 15 * mm, 18 * mm, 25 * mm],
    )
    lines_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8edf5")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1a3a6b")),
                ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#1a3a6b")),
                ("LINEABOVE", (0, -1), (-1, -1), 1, colors.HexColor("#1a3a6b")),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("ALIGN", (2, 0), (5, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(lines_table)
    story.append(Spacer(1, 20))


def _build_footer(customer_name, delivery_address, story, styles, note=None):
    """Build the delivery address + importance note at the bottom."""
    left_content = [
        Paragraph("<b>VIKTIG!</b>", styles["important"]),
        Paragraph(
            f"Dette er en innkjøpsordre fra {customer_name}",
            styles["important"],
        ),
    ]
    if note:
        left_content.append(Spacer(1, 6))
        left_content.append(Paragraph(note, styles["normal"]))

    right_content = [
        Paragraph("<b>Leveringsadresse:</b>", styles["normal"]),
    ]
    for line in delivery_address:
        right_content.append(Paragraph(line, styles["normal"]))

    footer = Table(
        [[left_content, right_content]],
        colWidths=[90 * mm, 85 * mm],
    )
    footer.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ]
        )
    )
    story.append(footer)


def generate_pdf(filename, order_number, customer_name, our_ref,
                 lines, delivery_address, note=None):
    """Generate a single test order PDF."""
    filepath = OUTPUT_DIR / filename
    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling {order_number}",
        author=customer_name,
    )

    styles = _styles()
    story = []

    _build_header(order_number, customer_name, TODAY_HUMAN, our_ref, story, styles)
    _build_line_table(lines, story)
    _build_footer(customer_name, delivery_address, story, styles, note=note)

    doc.build(story)
    print(f"  + {filename}  ({order_number})")


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Genererer test-PDF-er i {OUTPUT_DIR}/")
    print(f"Dato: {TODAY_HUMAN}\n")

    # ------------------------------------------------------------------
    # Test 1: Simple 3-line order, known products from Drevelin Sør
    # ------------------------------------------------------------------
    generate_pdf(
        filename=f"Bestilling_TEST-{TODAY}-001.pdf",
        order_number=f"TEST-{TODAY}-001",
        customer_name="Drevelin Ortopedi AS",
        our_ref="Grethe Golshani",
        lines=[
            ("111P170/12", "Streifytec Stiff, 400x400x12 mm", 5.0, "plate", 0, 3885.00),
            ("111P71/12",  "Streifyflex, 400x400x12 mm, black", 3.0, "plate", 0, 4119.00),
            ("60T18F30W",  "Loop Strap 30 mm, white 25m", 1.0, "rull", 0, 550.00),
        ],
        delivery_address=[
            "Drevelin Ortopedi AS",
            "Avd. Drevelin Sør",
            "Kongsgård Alle 53",
            "4632 Kristiansand S",
            "Norge",
        ],
    )

    # ------------------------------------------------------------------
    # Test 2: 4-line order from a different customer (larger order)
    # ------------------------------------------------------------------
    generate_pdf(
        filename=f"Bestilling_TEST-{TODAY}-002.pdf",
        order_number=f"TEST-{TODAY}-002",
        customer_name="Sophies Minde Ortopedi AS",
        our_ref="Lars Pedersen",
        lines=[
            ("111P170/12", "Streifytec Stiff, 400x400x12 mm", 10.0, "plate", 5, 7392.75),
            ("111P71/12",  "Streifyflex, 400x400x12 mm, black", 6.0, "plate", 0, 8238.00),
            ("60T18F30W",  "Loop Strap 30 mm, white 25m", 4.0, "rull", 0, 2200.00),
            ("60T18F50W",  "Loop Strap 50 mm, white 25m", 2.0, "rull", 0, 1400.00),
        ],
        delivery_address=[
            "Sophies Minde Ortopedi AS",
            "Trondheimsveien 132",
            "0570 Oslo",
            "Norge",
        ],
    )

    # ------------------------------------------------------------------
    # Test 3: Configuration sheet style with measurements
    # ------------------------------------------------------------------
    generate_pdf(
        filename=f"Bestilling_TEST-{TODAY}-003.pdf",
        order_number=f"TEST-{TODAY}-003",
        customer_name="Ortopro AS",
        our_ref="Marianne Olsen",
        lines=[
            ("KONFIG-LEGG", "Legg-ortose etter mål (se konfigurasjonsark)", 1.0, "stk", 0, 4250.00),
        ],
        delivery_address=[
            "Ortopro AS",
            "Storgata 45",
            "5008 Bergen",
            "Norge",
        ],
        note=(
            "<b>Konfigurasjonsark / Mål:</b><br/>"
            "Pasient: M.H. (f. 1985)<br/>"
            "Side: Høyre<br/>"
            "Midje: 82 cm &nbsp;&nbsp; Hofte: 96 cm<br/>"
            "Lår: 54 cm &nbsp;&nbsp; Kne: 38 cm &nbsp;&nbsp; Legg: 36 cm<br/>"
            "Ankel: 24 cm<br/>"
            "Farge: Svart &nbsp;&nbsp; Solokit: Ja"
        ),
    )

    # ------------------------------------------------------------------
    # Test 4: Order with unknown article number (fallback warning)
    # ------------------------------------------------------------------
    generate_pdf(
        filename=f"Bestilling_TEST-{TODAY}-004.pdf",
        order_number=f"TEST-{TODAY}-004",
        customer_name="Drevelin Ortopedi AS",
        our_ref="Grethe Golshani",
        lines=[
            ("111P170/12",     "Streifytec Stiff, 400x400x12 mm", 2.0, "plate", 0, 1554.00),
            ("UKJENT-XYZ-999", "Spesialplate type XYZ (ikke i katalog)", 1.0, "stk", 0, 1800.00),
        ],
        delivery_address=[
            "Drevelin Ortopedi AS",
            "Avd. Drevelin Nord",
            "Beddingen 14",
            "7042 Trondheim",
            "Norge",
        ],
    )

    # ------------------------------------------------------------------
    # Test 5: Order from new/unknown customer (partner creation warning)
    # ------------------------------------------------------------------
    generate_pdf(
        filename=f"Bestilling_TEST-{TODAY}-005.pdf",
        order_number=f"TEST-{TODAY}-005",
        customer_name="Testkunde Ortopedi Hæsj AS",
        our_ref="Erik Testesen",
        lines=[
            ("60T18F30W", "Loop Strap 30 mm, white 25m", 2.0, "rull", 0, 1100.00),
            ("60T18F50W", "Loop Strap 50 mm, white 25m", 1.0, "rull", 0, 700.00),
        ],
        delivery_address=[
            "Testkunde Ortopedi Hæsj AS",
            "Testveien 1",
            "9999 Testby",
            "Norge",
        ],
    )

    print(f"\n{'='*60}")
    print(f"  Genererte 5 test-PDF-er i {OUTPUT_DIR}")
    print(f"{'='*60}")
    print()
    print("Test-scenarier:")
    print(f"  001 — Enkel ordre, kjent kunde + kjente produkter")
    print(f"  002 — Stor ordre (4 linjer) fra annen kunde, med rabatt")
    print(f"  003 — Konfigurasjonsark med mål (utloser KONFIGURASJONSARK-tag)")
    print(f"  004 — Ukjent artikkelnummer (utloser fallback-advarsel)")
    print(f"  005 — Ny kunde (utloser 'Ny kunde opprettet'-advarsel)")
    print()
    print("Send test-PDF-ene som vedlegg til ordre@ortopartner.no")
    print("og klikk 'Sjekk e-post nå' i dashboardet.")


if __name__ == "__main__":
    main()
