"""Generate synthetic purchase order PDFs in three customer-specific formats.

Builds faithful reproductions of three real customer formats so we can
verify end-to-end Odoo registration:

  1. Blatchford Ortopedi  — multi-line, with "Int. ordre ref." column
  2. Bergen Mekaniske     — full metadata block, rabatt-%-kolonne, transport-linje
  3. Ortopediteknikk      — no price column at all (triggers price lookup)

Each order gets a unique order number based on a full timestamp
(HHMMSS), so repeated runs never collide with previous runs.

Usage:
    python scripts/generate_format_test_pdfs.py

Output: writes PDFs to test_pdfs/ folder.
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
    HRFlowable,
)

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "test_pdfs"
NOW = datetime.now()
TODAY_HUMAN_DOT = NOW.strftime("%d.%m.%Y")
TODAY_ISO = NOW.strftime("%Y-%m-%d")
# 6-digit suffix that changes every second — guarantees unique order numbers
UNIQ = NOW.strftime("%H%M%S")
DATESTAMP = NOW.strftime("%m%d")  # MMDD


def _fmt_nok(amount: float) -> str:
    """Format as Norwegian currency: '19 908,00'."""
    s = f"{amount:,.2f}"
    return s.replace(",", " ").replace(".", ",", 1)


def _styles():
    ss = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "title", parent=ss["Heading1"], fontSize=14, spaceAfter=10,
            textColor=colors.black, fontName="Helvetica-Bold",
        ),
        "h1": ParagraphStyle(
            "h1", parent=ss["Heading1"], fontSize=13, spaceAfter=6,
            fontName="Helvetica-Bold",
        ),
        "h2": ParagraphStyle(
            "h2", parent=ss["Heading2"], fontSize=11, spaceAfter=4,
            fontName="Helvetica-Bold",
        ),
        "normal": ParagraphStyle(
            "normal", parent=ss["Normal"], fontSize=9, leading=12,
        ),
        "normal_bold": ParagraphStyle(
            "normal_bold", parent=ss["Normal"], fontSize=9, leading=12,
            fontName="Helvetica-Bold",
        ),
        "small": ParagraphStyle(
            "small", parent=ss["Normal"], fontSize=8, leading=10,
            textColor=colors.HexColor("#555"),
        ),
        "small_right": ParagraphStyle(
            "small_right", parent=ss["Normal"], fontSize=8, leading=10,
            textColor=colors.HexColor("#555"), alignment=2,  # RIGHT
        ),
        "logo_blatchford": ParagraphStyle(
            "logo_blatchford", parent=ss["Heading1"], fontSize=22,
            textColor=colors.HexColor("#1f3864"), alignment=2,
            fontName="Helvetica-Bold",
        ),
        "logo_blatchford_sub": ParagraphStyle(
            "logo_blatchford_sub", parent=ss["Normal"], fontSize=14,
            textColor=colors.HexColor("#1f3864"), alignment=2,
            fontName="Helvetica",
        ),
        "logo_bm": ParagraphStyle(
            "logo_bm", parent=ss["Heading1"], fontSize=18,
            textColor=colors.HexColor("#0b5394"),
            fontName="Helvetica-Bold",
        ),
        "logo_ot": ParagraphStyle(
            "logo_ot", parent=ss["Heading1"], fontSize=18,
            textColor=colors.HexColor("#1a3a6b"), alignment=2,
            fontName="Helvetica-Bold",
        ),
        "logo_ot_sub": ParagraphStyle(
            "logo_ot_sub", parent=ss["Normal"], fontSize=9,
            textColor=colors.HexColor("#555"), alignment=2,
        ),
    }


# ------------------------------------------------------------------
# Format 1: Blatchford Ortopedi
# ------------------------------------------------------------------
def generate_blatchford(order_number: str, lines, total: float):
    styles = _styles()
    filename = f"Bestilling_Blatchford_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling {order_number}",
        author="Blatchford Ortopedi AS",
    )
    story = []

    # Logo block (top right)
    logo = [
        Paragraph("Blatchford<font color='#c00000'>:</font>", styles["logo_blatchford"]),
        Paragraph("Ortopedi", styles["logo_blatchford_sub"]),
    ]
    logo_row = Table([[None, logo]], colWidths=[90 * mm, 84 * mm])
    logo_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(logo_row)
    story.append(Spacer(1, 10))

    # Title
    story.append(Paragraph(f"<b>Bestilling {order_number}</b>", styles["h1"]))
    story.append(Spacer(1, 4))

    # Ortopartner block (left) + metadata box (right)
    left_block = [
        Paragraph("<b>Ortopartner AS</b>", styles["normal"]),
        Paragraph("Fabrikkvegen 1", styles["normal_bold"]),
        Spacer(1, 6),
        Paragraph("5265 &nbsp;&nbsp;&nbsp; YTRE ARNA", styles["normal_bold"]),
        Paragraph("Norge", styles["normal_bold"]),
        Spacer(1, 4),
        Paragraph("post@ortopartner.no", styles["normal_bold"]),
    ]

    meta_rows = [
        ["Bestillingsnr.", order_number],
        ["Dato", TODAY_HUMAN_DOT],
        ["Deres ref.", ""],
        ["Vår referanse", "Robin Larsen"],
        ["Int. ordre ref.", ""],
        ["Valuta", "NOK"],
        ["Side", "1"],
    ]
    meta_table = Table(meta_rows, colWidths=[35 * mm, 55 * mm])
    meta_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#555")),
                ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbb")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    header_layout = Table(
        [[left_block, meta_table]],
        colWidths=[75 * mm, 95 * mm],
    )
    header_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header_layout)
    story.append(Spacer(1, 16))

    # Line table
    header = ["Artikkelnr.", "Beskrivelse", "Int. ordre\nref.", "Rabatt", "Antall", "Enhet", "Beløp"]
    rows = [header]
    for item in lines:
        art, desc, ref, disc, qty, unit, amount = item
        rows.append([
            art, desc, ref,
            f"{disc}%" if disc else "",
            f"{qty:.1f}".replace(".", ","),
            unit,
            _fmt_nok(amount),
        ])

    lines_table = Table(
        rows,
        colWidths=[24 * mm, 60 * mm, 22 * mm, 16 * mm, 16 * mm, 14 * mm, 22 * mm],
    )
    lines_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
                ("ALIGN", (2, 0), (6, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(lines_table)
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.black))
    story.append(Spacer(1, 4))

    # Total row
    total_row = Table(
        [["", "Samlet ordreverdi (Valuta)", _fmt_nok(total)]],
        colWidths=[100 * mm, 50 * mm, 24 * mm],
    )
    total_row.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ]
        )
    )
    story.append(total_row)
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.black))
    story.append(Spacer(1, 14))

    # Legal text + delivery address
    legal = Paragraph(
        "For denne bestillingen gjelder Blatchford Ortopedi AS sine alminnelige innkjøpsbetingelser.<br/><br/>"
        "Vennligst bekreft denne bestillingen med pris og leveringsdato.",
        styles["small"],
    )
    delivery = [
        Paragraph("Leveringsadresse:", styles["small_right"]),
        Paragraph("Blatchford Ortopedi AS", styles["small_right"]),
        Paragraph("Avd. Arendal", styles["small_right"]),
        Paragraph("Langsæveien 4", styles["small_right"]),
        Paragraph("4846 Arendal", styles["small_right"]),
    ]
    legal_row = Table(
        [[legal, delivery]],
        colWidths=[105 * mm, 69 * mm],
    )
    legal_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(legal_row)
    story.append(Spacer(1, 20))

    # Signature
    story.append(Paragraph("Med vennlig hilsen", styles["normal"]))
    story.append(Paragraph("Robin Larsen", styles["normal"]))
    story.append(Paragraph(
        "45871812 | robin.larsen@blatchford.no | avd. Arendal",
        styles["normal"],
    ))

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


# ------------------------------------------------------------------
# Format 2: Bergen Mekaniske
# ------------------------------------------------------------------
def generate_bergen_mekaniske(order_number: str, lines, total: float):
    styles = _styles()
    filename = f"Bestilling_BergenMekaniske_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling {order_number}",
        author="Bergen Mekaniske AS",
    )
    story = []

    # Top row: Logo left, "Bestilling" right
    top_row = Table(
        [[
            Paragraph("<b>Bergen Mekaniske</b>", styles["logo_bm"]),
            Paragraph("<b>Bestilling</b>", styles["title"]),
        ]],
        colWidths=[90 * mm, 84 * mm],
    )
    top_row.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ]
        )
    )
    story.append(top_row)
    story.append(Spacer(1, 18))

    # Three-column row: sender, delivery address, metadata
    sender = [
        Paragraph("ORTOPARTNER AS", styles["normal_bold"]),
        Paragraph("Inngang 16, 3.etg", styles["normal"]),
        Paragraph("NO 5265 YTRE ARNA", styles["normal"]),
    ]
    delivery = [
        Paragraph("<b>Leveringsaddresse</b>", styles["normal_bold"]),
        Paragraph("Bergen Mekaniske AS", styles["normal"]),
        Paragraph("Hylkjeflaten 36", styles["normal"]),
        Paragraph("NO 5109 HYLKJE", styles["normal"]),
    ]
    meta_rows = [
        ["Best. nr.", order_number],
        ["Ordredato", TODAY_ISO],
        ["Innkjøper", "Jan-Steinar Sagstad"],
        ["Telefon", "92443056"],
        ["Epost", "jan.steinar.sagstad@bergenmek.no"],
        ["Transportør:", ""],
        ["Leveringsbetingelser", ""],
        ["Prosjekt", f"678 Water bracket Sotra Link ({UNIQ})"],
    ]
    meta_table = Table(meta_rows, colWidths=[35 * mm, 45 * mm])
    meta_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),  # right-align order number
            ]
        )
    )

    header_layout = Table(
        [[sender, delivery, meta_table]],
        colWidths=[55 * mm, 50 * mm, 80 * mm],
    )
    header_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header_layout)
    story.append(Spacer(1, 14))

    # Deres ref. / Merking / Melding
    extra = Table(
        [
            ["Deres ref.", ""],
            ["Merking", "678"],
            ["Melding til\nleverandør", "Ref. tilbud S00071"],
        ],
        colWidths=[30 * mm, 100 * mm],
    )
    extra.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(extra)
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#888")))
    story.append(Spacer(1, 4))

    # Line table
    header = [
        "Produkt\nID", "Produktnavn", "Lev.\nProdNr",
        "Ønsket lev.\n(lev.)", "Antall", "Enhet", "Pris", "%", "Beløp",
    ]
    rows = [header]
    for item in lines:
        pid, name, lev_prodnr, wish_date, qty, unit, price, disc, amount = item
        rows.append([
            pid,
            Paragraph(name, styles["normal"]),
            lev_prodnr,
            wish_date,
            f"{qty:.2f}".replace(".", ","),
            unit,
            _fmt_nok(price),
            f"{disc:.2f}".replace(".", ","),
            _fmt_nok(amount),
        ])

    lines_table = Table(
        rows,
        colWidths=[16 * mm, 52 * mm, 16 * mm, 18 * mm, 14 * mm, 12 * mm, 18 * mm, 10 * mm, 22 * mm],
    )
    lines_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (4, 0), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LINEBELOW", (0, 0), (-1, 0), 0.3, colors.HexColor("#888")),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(lines_table)
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#888")))
    story.append(Spacer(1, 6))

    # Total
    total_row = Table(
        [["", "NOK", _fmt_nok(total)]],
        colWidths=[130 * mm, 16 * mm, 28 * mm],
    )
    total_row.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ]
        )
    )
    story.append(total_row)

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


# ------------------------------------------------------------------
# Format 3: Ortopediteknikk (no price column)
# ------------------------------------------------------------------
def generate_ortopediteknikk(order_number: str, lines, internal_ref: str):
    styles = _styles()
    filename = f"Bestilling_Ortopediteknikk_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling {order_number}",
        author="Ortopediteknikk AS",
    )
    story = []

    # Logo (top right)
    logo = [
        Paragraph("<b>ORTOPEDITEKNIKK</b>", styles["logo_ot"]),
        Paragraph("Avdeling Oslo", styles["logo_ot_sub"]),
    ]
    logo_row = Table([[None, logo]], colWidths=[90 * mm, 84 * mm])
    logo_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(logo_row)
    story.append(Spacer(1, 14))

    # Title
    story.append(Paragraph("BESTILLING FRA ORTOPEDITEKNIKK", styles["h1"]))
    story.append(Spacer(1, 4))

    # Ortopartner block + metadata box
    left_block = [
        Paragraph("<b>Ortopartner AS</b>", styles["normal"]),
        Paragraph("Fabrikkvegen 1", styles["normal_bold"]),
        Spacer(1, 6),
        Paragraph("5265 &nbsp;&nbsp;&nbsp; YTRE ARNA", styles["normal_bold"]),
        Paragraph("Norge", styles["normal_bold"]),
        Spacer(1, 4),
        Paragraph("post@ortopartner.no", styles["normal_bold"]),
    ]

    meta_rows = [
        ["Bestillingsnr.", order_number],
        ["Dato", TODAY_HUMAN_DOT],
        ["Deres ref.", ""],
        ["Vår referanse", ""],
        ["Int. ordre ref.", internal_ref],
        ["Side", "1"],
    ]
    meta_table = Table(meta_rows, colWidths=[35 * mm, 55 * mm])
    meta_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#555")),
                ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbb")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    header_layout = Table(
        [[left_block, meta_table]],
        colWidths=[75 * mm, 95 * mm],
    )
    header_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header_layout)
    story.append(Spacer(1, 16))

    # Line table — NB: no price column
    header = ["Artikkelnr.", "Beskrivelse", "Int. ordre ref.", "Antall", "Enhet"]
    rows = [header]
    for item in lines:
        art, desc, ref, qty, unit = item
        rows.append([
            Paragraph(f"<b>{art}</b>", styles["normal_bold"]),
            Paragraph(f"<b>{desc}</b>", styles["normal_bold"]),
            ref,
            f"{qty:.1f}".replace(".", ","),
            unit,
        ])

    lines_table = Table(
        rows,
        colWidths=[28 * mm, 72 * mm, 32 * mm, 18 * mm, 18 * mm],
    )
    lines_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
                ("ALIGN", (3, 0), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(lines_table)
    story.append(Spacer(1, 20))

    # Footer: legal + delivery
    legal = Paragraph(
        "Vennligst bekreft bestillingen med en ordrebekreftelse. Faktura merkes med vårt ordrenummer "
        "og sendes til faktura@ortopediteknikk.no eller EHF-faktura til org.nr. 930614785. "
        "Telefonnr til bestiller er .",
        styles["small"],
    )
    delivery = [
        Paragraph("Leveringsadresse:", styles["small_right"]),
        Paragraph("Ortopediteknikk AS", styles["small_right"]),
        Paragraph("Avd. Oslo", styles["small_right"]),
        Paragraph("Ryensvingen 6", styles["small_right"]),
        Paragraph("0680 Oslo", styles["small_right"]),
    ]
    legal_row = Table(
        [[legal, delivery]],
        colWidths=[100 * mm, 74 * mm],
    )
    legal_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(legal_row)
    story.append(Spacer(1, 18))

    # Signature
    story.append(Paragraph("Med vennlig hilsen", styles["normal"]))
    story.append(Paragraph(" Aisha Hussain", styles["normal"]))
    story.append(Paragraph("(sign.)", styles["normal"]))
    story.append(Paragraph("Ortopediteknikk AS", styles["normal"]))

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Genererer format-test-PDF-er i {OUTPUT_DIR}/")
    print(f"Dato: {TODAY_HUMAN_DOT}   Unik suffiks: {UNIQ}\n")

    # ------------------------------------------------------------------
    # 1. Blatchford — 2 linjer, priser + Int. ordre ref
    # ------------------------------------------------------------------
    blatchford_number = f"8000 BO-2026-{DATESTAMP}{UNIQ[-2:]}"
    generate_blatchford(
        order_number=blatchford_number,
        lines=[
            # (art, desc, int.ref, rabatt, antall, enhet, beløp)
            ("SK1363-L/TI", "NEURO VARIO knee joint 16mm, ti., left",
             f"26{DATESTAMP}{UNIQ[-3:]}", 0, 1.0, "stk.", 9811.00),
            ("SK1383-R/TI", "NEURO VARIO knee joint 16mm, ti., left",
             f"26{DATESTAMP}{UNIQ[-3:]}", 0, 1.0, "stk.", 10097.00),
        ],
        total=19908.00,
    )

    # ------------------------------------------------------------------
    # 2. Bergen Mekaniske — 3 linjer inkl. transport, rabatt %-kolonne
    # ------------------------------------------------------------------
    bm_number = f"9{UNIQ}"
    generate_bergen_mekaniske(
        order_number=bm_number,
        lines=[
            # (pid, navn, lev.prodnr, ønsket lev., antall, enhet, pris, rabatt %, beløp)
            ("102633", "[TA4610-6X400ML] TA4610 6x400ml dual cartrdg+nozzles",
             "", "", 2.00, "stk", 6716.00, 0.00, 13432.00),
            ("102634", "[Z0100-MANUAL400ML1TO] 1:1Manual 400ml Dispensing Gun",
             "", "", 1.00, "stk", 1999.00, 0.00, 1999.00),
            ("145", "Transport",
             "", "", 1.00, "stk", 399.00, 0.00, 399.00),
        ],
        total=15830.00,
    )

    # ------------------------------------------------------------------
    # 3. Ortopediteknikk — 1 linje, ingen pris i tabellen
    # ------------------------------------------------------------------
    ot_number = f"OT-{DATESTAMP}{UNIQ[-3:]}"
    generate_ortopediteknikk(
        order_number=ot_number,
        internal_ref=f"25{DATESTAMP}{UNIQ[-4:]}",
        lines=[
            # (art, desc, int.ref, antall, enhet)
            ("KS3040-TI", "articulated side bar w. gear segments, lamin. left lateral",
             f"25{DATESTAMP}{UNIQ[-4:]}", 1.0, "stk."),
        ],
    )

    print(f"\n{'='*60}")
    print(f"  Genererte 3 format-test-PDF-er i {OUTPUT_DIR}")
    print(f"{'='*60}\n")
    print("Format-scenarier:")
    print("  Blatchford       — 'Int. ordre ref.' + 'Rabatt'-kolonne, 2 linjer, NOK-total")
    print("  Bergen Mekaniske — prosjekt-ref, rabatt-%-kolonne, transport som egen linje")
    print("  Ortopediteknikk  — INGEN priser i tabellen (Odoo må hente pris fra prisliste)")
    print()
    print("Ordrenumre er unike per kjøring (basert på tidsstempel),")
    print("så de vil ikke kollidere med tidligere testkjøringer.")


if __name__ == "__main__":
    main()
