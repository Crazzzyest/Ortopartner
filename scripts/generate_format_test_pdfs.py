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


# ------------------------------------------------------------------
# Format 4: Sophies Minde Ortopedi
# ------------------------------------------------------------------
def generate_sophies_minde(order_number: str, lines, total: float):
    """Sophies Minde format: metadata box, multiple lines, 10% discount,
    internal order ref, delivery address right-aligned."""
    styles = _styles()
    filename = f"Bestilling_SophiesMinde_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling {order_number}",
        author="Sophies Minde Ortopedi AS",
    )
    story = []

    # Logo placeholder (teal-ish like original)
    logo_style = ParagraphStyle(
        "sm_logo", parent=styles["title"], fontSize=20,
        textColor=colors.HexColor("#008080"), alignment=1,
        fontName="Helvetica-Bold",
    )
    logo_sub = ParagraphStyle(
        "sm_sub", parent=styles["normal"], fontSize=12,
        textColor=colors.HexColor("#008080"), alignment=1,
        fontName="Helvetica",
    )
    story.append(Paragraph("SOPHIES MINDE", logo_style))
    story.append(Paragraph("ORTOPEDI", logo_sub))
    story.append(Spacer(1, 14))

    # Title
    story.append(Paragraph(f"<b>Bestilling {order_number}</b>", styles["h1"]))
    story.append(Spacer(1, 4))

    # Ortopartner block (left) + metadata box (right)
    left_block = [
        Paragraph("<b>Ortopartner AS</b>", styles["normal"]),
        Paragraph("Fabrikkvegen 1, Inngang 16, 3. etg", styles["normal"]),
        Spacer(1, 6),
        Paragraph("5265 &nbsp;&nbsp;&nbsp; Ytre Arna", styles["normal"]),
        Paragraph("Norge", styles["normal"]),
        Spacer(1, 4),
        Paragraph("post@ortopartner.no", styles["normal"]),
    ]

    meta_rows = [
        ["Bestillingsnr.", order_number],
        ["Dato", TODAY_HUMAN_DOT],
        ["Deres ref.", ""],
        ["Vår referanse", "Knut Arild Mælum"],
        ["Int. ordre ref.", ""],
        ["Valuta", "NOK"],
        ["Side", "1 av 1"],
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
        colWidths=[26 * mm, 54 * mm, 20 * mm, 14 * mm, 14 * mm, 14 * mm, 24 * mm],
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

    # Total
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
    story.append(Spacer(1, 14))

    # Legal + delivery
    legal = Paragraph(
        "Se vedlagte bestilling.<br/>"
        "Vi imøteser deres ordrebekreftelse med priser og forventet leveringstid.<br/>"
        "Ordrebekreftelse sendes til: ordre@sophiesminde.no<br/>"
        "Faktura sendes til: faktura@sophiesminde.no",
        styles["small"],
    )
    delivery = [
        Paragraph("Leveringsadresse:", styles["small_right"]),
        Paragraph("Sophies Minde Ortopedi AS", styles["small_right"]),
        Paragraph("Avd. Bryn", styles["small_right"]),
        Paragraph("Brynsveien 14", styles["small_right"]),
        Paragraph("0667 Oslo", styles["small_right"]),
    ]
    legal_row = Table(
        [[legal, delivery]],
        colWidths=[105 * mm, 69 * mm],
    )
    legal_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(legal_row)
    story.append(Spacer(1, 14))

    # Signature
    story.append(Paragraph("Med vennlig hilsen/Best Regards", styles["normal"]))
    story.append(Paragraph("Knut Arild Mælum", styles["normal"]))
    story.append(Paragraph("(sign.)", styles["normal"]))
    story.append(Paragraph("Sophies Minde Ortopedi AS", styles["normal"]))

    # Footer
    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor("#aaa")))
    footer_row = Table(
        [["SOPHIES MINDE\nORTOPEDI AS", "POSTBOKS 493 Økern\n0512 Oslo",
          "Tlf 22 04 53 60\nOrg.nr. 986 116 710",
          "Post@sophiesminde.no\nwww.sophiesminde.no"]],
        colWidths=[40 * mm, 42 * mm, 42 * mm, 50 * mm],
    )
    footer_row.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(footer_row)

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


# ------------------------------------------------------------------
# Format 5: Norsk Teknisk Ortopedi (NTO)
# ------------------------------------------------------------------
def generate_nto(order_number: str, lines):
    """NTO format: simple table with 'Rest' column instead of quantity,
    no price column. Header says 'BESTILLING'."""
    styles = _styles()
    filename = f"Bestilling_NTO_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling {order_number}",
        author="Norsk Teknisk Ortopedi AS",
    )
    story = []

    # NTO Logo
    nto_logo = ParagraphStyle(
        "nto_logo", parent=styles["title"], fontSize=28,
        textColor=colors.HexColor("#1a3366"), alignment=2,
        fontName="Helvetica-Bold",
    )
    story.append(Paragraph("NTO", nto_logo))
    story.append(Spacer(1, 14))

    # Title
    story.append(Paragraph("<b>BESTILLING</b>", styles["h1"]))
    story.append(Spacer(1, 4))

    # Ortopartner block (left) + metadata box (right)
    left_block = [
        Paragraph("<b>Ortopartner AS</b>", styles["normal"]),
        Paragraph("Fabrikkveien 1, Inngang 16, 3 etg", styles["normal"]),
        Spacer(1, 6),
        Paragraph("5265 &nbsp;&nbsp;&nbsp; YTRE ARNA", styles["normal"]),
        Paragraph("Norge", styles["normal"]),
        Spacer(1, 4),
        Paragraph("post@ortopartner.no", styles["normal"]),
    ]

    meta_rows = [
        ["Bestillingsnr.", order_number],
        ["Dato", TODAY_HUMAN_DOT],
        ["Deres ref.", ""],
        ["Vår referanse", "Stein Erik Furulund"],
        ["Int. ordre ref.", ""],
        ["Side", "1 av 1"],
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

    # Line table — no price, uses "Enhet" and "Rest" columns
    header = ["Artikkelnr.", "Beskrivelse", "Int. ordre ref.", "Enhet", "Rest"]
    rows = [header]
    for item in lines:
        art, desc, qty, unit = item
        rows.append([
            art, desc, "",
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
    story.append(Spacer(1, 40))

    # Signature + delivery
    sig = [
        Paragraph("Med vennlig hilsen", styles["normal"]),
        Paragraph("(sign.)", styles["normal"]),
        Paragraph("Norsk Teknisk Ortopedi AS", styles["normal"]),
    ]
    delivery = [
        Paragraph("Leveringsadresse:", styles["small_right"]),
        Paragraph("Norsk Teknisk Ortopedi AS", styles["small_right"]),
        Paragraph("Vikavegen 17", styles["small_right"]),
        Paragraph("2312 OTTESTAD", styles["small_right"]),
        Paragraph("Norge", styles["small_right"]),
    ]
    bottom_row = Table(
        [[sig, delivery]],
        colWidths=[100 * mm, 74 * mm],
    )
    bottom_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(bottom_row)

    # Footer
    story.append(Spacer(1, 40))
    story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor("#aaa")))
    footer_row = Table(
        [["Norsk Teknisk Ortopedi AS\nVikavegen 17\n2312 Ottestad",
          "Telefon: 62 57 44 44\nE-mail: nto@ortonor.no",
          "Bankgiro 1813.05.07626\nOrg. nr. NO 954 472 299 MVA\nwww.ortonor.no"]],
        colWidths=[55 * mm, 55 * mm, 65 * mm],
    )
    footer_row.setStyle(
        TableStyle(
            [
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(footer_row)

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


# ------------------------------------------------------------------
# Format 6: ForMotion Norway / Össur (INNKJØPSORDRE)
# ------------------------------------------------------------------
def generate_formotion(order_number: str, lines, total: float):
    """ForMotion format: 'INNKJØPSORDRE' header, leverandørnr,
    planned receipt date, CRT units, reference column."""
    styles = _styles()
    filename = f"Innkjopsordre_ForMotion_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Innkjøpsordre {order_number}",
        author="ForMotion Norway AS",
    )
    story = []

    # Logo + order header
    fm_logo = ParagraphStyle(
        "fm_logo", parent=styles["title"], fontSize=18,
        textColor=colors.HexColor("#333"), fontName="Helvetica-Bold",
    )
    fm_sub = ParagraphStyle(
        "fm_sub", parent=styles["normal"], fontSize=8,
        textColor=colors.HexColor("#888"),
    )
    order_hdr = ParagraphStyle(
        "order_hdr", parent=styles["title"], fontSize=16,
        fontName="Helvetica-Bold", alignment=2,
    )

    top_row = Table(
        [[
            [Paragraph("ForMotion™", fm_logo), Paragraph("ORTOPEDI", fm_sub)],
            [
                Paragraph("Side 1 av 1", styles["small_right"]),
                Paragraph("<b>INNKJØPSORDRE</b>", order_hdr),
                Spacer(1, 4),
                Table(
                    [
                        ["Ordernr.", order_number],
                        ["Ordredato", TODAY_ISO],
                        ["Leverandørnr", "202537"],
                    ],
                    colWidths=[28 * mm, 40 * mm],
                    style=TableStyle([
                        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                        ("FONTSIZE", (0, 0), (-1, -1), 9),
                        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ]),
                ),
            ],
        ]],
        colWidths=[90 * mm, 84 * mm],
    )
    top_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(top_row)
    story.append(Spacer(1, 14))

    # Delivery + supplier addresses
    delivery = [
        Paragraph("<b>Leveringsadresse:</b>", styles["normal_bold"]),
        Paragraph("ForMotion Norway avd. Helsfyr", styles["normal"]),
        Paragraph("Barbro T. Foss / IMP: Tone/Elisabet +47 23288200", styles["normal"]),
        Paragraph("Innspurten 9", styles["normal"]),
        Paragraph("0663 OSLO", styles["normal"]),
        Paragraph("NORWAY", styles["normal"]),
    ]
    supplier = [
        Paragraph("Ortopartner As", styles["normal"]),
        Paragraph("Kosta Stanic", styles["normal"]),
        Paragraph("Fabrikkveien 1", styles["normal"]),
        Paragraph("5265 YTRE ARNA", styles["normal"]),
        Paragraph("NORWAY", styles["normal"]),
        Paragraph("post@ortopartner.no", styles["normal"]),
    ]
    addr_row = Table(
        [[delivery, supplier]],
        colWidths=[100 * mm, 74 * mm],
    )
    addr_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(addr_row)
    story.append(Spacer(1, 6))

    # Ref + invoice address
    story.append(Paragraph("<b>Ref:</b> 2210", styles["normal"]))
    story.append(Spacer(1, 4))
    story.append(Paragraph("<b>Fakturaadresse:</b>", styles["normal_bold"]))
    story.append(Paragraph("ForMotion Norway AS", styles["normal"]))
    story.append(Paragraph("PB 6180 Etterstad", styles["normal"]))
    story.append(Paragraph("0602 OSLO", styles["normal"]))
    story.append(Spacer(1, 14))

    # Line table
    header = ["Nr.", "Beskrivelse", "Ant.", "Enhet",
              "Direkte\nenhetskost\nEkskl. mva.", "Rab.\n%",
              "Planlagt\nmottaksda\nto", "Reference", "Sum"]
    rows = [header]
    for item in lines:
        nr, desc, qty, unit, price, disc, recv_date, ref, amount = item
        rows.append([
            nr,
            Paragraph(desc, styles["normal"]),
            str(qty),
            unit,
            _fmt_nok(price),
            f"{disc}%" if disc else "",
            recv_date,
            ref,
            _fmt_nok(amount),
        ])

    lines_table = Table(
        rows,
        colWidths=[22 * mm, 44 * mm, 10 * mm, 12 * mm, 22 * mm,
                   10 * mm, 20 * mm, 16 * mm, 22 * mm],
    )
    lines_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
                ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(lines_table)
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.black))
    story.append(Spacer(1, 4))

    # Total
    total_row = Table(
        [["", "Total NOK", _fmt_nok(total)]],
        colWidths=[120 * mm, 30 * mm, 24 * mm],
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
    story.append(Spacer(1, 30))

    # Footer instructions
    story.append(Paragraph(
        "Faktura (PDF) sendes til invoice.norway@formotion.com .<br/>"
        "Send ordrebekreftelser inkludert leveringsdato til "
        "clinicsupplies@ossur.com innen 24 timer.",
        styles["small"],
    ))
    story.append(Spacer(1, 14))
    story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor("#aaa")))

    footer_row = Table(
        [["Orderkontakt:\nClinic Supplies\nKistagången 12\nSE - 164 40 Kista\n"
          "+46 8-46 50 16 30\nclinicsupplies@formotion.com",
          "Bankgiro\n\nSwift/Bic\nNDEANOKK\nIBAN\nNO6860050668315",
          "MVA-nr.\nNO936787819MVA\nSelskapet har F-skatteseddel\n\nBank\nNordea Bank Abp"]],
        colWidths=[60 * mm, 55 * mm, 60 * mm],
    )
    footer_row.setStyle(
        TableStyle(
            [
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(footer_row)

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


# ------------------------------------------------------------------
# Format 7: Teknomed — freeform email order (plain text body in PDF)
# ------------------------------------------------------------------
def generate_teknomed_email(order_ref: str, lines_text: str):
    """Teknomed format: plain-text email printed as PDF.
    No structured table — just body text with article numbers inline."""
    styles = _styles()
    filename = f"Bestilling_Teknomed_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"bestilling - Teknomed {order_ref}",
        author="Teknomed AS",
    )
    story = []

    email_hdr_style = ParagraphStyle(
        "email_hdr", parent=styles["normal"], fontSize=9, leading=13,
        textColor=colors.HexColor("#333"),
    )
    email_body_style = ParagraphStyle(
        "email_body", parent=styles["normal"], fontSize=11, leading=16,
    )

    # Simulate email header
    story.append(Paragraph("<b>Fra:</b> Lager &lt;lager@teknomed.no&gt;", email_hdr_style))
    story.append(Paragraph("<b>Til:</b> Post &lt;post@ortopartner.no&gt;", email_hdr_style))
    story.append(Paragraph(f"<b>Dato:</b> {TODAY_HUMAN_DOT}", email_hdr_style))
    story.append(Paragraph(f"<b>Emne:</b> bestilling ({order_ref})", email_hdr_style))
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#ccc")))
    story.append(Spacer(1, 18))

    # Email body
    story.append(Paragraph("Hei, vi ønsker å bestille følgende", email_body_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(lines_text, email_body_style))
    story.append(Spacer(1, 18))
    story.append(Paragraph("Per-Erling", email_body_style))

    doc.build(story)
    print(f"  + {filename}  ({order_ref})")
    return filename


# ------------------------------------------------------------------
# Format 8: Østo Ortopedisenter
# ------------------------------------------------------------------
def generate_osto(order_number: str, lines):
    """Østo format: own logo area, 'Merket' field, separate 'Lev. adr.' line,
    Produktnr/Beskrivelse/Enhet/Antall columns, NO prices."""
    styles = _styles()
    filename = f"Bestilling_Osto_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling {order_number}",
        author="Østo Ortopedisenter AS",
    )
    story = []

    # Logo placeholder + order header (right)
    osto_logo = ParagraphStyle(
        "osto_logo", parent=styles["title"], fontSize=22,
        textColor=colors.HexColor("#e65100"), fontName="Helvetica-Bold",
    )
    order_title = ParagraphStyle(
        "order_title", parent=styles["title"], fontSize=16,
        fontName="Helvetica-Bold", alignment=2,
    )

    meta_rows = [
        ["Dato:", TODAY_HUMAN_DOT],
        ["Side:", "1"],
        ["Vår ref.", "TS"],
    ]
    meta_tbl = Table(meta_rows, colWidths=[20 * mm, 30 * mm])
    meta_tbl.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))

    top_left = [
        Paragraph("Østo", osto_logo),
        Paragraph("Ortopedisenter AS", styles["normal_bold"]),
        Paragraph("Gartnerveien 10", styles["normal"]),
        Paragraph("2312 OTTESTAD", styles["normal"]),
        Spacer(1, 6),
        Paragraph("Telefon: 62 57 39 00 &nbsp; Org.nr: 930087335MVA", styles["small"]),
    ]
    top_right = [
        Table(
            [["Bestilling", order_number]],
            colWidths=[30 * mm, 40 * mm],
            style=TableStyle([
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 14),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ]),
        ),
        Spacer(1, 4),
        meta_tbl,
    ]

    top_row = Table(
        [[top_left, top_right]],
        colWidths=[100 * mm, 74 * mm],
    )
    top_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(top_row)
    story.append(Spacer(1, 12))

    # Ortopartner address
    story.append(Paragraph("Ortopartner AS", styles["normal"]))
    story.append(Paragraph("Inngang 16, 3 etg.", styles["normal"]))
    story.append(Paragraph("5265 YTRE ARNA", styles["normal"]))
    story.append(Spacer(1, 8))

    # Merket + Lev. adr.
    merket_nr = str(int(order_number) - 2) if order_number.isdigit() else order_number
    merket_row = Table(
        [["Merket:", merket_nr]],
        colWidths=[20 * mm, 80 * mm],
    )
    merket_row.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    story.append(merket_row)
    story.append(Spacer(1, 4))

    # Delivery address (orange highlight)
    lev_style = ParagraphStyle(
        "lev_adr", parent=styles["normal"], fontSize=9,
        textColor=colors.HexColor("#e65100"), fontName="Helvetica-Bold",
    )
    story.append(Paragraph(
        "Lev. adr. &nbsp;&nbsp; Østo Ortopedisenter AS, Vestre Rosten 79, 7075 TILLER",
        lev_style,
    ))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "Vennligst send ordrebekreftelse til hilde.hansen@osto.no",
        styles["normal"],
    ))
    story.append(Paragraph(
        "Oppgi bestillingsnr på pakkseddel og faktura",
        styles["normal"],
    ))
    story.append(Spacer(1, 12))

    # Line table — no prices
    header = ["Produktnr", "Beskrivelse", "Enhet", "Antall"]
    rows = [header]
    for item in lines:
        art, desc, unit, qty = item
        rows.append([art, Paragraph(desc, styles["normal"]), unit, str(qty)])

    lines_table = Table(
        rows,
        colWidths=[30 * mm, 100 * mm, 18 * mm, 18 * mm],
    )
    lines_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
                ("ALIGN", (3, 0), (3, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(lines_table)

    # Footer
    story.append(Spacer(1, 30))
    story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor("#aaa")))
    footer_row = Table(
        [["Østo Ortopedisenter AS\nGartnerveien 10\n2312 OTTESTAD",
          "Telefon: 62 57 39 00\nE-post: post@osto.no",
          "Org.nr: 930087335MVA\nForetaksregisteret"]],
        colWidths=[55 * mm, 55 * mm, 65 * mm],
    )
    footer_row.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(footer_row)

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


# ------------------------------------------------------------------
# Format 9: Blatchford Bergen (multi-page, 10 lines)
# ------------------------------------------------------------------
def generate_blatchford_bergen(order_number: str, lines, total: float):
    """Blatchford Bergen variant: reuses Blatchford format but with Bergen
    address, Sjur Atle Pettersen, and many more lines (10)."""
    styles = _styles()
    filename = f"Bestilling_BlatchfordBergen_{UNIQ}.pdf"
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

    # Logo block
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

    # Address + metadata
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
        ["Vår referanse", "Sjur Atle Pettersen"],
        ["Int. ordre ref.", ""],
        ["Valuta", "NOK"],
        ["Side", "1"],
    ]
    meta_table = Table(meta_rows, colWidths=[35 * mm, 55 * mm])
    meta_table.setStyle(
        TableStyle([
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#555")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbb")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ])
    )

    header_layout = Table(
        [[left_block, meta_table]],
        colWidths=[75 * mm, 95 * mm],
    )
    header_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header_layout)
    story.append(Spacer(1, 16))

    # Lines table
    header = ["Artikkelnr.", "Beskrivelse", "Int. ordre\nref.", "Rabatt", "Antall", "Enhet", "Beløp"]
    rows = [header]
    for item in lines:
        art, desc, ref, disc, qty, unit, amount = item
        rows.append([
            art,
            Paragraph(desc, styles["normal"]),
            ref,
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
        TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
            ("ALIGN", (2, 0), (6, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
        ])
    )
    story.append(lines_table)
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.black))
    story.append(Spacer(1, 4))

    # Total
    total_row = Table(
        [["", "Samlet ordreverdi (Valuta)", _fmt_nok(total)]],
        colWidths=[100 * mm, 50 * mm, 24 * mm],
    )
    total_row.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
    ]))
    story.append(total_row)
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.black))
    story.append(Spacer(1, 14))

    # Legal + delivery (Bergen)
    legal = Paragraph(
        "For denne bestillingen gjelder Blatchford Ortopedi AS sine alminnelige innkjøpsbetingelser.<br/><br/>"
        "Vennligst bekreft denne bestillingen med pris og leveringsdato.",
        styles["small"],
    )
    delivery = [
        Paragraph("Leveringsadresse:", styles["small_right"]),
        Paragraph("Blatchford Ortopedi AS", styles["small_right"]),
        Paragraph("Avd. Bergen", styles["small_right"]),
        Paragraph("Fjøsangerveien 215", styles["small_right"]),
        Paragraph("5073 Bergen", styles["small_right"]),
        Paragraph("Norge", styles["small_right"]),
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
    story.append(Paragraph("Sjur Atle Pettersen", styles["normal"]))
    story.append(Paragraph(
        "97643502 | sjur.atle.pettersen@blatchford.no | avd. Bergen",
        styles["normal"],
    ))

    # Footer
    story.append(Spacer(1, 30))
    story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor("#aaa")))
    story.append(Paragraph(
        "Org.nummer: 914 120 349 MVA | Bank konto: 1503 34 68633 | 55 27 11 00 | blatchford.no",
        styles["small"],
    ))
    story.append(Paragraph(
        "post.bergen@blatchford.no | Fjøsangerveien 215 - 5073 Bergen",
        styles["small"],
    ))

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


# ------------------------------------------------------------------
# Format 10: Evomotion Konfigurationsblatt (Configuration Sheet)
# ------------------------------------------------------------------
def generate_evomotion_config(commission_no: str, measurements: dict):
    """Evomotion Configuration Sheet — Shorts. This is NOT a standard
    purchase order; it's a configuration form with body measurements,
    product choices, and detail options."""
    styles = _styles()
    filename = f"Konfigurationsblatt_{commission_no}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
        title=f"Configuration Sheet – Shorts {commission_no}",
        author="Evomotion GmbH",
    )
    story = []

    # Title bar
    title_style = ParagraphStyle(
        "evo_title", parent=styles["title"], fontSize=18,
        textColor=colors.white, fontName="Helvetica-Bold",
        backColor=colors.HexColor("#F5A623"),
    )
    story.append(Paragraph(
        "&nbsp;Configuration Sheet – <b>Shorts</b>", title_style,
    ))
    story.append(Spacer(1, 10))

    # Commission number
    story.append(Paragraph(
        f"<b>Commission no / Patient ID:</b> &nbsp; {commission_no}",
        styles["normal"],
    ))
    story.append(Spacer(1, 10))

    # Step 1
    step1_hdr = ParagraphStyle(
        "step_hdr", parent=styles["normal_bold"], fontSize=10,
        textColor=colors.HexColor("#F5A623"),
    )
    story.append(Paragraph("Step 1: What is ordered?", step1_hdr))
    story.append(Spacer(1, 4))

    m = measurements
    product = m.get("product", "Evomove Solokit")
    leg = m.get("leg", "right leg")
    gender = m.get("gender", "female")
    color = m.get("color", "black")
    supply = m.get("supply", "Continuous supply")

    choices_data = [
        ["Product:", product],
        ["Leg:", leg],
        ["Gender:", gender],
        ["Fabric color:", color],
        ["Supply type:", supply],
    ]
    choices_tbl = Table(choices_data, colWidths=[30 * mm, 80 * mm])
    choices_tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(choices_tbl)
    story.append(Spacer(1, 10))

    # Step 2: Measurements
    story.append(Paragraph("Step 2: Measurement (mandatory!)", step1_hdr))
    story.append(Spacer(1, 4))

    meas_data = [
        ["Circumference", "", "cm", "Lengths", "", "cm"],
        ["U1 Waist", "", str(m.get("U1", "")),
         "L1 Knee length", "", str(m.get("L1", ""))],
        ["U2 Buttocks", "", str(m.get("U2", "")),
         "L2 Buttocks", "", str(m.get("L2", ""))],
        ["U3 Thigh (left)", "", str(m.get("U3_left", "")),
         "L3 Seat length", "", str(m.get("L3", ""))],
        ["U3 Thigh (right)", "", str(m.get("U3_right", "")),
         "Body height", "", str(m.get("body_height", ""))],
        ["U4 Knee (left)", "", str(m.get("U4_left", "")),
         "", "", ""],
        ["U4 Knee (right)", "", str(m.get("U4_right", "")),
         "", "", ""],
    ]
    meas_tbl = Table(
        meas_data,
        colWidths=[32 * mm, 10 * mm, 18 * mm, 32 * mm, 10 * mm, 18 * mm],
    )
    meas_tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.HexColor("#F5A623")),
        ("ALIGN", (2, 0), (2, -1), "RIGHT"),
        ("ALIGN", (5, 0), (5, -1), "RIGHT"),
    ]))
    story.append(meas_tbl)
    story.append(Spacer(1, 10))

    # Step 4: Details
    story.append(Paragraph("Step 4: Details of the shorts", step1_hdr))
    story.append(Spacer(1, 4))

    details = m.get("details", {})
    detail_data = [
        ["Cord in waistband:", details.get("cord", "No")],
        ["Smartphone pocket:", details.get("smartphone_pocket", "none")],
        ["Zipper:", details.get("zipper", "No")],
        ["Pocket/cable outlet:", details.get("cable_outlet", "Standard (lateral)")],
        ["Wearing height:", details.get("wearing_height", "Waist shorts")],
    ]
    detail_tbl = Table(detail_data, colWidths=[40 * mm, 80 * mm])
    detail_tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(detail_tbl)
    story.append(Spacer(1, 20))

    # Footer
    story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor("#F5A623")))
    story.append(Paragraph(
        "© Evomotion GmbH, Version 3.2.0 &nbsp; Please send the completed configuration sheet "
        "with all images or scan data and any other comments to bestellung@evomotion.de.",
        styles["small"],
    ))

    doc.build(story)
    print(f"  + {filename}  ({commission_no})")
    return filename


# ------------------------------------------------------------------
# Format 11: Drevelin Ortopedi Sør (freeform email)
# ------------------------------------------------------------------
def generate_drevelin_email(order_ref: str, lines_text: str):
    """Drevelin format: freeform email order for prepreg material,
    non-standard units (kvm)."""
    styles = _styles()
    filename = f"Bestilling_Drevelin_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling preg ({order_ref})",
        author="Drevelin Ortopedi Sør",
    )
    story = []

    email_hdr_style = ParagraphStyle(
        "drev_hdr", parent=styles["normal"], fontSize=9, leading=13,
        textColor=colors.HexColor("#333"),
    )
    email_body_style = ParagraphStyle(
        "drev_body", parent=styles["normal"], fontSize=11, leading=16,
    )

    story.append(Paragraph("<b>Fra:</b> Grethe Helene Golshani &lt;grethe@drevelinsor.no&gt;", email_hdr_style))
    story.append(Paragraph("<b>Til:</b> Post &lt;post@ortopartner.no&gt;", email_hdr_style))
    story.append(Paragraph(f"<b>Dato:</b> {TODAY_HUMAN_DOT}", email_hdr_style))
    story.append(Paragraph(f"<b>Emne:</b> Bestilling preg ({order_ref})", email_hdr_style))
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#ccc")))
    story.append(Spacer(1, 18))

    story.append(Paragraph("Hei", email_body_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(lines_text, email_body_style))
    story.append(Spacer(1, 18))
    story.append(Paragraph("Med vennlig hilsen", email_body_style))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Grethe Helene Golshani", email_body_style))
    story.append(Paragraph("Logistikkansvarlig", styles["small"]))
    story.append(Paragraph("Drevelin Ortopedi Sør", styles["small"]))
    story.append(Paragraph("Mobil: +47 41263049", styles["small"]))
    story.append(Paragraph("E-post: grethe@drevelinsor.no", styles["small"]))

    doc.build(story)
    print(f"  + {filename}  ({order_ref})")
    return filename


# ------------------------------------------------------------------
# Format 12: Ortopediteknikk med priser (Lillestrøm)
# ------------------------------------------------------------------
def generate_ortopediteknikk_med_pris(order_number: str, lines, total: float):
    """Ortopediteknikk variant WITH prices and int.ordreref.
    Delivery to Lillestrøm instead of Oslo."""
    styles = _styles()
    filename = f"Bestilling_OrtopediteknikkLillestrom_{UNIQ}.pdf"
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

    # Logo
    logo = [
        Paragraph("<b>ORTOPEDITEKNIKK</b>", styles["logo_ot"]),
    ]
    logo_row = Table([[None, logo]], colWidths=[90 * mm, 84 * mm])
    logo_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(logo_row)
    story.append(Spacer(1, 14))

    # Title
    story.append(Paragraph("BESTILLING FRA ORTOPEDITEKNIKK", styles["h1"]))
    story.append(Spacer(1, 4))

    # Address + metadata
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
        ["Vår referanse", "Linn-Therese Østlie"],
        ["Int. ordre ref.", ""],
        ["Valuta", "NOK"],
        ["Side", "1"],
    ]
    meta_table = Table(meta_rows, colWidths=[35 * mm, 55 * mm])
    meta_table.setStyle(
        TableStyle([
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#555")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbb")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ])
    )

    header_layout = Table(
        [[left_block, meta_table]],
        colWidths=[75 * mm, 95 * mm],
    )
    header_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header_layout)
    story.append(Spacer(1, 16))

    # Lines with prices and int. ordreref
    header = ["Artikkelnr.", "Beskrivelse", "Int. ordre\nref.", "Rabatt", "Antall", "Enhet", "Beløp"]
    rows = [header]
    for item in lines:
        art, desc, ref, disc, qty, unit, amount = item
        rows.append([
            art,
            Paragraph(desc, styles["normal"]),
            ref,
            f"{disc}%" if disc else "",
            f"{qty:.1f}".replace(".", ","),
            unit,
            _fmt_nok(amount),
        ])

    lines_table = Table(
        rows,
        colWidths=[24 * mm, 58 * mm, 22 * mm, 14 * mm, 14 * mm, 14 * mm, 22 * mm],
    )
    lines_table.setStyle(
        TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
            ("ALIGN", (2, 0), (6, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
        ])
    )
    story.append(lines_table)
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.black))
    story.append(Spacer(1, 4))

    # Total
    total_row = Table(
        [["", "Samlet ordreverdi (Valuta)", _fmt_nok(total)]],
        colWidths=[100 * mm, 50 * mm, 24 * mm],
    )
    total_row.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
    ]))
    story.append(total_row)
    story.append(Spacer(1, 14))

    # Legal + delivery Lillestrøm
    legal = Paragraph(
        "Vennligst bekreft bestillingen med en ordrebekreftelse. Faktura merkes med vårt ordrenummer "
        "og sendes til faktura@ortopediteknikk.no eller EHF-faktura til org.nr. 930614785.",
        styles["small"],
    )
    delivery = [
        Paragraph("Leveringsadresse:", styles["small_right"]),
        Paragraph("Lillestrøm", styles["small_right"]),
        Paragraph("Dampsagveien 4, 2004", styles["small_right"]),
        Paragraph("Lillestrøm", styles["small_right"]),
    ]
    legal_row = Table(
        [[legal, delivery]],
        colWidths=[100 * mm, 74 * mm],
    )
    legal_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(legal_row)
    story.append(Spacer(1, 18))

    story.append(Paragraph("Med vennlig hilsen", styles["normal"]))
    story.append(Paragraph("Linn-Therese Østlie", styles["normal"]))
    story.append(Paragraph("(sign.)", styles["normal"]))
    story.append(Paragraph("Ortopediteknikk AS", styles["normal"]))

    doc.build(story)
    print(f"  + {filename}  ({order_number})")
    return filename


# ------------------------------------------------------------------
# Format 13: Atterås
# ------------------------------------------------------------------
def generate_atteras(order_number: str, lines):
    """Atterås format: own branding, simple table with art/desc/int.ref/antall/enhet,
    no prices. Delivery to Bergen."""
    styles = _styles()
    filename = f"Bestilling_Atteras_{UNIQ}.pdf"
    filepath = OUTPUT_DIR / filename

    doc = SimpleDocTemplate(
        str(filepath),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Bestilling {order_number}",
        author="Atterås AS",
    )
    story = []

    # Logo + contact info
    att_logo = ParagraphStyle(
        "att_logo", parent=styles["title"], fontSize=22,
        textColor=colors.HexColor("#2e7d9e"), fontName="Helvetica-Bold",
    )
    att_sub = ParagraphStyle(
        "att_sub", parent=styles["normal"], fontSize=9,
        textColor=colors.HexColor("#2e7d9e"), fontName="Helvetica-Oblique",
    )
    att_contact = ParagraphStyle(
        "att_contact", parent=styles["normal"], fontSize=9,
        textColor=colors.HexColor("#2e7d9e"), alignment=2,
    )

    top_row = Table(
        [[
            [Paragraph("Atterås", att_logo), Paragraph("Vi skaper bevegelse!", att_sub)],
            [
                Paragraph("<b>Atterås AS</b>", att_contact),
                Paragraph("Telefon: 936 86 000", att_contact),
                Paragraph("Epost: post@atteraas.no", att_contact),
            ],
        ]],
        colWidths=[90 * mm, 84 * mm],
    )
    top_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(top_row)
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#2e7d9e")))
    story.append(Spacer(1, 14))

    # Title
    story.append(Paragraph("<b>BESTILLING</b>", styles["h1"]))
    story.append(Spacer(1, 4))

    # Address + metadata
    left_block = [
        Paragraph("<b>Ortopartner</b>", styles["normal"]),
        Paragraph("Fabrikkvegen 1", styles["normal"]),
        Paragraph("Inngang 16", styles["normal"]),
        Paragraph("5265 &nbsp;&nbsp;&nbsp; YTRE ARNA", styles["normal"]),
        Paragraph("Norge", styles["normal"]),
        Spacer(1, 4),
        Paragraph("post@ortopartner.no", styles["normal"]),
    ]

    meta_rows = [
        ["Bestillingsnr.", order_number],
        ["Dato", TODAY_HUMAN_DOT],
        ["Deres ref.", ""],
        ["Vår referanse", "Stefan Bernsen"],
        ["Int. ordre ref.", ""],
        ["Side", "1"],
    ]
    meta_table = Table(meta_rows, colWidths=[35 * mm, 55 * mm])
    meta_table.setStyle(
        TableStyle([
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#555")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bbb")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ])
    )

    header_layout = Table(
        [[left_block, meta_table]],
        colWidths=[75 * mm, 95 * mm],
    )
    header_layout.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header_layout)
    story.append(Spacer(1, 16))

    # Lines — no prices
    header = ["Artikkelnr.", "Beskrivelse", "Int. ordre ref.", "Antall", "Enhet"]
    rows = [header]
    for item in lines:
        art, desc, qty, unit = item
        rows.append([art, desc, "", f"{qty:.1f}".replace(".", ","), unit])

    lines_table = Table(
        rows,
        colWidths=[28 * mm, 80 * mm, 28 * mm, 18 * mm, 18 * mm],
    )
    lines_table.setStyle(
        TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
            ("ALIGN", (3, 0), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
        ])
    )
    story.append(lines_table)
    story.append(Spacer(1, 20))

    # Legal + delivery
    legal = Paragraph(
        "For denne bestillingen gjelder Atterås sine Alminnelige innkjøpsbetingelser.<br/>"
        "Vennligst bekreft denne bestillingen ved å returnere en signert kopi.",
        styles["small"],
    )
    delivery = [
        Paragraph("Leveringsadresse:", styles["small_right"]),
        Paragraph("Atterås AS", styles["small_right"]),
        Paragraph("Møllendalsveien 1", styles["small_right"]),
        Paragraph("5009 Bergen", styles["small_right"]),
    ]
    legal_row = Table(
        [[legal, delivery]],
        colWidths=[100 * mm, 74 * mm],
    )
    legal_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(legal_row)
    story.append(Spacer(1, 18))

    story.append(Paragraph("Med vennlig hilsen", styles["normal"]))
    story.append(Paragraph("Stefan Bernsen", styles["normal"]))
    story.append(Paragraph("(sign.)", styles["normal"]))
    story.append(Paragraph("Atterås AS", styles["normal"]))

    # Footer line
    story.append(Spacer(1, 30))
    story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor("#aaa")))

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
    blatchford_number = f"TEST-BL-{DATESTAMP}{UNIQ[-4:]}"
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
    bm_number = f"TEST-BM-{DATESTAMP}{UNIQ[-4:]}"
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
    ot_number = f"TEST-OT-{DATESTAMP}{UNIQ[-4:]}"
    generate_ortopediteknikk(
        order_number=ot_number,
        internal_ref=f"25{DATESTAMP}{UNIQ[-4:]}",
        lines=[
            # (art, desc, int.ref, antall, enhet)
            ("KS3040-TI", "articulated side bar w. gear segments, lamin. left lateral",
             f"25{DATESTAMP}{UNIQ[-4:]}", 1.0, "stk."),
        ],
    )

    # ------------------------------------------------------------------
    # 4. Sophies Minde — 9 linjer, 10% rabatt, intern ordreref, NOK-total
    # ------------------------------------------------------------------
    sm_number = f"TEST-SM-{DATESTAMP}{UNIQ}"
    sm_ref = f"441{UNIQ[-3:]}"
    generate_sophies_minde(
        order_number=sm_number,
        lines=[
            # (art, desc, int.ref, rabatt, antall, enhet, beløp)
            ("107P31/100", "Sanding Net, 230x280Mm, grit 100 5Plate",
             "", 0, 10.0, "pakke", 2960.00),
            ("107P31/220", "Sanding Net, 230x280Mm, grit 220 5Plate",
             "", 0, 5.0, "pakke", 1480.00),
            ("69T11/20W", "Elastic Strap Button Holes, white, 20Mm 25m",
             "", 0, 1.0, "rull", 4040.00),
            ("FB5182-C/ST/1", "stirrup lamin./prepreg, 14mm, st. 50.7mm long,2.5mm,",
             sm_ref, "10,0", 2.0, "stk.", 1440.00),
            ("FB5282-C/ST/2", "stirrup thermoform, 14mm, st. 62mm long, 2.5mmk, straight",
             sm_ref, "10,0", 2.0, "stk.", 1305.00),
            ("SA2042-TI", "anchor, lamin. technic 14mm, G5 ti. straight",
             sm_ref, "10,0", 2.0, "stk.", 3074.40),
            ("SF5802-C/10/29", "spring unit H2O 14mm, yellow, very strong,max. 10°",
             sm_ref, "10,0", 2.0, "stk.", 3884.40),
            ("SF5802-C/15/05", "spring unit H2O 14mm, blue, normal,Max. 15°",
             sm_ref, "10,0", 2.0, "stk.", 3884.40),
            ("SF5802-C/15/11", "spring unit H2O 14mm, green,Medium, Max. 15°",
             sm_ref, "10,0", 2.0, "stk.", 3884.40),
        ],
        total=25952.60,
    )

    # ------------------------------------------------------------------
    # 5. NTO — 1 linje, ingen pris, "Rest"-kolonne
    # ------------------------------------------------------------------
    nto_number = f"TEST-NTO-2026-{UNIQ}"
    generate_nto(
        order_number=nto_number,
        lines=[
            # (art, desc, antall, enhet)
            ("119P2/M", "Latex Insulating Bag Medium", 20.0, "stk."),
        ],
    )

    # ------------------------------------------------------------------
    # 6. ForMotion — INNKJØPSORDRE, planlagt mottaksdato, CRT-enhet
    # ------------------------------------------------------------------
    fm_number = f"TEST-POOCH{UNIQ}"
    recv_date = NOW.strftime("%Y-%m-%d")
    generate_formotion(
        order_number=fm_number,
        lines=[
            # (nr, desc, qty, unit, price, disc, recv_date, ref, amount)
            ("V_ELGEL_400", "Elektrodengel Set 100ml (EA) x 4/CRT",
             5, "CRT", 603.00, 0, recv_date, "", 3015.00),
        ],
        total=3015.00,
    )

    # ------------------------------------------------------------------
    # 7. Teknomed — fritekst e-post-ordre (ingen tabell)
    # ------------------------------------------------------------------
    tek_ref = f"TEK-{UNIQ}"
    generate_teknomed_email(
        order_ref=tek_ref,
        lines_text="1 stk Ankeljoint SF0503-C.",
    )

    # ------------------------------------------------------------------
    # 8. Østo Ortopedisenter — 11 linjer, ingen priser, Merket-felt,
    #    separat leveringsadresse (Tiller)
    # ------------------------------------------------------------------
    osto_number = f"TEST-{UNIQ}86"
    generate_osto(
        order_number=osto_number,
        lines=[
            # (art, desc, unit, qty)
            ("FB5193-ST/1", "Prepreg stirrup 16 mm", "Stk.", 2),
            ("SF5203-TI/LR", "Neuro Swing 16 mm", "Stk.", 2),
            ("FB5293-LR/ST2", "Stirrup Neuro swing, thermoform, 16 mm", "Stk.", 2),
            ("SF5803-15/07", "Spring unit for NEURO SWING blue, normal", "Stk.", 1),
            ("SF5803-15/15", "Spring unit for NEURO SWING, 16mm green", "Stk.", 1),
            ("SF5803-10/21", "Spring unit for NEURO SWING, 16mm hvit", "Stk.", 1),
            ("SF5803-10/31", "Spring unit for NEURO SWING yellow very strong", "Stk.", 1),
            ("SF5803-05/63", "Spring unit for NEURO SWING red, extra strong", "Stk.", 1),
            ("SA1063-TI", "System anchor, 16mm TI straight", "Stk.", 2),
            ("FB2293-R/ST3", "Thermoformage stirrup 16mm right", "Stk.", 1),
            ("PE4000-LR", "Holder for ankle joints model technic, square 15x15x30mm", "Stk.", 4),
        ],
    )

    # ------------------------------------------------------------------
    # 9. Blatchford Bergen — 10 linjer, 2-siders potensial, int.ordreref
    # ------------------------------------------------------------------
    bb_number = f"TEST-3000-BO-2026-{UNIQ[-4:]}"
    bb_ref = f"25{UNIQ[-5:]}"
    generate_blatchford_bergen(
        order_number=bb_number,
        lines=[
            # (art, desc, int.ref, rabatt, antall, enhet, beløp)
            ("e01_ScEl-g", "Screening Electrodes Set big (CFF153) 4 Stk",
             f"26{UNIQ[-5:]}", 0, 2.0, "pakke", 336.00),
            ("FB5015-LR/ST5", "stirrup rivet, 20mm, st., L/R leg 165mm long, 3mm, bent",
             bb_ref, 0, 2.0, "stk.", 2550.00),
            ("PE1025-LR", "joint retainer for uniaxial joints 16mm &amp; 20mm",
             bb_ref, 0, 2.0, "stk.", 882.00),
            ("PE2000-LR", "holder for knee joints model technic, square 15x15x40mm",
             bb_ref, 0, 2.0, "stk.", 1758.00),
            ("PL3687-02/1", "xDRY 2mm PU foam layer, black, 1000x1400x4mm",
             "", 0, 3.0, "plate", 4059.00),
            ("SA1085-TI", "anchor, lamin. technic 20mm, G2 ti. straight",
             bb_ref, 0, 2.0, "stk.", 4694.00),
            ("SH5205-TI/LR", "NEURO SWING 2 ankle joint 20mm, ti., L/R leg straight",
             bb_ref, 0, 2.0, "stk.", 26032.00),
            ("SH5805-05/99", "spring unit 20mm, red, extra strong, max. 5°",
             bb_ref, 0, 2.0, "stk.", 3728.00),
            ("SH5805-15/25", "spring unit 20mm, green, medium, max. 15° range of motion",
             bb_ref, 0, 2.0, "stk.", 3728.00),
            ("SILI", "Silicone Release Agent (2L) - Holdbarhet 2 år",
             "", 0, 2.0, "stk.", 6452.00),
        ],
        total=54219.00,
    )

    # ------------------------------------------------------------------
    # 10. Evomotion Konfigurationsblatt — Configuration Sheet
    # ------------------------------------------------------------------
    evo_commission = f"JFNO26{UNIQ}"
    generate_evomotion_config(
        commission_no=evo_commission,
        measurements={
            "product": "Evomove Solokit (incl. pocket)",
            "leg": "right leg",
            "gender": "female",
            "color": "black",
            "supply": "Continuous supply",
            "U1": 70.5, "U2": 94.0,
            "U3_left": 54.0, "U3_right": 54.0,
            "U4_left": 34.0, "U4_right": 35.0,
            "L1": 60.0, "L2": 16.0, "L3": 22.0,
            "body_height": 178.0,
            "details": {
                "cord": "No",
                "smartphone_pocket": "none",
                "zipper": "No",
                "cable_outlet": "Standard (lateral)",
                "wearing_height": "Waist shorts",
            },
        },
    )

    # ------------------------------------------------------------------
    # 11. Drevelin Ortopedi Sør — fritekst e-post, prepreg-bestilling
    # ------------------------------------------------------------------
    drev_ref = f"DREV-{UNIQ}"
    generate_drevelin_email(
        order_ref=drev_ref,
        lines_text="Ønsker å bestille 20 kvm vevd preg:<br/><br/>"
                   "11C2 Carbon Fibre Pre-preg 280 g/sqm width: 1250",
    )

    # ------------------------------------------------------------------
    # 12. Ortopediteknikk Lillestrøm — 7 linjer MED priser + int.ordreref
    # ------------------------------------------------------------------
    otl_number = f"TEST-OT-{UNIQ[-4:]}"
    otl_ref = f"256{UNIQ[-5:]}"
    generate_ortopediteknikk_med_pris(
        order_number=otl_number,
        lines=[
            # (art, desc, ref, rabatt, antall, enhet, beløp)
            ("BK2303-L/AL", "NEURO ACTIVE articulated side bar 16mm, 22mm centre dist.",
             otl_ref, 0, 1.0, "stk.", 6417.00),
            ("BK9051-F020", "20° flexion stop for 22mm centre dist., st., 5mm",
             otl_ref, 0, 1.0, "stk.", 351.00),
            ("BK9051-F030", "30° flexion stop for 22mm centre dist., st., 5mm",
             otl_ref, 0, 1.0, "stk.", 351.00),
            ("FB5093-ST/4", "stirrup rivet, 16mm, st. 145mm long, 3mm, straight",
             "", 0, 1.0, "stk.", 999.00),
            ("PZ3100-LR", "joint retainers 16mm &amp; 20mm width 16mm und 20mm",
             otl_ref, 0, 1.0, "sett", 2246.00),
            ("SH4203-L/ST", "NEURO VARIO-SWING ankle joint 16mm, st., left",
             "", 0, 1.0, "stk.", 10172.00),
            ("SH4223-L/ST", "NEURO VARIO-SWING ankle joint 16mm, st., left",
             "", 0, 1.0, "stk.", 10666.00),
        ],
        total=31202.00,
    )

    # ------------------------------------------------------------------
    # 13. Atterås — 1 linje, ingen pris, eget format
    # ------------------------------------------------------------------
    att_number = f"TEST-ATT-{UNIQ[-4:]}"
    generate_atteras(
        order_number=att_number,
        lines=[
            # (art, desc, qty, unit)
            ("166P6/220", "Pinking Shears, coated handles, 220 mm", 3.0, "stk."),
        ],
    )

    print(f"\n{'='*60}")
    print(f"  Genererte 13 format-test-PDF-er i {OUTPUT_DIR}")
    print(f"{'='*60}\n")
    print("Format-scenarier:")
    print("   1. Blatchford Arendal — 'Int. ordre ref.' + 'Rabatt'-kolonne, 2 linjer")
    print("   2. Bergen Mekaniske   — prosjekt-ref, rabatt-%, transport-linje")
    print("   3. Ortopediteknikk    — INGEN priser (pris-lookup)")
    print("   4. Sophies Minde      — 9 linjer, 10% rabatt, int. ordreref")
    print("   5. NTO                — 'Rest'-kolonne, ingen pris")
    print("   6. ForMotion          — INNKJØPSORDRE, CRT-enhet, mottaksdato")
    print("   7. Teknomed           — Fritekst e-post (ingen tabell)")
    print("   8. Østo               — 11 linjer, Merket-felt, ingen pris, lev.adr. Tiller")
    print("   9. Blatchford Bergen  — 10 linjer, int.ordreref, avd. Bergen")
    print("  10. Evomotion          — Konfigurationsblatt (kroppsmål, ikke standard PO)")
    print("  11. Drevelin           — Fritekst e-post, prepreg kvm-bestilling")
    print("  12. Ortopediteknikk LS — 7 linjer MED priser, levering Lillestrøm")
    print("  13. Atterås            — 1 linje, ingen pris, eget format")
    print()
    print("Ordrenumre er unike per kjøring (basert på tidsstempel),")
    print("så de vil ikke kollidere med tidligere testkjøringer.")


if __name__ == "__main__":
    main()
