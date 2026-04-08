"""LLM-based order parsing using Claude API."""

from __future__ import annotations

import json
from pathlib import Path

import anthropic

from .config import load_config
from .models import ParsedOrder
from .pdf_extractor import extract_text, get_pdf_info, pdf_to_base64_images

load_config()

SYSTEM_PROMPT = """\
Du er en AI-agent som parser innkjøpsordre-PDF-er for Ortopartner AS.
Ortopartner mottar bestillinger fra ortopediske verksteder og klinikker i Norge.

Din oppgave er å ekstrahere strukturerte data fra PDF-teksten og returnere et JSON-objekt.

## Felter å ekstrahere

- **order_number**: Bestillingsnr / Ordernr / Best. nr. (f.eks. "DV-119464", "POOCH055385", "438435")
- **order_date**: Dato i ISO-format YYYY-MM-DD (konverter fra DD.MM.YYYY eller YYYY-MM-DD)
- **customer_name**: Firmaet som bestiller (IKKE Ortopartner – det er mottaker). Se etter logo, header, bunntekst eller avsender.
- **customer_reference**: "Vår referanse" / kontaktperson hos bestiller
- **internal_order_ref**: "Int. ordre ref." på hodenivå (om det finnes)
- **currency**: Valuta (default NOK)
- **line_items**: Liste med ordrelinjer:
  - article_number: Artikkelnr / Produktnr / Nr
  - description: Beskrivelse / Produktnavn
  - quantity: Antall / Ant. / Rest (numerisk)
  - unit: Enhet (stk, pakke, plate, rull, Kg, CRT, sett, etc.)
  - internal_order_ref: Int. ordre ref. per linje (kan avvike fra hode)
  - discount_percent: Rabatt % (null hvis ingen)
  - unit_price: Enhetspris (null hvis ikke oppgitt)
  - line_total: Beløp / Sum per linje (null hvis ikke oppgitt)
- **total_amount**: Samlet ordreverdi / Total (null hvis ikke oppgitt)
- **delivery_address**: Leveringsadresse med name, street, postal_code, city, country
- **special_instructions**: Viktige meldinger, fakturainstruksjoner, merknader
- **document_type**: "purchase_order" for vanlige bestillinger, "configuration_sheet" for Konfigurationsblatt/Configuration Sheet
- **confidence**: 0.0-1.0 – hvor sikker du er på ekstraksjonen
- **warnings**: Liste med advarsler (f.eks. "Ingen priser oppgitt", "Uvanlig format", "Flersidig ordre")

## Viktige regler

1. customer_name er ALDRI "Ortopartner AS" – det er leverandøren som mottar ordren.
2. Datoer skal alltid konverteres til YYYY-MM-DD format.
3. Beløp bruker norsk format (mellomrom som tusenskille, komma som desimaltegn): "1 927,00" = 1927.00
4. Hvis et felt mangler helt i PDF-en, sett det til null.
5. Ved tvil, sett lavere confidence og legg til en warning.
6. For Konfigurationsblatt: ekstraher Commission no som order_number og hva som er bestilt som line_items.
7. Noen bestillinger har "Rest" i stedet for "Antall" – bruk denne verdien som quantity.
8. Artikkelnummer i hakeparentes som [TA4610-6X400ML] – ekstraher uten hakeparentes.

Returner KUN gyldig JSON, ingen annen tekst.
"""

CONFIG_SHEET_PROMPT = """\
Du er en AI-agent som parser konfigurasjonsark (Configuration Sheet / Konfigurationsblatt)
for ortopediske hjelpemidler, sendt til Ortopartner AS.

Disse dokumentene er bildebaserte skjemaer med avkrysningsbokser, måleskjemaer og produktvalg.
De inneholder IKKE ordrelinjer med priser/artikkelnummer slik vanlige bestillinger gjør.

## Felter å ekstrahere

- **order_number**: Commission no / Patient ID / Referanse (f.eks. "JFNO2601952")
- **order_date**: Dato hvis oppgitt (ISO YYYY-MM-DD), ellers null
- **customer_name**: Klinikken/verkstedet som sender inn skjemaet (se etter logo, bunntekst, avsender). IKKE "Ortopartner" eller "Evomotion".
- **customer_reference**: Pasient-ID, kontaktperson, eller referansenavn
- **line_items**: Lag EN ordrelinje som beskriver produktet:
  - article_number: null (konfigurasjonsark har ikke artikkelnummer)
  - description: Fullt produktnavn med alle valgte opsjoner (f.eks. "Evomotion Solokit, høyre ben, dame, sort, kontinuerlig forsyning")
  - quantity: 1
  - unit: "stk"
- **delivery_address**: Leveringsadresse hvis oppgitt, ellers null
- **special_instructions**: VIKTIG — her skal ALL måledata og konfigurasjonsdetaljer inn:
  - Alle kroppsmål med verdier (midje, hofte, lår, kne, lengder, kroppshøyde osv.)
  - Valgte opsjoner (kontrollenhet, lommetype, glidelås, kabelutgang osv.)
  - Eventuelle håndskrevne notater eller merknader
  - Elektrodeposisjon-info
  Formater som lesbar tekst, f.eks.:
  "Produktvalg: Evomotion Solokit, høyre ben, dame, sort. Mål: Midje 70.5cm, Hofte 94.0cm, Lår 54.0cm (V+H), Kne 35.0cm (V: 34.3, H: 35.0), Knelengde 60.0cm, Setelengde 22.0cm, Kroppshøyde 178.0cm. Kontinuerlig forsyning."
- **document_type**: Alltid "configuration_sheet"
- **confidence**: 0.0-1.0 — sett høyere verdi (0.7-0.9) hvis du klarer å lese måledata og produktvalg tydelig fra bildet
- **warnings**: Advarsler (f.eks. "Håndskrift vanskelig å lese", "Mål mangler for venstre ben")

## Viktige regler

1. Bruk ALLTID bildet/bildene som primærkilde — tekst-ekstraksjonen er ofte ufullstendig for disse skjemaene.
2. Se nøye etter avkrysninger (X eller fylte sirkler) for produktvalg.
3. Mål står ofte i bokser til høyre for illustrasjoner — les tallene nøye.
4. Hvis venstre/høyre har ulike mål, noter begge.
5. customer_name er ALDRI "Ortopartner AS" eller "Evomotion" — det er leverandør/produsent.

Returner KUN gyldig JSON, ingen annen tekst.
"""


def parse_order_pdf(pdf_path: str | Path) -> ParsedOrder:
    """Parse a single order PDF using Claude API."""
    pdf_path = Path(pdf_path)
    info = get_pdf_info(pdf_path)

    # Extract text
    text = extract_text(pdf_path)

    if not text.strip():
        # Fallback to vision for image-based PDFs
        return _parse_with_vision(pdf_path, info)

    # Check if this might be a configuration sheet (image-heavy)
    is_config_sheet = "konfiguration" in pdf_path.name.lower() or "configuration" in text.lower()

    if is_config_sheet:
        return _parse_with_vision(pdf_path, info, is_config_sheet=True)

    return _parse_with_text(text, info)


def _parse_with_text(text: str, info: dict) -> ParsedOrder:
    """Parse using text extraction."""
    client = anthropic.Anthropic()

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": f"Parser denne bestillings-PDF-en:\n\nFilnavn: {info['filename']}\nAntall sider: {info['num_pages']}\n\n{text}",
            }
        ],
    )

    raw_json = response.content[0].text
    # Strip markdown code fences if present
    if raw_json.startswith("```"):
        raw_json = raw_json.split("\n", 1)[1]
        if raw_json.endswith("```"):
            raw_json = raw_json[:-3]

    data = json.loads(raw_json)
    data["source_file"] = info["filename"]
    return ParsedOrder(**data)


def _parse_with_vision(pdf_path: Path, info: dict, is_config_sheet: bool = False) -> ParsedOrder:
    """Parse using vision (for image-heavy PDFs like configuration sheets)."""
    client = anthropic.Anthropic()

    # Pick prompt based on document type
    system_prompt = CONFIG_SHEET_PROMPT if is_config_sheet else SYSTEM_PROMPT

    # Also get text as supplementary
    text = extract_text(pdf_path)
    images = pdf_to_base64_images(pdf_path, max_pages=3)

    intro = (
        "Parser dette konfigurasjonsarket (Configuration Sheet):"
        if is_config_sheet
        else "Parser denne bestillings-PDF-en:"
    )

    content: list[dict] = []
    content.append(
        {
            "type": "text",
            "text": f"{intro}\n\nFilnavn: {info['filename']}\nAntall sider: {info['num_pages']}\n\nEkstrahert tekst:\n{text[:3000]}",
        }
    )

    for i, img_b64 in enumerate(images):
        content.append(
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": img_b64,
                },
            }
        )
        content.append({"type": "text", "text": f"(Side {i + 1})"})

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        system=system_prompt,
        messages=[{"role": "user", "content": content}],
    )

    raw_json = response.content[0].text
    if raw_json.startswith("```"):
        raw_json = raw_json.split("\n", 1)[1]
        if raw_json.endswith("```"):
            raw_json = raw_json[:-3]

    data = json.loads(raw_json)
    data["source_file"] = info["filename"]
    return ParsedOrder(**data)
