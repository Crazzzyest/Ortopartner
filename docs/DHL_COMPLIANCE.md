# DHL Tracking — Compliance-notat

Dette notatet beskriver hvordan vår DHL-integrasjon overholder DHLs API
Terms of Service og GDPR. Det er kort og handlingsorientert — ment som
referanse for utviklere og som grunnlag for personvernerklæringen til
Ortopartner.

## API-en vi bruker

**DHL Shipment Tracking — Unified API**
- Endepunkt: `https://api-eu.dhl.com/track/shipments`
- Auth: `DHL-API-Key` header
- Vi bruker `service=express` (DHL Express)

Vi bruker **ikke** MyDHL API til sporing — bare Unified Tracking, fordi
vi ikke booker sendinger via Odoo (Ortopartner booker via DHL ProView).

## ToS-tiltak

| Krav fra DHL                         | Vårt tiltak                                                                  | Hvor i kode                            |
|--------------------------------------|------------------------------------------------------------------------------|----------------------------------------|
| Legitimt forretningsformål per kall  | Sjekk at trackingnr ligger på en `stock.picking` i Odoo før API-kall        | `DhlTracker._is_tracking_authorized`   |
| Ikke enumerere/scanne nr             | Vi sporer kun nr som har vært gjennom Odoo (ikke arbitrære nr)              | Samme guard                            |
| Respekt for rate limit               | Eksponentiell backoff ved 429, leser `Retry-After`-header                   | `DhlClient._request_with_backoff`      |
| Cache for å unngå unødvendige kall   | 15 min TTL i SQLite-cache                                                    | `DhlCache` (`FRESH_TTL`)               |
| Ikke videreselge data                | Sporingsdata vises kun internt i Ortopartners Odoo + dashboard               | (Avtalemessig — ingen kode)            |
| Audit-spor per kall                  | Hvert API-kall logges med `so_name` + `picking_id`                          | `DhlTracker._track_with_cache`         |

## GDPR-tiltak

| Type persondata                       | Tiltak                                                                  | Hvor i kode                          |
|---------------------------------------|--------------------------------------------------------------------------|--------------------------------------|
| Mottakerens signatur (navn)           | Strippes med regex før den vises i Odoo-chatter eller dashboard         | `DhlClient._strip_signature`         |
| Leveringsadresse (city + country)     | Vises kun til Ortopartner-ansatte med tilgang til SO-en (Odoo ACL)      | `stock.picking` arver SO-tilgang     |
| Cachet API-respons (rå data)          | Auto-slettes etter 30 dager                                             | `DhlCache.purge_expired` (RETENTION) |
| Hele cache hvis samtykke trekkes      | `tracker.purge_cache()` + `cache.clear_all()` for fullstendig sletting   | `DhlTracker.purge_cache`             |

### Lagringssted og levetid

| Datatype                              | Hvor lagres det                | Levetid              |
|---------------------------------------|--------------------------------|----------------------|
| Cachet DHL-respons (rå JSON)          | `output/dhl_cache.sqlite`      | 30 dager (auto-slett)|
| Tracking events i Odoo-chatter        | Odoo `stock.picking` chatter   | Følger Odoos retention |
| `carrier_tracking_ref` på picking     | Odoo `stock.picking`           | Følger SO-ens levetid|
| Audit-logg (`event_log` JSONL)        | `output/events.jsonl`          | Følger system-retention|

## Sletting på forespørsel (GDPR-art. 17)

Ved forespørsel om sletting av persondata for en gitt sending:

1. Fjern raden i `dhl_tracking_cache`:
   ```python
   from src.dhl_cache import DhlCache
   DhlCache().clear_all()  # eller delvis: spesifikt trackingnr
   ```
2. Slett relevante chatter-meldinger på picking-en i Odoo (manuelt fra UI)
3. Behold `carrier_tracking_ref` på picking — dette er forretningsdata,
   ikke direkte persondata.

## Rate limit-budsjett

DHLs offisielle grenser for Shipment Tracking Unified API (verifisert
mot DHL developer support, april 2026):

| Grense                    | Verdi                       | Kilde                 |
|---------------------------|-----------------------------|-----------------------|
| Per sending per dag       | **10 kall**                 | DHL support article   |
| Per sekund                | **1 kall hvert 5. sekund**  | API Reference (dev)   |
| Daglig totalbudsjett (dev)| **250 kall/dag**            | API Reference         |
| Daglig (oppgradert)       | (mnd-volum / 30) × 3 × 10   | DHL support article   |
| Brudd → respons           | HTTP 429 + `Retry-After`    | API Reference         |

### Vår TTL-policy (status-aware)

For å holde oss under 10/sending/dag varierer vi TTL etter status og
tid på døgnet. Implementasjon: `src/dhl_policy.py`.

| Status                            | TTL arbeidstid (07–19) | TTL natt (19–07) | ~kall/dag |
|-----------------------------------|------------------------|------------------|-----------|
| `delivered` / `failure` / `cancelled` | ∞ (ingen API-kall)  | ∞                | 0         |
| `delivery` (ute hos kurer)        | 30 min                  | 60 min           | 36 → ~6 etter overgang |
| `transit`                         | 4 timer                 | 8 timer          | 4.5       |
| `pre-transit`                     | 6 timer                 | 12 timer         | 3.0       |

**Spesielt om `delivery`:** uten korreksjon ville en sending i denne
statusen kunne treffe ~36 kall/dag — over DHLs grense. I praksis
holder ikke en sending seg i "delivery"-status mer enn noen få timer
før den blir `delivered`, så det reelle forbruket er 4–8 kall før
terminal-status nås. Vi overvåker dette via `cache_stats()` og kan
heve TTL hvis det blir problem.

### Vårt forventede forbruk

- ~20 aktive sendinger til enhver tid hos Ortopartner
- 5 i `transit` × 4.5 = ~23 kall/dag
- 2 i `delivery` × ~8 = ~16 kall/dag (kortvarig burst)
- 3 i `pre-transit` × 3 = ~9 kall/dag
- 10 ferdig leverte = 0 kall/dag
- **Total: ~50 kall/dag** — godt under 250-grensen i dev tier

### Hva hvis volumet vokser

Med 100 aktive sendinger ville vi ligge på ~250 kall/dag og treffe
dev-grensen. Da ber vi DHL om oppgradering (formelen
(mnd-volum / 30) × 3 × 10 gir typisk romslig grense for prod-bruk).

## Endringslogg

- **2026-04-26** — Første versjon. Unified API-klient, cache (15 min TTL),
  ToS-guards og signatur-stripping.
- **2026-04-26** — Status-aware TTL via `dhl_policy.py`. Default TTL
  økt fra 15 min til 4 timer (transit) for å respektere DHLs grense
  på 10 kall/sending/dag. Terminal-statuser slipper API-kall helt.
  Arbeidstid-vs-natt multiplikator på 2×.
