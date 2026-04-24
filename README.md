# Indgangen – Hjælpemiddelansøgninger

Automatisering der behandler indkomne ansøgninger om hjælpemidler, forbrugsgoder og boligindretning fra Odense Kommunes postkasse.

## Hvad gør robotten?

1. **Henter e-mails** fra indbakken `hjaelpemidler@odense.dk` via Microsoft Graph API
2. **Filtrerer** på mails med emnet *"Ansøgning om hjælpemiddel, forbrugsgode eller boligindretning"* fra XFlow eller hjælpemidler-postkassen
3. **Udtrækker data** fra den vedhæftede PDF-ansøgning (CPR-nummer, telefonnummer, funktionsnedsættelse, ønskede hjælpemidler)
4. **Matcher hjælpemidler** mod et regelopslagssæt (`Regelsæt.xlsx`) for at finde relevante paragraffer (§112, §113 m.fl.)
5. **Opretter eller finder borgeren** i KMD Nexus
6. **Opretter forløb, skema og opgave** i Nexus baseret på de matchede paragraffer
7. **Tilknytter e-mailen** til det relevante forløb i Nexus
8. **Sletter den behandlede e-mail** fra indbakken

## Forudsætninger

- Python ≥ 3.13
- [`uv`](https://docs.astral.sh/uv/) til pakkehåndtering
- Adgang til **Automation Server** (arbejdskø)
- Adgang til **KMD Nexus** (produktion)
- En **RoboA**-konto med adgang til Microsoft Graph (e-mail)
- En **Odense SQL Server**-konto til tracking

## Installation

```sh
uv sync
```

## Konfiguration

Kopiér `.env.example` til `.env` og udfyld følgende:

| Variabel | Beskrivelse |
|---|---|
| `EXCEL_MAPPING_PATH` | Sti til `Regelsæt.xlsx` (kan også angives via `--excel-file`) |
| *(Automation Server-variabler)* | Ifølge `automation-server-client`-dokumentationen |

## Kørsel

```sh
# Fyld arbejdskøen med nye mails
uv run python main.py --queue

# Behandl arbejdskøen
uv run python main.py
```

### Argumenter

| Argument | Beskrivelse |
|---|---|
| `--excel-file <sti>` | Tilsidesæt stien til `Regelsæt.xlsx` |
| `--queue` | Fyld arbejdskøen og afslut (kør ingen behandling) |

## Regelopsætning (`Regelsæt.xlsx`)

Excel-filen indeholder to ark:

- **Placeringer** – Nøgleord per paragraf. Robotten søger fritekst fra ansøgningen efter disse nøgleord for at afgøre hvilke paragraffer der er relevante.
- **Forløb** – Mapping fra paragraf til forløbsnavn, skematype, opgavetype, tag og ansvarlig organisation i Nexus.

## Afhængigheder

| Pakke | Formål |
|---|---|
| `automation-server-client` | Arbejdskø-håndtering |
| `kmd-nexus-client` | Integration med KMD Nexus |
| `odk-tools` | Aktivitetssporing |
| `pymupdf` | PDF-parsing |
| `openpyxl` | Læsning af Excel-regelsæt |
| `msgraph-sdk` / `azure-identity` | E-mail via Microsoft Graph |

## Persondatasikkerhed

Robotten behandler følsomme personoplysninger på vegne af Odense Kommune, herunder CPR-numre og helbredsoplysninger (særlige kategorier jf. GDPR art. 9).

- Ingen personoplysninger må lægges i dette repository — hverken som testdata, i kode eller i kommentarer
- `input/`-mappen er ekskluderet via `.gitignore` og må aldrig committes
- Legitimationsoplysninger håndteres udelukkende via miljøvariabler (`.env`) og Automation Server Credentials

