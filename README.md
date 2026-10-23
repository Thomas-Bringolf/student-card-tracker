# student-card-tracker

Automates fetching spendings from the SW-KA KartenService API
(`autoload.sw-ka.de`) into an Excel workbook: data tables, a statistics
dashboard with charts, and receipt PDFs.

## How it works

1. **[A] Read workbook** – finds the last saved day in the `Transactions` sheet
2. **[B] Login** – `ClientReg` handshake + `LOGIN` call, returns a session token
3. **[C] Fetch** – `TRANS` (purchase headers) + `TRANSPOS` (line items) for the
   missing date range (server keeps ~180 days of history)
4. **[D] Append + dedupe** – new rows are keyed by `transFullId` /
   (`transFullId`, `posId`) so re-runs never create duplicates; OCR-derived
   positions are replaced as soon as the API provides the real ones
5. **[E] Receipts** – downloads new Kassenbon PDFs (`KATRPDF`) into `receipts/`
6. **[F] Receipt text** – for transactions without API positions, the items are
   read from the receipt PDF text (`pypdfium2`, no image OCR needed) and added
   to `Positions` with `Quelle = ocr`
7. **[G] Save** – rebuilds all three sheets and writes the workbook atomically
8. **[H] Close** – ends the session (the API has no server-side logout)

## Setup

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp config.example.toml config.toml   # then fill in password
```

`config.toml` and `receipts/` are gitignored – your password and personal
receipts never end up in git.

## Usage

```bash
.venv/bin/python swka_spendings.py           # single run
.venv/bin/python swka_spendings.py --loop    # scrape every hour (interval_seconds)
```

Options:

| Flag | Meaning |
|---|---|
| `--since DD.MM.YYYY` | force a start date (dedupe keeps it safe) |
| `--auth-token TOKEN` | skip login, use an existing session token |
| `--ocr-report` | parse all receipts and compare items against the API data |
| `--config PATH` | alternate config file |
| `--loop` | run forever, scraping every `interval_seconds` (default 3600) |

All config values can be set via environment variables (override the TOML):
`SWKA_KARTE_NR`, `SWKA_PASSWORD`, `SWKA_EXCEL_PATH`, `SWKA_LOOKBACK_DAYS`,
`SWKA_RECEIPTS_FOLDER`, `SWKA_RECEIPTS_DOWNLOAD`, `SWKA_WEBDAV_URL`,
`SWKA_WEBDAV_USER`, `SWKA_WEBDAV_PASSWORD`, `SWKA_WEBDAV_PATH`,
`SWKA_INTERVAL_SECONDS`, `SWKA_HEALTH_FILE`.

## WebDAV upload

If `[webdav]` is configured (or the `SWKA_WEBDAV_*` variables are set), the
workbook is uploaded via WebDAV after every successful run – receipts stay
local. Missing folders are created with `MKCOL`. With `--loop`, a health file
(`SWKA_HEALTH_FILE`, default `.health`) is touched after each successful run –
the Docker `HEALTHCHECK` uses it to detect a stuck scraper.

## Docker

```bash
cp .env.example .env   # fill in credentials + WebDAV target
docker compose up -d --build
```

The container runs `swka_spendings.py --loop` (hourly scrape), keeps
`spendings.xlsx`, `receipts/` and the health file in the bind-mounted `./data/`
directory, and pushes the workbook to ownCloud via WebDAV. Logs go to stdout
(`docker compose logs -f`). The built-in `HEALTHCHECK` turns the container
unhealthy if no successful run happened for ~2.5 hours.

## Workbook layout

- **Übersicht** (regenerated on every run): key metrics (total spend, average
  purchase, current vs. previous month, …), monthly summary with share,
  top locations, top articles, weekday distribution, daily cumulative spend –
  each with a chart – plus a section of notes explaining the whole workflow.
- **Transactions**: `transFullId, Datum, Uhrzeit, Ort, Kasse, Typ, Betrag, Bon-ID`
  – title row, explanatory note, frozen header, filters, banded rows.
- **Positions**: `transFullId, posId, Bezeichnung, Menge, Einzelpreis, Gesamtpreis, Quelle`
  – line items per receipt, join key `transFullId`. `Quelle` is `api` for
  KartenService data, `ocr` for items read from the receipt PDF.

## Receipts & item text

PDFs are stored as `receipts/YYYY-MM-DD_transFullId.pdf` (only new ones are
downloaded). The Bon-ID column holds the server's `dateiablageId`.

The receipts contain a text layer (no image OCR required). If the API has no
positions for a transaction, the script parses the receipt text (including
wrapped names and weighed items like "Teller 545g") and adds the items to
`Positions` with `Quelle = ocr`. Totals are validated against the transaction
amount; failed parses are skipped. `--ocr-report` prints a comparison of all
receipts against the API positions.
