#!/usr/bin/env python3
"""Fetch SW-KA KartenService spending data and keep an Excel workbook up to date."""

from __future__ import annotations

import argparse
import base64
import json
import logging
import os
import re
import sys
import time
import tomllib
import zipfile
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote, unquote

import requests
import pypdfium2 as pdfium
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_URL = "https://autoload.sw-ka.de"

log = logging.getLogger("swka")
API_ROOT = f"{BASE_URL}/TL1/TLM/KASVC/"
CLIENT_REG_URL = f"{BASE_URL}/TL1/TLA/ClientReg"
BASIC_AUTH = "Basic S0FTVkM6ekt2NXlFMUxaVW12VzI5SQ=="
LOGIN_TOKEN = base64.b64decode(BASIC_AUTH.split()[1]).decode()
CLIENT_ID = 20
REG_KEY = "4G3hQj7E5ivIzbXNlt4Eu8rLvRb3Dy4g"
DEFAULT_HISTORY_DAYS = 180

TX_SHEET = "Transactions"
POS_SHEET = "Positions"
OV_SHEET = "Übersicht"
TX_HEADERS = ["transFullId", "Datum", "Uhrzeit", "Ort", "Kasse", "Typ", "Betrag", "Bon-ID"]
POS_HEADERS = ["transFullId", "posId", "Bezeichnung", "Menge", "Einzelpreis", "Gesamtpreis", "Quelle"]
TX_WIDTHS = [16, 12, 9, 30, 26, 12, 12, 12]
POS_WIDTHS = [16, 8, 42, 8, 14, 14, 10]

TX_TITLE = "Transaktionen"
TX_NOTES = (
    "Jede Zeile = ein Kartenumsatz. Schlüssel: transFullId – bereits vorhandene Zeilen "
    "werden beim Import übersprungen (keine Duplikate). Datum, Uhrzeit, Ort, Kasse und "
    "Typ stammen direkt vom SW-KA KartenService. Betrag: negative Werte sind Ausgaben, "
    "positive sind Aufladungen/Gutschriften. Bon-ID = Kennung des Kassenbons, die PDFs "
    "liegen automatisch im receipts/-Ordner. Das Skript hängt nur neue Zeilen an."
)
POS_TITLE = "Positionen"
POS_NOTES = (
    "Einzelne Posten eines Kassenbons, verknüpft mit dem Blatt Transactions über "
    "transFullId (posId = Position auf dem Bon). Basis für die Artikelstatistik auf der "
    "Übersicht. Dubletten (transFullId + posId) werden übersprungen. Quelle: 'api' = "
    "vom KartenService geliefert, 'ocr' = automatisch aus dem Bon-PDF ergänzt, wenn die "
    "API keine Positionen hat."
)

DARK = "1F4E78"
MID = "2E75B6"
BAND = "EAF1F8"
BORDER_CLR = "D9E2EC"
GREY = "595959"
LABEL_CLR = "44546A"

MONEY_FMT = '#,##0.00 "€"'
DATE_FMT = "DD.MM.YYYY"
TIME_FMT = "HH:MM"
INT_FMT = "#,##0"
QTY_FMT = "#,##0.##"
PCT_FMT = "0.0%"

FMT_BY_HEADER = {
    "Datum": DATE_FMT,
    "Uhrzeit": TIME_FMT,
    "Betrag": MONEY_FMT,
    "Einzelpreis": MONEY_FMT,
    "Gesamtpreis": MONEY_FMT,
    "Menge": QTY_FMT,
    "posId": "0",
    "Bon-ID": "0",
}

_HEADER_FILL = PatternFill("solid", fgColor=DARK)
_SECTION_FILL = PatternFill("solid", fgColor=MID)
_BAND_FILL = PatternFill("solid", fgColor=BAND)
_THIN = Side(style="thin", color=BORDER_CLR)
CELL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_TOTAL_BORDER = Border(left=_THIN, right=_THIN, top=Side(style="medium", color=DARK), bottom=_THIN)

FONT_NAME = "Calibri"


def _font(size: int = 10, *, bold: bool = False, italic: bool = False,
          color: str | None = None) -> Font:
    return Font(name=FONT_NAME, size=size, bold=bold, italic=italic, color=color)

_TS_RE = re.compile(
    r"^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?\s*$"
)
_MONTH_ABBR = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun",
               "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
_WEEKDAYS = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]


class ApiError(RuntimeError):
    pass


class WorkbookError(RuntimeError):
    pass


def parse_api_datetime(value: str) -> datetime:
    m = _TS_RE.match(value or "")
    if not m:
        raise ValueError(f"unrecognized date: {value!r}")
    d, mo, y, h, mi, s = (int(g) if g else 0 for g in m.groups())
    return datetime(y, mo, d, h, mi, s)


def parse_config_date(value: str) -> date:
    return datetime.strptime(value.strip(), "%d.%m.%Y").date()


def _parse_json(body: str):
    data = json.loads(body)
    while isinstance(data, str):
        data = json.loads(data)
    return data


def _find_auth_token(data):
    if isinstance(data, dict):
        token = data.get("authToken")
        return str(token) if token else None
    if isinstance(data, list):
        for item in data:
            token = _find_auth_token(item)
            if token:
                return token
    return None


def _to_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return parse_api_datetime(value).date()
        except ValueError:
            return None
    return None


class KartenService:
    def __init__(self, karte_nr: str, password: str, timeout: float = 30.0):
        self.karte_nr = karte_nr
        self.password = password
        self.timeout = timeout
        self.auth_token: str | None = None
        self.server_config: dict = {}
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:158.0) "
                "Gecko/20100101 Firefox/158.0",
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Accept-Language": "en-US,en;q=0.9",
                "Authorization": BASIC_AUTH,
                "X-Requested-With": "XMLHttpRequest",
                "Referer": f"{BASE_URL}/KartenService/index.html",
            }
        )

    @property
    def history_days(self) -> int:
        try:
            return int(self.server_config.get("lTransTage") or DEFAULT_HISTORY_DAYS)
        except (TypeError, ValueError):
            return DEFAULT_HISTORY_DAYS

    def connect(self) -> None:
        resp = self.session.post(
            CLIENT_REG_URL,
            params={"ClientID": CLIENT_ID, "RegKey": REG_KEY},
            timeout=self.timeout,
        )
        if not resp.ok:
            raise ApiError(f"ClientReg handshake failed (HTTP {resp.status_code})")

    def login(self) -> None:
        resp = self.session.post(
            API_ROOT + "LOGIN",
            params={"karteNr": self.karte_nr, "format": "JSON", "datenformat": "JSON"},
            json={"BenutzerID": self.karte_nr, "Passwort": self.password},
            headers={"Origin": BASE_URL},
            timeout=self.timeout,
        )
        body = (resp.text or "").strip()
        if not resp.ok:
            raise ApiError(f"login failed (HTTP {resp.status_code}): {body[:300]}")
        try:
            data = _parse_json(body)
        except (ValueError, json.JSONDecodeError) as exc:
            raise ApiError(f"unexpected login response: {body[:300]}") from exc
        token = _find_auth_token(data)
        if not token:
            hint = " (card blocked or internet access not enabled?)" if "inactiv" in body.lower() else ""
            raise ApiError(f"login rejected: {body[:300]}{hint}")
        self.auth_token = token
        if isinstance(data, dict):
            self.server_config = data
        elif isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    self.server_config = item
                    break

    def use_token(self, token: str) -> None:
        self.auth_token = unquote(token)

    def fetch(self, date_from: date, date_to: date) -> tuple[list[dict], list[dict]]:
        if not self.auth_token:
            raise ApiError("not logged in")
        params = {
            "format": "JSON",
            "authToken": self.auth_token,
            "karteNr": self.karte_nr,
            "datumVon": date_from.strftime("%d.%m.%Y"),
            "datumBis": (date_to + timedelta(days=1)).strftime("%d.%m.%Y"),
        }
        positions = self._get("TRANSPOS", params)
        transactions = self._get("TRANS", params)
        return transactions, positions

    def _get(self, endpoint: str, params: dict) -> list[dict]:
        resp = self.session.get(API_ROOT + endpoint, params=params, timeout=self.timeout)
        body = (resp.text or "").strip()
        if not resp.ok:
            raise ApiError(f"{endpoint} failed (HTTP {resp.status_code}): {body[:300]}")
        try:
            data = _parse_json(body)
        except (ValueError, json.JSONDecodeError) as exc:
            raise ApiError(f"unexpected {endpoint} response: {body[:300]}") from exc
        if not isinstance(data, list):
            raise ApiError(f"unexpected {endpoint} response: {body[:300]}")
        return data

    def close(self) -> None:
        self.auth_token = None
        self.session.close()


def tx_to_row(t: dict) -> dict:
    ts = parse_api_datetime(t.get("datum") or "")
    return {
        "transFullId": t.get("transFullId"),
        "Datum": ts.date(),
        "Uhrzeit": ts.time().replace(second=0),
        "Ort": t.get("ortName"),
        "Kasse": t.get("kaName"),
        "Typ": t.get("typName"),
        "Betrag": t.get("zahlBetrag"),
        "Bon-ID": t.get("dateiablageId"),
    }


def pos_to_row(p: dict) -> dict:
    return {
        "transFullId": p.get("transFullId"),
        "posId": p.get("posId"),
        "Bezeichnung": p.get("name"),
        "Menge": p.get("menge"),
        "Einzelpreis": p.get("epreis"),
        "Gesamtpreis": p.get("gpreis"),
        "Quelle": "api",
    }


def _read_sheet_dicts(wb, title: str, headers: list[str]) -> list[dict]:
    if title not in wb.sheetnames:
        return []
    ws = wb[title]
    header_row = None
    for i, (value,) in enumerate(
        ws.iter_rows(min_row=1, max_row=12, min_col=1, max_col=1, values_only=True),
        start=1,
    ):
        if value == headers[0]:
            header_row = i
            break
    if header_row is None:
        if ws.max_row <= 1:
            return []
        raise WorkbookError(
            f"sheet '{title}': header row not found – unexpected layout, "
            "move the file aside to rebuild it"
        )
    rows = []
    for row in ws.iter_rows(
        min_row=header_row + 1, max_row=ws.max_row, max_col=len(headers), values_only=True
    ):
        if row[0] in (None, ""):
            continue
        rows.append(dict(zip(headers, row)))
    return rows


def load_workbook_data(path: Path) -> tuple[Workbook, list[dict], list[dict]]:
    if not path.exists():
        wb = Workbook()
        wb.remove(wb.active)
        return wb, [], []
    wb = load_workbook(path)
    tx_rows = _read_sheet_dicts(wb, TX_SHEET, TX_HEADERS)
    pos_rows = _read_sheet_dicts(wb, POS_SHEET, POS_HEADERS)
    for r in pos_rows:
        if not r.get("Quelle"):
            r["Quelle"] = "api"
    return wb, tx_rows, pos_rows


def _title_bar(ws, row: int, ncols: int, text: str, font: Font, fill,
               height: float, alignment: Alignment) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = font
        if fill is not None:
            cell.fill = fill
    ws.cell(row, 1).value = text
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    ws.row_dimensions[row].height = height
    ws.cell(row, 1).alignment = alignment


def write_data_sheet(wb, index: int, title: str, headers: list[str], widths: list[int],
                     title_text: str, notes: str, rows: list[dict]) -> None:
    ws = wb.create_sheet(title, index)
    ws.sheet_properties.tabColor = MID
    ws.sheet_view.showGridLines = False
    ncols = len(headers)

    _title_bar(
        ws, 1, ncols, title_text,
        _font(size=14, bold=True, color="FFFFFF"), _HEADER_FILL, 26,
        Alignment(vertical="center", indent=1),
    )
    lines = max(1, -(-len(notes) // max(20, sum(widths))))
    _title_bar(
        ws, 2, ncols, notes,
        _font(size=9, italic=True, color=GREY), None, 12 * lines + 6,
        Alignment(wrap_text=True, vertical="center", indent=1),
    )

    for c, text in enumerate(headers, start=1):
        cell = ws.cell(3, c, text)
        cell.font = _font(bold=True, color="FFFFFF")
        cell.fill = _SECTION_FILL
        cell.border = CELL_BORDER
        cell.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[3].height = 20

    for i, row in enumerate(rows):
        r = 4 + i
        band = i % 2 == 1
        for c, key in enumerate(headers, start=1):
            cell = ws.cell(r, c, row.get(key))
            cell.font = _font()
            cell.border = CELL_BORDER
            if band:
                cell.fill = _BAND_FILL
            fmt = FMT_BY_HEADER.get(key)
            if fmt:
                cell.number_format = fmt

    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A4"
    if rows:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{3 + len(rows)}"


def save_workbook(wb, path: Path) -> None:
    tmp = path.with_name(path.name + ".tmp")
    wb.save(tmp)
    tmp.replace(path)


def _section(ws, row: int, text: str) -> int:
    _title_bar(
        ws, row, 5, text, _font(size=11, bold=True, color="FFFFFF"), _SECTION_FILL, 20,
        Alignment(vertical="center", indent=1),
    )
    return row + 1


def _table_header(ws, row: int, labels: list[str]) -> None:
    for c, text in enumerate(labels, start=1):
        cell = ws.cell(row, c, text)
        cell.font = _font(bold=True, color="FFFFFF")
        cell.fill = _SECTION_FILL
        cell.border = CELL_BORDER
        cell.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 18


def _table_row(ws, row: int, values: list, fmts: list[str | None], band: bool = False,
               bold: bool = False, top_border: bool = False) -> None:
    for c, value in enumerate(values, start=1):
        cell = ws.cell(row, c, value)
        cell.font = _font(bold=bold)
        cell.border = _TOTAL_BORDER if top_border else CELL_BORDER
        if band:
            cell.fill = _BAND_FILL
        fmt = fmts[c - 1]
        if fmt:
            cell.number_format = fmt


def _month_range(d0: date, d1: date):
    y, m = d0.year, d0.month
    while (y, m) <= (d1.year, d1.month):
        yield y, m
        m += 1
        if m == 13:
            y, m = y + 1, 1


def _collect(tx_rows: list[dict]) -> list[dict]:
    records = []
    for row in tx_rows:
        day = _to_date(row.get("Datum"))
        if not day:
            continue
        try:
            amt = float(row.get("Betrag") or 0)
        except (TypeError, ValueError):
            continue
        records.append({
            "day": day,
            "amt": amt,
            "ort": str(row.get("Ort") or "Unbekannt"),
            "tfid": str(row.get("transFullId") or ""),
        })
    return records


def _bar_chart(ws, anchor: str, title: str, hdr_row: int, first_row: int,
               last_row: int, val_col: int, cat_col: int = 1,
               horizontal: bool = False, show_values: bool = False) -> None:
    chart = BarChart()
    chart.type = "bar" if horizontal else "col"
    chart.style = 10
    chart.title = title
    chart.add_data(
        Reference(ws, min_col=val_col, min_row=hdr_row, max_row=last_row),
        titles_from_data=True,
    )
    chart.set_categories(Reference(ws, min_col=cat_col, min_row=first_row, max_row=last_row))
    chart.series[0].graphicalProperties = GraphicalProperties(solidFill=MID)
    chart.legend = None
    chart.height = 7
    chart.width = 15
    chart.gapWidth = 60
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.title = "Euro"
    if horizontal:
        chart.x_axis.tickLblSkip = 1
    if show_values:
        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True
        chart.dLbls.showCatName = False
        chart.dLbls.showSerName = False
        chart.dLbls.showLegendKey = False
        chart.dLbls.showPercent = False
        chart.dLbls.showBubbleSize = False
        chart.dLbls.numFmt = '#,##0.00 "€"'
    ws.add_chart(chart, anchor)


def _overview_notes(folder_name: str, stats: dict | None) -> list[str]:
    if stats is None:
        receipts_line = (
            f"Kassenbons: Download in diesem Lauf deaktiviert (Konfiguration, Ziel {folder_name}/)."
        )
    elif "error" in stats:
        receipts_line = f"Kassenbons: Download fehlgeschlagen ({stats['error']})."
    else:
        receipts_line = (
            f"Kassenbons: {stats['new']} in diesem Lauf neu geladen, {stats['have']} bereits "
            f"vorhanden, {stats['failed']} Fehler, {stats['none']} ohne PDF-Verweis "
            f"→ Ordner {folder_name}/ (Dateiname JJJJ-MM-TT_transFullId.pdf)."
        )
    return [
        "Übersicht: Dieses Blatt wird bei jedem Lauf automatisch komplett neu erzeugt – "
        "eigene Änderungen hier gehen verloren.",
        "Transaktionen/Positionen: Rohdaten-Tabellen. Neue Zeilen werden angehängt, "
        "vorhandene Datensätze (transFullId bzw. transFullId+posId) übersprungen.",
        "Ablauf eines Laufs: [A] letzten Tag im Buch lesen → [B] KartenService-Login → "
        "[C] Daten ab diesem Tag holen → [D] dedupliziert anhängen → [E] Kassenbons laden → "
        "[F] fehlende Positionen aus dem Bon-Text ergänzen → [G] Datei speichern → "
        "[H] Sitzung beenden.",
        receipts_line,
        "Bon-Auslesung: Fehlen die Positionen einer Transaktion in der API, werden die "
        "Waren automatisch aus dem Text des Bon-PDFs übernommen (Spalte Quelle='ocr'). "
        "Mit --ocr-report lässt sich die Auslesung gegen die API-Daten prüfen.",
        "Datenquelle: autoload.sw-ka.de (SW-KA KartenService); der Server hält rund 180 Tage "
        "Transaktionen – ältere Daten existieren nur noch in dieser Datei.",
        "Beträge: In den Auswertungen werden Ausgaben positiv gezählt (intern sind sie "
        "negativ), Aufladungen/Gutschriften erscheinen separat.",
        "Aktualisierung: Skript erneut ausführen (optional --since DD.MM.YYYY für ältere "
        "Zeiträume). Zugangsdaten liegen lokal in config.toml.",
    ]


def build_overview(wb, tx_rows: list[dict], pos_rows: list[dict],
                   receipt_stats: dict | None, karte_nr: str, receipts_folder: str) -> None:
    if OV_SHEET in wb.sheetnames:
        del wb[OV_SHEET]
    ws = wb.create_sheet(OV_SHEET, 0)
    ws.sheet_properties.tabColor = DARK
    ws.sheet_view.showGridLines = False
    for col, width in {"A": 32, "B": 16, "C": 16, "D": 15, "E": 15, "F": 3, "G": 3}.items():
        ws.column_dimensions[col].width = width

    _title_bar(
        ws, 1, 5, "Finanzübersicht – Studierendenwerks-Karte",
        _font(size=16, bold=True, color="FFFFFF"), _HEADER_FILL, 30,
        Alignment(vertical="center", indent=1),
    )

    records = _collect(tx_rows)
    if not records:
        ws.cell(3, 1, "Keine Transaktionen vorhanden – erst Daten importieren.")
        return

    days = [r["day"] for r in records]
    min_d, max_d = min(days), max(days)
    _title_bar(
        ws, 2, 5,
        f"KartenNr. {karte_nr} · Zeitraum {min_d:%d.%m.%Y} – {max_d:%d.%m.%Y} · "
        f"Datenstand {datetime.now():%d.%m.%Y %H:%M}",
        _font(size=10, italic=True, color=GREY), None, 18,
        Alignment(vertical="center", indent=1),
    )

    purchases = [r for r in records if r["amt"] < 0]
    spend = -sum(r["amt"] for r in purchases)
    credits = sum(r["amt"] for r in records if r["amt"] > 0)
    n = len(purchases)
    months = list(_month_range(min_d, max_d))
    avg = spend / n if n else 0.0
    avg_month = spend / len(months) if months else 0.0
    biggest = max((-r["amt"] for r in purchases), default=0.0)
    last_day = max(days)
    ort_counts = Counter(r["ort"] for r in purchases)
    top_ort, top_ort_n = ort_counts.most_common(1)[0] if ort_counts else ("–", 0)
    today = date.today()
    cur_month_spend = -sum(
        r["amt"] for r in purchases
        if (r["day"].year, r["day"].month) == (today.year, today.month)
    )
    prev_anchor = today.replace(day=1) - timedelta(days=1)
    prev_month_spend = -sum(
        r["amt"] for r in purchases
        if (r["day"].year, r["day"].month) == (prev_anchor.year, prev_anchor.month)
    )

    row = _section(ws, 4, "Kennzahlen")
    metrics = [
        ("Gesamtausgaben", spend, MONEY_FMT),
        ("Einkäufe", n, INT_FMT),
        ("Ø pro Einkauf", avg, MONEY_FMT),
        ("Ø pro Monat", avg_month, MONEY_FMT),
        ("Größter Einkauf", biggest, MONEY_FMT),
        ("Aufladungen / Gutschriften", credits, MONEY_FMT),
        ("Laufender Monat", cur_month_spend, MONEY_FMT),
        ("Vormonat", prev_month_spend, MONEY_FMT),
        ("Letzter Einkauf", last_day, DATE_FMT),
        ("Häufigster Ort", f"{top_ort} ({top_ort_n}×)", None),
    ]
    for label, value, fmt in metrics:
        label_cell = ws.cell(row, 1, label)
        label_cell.font = _font(color=LABEL_CLR)
        value_cell = ws.cell(row, 2, value)
        value_cell.font = _font(bold=True, size=11)
        if fmt:
            value_cell.number_format = fmt
        row += 1
    row += 1

    row = _section(ws, row, "Pro Monat")
    monthly_hdr = row
    _table_header(ws, row, ["Monat", "Ausgaben", "Einkäufe", "Ø Einkauf", "Anteil"])
    row += 1
    monthly_first = row
    for i, (y, m) in enumerate(months):
        m_spend = -sum(
            r["amt"] for r in purchases if (r["day"].year, r["day"].month) == (y, m)
        )
        m_cnt = sum(1 for r in purchases if (r["day"].year, r["day"].month) == (y, m))
        label = f"{_MONTH_ABBR[m - 1]} {y}"
        if (y, m) == (today.year, today.month):
            label += " *"
        _table_row(
            ws, row,
            [label, m_spend, m_cnt, (m_spend / m_cnt if m_cnt else 0),
             (m_spend / spend if spend else 0)],
            [None, MONEY_FMT, INT_FMT, MONEY_FMT, PCT_FMT],
            band=i % 2 == 1,
        )
        row += 1
    monthly_last = row - 1
    _table_row(
        ws, row, ["Gesamt", spend, n, avg, 1.0],
        [None, MONEY_FMT, INT_FMT, MONEY_FMT, PCT_FMT],
        bold=True, top_border=True,
    )
    row += 1
    ws.cell(row, 1, "* laufender Monat").font = _font(size=9, italic=True, color="808080")
    row += 2

    ort_spend: Counter = Counter()
    ort_cnt: Counter = Counter()
    for r in purchases:
        ort_spend[r["ort"]] += -r["amt"]
        ort_cnt[r["ort"]] += 1
    top_orten = sorted(ort_spend.items(), key=lambda kv: -kv[1])[:10]

    row = _section(ws, row, "Top-Orte nach Ausgaben")
    ort_hdr = row
    _table_header(ws, row, ["Ort", "Ausgaben", "Einkäufe"])
    row += 1
    ort_first = row
    for i, (name, value) in enumerate(top_orten):
        _table_row(
            ws, row, [name, value, ort_cnt[name]],
            [None, MONEY_FMT, INT_FMT], band=i % 2 == 1,
        )
        row += 1
    ort_last = row - 1
    row += 1

    art_spend: Counter = Counter()
    art_qty: Counter = Counter()
    for p in pos_rows:
        name = p.get("Bezeichnung")
        if not name:
            continue
        try:
            gross = float(p.get("Gesamtpreis") or 0)
        except (TypeError, ValueError):
            continue
        if gross <= 0:
            continue
        key = str(name)
        art_spend[key] += gross
        try:
            art_qty[key] += float(p.get("Menge") or 0)
        except (TypeError, ValueError):
            pass
    top_art = sorted(art_spend.items(), key=lambda kv: -kv[1])[:10]

    row = _section(ws, row, "Top-Artikel nach Ausgaben")
    art_hdr = row
    _table_header(ws, row, ["Artikel", "Ausgaben", "Menge gesamt"])
    row += 1
    art_first = row
    for i, (name, value) in enumerate(top_art):
        _table_row(
            ws, row, [name, value, art_qty[name]],
            [None, MONEY_FMT, QTY_FMT], band=i % 2 == 1,
        )
        row += 1
    art_last = row - 1
    row += 1

    wd_spend = [0.0] * 7
    wd_cnt = [0] * 7
    for r in purchases:
        idx = r["day"].weekday()
        wd_spend[idx] += -r["amt"]
        wd_cnt[idx] += 1

    row = _section(ws, row, "Wochentage")
    wd_hdr = row
    _table_header(ws, row, ["Wochentag", "Ausgaben", "Einkäufe"])
    row += 1
    wd_first = row
    for i in range(7):
        _table_row(
            ws, row, [_WEEKDAYS[i], wd_spend[i], wd_cnt[i]],
            [None, MONEY_FMT, INT_FMT], band=i % 2 == 1,
        )
        row += 1
    wd_last = row - 1
    row += 1

    daily: dict = {}
    for r in purchases:
        daily[r["day"]] = daily.get(r["day"], 0.0) + -r["amt"]
    all_days = [min_d + timedelta(days=i) for i in range((max_d - min_d).days + 1)]

    row = _section(ws, row, "Kumuliert (täglich)")
    cum_hdr = row
    _table_header(ws, row, ["Datum", "Ausgaben/Tag", "Kumuliert"])
    row += 1
    cum_first = row
    running = 0.0
    for i, day in enumerate(all_days):
        day_spend = daily.get(day, 0.0)
        running += day_spend
        _table_row(
            ws, row, [day, day_spend, running],
            [DATE_FMT, MONEY_FMT, MONEY_FMT], band=i % 2 == 1,
        )
        row += 1
    cum_last = row - 1
    row += 1

    row = _section(ws, row, "Hinweise & Erklärungen")
    notes_width = 32 + 16 + 16 + 15 + 15
    for bullet in _overview_notes(receipts_folder, receipt_stats):
        text = "•  " + bullet
        for c in range(1, 6):
            cell = ws.cell(row, c)
            cell.font = _font(size=9, color=LABEL_CLR)
        cell = ws.cell(row, 1, text)
        cell.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        est_lines = max(1, -(-len(text) // (notes_width - 6)))
        ws.row_dimensions[row].height = est_lines * 12 + 4
        row += 1

    if len(all_days) > 1:
        line = LineChart()
        line.title = "Kumulierte Ausgaben"
        line.style = 12
        line.add_data(
            Reference(ws, min_col=3, min_row=cum_hdr, max_row=cum_last),
            titles_from_data=True,
        )
        line.set_categories(Reference(ws, min_col=1, min_row=cum_first, max_row=cum_last))
        series = line.series[0]
        series.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill="C00000", w=24000)
        )
        series.smooth = False
        line.height = 7
        line.width = 15
        line.x_axis.delete = False
        line.y_axis.delete = False
        line.y_axis.title = "Euro"
        ws.add_chart(line, "H4")

    if monthly_last >= monthly_first:
        _bar_chart(ws, "H18", "Ausgaben pro Monat", monthly_hdr, monthly_first,
                   monthly_last, val_col=2, show_values=True)
    if top_orten and ort_last >= ort_first:
        pie = PieChart()
        pie.title = "Anteile nach Ort (Top 10)"
        pie.style = 10
        pie.add_data(
            Reference(ws, min_col=2, min_row=ort_hdr, max_row=ort_last),
            titles_from_data=True,
        )
        pie.set_categories(Reference(ws, min_col=1, min_row=ort_first, max_row=ort_last))
        pie.dataLabels = DataLabelList(
            showPercent=True, showCatName=False, showVal=False,
            showSerName=False, showLegendKey=False,
        )
        pie.legend.position = "r"
        pie.height = 7
        pie.width = 15
        ws.add_chart(pie, "H32")
    if top_art and art_last >= art_first:
        _bar_chart(ws, "H46", "Top-Artikel nach Ausgaben", art_hdr, art_first,
                   art_last, val_col=2, horizontal=True)
    if sum(wd_spend) > 0:
        _bar_chart(ws, "H60", "Ausgaben nach Wochentag", wd_hdr, wd_first,
                   wd_last, val_col=2)


def download_receipts(client: KartenService, tx_rows: list[dict], folder: Path) -> dict:
    stats = {"new": 0, "have": 0, "failed": 0, "none": 0}
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return {"error": str(exc)}
    for row in tx_rows:
        try:
            bon_id = int(row.get("Bon-ID"))
        except (TypeError, ValueError):
            stats["none"] += 1
            continue
        path = _receipt_path(folder, row)
        if path.exists():
            stats["have"] += 1
            continue
        try:
            resp = client.session.get(
                API_ROOT + "KATRPDF",
                params={
                    "authToken": client.auth_token,
                    "karteNr": client.karte_nr,
                    "loginToken": LOGIN_TOKEN,
                    "dateiablageId": bon_id,
                },
                timeout=client.timeout,
            )
        except requests.RequestException:
            stats["failed"] += 1
            continue
        if resp.ok and resp.content[:4] == b"%PDF":
            tmp = path.with_name(path.name + ".tmp")
            tmp.write_bytes(resp.content)
            tmp.replace(path)
            stats["new"] += 1
            if stats["new"] % 25 == 0:
                log.info(f"  ... {stats['new']} Bons geladen")
        else:
            stats["failed"] += 1
        time.sleep(0.05)
    return stats


def _receipt_path(folder: Path, row: dict) -> Path:
    day = _to_date(row.get("Datum"))
    tfid = re.sub(r"[^\w.-]", "_", str(row.get("transFullId") or "unknown"))
    prefix = f"{day:%Y-%m-%d}_" if day else ""
    return folder / f"{prefix}{tfid}.pdf"


def extract_receipt_text(path: Path) -> str:
    pdf = pdfium.PdfDocument(str(path))
    try:
        return "\n".join(
            pdf[i].get_textpage().get_text_range() for i in range(len(pdf))
        )
    finally:
        pdf.close()


_EUR_RE = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$|^\d+,\d{2}$")
_ITEM_LINE_RE = re.compile(r"^(?P<name>.*?)\s*\(o\)\s*(?P<price>\d+,\d{2})?\s*$")
_SUMME_RE = re.compile(r"^Summe\s+EUR\s+(\d+,\d{2})\s*$")
_TARIFF_RE = re.compile(r"^[^\d()\s]+\s+\(\d+\)$")
_DATUM_RE = re.compile(
    r"^(?:Datum:\s*)?\d{2}[./]\d{2}[./]\d{4}\s+\d{2}:\d{2}(?::\d{2})?\s*$"
)
_WEIGHT_RE = re.compile(r"^\d+,\d{3}$")


def _parse_eur(value: str) -> float:
    if "," in value:
        return float(value.replace(".", "").replace(",", "."))
    return float(value)


def parse_receipt(text: str) -> list[dict] | None:
    lines = [l.strip() for l in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    lines = [l for l in lines if l]
    summe_idx = next((i for i, l in enumerate(lines) if l.startswith("Summe")), None)
    if summe_idx is None:
        return None
    m = _SUMME_RE.match(lines[summe_idx])
    if not m:
        return None
    total = _parse_eur(m.group(1))
    datum_idx = next(
        (i for i, l in enumerate(lines[:summe_idx]) if _DATUM_RE.match(l)), None
    )
    if datum_idx is None:
        return None
    region = lines[datum_idx + 1:summe_idx]
    while region and _TARIFF_RE.match(region[0]):
        region = region[1:]

    items: list[dict] = []
    buf: list[str] = []
    pending = False
    i = 0
    while i < len(region):
        line = region[i]
        if pending:
            pm = _EUR_RE.match(line)
            if pm:
                items.append({"name": " ".join(buf), "price": _parse_eur(line)})
                buf = []
                pending = False
                i += 1
                continue
            items.append({"name": " ".join(buf), "price": None})
            buf = []
            pending = False
        if line == "kg €/kg" and buf:
            weight_idx = next(
                (j for j, part in enumerate(buf) if _WEIGHT_RE.match(part)), None
            )
            name = " ".join(buf[:weight_idx]) if weight_idx is not None else " ".join(buf)
            try:
                j = region.index("Preis €", i)
            except ValueError:
                return None
            if j + 1 >= len(region) or not _EUR_RE.match(region[j + 1]):
                return None
            items.append({"name": name, "price": _parse_eur(region[j + 1])})
            buf = []
            i = j + 2
            if i < len(region) and region[i] == "+/- €":
                i += 1
                if i < len(region) and _EUR_RE.match(region[i]):
                    i += 1
            continue
        om = _ITEM_LINE_RE.match(line)
        if om:
            name = om.group("name").strip()
            if buf:
                name = (" ".join(buf) + " " + name).strip()
                buf = []
            price = _parse_eur(om.group("price")) if om.group("price") else None
            if price is not None:
                items.append({"name": name, "price": price})
            else:
                buf = [name]
                pending = True
            i += 1
            continue
        if _EUR_RE.match(line):
            return None
        buf.append(line)
        i += 1
    if pending or buf:
        items.append({"name": " ".join(buf), "price": None})
    if not items:
        return None
    if any(not it["name"] or it["price"] is None for it in items):
        return None
    if abs(sum(it["price"] for it in items) - total) > 0.011:
        return None
    return items


def fill_positions_from_receipts(all_tx: list[dict], all_pos: list[dict],
                                 folder: Path) -> tuple[list[dict], dict]:
    stats = {"parsed": 0, "items": 0, "filled": 0, "failed": 0}
    with_pos = {str(r.get("transFullId")) for r in all_pos}
    added = []
    for row in all_tx:
        tfid = str(row.get("transFullId"))
        if tfid in with_pos:
            continue
        pdf = _receipt_path(folder, row)
        if not pdf.exists():
            stats["failed"] += 1
            continue
        try:
            items = parse_receipt(extract_receipt_text(pdf))
        except Exception:
            items = None
        try:
            betrag = abs(float(row.get("Betrag") or 0))
        except (TypeError, ValueError):
            betrag = None
        if items is None or (
            betrag is not None
            and abs(sum(i["price"] for i in items) - betrag) > 0.011
        ):
            stats["failed"] += 1
            continue
        stats["parsed"] += 1
        stats["filled"] += 1
        stats["items"] += len(items)
        for i, item in enumerate(items, start=1):
            added.append({
                "transFullId": row.get("transFullId"),
                "posId": i,
                "Bezeichnung": item["name"],
                "Menge": 1,
                "Einzelpreis": item["price"],
                "Gesamtpreis": item["price"],
                "Quelle": "ocr",
            })
    all_pos.extend(added)
    return all_pos, stats


def run_ocr_report(all_tx: list[dict], all_pos: list[dict], folder: Path) -> None:
    api_by_tfid: dict = {}
    for r in all_pos:
        if str(r.get("Quelle") or "api").lower() != "ocr":
            api_by_tfid.setdefault(str(r.get("transFullId")), []).append(r)

    parsed = total_ok = names_ok = prices_ok = 0
    problems: list[str] = []
    for row in all_tx:
        tfid = str(row.get("transFullId"))
        pdf = _receipt_path(folder, row)
        if not pdf.exists():
            problems.append(f"{tfid}: no PDF")
            continue
        try:
            items = parse_receipt(extract_receipt_text(pdf))
        except Exception as exc:
            problems.append(f"{tfid}: extract error {exc}")
            continue
        if items is None:
            problems.append(f"{tfid}: parse failed")
            continue
        parsed += 1
        try:
            betrag = abs(float(row.get("Betrag") or 0))
        except (TypeError, ValueError):
            betrag = None
        if betrag is not None and abs(sum(i["price"] for i in items) - betrag) <= 0.011:
            total_ok += 1
        else:
            problems.append(
                f"{tfid}: total {sum(i['price'] for i in items):.2f} != Betrag {betrag}"
            )
        api = sorted(
            api_by_tfid.get(tfid, []), key=lambda r: int(r.get("posId") or 0)
        )
        if not api:
            continue
        ocr_names = [i["name"] for i in items]
        api_names = [str(r.get("Bezeichnung") or "") for r in api]
        ocr_prices = [round(i["price"], 2) for i in items]
        api_prices = [
            round(float(r.get("Gesamtpreis") or 0), 2) for r in api
        ]
        if ocr_names == api_names:
            names_ok += 1
        else:
            problems.append(f"{tfid}: names {ocr_names} != API {api_names}")
        if ocr_prices == api_prices:
            prices_ok += 1
        else:
            problems.append(f"{tfid}: prices {ocr_prices} != API {api_prices}")

    n_tx = len(all_tx)
    log.info(f"[OCR-Report] parsed {parsed}/{n_tx} receipts, "
          f"total vs Betrag ok: {total_ok}, names identical to API: {names_ok}, "
          f"prices identical: {prices_ok}")
    for line in problems[:10]:
        log.info(f"  ! {line}")
    if len(problems) > 10:
        log.info(f"  ... {len(problems) - 10} more problems")


def load_config(path: Path, require_password: bool = True) -> dict:
    cfg: dict = {}
    if path.exists():
        with path.open("rb") as fh:
            cfg = tomllib.load(fh)
    api = cfg.get("api", {})
    excel = cfg.get("excel", {})
    receipts = cfg.get("receipts", {})
    webdav = cfg.get("webdav", {})

    def env_or(key: str, fallback) -> str:
        return os.environ.get(key) or str(fallback or "")

    karte_nr = env_or("SWKA_KARTE_NR", api.get("karte_nr"))
    password = env_or("SWKA_PASSWORD", api.get("password"))
    if not karte_nr.isdigit():
        sys.exit("config: karte_nr must be the numeric card number "
                 "(config.toml [api] or SWKA_KARTE_NR)")
    if require_password and not password:
        sys.exit("config: password is missing "
                 "(config.toml [api] or SWKA_PASSWORD)")

    def env_int(key: str, fallback, minimum: int = 0) -> int:
        raw = os.environ.get(key) or (str(fallback) if fallback is not None else "")
        try:
            value = int(raw)
        except (TypeError, ValueError):
            value = 0
        return max(minimum, value) if raw else fallback

    def env_bool(key: str, fallback: bool) -> bool:
        raw = os.environ.get(key)
        if raw is None:
            return fallback
        return raw.strip().lower() not in ("0", "false", "no", "off", "")

    return {
        "karte_nr": karte_nr,
        "password": password,
        "path": Path(env_or("SWKA_EXCEL_PATH", excel.get("path"))
                     or "spendings.xlsx").expanduser(),
        "lookback_days": env_int("SWKA_LOOKBACK_DAYS", excel.get("lookback_days") or 60),
        "receipts_folder": Path(env_or("SWKA_RECEIPTS_FOLDER", receipts.get("folder"))
                                or "receipts").expanduser(),
        "receipts_download": env_bool("SWKA_RECEIPTS_DOWNLOAD",
                                      bool(receipts.get("download", True))),
        "webdav_url": env_or("SWKA_WEBDAV_URL", webdav.get("url")).rstrip("/"),
        "webdav_user": env_or("SWKA_WEBDAV_USER", webdav.get("user")),
        "webdav_password": env_or("SWKA_WEBDAV_PASSWORD", webdav.get("password")),
        "webdav_path": env_or("SWKA_WEBDAV_PATH", webdav.get("path")
                              or "/spendings.xlsx"),
        "interval_seconds": env_int("SWKA_INTERVAL_SECONDS",
                                    cfg.get("interval_seconds") or 3600, 60),
        "health_file": Path(env_or("SWKA_HEALTH_FILE", cfg.get("health_file"))
                            or ".health").expanduser(),
    }


def _resolve(p: Path) -> Path:
    return p if p.is_absolute() else Path(__file__).resolve().parent / p


def upload_workbook(cfg: dict, workbook: Path) -> None:
    """PUT the workbook to an ownCloud WebDAV endpoint (receipts stay local)."""
    base = cfg["webdav_url"]
    if not base or not cfg["webdav_user"]:
        return
    user, pw = cfg["webdav_user"], cfg["webdav_password"]
    segments = [s for s in cfg["webdav_path"].split("/") if s]
    if not segments:
        return
    quoted = [quote(s, safe="") for s in segments]
    dav_root = f"{base}/remote.php/dav/files/{quote(user, safe='')}"
    auth = (user, pw)
    for i in range(len(quoted) - 1):
        requests.request("MKCOL", dav_root + "/" + "/".join(quoted[:i + 1]),
                         auth=auth, timeout=(10, 30))
    url = dav_root + "/" + "/".join(quoted)
    with workbook.open("rb") as fh:
        resp = requests.put(url, data=fh, auth=auth, timeout=(10, 120))
    resp.raise_for_status()
    log.info(f"[I] uploaded to WebDAV: {'/'.join(segments)}")


def touch_health(cfg: dict) -> None:
    path = _resolve(cfg["health_file"])
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(datetime.now().isoformat())


def run_once(cfg: dict, args) -> int:
    workbook = _resolve(cfg["path"])
    receipts_dir = _resolve(cfg["receipts_folder"])

    try:
        wb, tx_rows, pos_rows = load_workbook_data(workbook)
    except (WorkbookError, zipfile.BadZipFile, OSError) as exc:
        sys.exit(f"error: cannot read workbook: {exc}")

    known_days = [d for d in (_to_date(r.get("Datum")) for r in tx_rows) if d]
    last_day = max(known_days) if known_days else None
    if args.since:
        try:
            date_from = parse_config_date(args.since)
        except ValueError:
            sys.exit(f"invalid --since date: {args.since!r} (expected DD.MM.YYYY)")
        log.info(f"[A] workbook: {workbook} | start overridden to {date_from:%d.%m.%Y}")
    elif last_day:
        date_from = last_day - timedelta(days=1)
        log.info(f"[A] workbook: {workbook} | last saved day {last_day:%d.%m.%Y} -> from {date_from:%d.%m.%Y}")
    else:
        date_from = date.today() - timedelta(days=cfg["lookback_days"])
        log.info(f"[A] workbook: {workbook} | new file -> last {cfg['lookback_days']} days from {date_from:%d.%m.%Y}")

    client = KartenService(cfg["karte_nr"], cfg["password"])
    try:
        if args.auth_token:
            client.use_token(args.auth_token)
            log.info("[B] using provided auth token (login skipped)")
        else:
            client.connect()
            client.login()
            log.info("[B] login OK, session token acquired")

        earliest = date.today() - timedelta(days=client.history_days - 1)
        if date_from < earliest:
            log.info(f"[A] clamped to server history limit: {earliest:%d.%m.%Y}")
            date_from = earliest
        date_to = date.today()
        if date_from > date_to:
            log.info("[C] nothing to fetch (workbook is up to date)")
            upload_workbook(cfg, workbook)
            return 0

        transactions, positions = client.fetch(date_from, date_to)
        log.info(f"[C] fetched {len(transactions)} transactions, {len(positions)} positions "
              f"({date_from:%d.%m.%Y} - {date_to:%d.%m.%Y})")

        tx_keys = {str(r.get("transFullId")) for r in tx_rows}
        ocr_rows = [
            r for r in pos_rows if str(r.get("Quelle") or "api").lower() == "ocr"
        ]
        if ocr_rows:
            api_tfids = {str(p.get("transFullId")) for p in positions}
            replaced = {
                str(r.get("transFullId")) for r in ocr_rows
            } & api_tfids
            if replaced:
                pos_rows = [
                    r for r in pos_rows
                    if not (
                        str(r.get("transFullId")) in replaced
                        and str(r.get("Quelle") or "api").lower() == "ocr"
                    )
                ]
                log.info(f"[D] replaced OCR positions with API data "
                      f"for {len(replaced)} transaction(s)")
        pos_keys = {
            (str(r.get("transFullId")), int(r.get("posId") or 0)) for r in pos_rows
        }
        new_tx = sorted(
            (t for t in transactions if str(t.get("transFullId")) not in tx_keys),
            key=lambda t: parse_api_datetime(t.get("datum") or ""),
        )
        tx_order = {
            str(t.get("transFullId")): parse_api_datetime(t.get("datum") or "")
            for t in transactions
        }
        new_pos = sorted(
            (
                p for p in positions
                if (str(p.get("transFullId")), int(p.get("posId") or 0)) not in pos_keys
            ),
            key=lambda p: (
                tx_order.get(str(p.get("transFullId")), datetime.max),
                int(p.get("posId") or 0),
            ),
        )
        log.info(f"[D] added {len(new_tx)} new transactions, {len(new_pos)} new positions "
              f"({len(transactions) - len(new_tx)} duplicates skipped)")

        all_tx = tx_rows + [tx_to_row(t) for t in new_tx]
        all_pos = pos_rows + [pos_to_row(p) for p in new_pos]

        if cfg["receipts_download"]:
            receipt_stats = download_receipts(client, all_tx, receipts_dir)
            if "error" in receipt_stats:
                log.info(f"[E] receipts failed: {receipt_stats['error']}")
            else:
                log.info(f"[E] receipts: {receipt_stats['new']} new, "
                      f"{receipt_stats['have']} already present, "
                      f"{receipt_stats['failed']} failed, "
                      f"{receipt_stats['none']} without PDF -> {receipts_dir}/")
        else:
            receipt_stats = None
            log.info("[E] receipts: skipped (disabled in config)")

        all_pos, ocr_stats = fill_positions_from_receipts(
            all_tx, all_pos, receipts_dir
        )
        if ocr_stats["parsed"]:
            log.info(f"[F] receipt text: {ocr_stats['parsed']} receipts parsed "
                  f"({ocr_stats['items']} items), {ocr_stats['failed']} failed")
        else:
            note = (f", {ocr_stats['failed']} failed" if ocr_stats["failed"] else "")
            log.info(f"[F] receipt text: no gaps to fill "
                  f"(all transactions have API positions{note})")
        if args.ocr_report:
            run_ocr_report(all_tx, all_pos, receipts_dir)

        for name in (TX_SHEET, POS_SHEET, OV_SHEET):
            if name in wb.sheetnames:
                del wb[name]
        write_data_sheet(wb, 0, TX_SHEET, TX_HEADERS, TX_WIDTHS, TX_TITLE, TX_NOTES, all_tx)
        write_data_sheet(wb, 1, POS_SHEET, POS_HEADERS, POS_WIDTHS, POS_TITLE, POS_NOTES, all_pos)
        build_overview(wb, all_tx, all_pos, receipt_stats, cfg["karte_nr"],
                       str(cfg["receipts_folder"]))
        save_workbook(wb, workbook)
        log.info(f"[G] saved {workbook}")
        upload_workbook(cfg, workbook)
    except (ApiError, WorkbookError, requests.RequestException, ValueError,
            zipfile.BadZipFile, OSError) as exc:
        log.error("%s", exc)
        return 1
    finally:
        client.close()
        log.info("[H] session closed")
    return 0


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stdout,
    )
    parser = argparse.ArgumentParser(
        description="Fetch KartenService transactions and append them to an Excel workbook."
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path(__file__).with_name("config.toml"),
        help="path to config.toml (default: next to this script; env vars override)",
    )
    parser.add_argument("--since", metavar="DD.MM.YYYY", help="override the start date")
    parser.add_argument(
        "--auth-token",
        help="use an existing session token instead of logging in (testing)",
    )
    parser.add_argument(
        "--ocr-report",
        action="store_true",
        help="compare receipt text extraction against API positions and print a report",
    )
    parser.add_argument(
        "--loop",
        action="store_true",
        help="run forever, scraping every interval_seconds (default 3600)",
    )
    args = parser.parse_args(argv)

    cfg = load_config(args.config, require_password=not args.auth_token)

    if not args.loop:
        return run_once(cfg, args)

    interval = cfg["interval_seconds"]
    log.info(f"[loop] starting, scraping every {interval}s "
          f"(health file: {_resolve(cfg['health_file'])})")
    while True:
        started = time.monotonic()
        rc = run_once(cfg, args)
        if rc == 0:
            touch_health(cfg)
        elapsed = time.monotonic() - started
        wait = max(30.0, interval - elapsed)
        log.info(f"[loop] next run in {wait:.0f}s")
        try:
            time.sleep(wait)
        except KeyboardInterrupt:
            log.info("loop stopped")
            return 0


if __name__ == "__main__":
    sys.exit(main())
