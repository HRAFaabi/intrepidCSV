"""Business logic for the Streamlit transfer-sheet application.

Designed to:
- consolidate overlapping Excel sheets as successive snapshots;
- keep only operationally current rows for each Start Date;
- parse and normalize transfer data safely;
- calculate pickup dates before period filtering (including J-1 pickups);
- surface anomalies instead of silently dropping bad data;
- build the operational PDF in memory.
"""

from __future__ import annotations

import html
import re
from datetime import date, datetime, time, timedelta
from io import BytesIO
from typing import Iterable

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

MONTHS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

REQUIRED_COLUMNS = {
    "Txn ID",
    "First Name",
    "Surname",
    "Component Name",
    "Start Date",
    "Flight No",
    "Flight Time",
    "Current Status",
}

CSV_FIELDNAMES = [
    "Txn ID", "Passengers", "Group Size", "Transfer Type",
    "Start", "Destination", "Flight No", "Flight Time",
    "Pickup Date", "Pickup Time", "Driver", "Car",
]

# Pickup rules used by the agency. Add other airports only when the business
# rule is known and validated.
AIRPORT_PICKUP_OFFSETS = {
    "RAK": timedelta(hours=3),
    "CMN": timedelta(hours=7),
}

AIRPORT_ALIASES = {
    "RAK": (
        "(RAK)",
        "MARRAKECH AIRPORT",
        "MARRAKESH AIRPORT",
        "MARRAKECH MENARA AIRPORT",
        "MARRAKESH MENARA AIRPORT",
        "MENARA AIRPORT",
    ),
    "CMN": (
        "(CMN)",
        "CASABLANCA AIRPORT",
        "MOHAMMED V AIRPORT",
        "MOHAMED V AIRPORT",
        "MOHAMMED 5 AIRPORT",
    ),
}

ACCOMMODATION_KEYWORDS = (
    "RIAD", "HOTEL", "VILLA", "KASBAH", "MAISON", "LODGE", "RESORT",
)


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------

def _text(value) -> str:
    """Convert Excel/pandas values to a safe trimmed string."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def normalize_text(value) -> str:
    text = _text(value).upper()
    text = text.replace("’", "'")
    return re.sub(r"\s+", " ", text).strip()


def parse_row_date(raw) -> datetime:
    """Parse source dates such as 01SEP26, 01-SEP-26, 01/09/2026 or Excel dates."""
    if isinstance(raw, pd.Timestamp):
        return raw.to_pydatetime().replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(raw, datetime):
        return raw.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(raw, date):
        return datetime(raw.year, raw.month, raw.day)

    value = _text(raw)
    if not value:
        raise ValueError("empty date")

    cleaned = normalize_text(value).replace("-", "").replace(" ", "")
    m = re.fullmatch(r"(\d{1,2})([A-Z]{3})(\d{2,4})", cleaned)
    if m:
        day_s, month_s, year_s = m.groups()
        if month_s not in MONTHS:
            raise ValueError(f"unknown month in '{value}'")
        year = int(year_s)
        if year < 100:
            year += 2000
        return datetime(year, MONTHS[month_s], int(day_s))

    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass

    raise ValueError(f"Could not parse Start Date value '{value}'")


def parse_display_date(value: str) -> datetime | None:
    value = _text(value)
    if not value:
        return None
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return None


def parse_flight_time(raw) -> tuple[int, int] | None:
    """Parse 24-hour source times, including Excel time objects."""
    if isinstance(raw, pd.Timestamp):
        return raw.hour, raw.minute
    if isinstance(raw, datetime):
        return raw.hour, raw.minute
    if isinstance(raw, time):
        return raw.hour, raw.minute

    value = _text(raw)
    if not value:
        return None

    # Accept H:MM or HH:MM, optionally with seconds.
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?", value)
    if not m:
        return None
    hour, minute = int(m.group(1)), int(m.group(2))
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return hour, minute


def format_date(value: datetime) -> str:
    return value.strftime("%d-%b-%y").upper()


def normalize_status(raw) -> str:
    value = re.sub(r"[\s_-]+", "", normalize_text(raw)).lower()
    if value == "confirmed":
        return "confirmed"
    if value in {"onrequest", "pending", "request"}:
        return "onrequest"
    if value in {"cancelled", "canceled"}:
        return "cancelled"
    return value


def _make_anomaly(
    severity: str,
    code: str,
    message: str,
    *,
    txn_id: str = "",
    start_date: str = "",
    pickup_date: str = "",
    component: str = "",
) -> dict:
    return {
        "Severity": severity,
        "Code": code,
        "Txn ID": txn_id,
        "Start Date": start_date,
        "Pickup Date": pickup_date,
        "Component": component,
        "Message": message,
    }


def deduplicate_anomalies(anomalies: Iterable[dict]) -> list[dict]:
    seen = set()
    out = []
    for item in anomalies:
        key = tuple(item.get(k, "") for k in (
            "Severity", "Code", "Txn ID", "Start Date", "Pickup Date", "Component", "Message"
        ))
        if key not in seen:
            seen.add(key)
            out.append(item)
    severity_order = {"ERROR": 0, "WARNING": 1, "INFO": 2}
    out.sort(key=lambda x: (severity_order.get(x.get("Severity"), 9), x.get("Start Date", ""), x.get("Txn ID", "")))
    return out


# ---------------------------------------------------------------------------
# Source loading + snapshot consolidation
# ---------------------------------------------------------------------------

def _validate_columns(df: pd.DataFrame) -> None:
    missing = sorted(REQUIRED_COLUMNS - set(df.columns))
    if missing:
        raise ValueError(
            "Colonnes obligatoires manquantes : " + ", ".join(missing)
        )


def _canonical_start_date_key(value) -> str | None:
    try:
        return parse_row_date(value).date().isoformat()
    except ValueError:
        return None


def load_source_dataframe(uploaded_file, consolidate_snapshots: bool = True):
    """Load XLSX/XLS/CSV and optionally consolidate overlapping Excel snapshots.

    Snapshot rule: for each valid Start Date, the right-most/latest worksheet that
    contains that date is considered authoritative for that date.

    Returns: (dataframe, load_report)
    """
    name = getattr(uploaded_file, "name", "uploaded_file").lower()
    report = {
        "file_name": getattr(uploaded_file, "name", "uploaded_file"),
        "sheet_count": 1,
        "raw_rows": 0,
        "current_rows": 0,
        "exact_duplicate_rows_removed": 0,
        "snapshot_rows_removed": 0,
        "dates_consolidated": 0,
        "invalid_start_date_rows": 0,
    }

    if name.endswith((".xlsx", ".xls")):
        sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype=object)
        if not sheets:
            raise ValueError("Le classeur Excel ne contient aucune feuille lisible.")

        report["sheet_count"] = len(sheets)
        frames = []
        for sheet_order, (sheet_name, frame) in enumerate(sheets.items()):
            frame = frame.copy()
            frame.columns = [_text(c) for c in frame.columns]
            frame = frame.dropna(how="all")
            if frame.empty:
                continue
            _validate_columns(frame)
            frame["_source_sheet_order"] = sheet_order
            frame["_source_sheet_name"] = sheet_name
            frames.append(frame)

        if not frames:
            raise ValueError("Le classeur ne contient aucune ligne de réservation.")
        combined = pd.concat(frames, ignore_index=True)
    else:
        combined = pd.read_csv(uploaded_file, dtype=object, encoding="utf-8-sig")
        combined.columns = [_text(c) for c in combined.columns]
        combined = combined.dropna(how="all")
        _validate_columns(combined)
        combined["_source_sheet_order"] = 0
        combined["_source_sheet_name"] = "CSV"

    report["raw_rows"] = len(combined)

    # Remove exact duplicates while ignoring our source metadata.
    business_columns = [
        c for c in combined.columns
        if c not in {"_source_sheet_order", "_source_sheet_name"}
    ]
    before_exact_dedupe = len(combined)
    combined = combined.drop_duplicates(subset=business_columns, keep="last").copy()
    report["exact_duplicate_rows_removed"] = before_exact_dedupe - len(combined)
    combined["_start_date_key"] = combined["Start Date"].map(_canonical_start_date_key)
    report["invalid_start_date_rows"] = int(combined["_start_date_key"].isna().sum())

    if consolidate_snapshots and report["sheet_count"] > 1:
        valid_mask = combined["_start_date_key"].notna()
        valid = combined.loc[valid_mask].copy()
        invalid = combined.loc[~valid_mask].copy()

        if not valid.empty:
            latest_order_by_date = valid.groupby("_start_date_key")["_source_sheet_order"].transform("max")
            latest_valid = valid.loc[valid["_source_sheet_order"] == latest_order_by_date].copy()
            report["snapshot_rows_removed"] = len(valid) - len(latest_valid)
            report["dates_consolidated"] = int(valid["_start_date_key"].nunique())
            combined = pd.concat([latest_valid, invalid], ignore_index=True)

    combined = combined.sort_values(
        by=["_source_sheet_order", "_start_date_key"],
        kind="stable",
        na_position="last",
    ).reset_index(drop=True)

    report["current_rows"] = len(combined)
    return combined, report


# ---------------------------------------------------------------------------
# Transfer parsing / inference / pickup calculation
# ---------------------------------------------------------------------------

def detect_airport(location: str) -> str | None:
    text = normalize_text(location)
    if not text:
        return None

    for code, aliases in AIRPORT_ALIASES.items():
        if any(alias in text for alias in aliases):
            return code

    # Generic IATA-like code: useful for warning about airports whose offset is unknown.
    code_match = re.search(r"\(([A-Z]{3})\)", text)
    if "AIRPORT" in text and code_match:
        return code_match.group(1)

    return None


def looks_like_airport(location: str) -> bool:
    text = normalize_text(location)
    return "AIRPORT" in text or bool(re.search(r"\([A-Z]{3}\)", text))


def split_component_name(component: str):
    """Return (explicit_type, start, destination).

    The type may appear anywhere in the component text. Route extraction still
    requires a recognizable 'A to B' portion.
    """
    text = _text(component)
    explicit_type = ""
    type_match = re.search(r"\b(Arrival|Departure)\s+Transfer\b", text, re.IGNORECASE)
    if type_match:
        explicit_type = type_match.group(1).capitalize()

    # Prefer text after "Arrival/Departure Transfer -" when available.
    route_text = text
    prefix_match = re.search(
        r"\b(?:Arrival|Departure)\s+Transfer\s*-\s*(.+)$",
        text,
        re.IGNORECASE,
    )
    if prefix_match:
        route_text = prefix_match.group(1).strip()

    parts = re.split(r"\s+to\s+", route_text, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) == 2:
        return explicit_type, parts[0].strip(), parts[1].strip()

    return explicit_type, route_text.strip(), ""


def infer_transfer_type(component: str, start: str, destination: str):
    component_upper = normalize_text(component)
    if re.search(r"\bARRIVAL\s+TRANSFER\b", component_upper):
        return "Arrival", "explicit"
    if re.search(r"\bDEPARTURE\s+TRANSFER\b", component_upper):
        return "Departure", "explicit"

    start_airport = detect_airport(start) or ("AIRPORT" if looks_like_airport(start) else None)
    dest_airport = detect_airport(destination) or ("AIRPORT" if looks_like_airport(destination) else None)

    if start_airport and not dest_airport:
        return "Arrival", "airport_route"
    if dest_airport and not start_airport:
        return "Departure", "airport_route"

    # Weak fallback used only if one side clearly looks like accommodation.
    start_upper = normalize_text(start)
    dest_upper = normalize_text(destination)
    start_accommodation = any(k in start_upper for k in ACCOMMODATION_KEYWORDS)
    dest_accommodation = any(k in dest_upper for k in ACCOMMODATION_KEYWORDS)
    if dest_accommodation and not start_accommodation:
        return "Arrival", "accommodation_fallback"
    if start_accommodation and not dest_accommodation:
        return "Departure", "accommodation_fallback"

    return "", "unknown"


def compute_pickup(flight_dt: datetime, transfer_type: str, destination: str):
    """Return (pickup_datetime, optional_warning_code)."""
    if transfer_type.lower() != "departure":
        return flight_dt, None

    airport = detect_airport(destination)
    if airport in AIRPORT_PICKUP_OFFSETS:
        return flight_dt - AIRPORT_PICKUP_OFFSETS[airport], None

    if looks_like_airport(destination):
        return flight_dt, "UNKNOWN_AIRPORT_OFFSET"

    # Inter-city / hotel destination: do not apply an airport offset.
    return flight_dt, None


def _group_transfer_records(records: list[dict]) -> list[dict]:
    groups = {}
    order = []

    for r in records:
        # If Txn ID is absent, keep the row isolated to avoid merging unrelated bookings.
        txn_for_key = r["Txn ID"] or r["_row_key"]
        key = (
            txn_for_key,
            r["Transfer Type"],
            r["Start"],
            r["Destination"],
            r["Flight No"],
            r["Flight Time"],
            r["Pickup Date"],
            r["Pickup Time"],
        )
        name = " ".join(
            p for p in (r["Title"], r["First Name"], r["Surname"]) if p
        ).strip() or "Nom manquant"

        if key not in groups:
            groups[key] = {
                "Txn ID": r["Txn ID"],
                "Passengers": [name],
                "Group Size": 0,
                "Transfer Type": r["Transfer Type"],
                "Start": r["Start"],
                "Destination": r["Destination"],
                "Flight No": r["Flight No"],
                "Flight Time": r["Flight Time"],
                "Pickup Date": r["Pickup Date"],
                "Pickup Time": r["Pickup Time"],
                "Driver": "",
                "Car": "",
                "_Status": r["_Status"],
            }
            order.append(key)
        else:
            groups[key]["Passengers"].append(name)

    out = []
    for key in order:
        item = groups[key]
        item["Group Size"] = len(item["Passengers"])
        item["Passengers"] = "; ".join(item["Passengers"])
        out.append(item)

    out.sort(
        key=lambda r: (
            parse_display_date(r["Pickup Date"]) or datetime.max,
            parse_time_minutes(r["Pickup Time"]),
            r["Txn ID"],
        )
    )
    return out


def build_transfer_dataset(df: pd.DataFrame):
    """Parse the consolidated source into confirmed/pending grouped transfers.

    Returns a dict with confirmed_rows, pending_rows and anomalies.
    Cancelled rows are intentionally excluded from operational outputs.
    """
    confirmed_records = []
    pending_records = []
    anomalies = []

    for idx, row in df.iterrows():
        status_raw = _text(row.get("Current Status"))
        status = normalize_status(status_raw)

        if status == "cancelled":
            continue
        if status not in {"confirmed", "onrequest"}:
            anomalies.append(_make_anomaly(
                "WARNING",
                "UNKNOWN_STATUS",
                f"Statut non reconnu : '{status_raw or 'vide'}'. Ligne exclue.",
                txn_id=_text(row.get("Txn ID")),
                start_date=_text(row.get("Start Date")),
                component=_text(row.get("Component Name")),
            ))
            continue

        txn_id = _text(row.get("Txn ID"))
        component = _text(row.get("Component Name"))
        start_date_raw = _text(row.get("Start Date"))

        if not txn_id:
            anomalies.append(_make_anomaly(
                "WARNING", "MISSING_TXN_ID",
                "Txn ID manquant : la ligne ne sera pas fusionnée avec d'autres passagers.",
                start_date=start_date_raw,
                component=component,
            ))

        try:
            row_date = parse_row_date(row.get("Start Date"))
        except ValueError:
            anomalies.append(_make_anomaly(
                "ERROR", "INVALID_START_DATE",
                f"Date de départ illisible : '{start_date_raw or 'vide'}'. Ligne exclue.",
                txn_id=txn_id,
                start_date=start_date_raw,
                component=component,
            ))
            continue

        explicit_type, start_place, destination_place = split_component_name(component)
        transfer_type, type_source = infer_transfer_type(component, start_place, destination_place)
        if explicit_type and not transfer_type:
            transfer_type = explicit_type

        if not destination_place:
            anomalies.append(_make_anomaly(
                "WARNING", "ROUTE_UNPARSED",
                "Le trajet 'départ → destination' n'a pas pu être extrait du Component Name.",
                txn_id=txn_id,
                start_date=format_date(row_date),
                component=component,
            ))

        if not transfer_type:
            anomalies.append(_make_anomaly(
                "WARNING", "UNCLASSIFIED_TRANSFER",
                "Impossible de déterminer Arrival/Departure. Le transfert ira dans 'Unclassified'.",
                txn_id=txn_id,
                start_date=format_date(row_date),
                component=component,
            ))
        elif type_source == "accommodation_fallback":
            anomalies.append(_make_anomaly(
                "INFO", "TYPE_INFERRED_WEAKLY",
                f"Type {transfer_type} déduit à partir du type de lieu, à vérifier.",
                txn_id=txn_id,
                start_date=format_date(row_date),
                component=component,
            ))

        raw_time = _text(row.get("Flight Time"))
        time_parts = parse_flight_time(row.get("Flight Time"))
        pickup_date_str = format_date(row_date)
        pickup_time_str = raw_time
        flight_time_str = raw_time

        if time_parts is None:
            anomalies.append(_make_anomaly(
                "WARNING", "INVALID_FLIGHT_TIME",
                f"Heure de vol illisible : '{raw_time or 'vide'}'. Pickup non recalculé.",
                txn_id=txn_id,
                start_date=format_date(row_date),
                pickup_date=pickup_date_str,
                component=component,
            ))
        else:
            hour, minute = time_parts
            flight_dt = row_date.replace(hour=hour, minute=minute)
            pickup_dt, warning_code = compute_pickup(
                flight_dt, transfer_type, destination_place
            )
            flight_time_str = flight_dt.strftime("%H:%M")
            pickup_date_str = format_date(pickup_dt)
            pickup_time_str = pickup_dt.strftime("%H:%M")

            if warning_code == "UNKNOWN_AIRPORT_OFFSET":
                anomalies.append(_make_anomaly(
                    "WARNING", "UNKNOWN_AIRPORT_OFFSET",
                    "Aéroport détecté mais aucune règle d'anticipation n'est configurée ; pickup laissé à l'heure du vol.",
                    txn_id=txn_id,
                    start_date=format_date(row_date),
                    pickup_date=pickup_date_str,
                    component=component,
                ))

            if pickup_dt.date() != row_date.date():
                anomalies.append(_make_anomaly(
                    "INFO", "CROSS_DAY_PICKUP",
                    f"Pickup calculé la veille : {pickup_date_str} à {pickup_time_str} pour un vol du {format_date(row_date)} à {flight_time_str}.",
                    txn_id=txn_id,
                    start_date=format_date(row_date),
                    pickup_date=pickup_date_str,
                    component=component,
                ))

        if not _text(row.get("First Name")) and not _text(row.get("Surname")):
            anomalies.append(_make_anomaly(
                "WARNING", "MISSING_PASSENGER_NAME",
                "Nom du passager manquant.",
                txn_id=txn_id,
                start_date=format_date(row_date),
                pickup_date=pickup_date_str,
                component=component,
            ))

        record = {
            "Txn ID": txn_id,
            "Title": _text(row.get("Title")),
            "First Name": _text(row.get("First Name")),
            "Surname": _text(row.get("Surname")),
            "Transfer Type": transfer_type,
            "Start": start_place,
            "Destination": destination_place,
            "Flight No": _text(row.get("Flight No")),
            "Flight Time": flight_time_str,
            "Pickup Date": pickup_date_str,
            "Pickup Time": pickup_time_str,
            "Driver": "",
            "Car": "",
            "_Status": status,
            "_row_key": f"ROW-{idx}",
        }

        if status == "confirmed":
            confirmed_records.append(record)
        else:
            pending_records.append(record)

    confirmed_rows = _group_transfer_records(confirmed_records)
    pending_rows = _group_transfer_records(pending_records)

    return {
        "confirmed_rows": confirmed_rows,
        "pending_rows": pending_rows,
        "anomalies": deduplicate_anomalies(anomalies),
    }


def rows_in_period(rows: Iterable[dict], start_date: date, end_date: date) -> list[dict]:
    out = []
    for row in rows:
        parsed = parse_display_date(row.get("Pickup Date", ""))
        if parsed and start_date <= parsed.date() <= end_date:
            out.append(dict(row))
    return out


def anomalies_in_period(anomalies: Iterable[dict], start_date: date, end_date: date) -> list[dict]:
    """Keep anomalies linked to either pickup or source dates in the requested period.

    Global/unparseable anomalies are kept so they are never hidden silently.
    """
    out = []
    for item in anomalies:
        date_value = item.get("Pickup Date") or item.get("Start Date") or ""
        parsed = parse_display_date(date_value)
        if parsed is None:
            # Start Date can also be source style 01SEP26.
            try:
                parsed = parse_row_date(date_value)
            except ValueError:
                parsed = None
        if parsed is None or start_date <= parsed.date() <= end_date:
            out.append(dict(item))
    return out


# ---------------------------------------------------------------------------
# Operational assignment validation
# ---------------------------------------------------------------------------

def validate_assignments(rows: Iterable[dict], vehicle_capacities: dict[str, int]):
    anomalies = []
    driver_slots = {}
    vehicle_slots = {}

    rows = list(rows)
    for row in rows:
        txn_id = _text(row.get("Txn ID"))
        pickup_date = _text(row.get("Pickup Date"))
        pickup_time = _text(row.get("Pickup Time"))
        driver = _text(row.get("Driver"))
        car = _text(row.get("Car"))
        group_size = row.get("Group Size", 0)
        try:
            group_size = int(group_size)
        except (TypeError, ValueError):
            group_size = 0

        if not driver or not car:
            anomalies.append(_make_anomaly(
                "WARNING", "UNASSIGNED",
                "Chauffeur ou véhicule non assigné. La ligne sera surlignée dans le PDF.",
                txn_id=txn_id,
                pickup_date=pickup_date,
            ))

        if car:
            capacity = vehicle_capacities.get(car)
            if capacity is not None and group_size > capacity:
                anomalies.append(_make_anomaly(
                    "ERROR", "OVER_CAPACITY",
                    f"{group_size} passagers pour un véhicule configuré à {capacity} places.",
                    txn_id=txn_id,
                    pickup_date=pickup_date,
                ))

        # Exact-time conflict only. Without mission duration/travel time we avoid
        # pretending to detect broader overlaps.
        if pickup_date and pickup_time and driver:
            driver_slots.setdefault((pickup_date, pickup_time, driver), []).append(txn_id or "sans Txn ID")
        if pickup_date and pickup_time and car:
            vehicle_slots.setdefault((pickup_date, pickup_time, car), []).append(txn_id or "sans Txn ID")

    for (d, t, driver), txn_ids in driver_slots.items():
        if len(txn_ids) > 1:
            anomalies.append(_make_anomaly(
                "ERROR", "DRIVER_CONFLICT",
                f"{driver} est affecté à {len(txn_ids)} transferts à {t} ({', '.join(txn_ids)}).",
                pickup_date=d,
            ))

    for (d, t, car), txn_ids in vehicle_slots.items():
        if len(txn_ids) > 1:
            anomalies.append(_make_anomaly(
                "ERROR", "VEHICLE_CONFLICT",
                f"{car} est affecté à {len(txn_ids)} transferts à {t} ({', '.join(txn_ids)}).",
                pickup_date=d,
            ))

    return deduplicate_anomalies(anomalies)


# ---------------------------------------------------------------------------
# PDF generation
# ---------------------------------------------------------------------------

COLOR_DEP = colors.HexColor("#1F4E79")
COLOR_ARR = colors.HexColor("#2E7D32")
COLOR_OTHER = colors.HexColor("#6B7280")
COLOR_ROW_ALT = colors.HexColor("#F2F2F2")
COLOR_UNASSIGNED = colors.HexColor("#FFF3E0")
COLOR_GRID = colors.HexColor("#CCCCCC")

COLUMNS = [
    ("Time", 1.6, "cell_center"),
    ("Hotel / Riad", 5.4, "cell"),
    ("Passengers", 6.0, "cell"),
    ("Grp", 1.0, "cell_center"),
    ("Flight No", 2.3, "cell_center"),
    ("Flt Time", 1.6, "cell_center"),
    ("Driver", 3.4, "cell"),
    ("Car", 2.6, "cell"),
    ("Txn ID", 1.9, "cell_center"),
]


def parse_time_minutes(value):
    parts = parse_flight_time(value)
    if parts is None:
        return 24 * 60 + 1
    return parts[0] * 60 + parts[1]


def group_by_date(rows):
    groups = {}
    for row in rows:
        d = _text(row.get("Pickup Date"))
        groups.setdefault(d, []).append(row)

    def sort_key(item):
        parsed = parse_display_date(item[0])
        return (parsed is None, parsed or datetime.max, item[0])

    return dict(sorted(groups.items(), key=sort_key))


def row_transfer_type(row):
    t = _text(row.get("Transfer Type"))
    if t in {"Arrival", "Departure"}:
        return t
    start = _text(row.get("Start"))
    destination = _text(row.get("Destination"))
    inferred, _ = infer_transfer_type("", start, destination)
    return inferred or "Unknown"


def location_cell(row, kind):
    if kind == "Departure":
        return _text(row.get("Start")) or "—"
    if kind == "Arrival":
        return _text(row.get("Destination")) or "—"
    start = _text(row.get("Start")) or "—"
    destination = _text(row.get("Destination")) or "—"
    return f"{start} → {destination}"


def _safe_paragraph(value, style):
    return Paragraph(html.escape(str(value or "")), style)


def make_styles():
    styles = getSampleStyleSheet()
    cell = ParagraphStyle("cell", parent=styles["Normal"], fontSize=8, leading=9.5)
    cell_center = ParagraphStyle("cell_center", parent=cell, alignment=TA_CENTER)
    header_cell = ParagraphStyle(
        "header_cell", parent=styles["Normal"], fontSize=8.5, leading=10,
        textColor=colors.white, alignment=TA_CENTER, fontName="Helvetica-Bold",
    )
    section_title = ParagraphStyle(
        "section_title", parent=styles["Heading2"], fontSize=13,
        textColor=colors.white, spaceAfter=0, spaceBefore=0, leading=16,
    )
    day_title = ParagraphStyle(
        "day_title", parent=styles["Title"], fontSize=18,
        alignment=TA_CENTER, spaceAfter=2,
    )
    subtitle = ParagraphStyle(
        "subtitle", parent=styles["Normal"], fontSize=10, alignment=TA_CENTER,
        textColor=colors.HexColor("#555555"), spaceAfter=12,
    )
    empty_note = ParagraphStyle(
        "empty_note", parent=styles["Normal"], fontSize=9,
        textColor=colors.HexColor("#777777"), leftIndent=4,
    )
    return {
        "cell": cell,
        "cell_center": cell_center,
        "header_cell": header_cell,
        "section_title": section_title,
        "day_title": day_title,
        "subtitle": subtitle,
        "empty_note": empty_note,
    }


def build_section_table(rows, kind, styles):
    col_widths = [w * cm for _, w, _ in COLUMNS]
    header = [_safe_paragraph(label, styles["header_cell"]) for label, _, _ in COLUMNS]
    data = [header]
    unassigned_row_indexes = []

    for i, row in enumerate(rows, start=1):
        driver = _text(row.get("Driver"))
        car = _text(row.get("Car"))
        if not driver or not car:
            unassigned_row_indexes.append(i)

        values = [
            _text(row.get("Pickup Time")),
            location_cell(row, kind),
            _text(row.get("Passengers")),
            row.get("Group Size", ""),
            _text(row.get("Flight No")) or "—",
            _text(row.get("Flight Time")),
            driver or "To assign",
            car or "To assign",
            _text(row.get("Txn ID")),
        ]
        data.append([
            _safe_paragraph(v, styles[align])
            for v, (_, _, align) in zip(values, COLUMNS)
        ])

    table = Table(data, colWidths=col_widths, repeatRows=1)
    if kind == "Departure":
        section_color = COLOR_DEP
    elif kind == "Arrival":
        section_color = COLOR_ARR
    else:
        section_color = COLOR_OTHER

    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), section_color),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_GRID),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), COLOR_ROW_ALT))
    for i in unassigned_row_indexes:
        style_cmds.append(("BACKGROUND", (6, i), (7, i), COLOR_UNASSIGNED))

    table.setStyle(TableStyle(style_cmds))
    return table


def section_header_bar(title, count, kind, styles):
    if kind == "Departure":
        color = COLOR_DEP
    elif kind == "Arrival":
        color = COLOR_ARR
    else:
        color = COLOR_OTHER

    label = f"{title} ({count})"
    bar = Table(
        [[_safe_paragraph(label, styles["section_title"])]],
        colWidths=[sum(w for _, w, _ in COLUMNS) * cm],
    )
    bar.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), color),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return bar


def build_day_flowables(date_label, rows, styles):
    flow = []
    parsed = parse_display_date(date_label)
    pretty_date = parsed.strftime("%A %d %B %Y") if parsed else date_label

    flow.append(_safe_paragraph(pretty_date, styles["day_title"]))
    total = len(rows)
    flow.append(_safe_paragraph(
        f"{total} transfer{'s' if total != 1 else ''} scheduled",
        styles["subtitle"],
    ))

    sections = (
        ("Departure", "Departures"),
        ("Arrival", "Arrivals"),
        ("Unknown", "Unclassified"),
    )
    for kind, title in sections:
        if kind == "Unknown":
            section_rows = [r for r in rows if row_transfer_type(r) not in {"Departure", "Arrival"}]
        else:
            section_rows = [r for r in rows if row_transfer_type(r) == kind]
        section_rows.sort(key=lambda r: parse_time_minutes(r.get("Pickup Time", "")))

        # Do not print an empty Unclassified section.
        if kind == "Unknown" and not section_rows:
            continue

        flow.append(section_header_bar(title, len(section_rows), kind, styles))
        if section_rows:
            flow.append(build_section_table(section_rows, kind, styles))
        else:
            flow.append(_safe_paragraph(f"No {title.lower()} on this day.", styles["empty_note"]))
        flow.append(Spacer(1, 14))

    return flow


def add_page_furniture(canvas_obj, doc, report_title):
    canvas_obj.saveState()
    width, _ = landscape(A4)
    canvas_obj.setFont("Helvetica", 8)
    canvas_obj.setFillColor(colors.HexColor("#888888"))
    canvas_obj.drawString(1.2 * cm, 0.7 * cm, report_title)
    canvas_obj.drawRightString(
        width - 1.2 * cm,
        0.7 * cm,
        f"Page {doc.page} — generated {datetime.now().strftime('%d %b %Y %H:%M')}",
    )
    canvas_obj.restoreState()


def build_pdf_from_rows(rows, title="Transfer Schedule") -> BytesIO:
    rows = list(rows)
    if not rows:
        raise ValueError("No rows to build a PDF from.")

    groups = group_by_date(rows)
    styles = make_styles()
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=1.2 * cm,
        rightMargin=1.2 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.2 * cm,
        title=title,
    )

    story = []
    for idx, (date_label, day_rows) in enumerate(groups.items()):
        if idx > 0:
            story.append(PageBreak())
        story.extend(build_day_flowables(date_label, day_rows, styles))

    doc.build(
        story,
        onFirstPage=lambda c, d: add_page_furniture(c, d, title),
        onLaterPages=lambda c, d: add_page_furniture(c, d, title),
    )
    buffer.seek(0)
    return buffer
