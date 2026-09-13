"""
transfer_logic.py

Shared logic for the transfer sheet tool, refactored from the original
build_transfer_sheet.py and transfers_pdf.py scripts so a Streamlit app
can call everything in-memory (no temp CSV files, no subprocess calls).
"""

import re
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak,
)

MONTHS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

# ------------------------------------------------------------------
# Parsing helpers (from build_transfer_sheet.py)
# ------------------------------------------------------------------

def parse_target_date(raw: str) -> datetime:
    cleaned = raw.strip().upper().replace(" ", "")
    m = re.match(r"^(\d{1,2})-?([A-Z]{3})-?(\d{2,4})$", cleaned)
    if not m:
        raise ValueError(f"Could not understand date '{raw}'.")
    day, mon_abbr, year = m.groups()
    if mon_abbr not in MONTHS:
        raise ValueError(f"Unknown month abbreviation '{mon_abbr}' in '{raw}'")
    day, month, year = int(day), MONTHS[mon_abbr], int(year)
    if year < 100:
        year += 2000
    return datetime(year, month, day)


def parse_row_date(raw: str) -> datetime:
    cleaned = raw.strip().upper().replace("-", "").replace(" ", "")
    m = re.match(r"^(\d{1,2})([A-Z]{3})(\d{2,4})$", cleaned)
    if not m:
        raise ValueError(f"Could not parse Start Date value '{raw}'")
    day, mon_abbr, year = m.groups()
    day, month, year = int(day), MONTHS[mon_abbr], int(year)
    if year < 100:
        year += 2000
    return datetime(year, month, day)


def parse_flight_time(raw: str):
    raw = (raw or "").strip()
    m = re.match(r"^(\d{1,2}):(\d{2})$", raw)
    if not m:
        return None
    hour, minute = int(m.group(1)), int(m.group(2))
    if hour > 23 or minute > 59:
        return None
    return hour, minute


def split_component_name(component: str):
    text = component.strip()
    transfer_type = ""
    m = re.match(r"^(Arrival|Departure)\s+Transfer\s*-\s*(.+)$", text, re.IGNORECASE)
    if m:
        transfer_type = m.group(1).capitalize()
        rest = m.group(2)
    else:
        rest = text
    parts = re.split(r"\s+to\s+", rest, flags=re.IGNORECASE)
    if len(parts) >= 2:
        start = parts[0].strip()
        destination = " to ".join(parts[1:]).strip()
    else:
        start = rest.strip()
        destination = ""
    return transfer_type, start, destination


def compute_pickup(flight_dt: datetime, transfer_type: str, destination: str):
    dest_upper = destination.upper()
    if transfer_type.lower() == "departure":
        if "(RAK)" in dest_upper or "MARRAKECH AIRPORT" in dest_upper:
            return flight_dt - timedelta(hours=3)
        if "(CMN)" in dest_upper or "CASABLANCA" in dest_upper:
            return flight_dt - timedelta(hours=7)
    return flight_dt


# Mots-clés utilisés pour deviner le type de transfert quand le texte
# source ne commence pas par "Arrival Transfer -" / "Departure Transfer -".
AIRPORT_KEYWORDS = ("(RAK)", "(CMN)", "AIRPORT")
ACCOMMODATION_KEYWORDS = ("RIAD", "HOTEL", "VILLA", "KASBAH", "MAISON")


def infer_transfer_type_from_destination(start: str, destination: str) -> str:
    """Devine Departure/Arrival quand le texte source ne le précise pas
    explicitement, en se basant uniquement sur la destination (comme
    compute_pickup le fait déjà pour le décalage horaire).
    Renvoie "" si aucun mot-clé connu n'est trouvé (reste non classé)."""
    dest_upper = destination.upper()
    if any(kw in dest_upper for kw in AIRPORT_KEYWORDS):
        return "Departure"
    if any(kw in dest_upper for kw in ACCOMMODATION_KEYWORDS):
        return "Arrival"
    return ""


# ------------------------------------------------------------------
# Reading the source file (xlsx or csv) into one combined DataFrame
# ------------------------------------------------------------------

def load_source_dataframe(uploaded_file) -> pd.DataFrame:
    """Accepts a Streamlit UploadedFile (xlsx or csv) and returns one
    combined DataFrame of all rows, all columns as strings."""
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)
        combined = pd.concat(all_sheets.values(), ignore_index=True)
    else:
        combined = pd.read_csv(uploaded_file, dtype=str, encoding="utf-8-sig")
    combined = combined.dropna(how="all")
    combined = combined.drop_duplicates()
    return combined


# ------------------------------------------------------------------
# Building the grouped transfer rows for one date (from build_transfer_sheet.py)
# ------------------------------------------------------------------

def build_transfers_for_date(df: pd.DataFrame, target_date: datetime):
    """Filters/parses df for one pickup date, returns a list of grouped
    dict rows with the same fields the original CSV output had."""
    results = []
    for _, row in df.iterrows():
        status = (row.get("Current Status") or "").strip().lower()
        if status != "confirmed":
            continue

        start_date_raw = row.get("Start Date") or ""
        try:
            row_date = parse_row_date(start_date_raw)
        except ValueError:
            continue
        if row_date.date() != target_date.date():
            continue

        transfer_type, start_place, destination_place = split_component_name(
            row.get("Component Name") or ""
        )
        if not transfer_type:
            transfer_type = infer_transfer_type_from_destination(start_place, destination_place)

        time_parts = parse_flight_time(row.get("Flight Time") or "")
        if time_parts is None:
            flight_time_str = (row.get("Flight Time") or "").strip()
            pickup_time_str = flight_time_str
            pickup_date_str = row_date.strftime("%d-%b-%y").upper()
        else:
            hour, minute = time_parts
            flight_dt = row_date.replace(hour=hour, minute=minute)
            pickup_dt = compute_pickup(flight_dt, transfer_type, destination_place)
            flight_time_str = flight_dt.strftime("%H:%M")
            pickup_time_str = pickup_dt.strftime("%H:%M")
            pickup_date_str = pickup_dt.strftime("%d-%b-%y").upper()

        results.append({
            "Txn ID": row.get("Txn ID", "") or "",
            "Title": row.get("Title", "") or "",
            "First Name": row.get("First Name", "") or "",
            "Surname": row.get("Surname", "") or "",
            "Transfer Type": transfer_type,
            "Start": start_place,
            "Destination": destination_place,
            "Flight No": (row.get("Flight No") or "").strip(),
            "Flight Time": flight_time_str,
            "Pickup Date": pickup_date_str,
            "Pickup Time": pickup_time_str,
            "Driver": "",
            "Car": "",
        })

    groups = {}
    order = []
    for r in results:
        key = (
            r["Txn ID"], r["Transfer Type"], r["Start"], r["Destination"],
            r["Flight No"], r["Pickup Date"], r["Pickup Time"],
        )
        name = " ".join(p for p in (r["Title"], r["First Name"], r["Surname"]) if p).strip()
        if key not in groups:
            groups[key] = {**r, "Passengers": [name]}
            order.append(key)
        else:
            groups[key]["Passengers"].append(name)

    grouped_results = []
    for key in order:
        g = groups[key]
        passengers = g.pop("Passengers")
        g["Passengers"] = "; ".join(passengers)
        g["Group Size"] = len(passengers)
        del g["Title"]
        del g["First Name"]
        del g["Surname"]
        grouped_results.append(g)

    grouped_results.sort(key=lambda r: (r["Pickup Date"], r["Pickup Time"]))
    return grouped_results


CSV_FIELDNAMES = [
    "Txn ID", "Passengers", "Group Size", "Transfer Type",
    "Start", "Destination", "Flight No", "Flight Time",
    "Pickup Date", "Pickup Time", "Driver", "Car",
]

# ------------------------------------------------------------------
# PDF building (from transfers_pdf.py, adapted to take rows directly)
# ------------------------------------------------------------------

COLOR_DEP = colors.HexColor("#1F4E79")
COLOR_ARR = colors.HexColor("#2E7D32")
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


def infer_transfer_type(row):
    t = (row.get("Transfer Type") or "").strip()
    if t:
        return t
    dest = (row.get("Destination") or "").lower()
    start = (row.get("Start") or "").lower()
    if "airport" in dest:
        return "Departure"
    if "airport" in start:
        return "Arrival"
    return "Unknown"


def parse_date(value):
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def parse_time(value):
    try:
        h, m = value.split(":")
        return int(h) * 60 + int(m)
    except (ValueError, AttributeError):
        return 24 * 60 + 1


def group_by_date(rows):
    groups = {}
    for row in rows:
        d = (row.get("Pickup Date") or "").strip()
        groups.setdefault(d, []).append(row)

    def sort_key(item):
        parsed = parse_date(item[0])
        return (parsed is None, parsed or datetime.max, item[0])

    return dict(sorted(groups.items(), key=sort_key))


def location_cell(row, kind):
    if kind == "Departure":
        return row.get("Start", "") or "—"
    return row.get("Destination", "") or "—"


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
    day_title = ParagraphStyle("day_title", parent=styles["Title"], fontSize=18,
                                alignment=TA_CENTER, spaceAfter=2)
    subtitle = ParagraphStyle(
        "subtitle", parent=styles["Normal"], fontSize=10, alignment=TA_CENTER,
        textColor=colors.HexColor("#555555"), spaceAfter=12,
    )
    empty_note = ParagraphStyle(
        "empty_note", parent=styles["Normal"], fontSize=9,
        textColor=colors.HexColor("#777777"), leftIndent=4,
    )
    return {
        "cell": cell, "cell_center": cell_center, "header_cell": header_cell,
        "section_title": section_title, "day_title": day_title,
        "subtitle": subtitle, "empty_note": empty_note,
    }


def build_section_table(rows, kind, styles):
    col_widths = [w * cm for _, w, _ in COLUMNS]
    header = [Paragraph(label, styles["header_cell"]) for label, _, _ in COLUMNS]
    data = [header]
    unassigned_row_indexes = []

    for i, row in enumerate(rows, start=1):
        driver = (row.get("Driver") or "").strip()
        car = (row.get("Car") or "").strip()
        if not driver and not car:
            unassigned_row_indexes.append(i)

        values = [
            row.get("Pickup Time", ""),
            location_cell(row, kind),
            row.get("Passengers", ""),
            row.get("Group Size", ""),
            row.get("Flight No", "") or "—",
            row.get("Flight Time", ""),
            driver or "To assign",
            car or "To assign",
            row.get("Txn ID", ""),
        ]
        data.append([
            Paragraph(str(v), styles[align])
            for v, (_, _, align) in zip(values, COLUMNS)
        ])

    table = Table(data, colWidths=col_widths, repeatRows=1)
    section_color = COLOR_DEP if kind == "Departure" else COLOR_ARR
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
    color = COLOR_DEP if kind == "Departure" else COLOR_ARR
    label = f"\u2708  {title}  ({count})"
    bar = Table([[Paragraph(label, styles["section_title"])]],
                colWidths=[sum(w for _, w, _ in COLUMNS) * cm])
    bar.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), color),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return bar


def build_day_flowables(date_label, rows, styles):
    flow = []
    parsed = parse_date(date_label)
    pretty_date = parsed.strftime("%A %d %B %Y") if parsed else date_label

    flow.append(Paragraph(pretty_date, styles["day_title"]))
    total = len(rows)
    flow.append(Paragraph(f"{total} transfer{'s' if total != 1 else ''} scheduled",
                           styles["subtitle"]))

    for kind, title in (("Departure", "Departures"), ("Arrival", "Arrivals")):
        section_rows = [r for r in rows if infer_transfer_type(r) == kind]
        section_rows.sort(key=lambda r: parse_time(r.get("Pickup Time", "")))

        flow.append(section_header_bar(title, len(section_rows), kind, styles))
        if section_rows:
            flow.append(build_section_table(section_rows, kind, styles))
        else:
            flow.append(Paragraph(f"No {title.lower()} on this day.", styles["empty_note"]))
        flow.append(Spacer(1, 14))

    other_rows = [r for r in rows if infer_transfer_type(r) not in ("Departure", "Arrival")]
    if other_rows:
        other_rows.sort(key=lambda r: parse_time(r.get("Pickup Time", "")))
        flow.append(section_header_bar("Unclassified", len(other_rows), "Departure", styles))
        flow.append(build_section_table(other_rows, "Departure", styles))
        flow.append(Spacer(1, 14))

    return flow


def add_page_furniture(canvas_obj, doc, report_title):
    canvas_obj.saveState()
    width, height = landscape(A4)
    canvas_obj.setFont("Helvetica", 8)
    canvas_obj.setFillColor(colors.HexColor("#888888"))
    canvas_obj.drawString(1.2 * cm, 0.7 * cm, report_title)
    canvas_obj.drawRightString(
        width - 1.2 * cm, 0.7 * cm,
        f"Page {doc.page} — generated {datetime.now().strftime('%d %b %Y %H:%M')}"
    )
    canvas_obj.restoreState()


def build_pdf_from_rows(rows, title="Transfer Schedule") -> BytesIO:
    """Builds the multi-day PDF in memory and returns a BytesIO buffer."""
    if not rows:
        raise ValueError("No rows to build a PDF from.")

    groups = group_by_date(rows)
    styles = make_styles()

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=1.2 * cm, rightMargin=1.2 * cm,
        topMargin=1.0 * cm, bottomMargin=1.2 * cm,
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
