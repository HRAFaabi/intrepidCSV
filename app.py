import csv
import hashlib
import io
import zipfile

import pandas as pd
import streamlit as st

from transfer_logic import (
    CSV_FIELDNAMES,
    anomalies_in_period,
    build_pdf_from_rows,
    build_transfer_dataset,
    load_source_dataframe,
    rows_in_period,
    validate_assignments,
)

st.set_page_config(
    page_title="Fiches de transfert",
    page_icon="🚐",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Fleet configuration
# ---------------------------------------------------------------------------
# IMPORTANT: adjust capacity if your "9 places" includes the driver.
FLEET = [
    {"driver": "Khalid", "vehicle": "Ford Tourneo 9 places", "capacity": 9},
    {"driver": "Lhoussaine Dades", "vehicle": "Mercedes Vito 9 places", "capacity": 9},
    {"driver": "Bilal", "vehicle": "Toyota TX 4 places", "capacity": 4},
    {"driver": "Lahcen", "vehicle": "Toyota Prado 4 places", "capacity": 4},
]

FLEET_LABELS = [f"{x['driver']} — {x['vehicle']}" for x in FLEET]
LABEL_TO_PAIR = {
    f"{x['driver']} — {x['vehicle']}": (x["driver"], x["vehicle"])
    for x in FLEET
}
VEHICLE_CAPACITIES = {x["vehicle"]: x["capacity"] for x in FLEET}
UNASSIGNED = ""


def pair_to_label(driver: str, car: str) -> str:
    driver = (driver or "").strip()
    car = (car or "").strip()
    for item in FLEET:
        if driver == item["driver"] and car == item["vehicle"]:
            return f"{item['driver']} — {item['vehicle']}"
    return UNASSIGNED


def label_to_pair(label: str):
    return LABEL_TO_PAIR.get(label, ("", ""))


def dataframe_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


st.title("🚐 Générateur de fiches de transfert")
st.caption(
    "Desert Evasion — consolidation des snapshots, calcul du pickup réel, "
    "contrôle des anomalies, affectation chauffeur/véhicule et génération PDF."
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
if "transfer_dataset" not in st.session_state:
    st.session_state.transfer_dataset = None
if "load_report" not in st.session_state:
    st.session_state.load_report = None
if "loaded_key" not in st.session_state:
    st.session_state.loaded_key = None
if "loaded_file_hash" not in st.session_state:
    st.session_state.loaded_file_hash = None

# ---------------------------------------------------------------------------
# Upload / period / loading
# ---------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Fichier réservations (.xlsx, .xls ou .csv)",
    type=["xlsx", "xls", "csv"],
)

col1, col2, col3 = st.columns([1, 1, 2])
with col1:
    start_date = st.date_input("Date de début")
with col2:
    end_date = st.date_input("Date de fin")
with col3:
    pdf_title = st.text_input("Titre affiché sur le PDF", value="Desert Evasion")

consolidate_snapshots = st.checkbox(
    "Consolider les feuilles Excel comme snapshots successifs",
    value=True,
    help=(
        "Pour chaque Start Date, la dernière feuille du classeur contenant cette date "
        "est considérée comme la version la plus récente. Cela évite de ressortir "
        "d'anciennes réservations ensuite annulées ou modifiées."
    ),
)

current_file_hash = None
if uploaded_file is not None:
    current_file_hash = hashlib.sha256(uploaded_file.getvalue()).hexdigest()[:16]

load_clicked = st.button(
    "📂 Charger et analyser le fichier",
    type="primary",
    disabled=uploaded_file is None,
)

if load_clicked:
    if start_date > end_date:
        st.error("La date de début doit être avant ou égale à la date de fin.")
        st.stop()

    with st.spinner("Lecture, consolidation et analyse du fichier..."):
        try:
            source_df, load_report = load_source_dataframe(
                uploaded_file,
                consolidate_snapshots=consolidate_snapshots,
            )
            dataset = build_transfer_dataset(source_df)
        except Exception as exc:
            st.error(f"Impossible de traiter le fichier : {exc}")
            st.stop()

    st.session_state.transfer_dataset = dataset
    st.session_state.load_report = load_report
    st.session_state.loaded_file_hash = current_file_hash
    st.session_state.loaded_key = (
        f"{current_file_hash}-snap{int(consolidate_snapshots)}"
    )

if (
    uploaded_file is not None
    and st.session_state.loaded_file_hash is not None
    and current_file_hash != st.session_state.loaded_file_hash
):
    st.info("Un nouveau fichier a été sélectionné. Clique sur « Charger et analyser le fichier » pour l'utiliser.")

# ---------------------------------------------------------------------------
# Analysis + period filtering
# ---------------------------------------------------------------------------
if st.session_state.transfer_dataset is not None:
    if start_date > end_date:
        st.error("La date de début doit être avant ou égale à la date de fin.")
        st.stop()

    dataset = st.session_state.transfer_dataset
    load_report = st.session_state.load_report or {}

    confirmed_rows = rows_in_period(
        dataset["confirmed_rows"], start_date, end_date
    )
    pending_rows = rows_in_period(
        dataset["pending_rows"], start_date, end_date
    )
    source_anomalies = anomalies_in_period(
        dataset["anomalies"], start_date, end_date
    )

    total_groups = len(confirmed_rows)
    total_travelers = sum(int(r.get("Group Size", 0) or 0) for r in confirmed_rows)
    pending_travelers = sum(int(r.get("Group Size", 0) or 0) for r in pending_rows)

    error_count = sum(1 for x in source_anomalies if x.get("Severity") == "ERROR")
    warning_count = sum(1 for x in source_anomalies if x.get("Severity") == "WARNING")
    info_count = sum(1 for x in source_anomalies if x.get("Severity") == "INFO")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Confirmés", total_groups)
    m2.metric("Voyageurs", total_travelers)
    m3.metric("OnRequest", len(pending_rows))
    m4.metric("Warnings", warning_count)
    m5.metric("Erreurs source", error_count)

    # Snapshot audit
    if load_report:
        with st.expander("🔎 Rapport de lecture / consolidation"):
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Feuilles", load_report.get("sheet_count", 1))
            c2.metric("Lignes brutes", load_report.get("raw_rows", 0))
            c3.metric("Doublons identiques retirés", load_report.get("exact_duplicate_rows_removed", 0))
            c4.metric("Anciennes variantes retirées", load_report.get("snapshot_rows_removed", 0))
            c5.metric("Lignes courantes", load_report.get("current_rows", 0))
            if load_report.get("invalid_start_date_rows", 0):
                st.warning(
                    f"{load_report['invalid_start_date_rows']} ligne(s) ont une Start Date illisible. "
                    "Elles apparaissent dans les anomalies."
                )

    # Source anomalies: never silently hidden.
    if source_anomalies:
        anomalies_df = pd.DataFrame(source_anomalies)
        with st.expander(
            f"⚠️ Anomalies détectées ({len(source_anomalies)})",
            expanded=(error_count > 0),
        ):
            severity_filter = st.multiselect(
                "Afficher les niveaux",
                options=["ERROR", "WARNING", "INFO"],
                default=["ERROR", "WARNING", "INFO"],
            )
            filtered_anomalies = anomalies_df[
                anomalies_df["Severity"].isin(severity_filter)
            ]
            st.dataframe(
                filtered_anomalies,
                use_container_width=True,
                hide_index=True,
            )
            st.download_button(
                "⬇️ Télécharger les anomalies en CSV",
                data=dataframe_csv_bytes(anomalies_df),
                file_name="anomalies_transferts.csv",
                mime="text/csv",
            )
    else:
        st.success("Aucune anomalie source détectée sur la période sélectionnée.")

    # Pending transfers are visible but not operationally exported.
    if pending_rows:
        pending_df = pd.DataFrame(pending_rows)
        pending_cols = [
            "Pickup Date", "Pickup Time", "Flight Time", "Transfer Type",
            "Passengers", "Group Size", "Start", "Destination",
            "Flight No", "Txn ID",
        ]
        with st.expander(
            f"🟠 OnRequest — {len(pending_rows)} groupe(s), {pending_travelers} voyageur(s)",
            expanded=False,
        ):
            st.info(
                "Ces transferts sont affichés pour contrôle, mais ils ne sont pas ajoutés au PDF tant qu'ils ne sont pas Confirmed."
            )
            st.dataframe(
                pending_df[pending_cols],
                use_container_width=True,
                hide_index=True,
            )

    if not confirmed_rows:
        st.warning(
            "Aucun transfert confirmé dont le PICKUP tombe dans cette période. "
            "Le filtrage se fait maintenant sur Pickup Date, pas seulement sur Start Date."
        )
        st.stop()

    st.success(
        f"{total_groups} groupe(s) confirmé(s) — {total_travelers} voyageur(s). "
        "Affecte chauffeur/véhicule directement dans les tableaux."
    )

    # -----------------------------------------------------------------------
    # Editable operational table, grouped by real pickup date
    # -----------------------------------------------------------------------
    per_date_rows = {}
    for row in confirmed_rows:
        per_date_rows.setdefault(row["Pickup Date"], []).append(row)

    display_cols = [
        "Pickup Date", "Pickup Time", "Flight Time", "Transfer Type",
        "Passengers", "Group Size", "Chauffeur",
        "Start", "Destination", "Flight No", "Txn ID",
    ]

    tabs = st.tabs(list(per_date_rows.keys()))
    edited_per_date = {}

    for tab, (date_label, rows) in zip(tabs, per_date_rows.items()):
        with tab:
            df_day = pd.DataFrame(rows).copy()
            df_day["Chauffeur"] = df_day.apply(
                lambda r: pair_to_label(
                    r.get("Driver", ""),
                    r.get("Car", ""),
                ),
                axis=1,
            )
            df_day = df_day[display_cols]

            column_config = {
                col: st.column_config.TextColumn(disabled=True)
                for col in display_cols
                if col not in {"Chauffeur", "Group Size"}
            }
            column_config["Group Size"] = st.column_config.NumberColumn(
                "Grp",
                disabled=True,
                format="%d",
            )
            column_config["Chauffeur"] = st.column_config.SelectboxColumn(
                "Chauffeur / Véhicule",
                options=[UNASSIGNED] + FLEET_LABELS,
                required=False,
                width="medium",
            )

            edited_df = st.data_editor(
                df_day,
                key=f"editor_{st.session_state.loaded_key}_{date_label}",
                column_config=column_config,
                use_container_width=True,
                num_rows="fixed",
                hide_index=True,
            )

            records = edited_df.to_dict("records")
            for rec in records:
                driver, car = label_to_pair(rec.pop("Chauffeur", ""))
                rec["Driver"] = driver
                rec["Car"] = car
            edited_per_date[date_label] = records

    st.divider()

    all_rows = []
    for rows in edited_per_date.values():
        all_rows.extend(rows)

    operational_anomalies = validate_assignments(
        all_rows,
        VEHICLE_CAPACITIES,
    )
    blocking_errors = [
        x for x in operational_anomalies
        if x.get("Severity") == "ERROR"
    ]
    assignment_warnings = [
        x for x in operational_anomalies
        if x.get("Severity") == "WARNING"
    ]

    if operational_anomalies:
        with st.expander(
            f"🚦 Contrôle des affectations ({len(operational_anomalies)})",
            expanded=bool(blocking_errors),
        ):
            st.dataframe(
                pd.DataFrame(operational_anomalies),
                use_container_width=True,
                hide_index=True,
            )
            if blocking_errors:
                st.error(
                    "Corrige les erreurs rouges avant de générer le PDF : capacité dépassée ou conflit exact d'affectation."
                )
            elif assignment_warnings:
                st.warning(
                    "Les transferts non assignés sont autorisés, mais seront surlignés dans le PDF."
                )

    gen_col, info_col = st.columns([1, 3])
    with gen_col:
        generate_clicked = st.button(
            "📄 Générer le PDF",
            type="primary",
            disabled=bool(blocking_errors),
        )
    with info_col:
        if blocking_errors:
            st.caption(
                "Génération bloquée jusqu'à correction des conflits/capacités."
            )
        else:
            missing = sum(
                1 for r in all_rows
                if not (r.get("Driver") or "").strip()
                or not (r.get("Car") or "").strip()
            )
            if missing:
                st.caption(
                    f"{missing} transfert(s) restent sans affectation complète ; ils seront surlignés."
                )

    if generate_clicked:
        with st.spinner("Construction du PDF..."):
            pdf_buffer = build_pdf_from_rows(all_rows, title=pdf_title)

        file_suffix = (
            f"{start_date.strftime('%d%b%y').upper()}_"
            f"{end_date.strftime('%d%b%y').upper()}"
        )

        st.download_button(
            "⬇️ Télécharger le PDF",
            data=pdf_buffer,
            file_name=f"transferts_{file_suffix}.pdf",
            mime="application/pdf",
        )

        # CSV archive, one file per pickup day.
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for date_label, rows in edited_per_date.items():
                if not rows:
                    continue
                csv_buffer = io.StringIO()
                writer = csv.DictWriter(
                    csv_buffer,
                    fieldnames=CSV_FIELDNAMES,
                    extrasaction="ignore",
                )
                writer.writeheader()
                writer.writerows(rows)
                zf.writestr(
                    f"{date_label.replace('-', '')}.csv",
                    csv_buffer.getvalue(),
                )
        zip_buffer.seek(0)

        st.download_button(
            "⬇️ Télécharger les CSV édités (ZIP)",
            data=zip_buffer,
            file_name=f"csv_{file_suffix}.zip",
            mime="application/zip",
        )
