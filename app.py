import io
import zipfile
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

from transfer_logic import (
    load_source_dataframe,
    build_transfers_for_date,
    build_pdf_from_rows,
    CSV_FIELDNAMES,
)

st.set_page_config(page_title="Fiches de transfert", page_icon="🚐", layout="wide")

# ------------------------------------------------------------------
# Flotte — modifie cette liste quand un chauffeur/véhicule change
# ------------------------------------------------------------------
FLEET = [
    ("Khalid", "Ford Tourneo 9 places"),
    ("Lhoussaine Dades", "Mercedes Vito 9 places"),
    ("Bilal", "Toyota TX 4 places"),
    ("Lahcen", "Toyota Prado 4 places"),
]
FLEET_LABELS = [f"{driver} — {vehicle}" for driver, vehicle in FLEET]
LABEL_TO_PAIR = {f"{driver} — {vehicle}": (driver, vehicle) for driver, vehicle in FLEET}
UNASSIGNED = ""


def pair_to_label(driver: str, car: str) -> str:
    """Reconstitue le label 'Chauffeur — Véhicule' si ça matche la flotte connue."""
    for d, v in FLEET:
        if driver.strip() == d and car.strip() == v:
            return f"{d} — {v}"
    return UNASSIGNED


def label_to_pair(label: str):
    return LABEL_TO_PAIR.get(label, ("", ""))

st.title("🚐 Générateur de fiches de transfert")
st.caption(
    "Desert Evasion — upload le fichier de réservations, charge une période, "
    "ajoute les chauffeurs/voitures directement dans le tableau, puis génère le PDF."
)

# ------------------------------------------------------------------
# Session state init
# ------------------------------------------------------------------
if "per_date_rows" not in st.session_state:
    st.session_state.per_date_rows = None  # dict[str, list[dict]]
if "loaded_key" not in st.session_state:
    st.session_state.loaded_key = None

# ------------------------------------------------------------------
# Step 1 : upload + période + chargement
# ------------------------------------------------------------------
uploaded_file = st.file_uploader("Fichier réservations (.xlsx ou .csv)", type=["xlsx", "xls", "csv"])

col1, col2, col3 = st.columns([1, 1, 2])
with col1:
    start_date = st.date_input("Date de début")
with col2:
    end_date = st.date_input("Date de fin")
with col3:
    pdf_title = st.text_input("Titre affiché sur le PDF", value="Desert Evasion")

load_clicked = st.button(
    "📂 Charger les transferts",
    type="primary",
    disabled=uploaded_file is None,
)

if load_clicked:
    if start_date > end_date:
        st.error("La date de début doit être avant (ou égale à) la date de fin.")
        st.stop()

    with st.spinner("Lecture du fichier..."):
        try:
            df = load_source_dataframe(uploaded_file)
        except Exception as e:
            st.error(f"Impossible de lire le fichier : {e}")
            st.stop()

    per_date_rows = {}
    d = start_date
    while d <= end_date:
        target_dt = datetime(d.year, d.month, d.day)
        day_rows = build_transfers_for_date(df, target_dt)
        per_date_rows[d.strftime("%d-%b-%y").upper()] = day_rows
        d += timedelta(days=1)

    st.session_state.per_date_rows = per_date_rows
    # unique key per (fichier, période) pour repartir à zéro si ça change
    st.session_state.loaded_key = f"{uploaded_file.name}-{start_date}-{end_date}"

# ------------------------------------------------------------------
# Step 2 : preview + édition (Driver / Car) + merge + PDF
# ------------------------------------------------------------------
if st.session_state.per_date_rows is not None:
    per_date_rows = st.session_state.per_date_rows
    days_with_rows = {d: r for d, r in per_date_rows.items() if r}

    total_groups = sum(len(r) for r in days_with_rows.values())
    total_travelers = sum(sum(row["Group Size"] for row in r) for r in days_with_rows.values())

    if total_groups == 0:
        st.warning("Aucun transfert confirmé trouvé sur cette période. Vérifie les dates et le statut des réservations.")
        st.stop()

    st.success(
        f"{total_groups} groupe(s) de transfert trouvé(s) "
        f"({total_travelers} voyageur(s) au total) sur la période choisie. "
        f"Modifie **Driver** et **Car** directement dans les tableaux ci-dessous."
    )

    display_cols = [
        "Pickup Date", "Pickup Time", "Flight Time", "Transfer Type",
        "Passengers", "Group Size", "Chauffeur",
        "Start", "Destination", "Flight No", "Txn ID",
    ]

    tabs = st.tabs(list(days_with_rows.keys()))
    edited_per_date = {}

    for tab, (date_label, rows) in zip(tabs, days_with_rows.items()):
        with tab:
            df_day = pd.DataFrame(rows)
            df_day["Chauffeur"] = df_day.apply(
                lambda r: pair_to_label(r.get("Driver", ""), r.get("Car", "")), axis=1
            )
            df_day = df_day[display_cols]

            column_config = {
                col: st.column_config.TextColumn(disabled=True)
                for col in display_cols if col != "Chauffeur"
            }
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

    gen_col, dl_col = st.columns([1, 3])
    with gen_col:
        generate_clicked = st.button("📄 Fusionner et générer le PDF", type="primary")

    if generate_clicked:
        all_rows = []
        for rows in edited_per_date.values():
            all_rows.extend(rows)

        missing = sum(1 for r in all_rows if not (r.get("Driver") or "").strip() or not (r.get("Car") or "").strip())
        if missing:
            st.info(f"{missing} transfert(s) sans chauffeur/voiture assigné — ils apparaîtront surlignés dans le PDF.")

        with st.spinner("Construction du PDF..."):
            pdf_buffer = build_pdf_from_rows(all_rows, title=pdf_title)

        with dl_col:
            st.download_button(
                "⬇️ Télécharger le PDF",
                data=pdf_buffer,
                file_name=f"transferts_{start_date.strftime('%d%b%y').upper()}_{end_date.strftime('%d%b%y').upper()}.pdf",
                mime="application/pdf",
            )

        # Bonus : zip des CSV édités (avec Driver/Car remplis), un par jour
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for date_label, rows in edited_per_date.items():
                if not rows:
                    continue
                csv_buffer = io.StringIO()
                import csv as csv_module
                writer = csv_module.DictWriter(csv_buffer, fieldnames=CSV_FIELDNAMES)
                writer.writeheader()
                writer.writerows(rows)
                zf.writestr(f"{date_label.replace('-', '')}.csv", csv_buffer.getvalue())
        zip_buffer.seek(0)

        st.download_button(
            "⬇️ Télécharger les CSV édités (un par jour, en zip)",
            data=zip_buffer,
            file_name=f"csv_{start_date.strftime('%d%b%y').upper()}_{end_date.strftime('%d%b%y').upper()}.zip",
            mime="application/zip",
        )
