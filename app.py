import io
import zipfile
from datetime import timedelta

import streamlit as st

from transfer_logic import (
    load_source_dataframe,
    build_transfers_for_date,
    build_pdf_from_rows,
    CSV_FIELDNAMES,
)
import csv

st.set_page_config(page_title="Fiches de transfert", page_icon="🚐", layout="centered")

st.title("🚐 Générateur de fiches de transfert")
st.caption("Desert Evasion — upload le fichier de réservations, choisis une période, récupère un PDF prêt à imprimer.")

uploaded_file = st.file_uploader("Fichier réservations (.xlsx ou .csv)", type=["xlsx", "xls", "csv"])

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Date de début")
with col2:
    end_date = st.date_input("Date de fin")

pdf_title = st.text_input("Titre affiché sur le PDF", value="Desert Evasion")

generate = st.button("Générer le PDF", type="primary", disabled=uploaded_file is None)

if generate:
    if start_date > end_date:
        st.error("La date de début doit être avant (ou égale à) la date de fin.")
        st.stop()

    with st.spinner("Lecture du fichier..."):
        try:
            df = load_source_dataframe(uploaded_file)
        except Exception as e:
            st.error(f"Impossible de lire le fichier : {e}")
            st.stop()

    all_rows = []
    per_date_rows = {}
    d = start_date
    while d <= end_date:
        target_dt = __import__("datetime").datetime(d.year, d.month, d.day)
        day_rows = build_transfers_for_date(df, target_dt)
        per_date_rows[d.strftime("%d-%b-%y").upper()] = day_rows
        all_rows.extend(day_rows)
        d += timedelta(days=1)

    total_groups = len(all_rows)
    total_travelers = sum(r.get("Group Size", 0) for r in all_rows)
    st.success(
        f"{total_groups} groupe(s) de transfert trouvé(s) "
        f"({total_travelers} voyageur(s) au total) sur la période choisie."
    )

    if total_groups == 0:
        st.warning("Aucun transfert confirmé trouvé sur cette période. Vérifie les dates et le statut des réservations.")
        st.stop()

    # Aperçu par jour
    for date_label, rows in per_date_rows.items():
        if rows:
            with st.expander(f"{date_label} — {len(rows)} transfert(s)"):
                st.dataframe(rows, use_container_width=True)

    # PDF
    with st.spinner("Construction du PDF..."):
        pdf_buffer = build_pdf_from_rows(all_rows, title=pdf_title)

    st.download_button(
        "⬇️ Télécharger le PDF",
        data=pdf_buffer,
        file_name=f"transferts_{start_date.strftime('%d%b%y').upper()}_{end_date.strftime('%d%b%y').upper()}.pdf",
        mime="application/pdf",
    )

    # Bonus : zip des CSV par jour, si jamais utile pour Driver/Car offline
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for date_label, rows in per_date_rows.items():
            if not rows:
                continue
            csv_buffer = io.StringIO()
            writer = csv.DictWriter(csv_buffer, fieldnames=CSV_FIELDNAMES)
            writer.writeheader()
            writer.writerows(rows)
            zf.writestr(f"{date_label.replace('-', '')}.csv", csv_buffer.getvalue())
    zip_buffer.seek(0)

    st.download_button(
        "⬇️ Télécharger les CSV (un par jour, en zip)",
        data=zip_buffer,
        file_name=f"csv_{start_date.strftime('%d%b%y').upper()}_{end_date.strftime('%d%b%y').upper()}.zip",
        mime="application/zip",
    )
