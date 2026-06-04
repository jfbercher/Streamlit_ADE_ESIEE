#!/usr/bin/env python3
"""
ade_app.py - Interface Streamlit pour analyser les heures d'enseignement ADE.

Usage: streamlit run ade_app.py
Dépendances: streamlit, pandas, openpyxl (+ ade_heures.py dans le même dossier)
"""

import io
import os
import re
import tempfile
from collections import defaultdict

import pandas as pd
import streamlit as st

from ade_heures import (
    HETD_COEFFICIENTS_ESIEE,
    HETP_COEFFICIENTS,
    MODALITY_COLORS,
    MODALITY_ORDER,
    generate_excel,
    hetd,
    hetp,
    parse_ics,
    process_events,
    extract_codes,
)

# ---------------------------------------------------------------------------
# Config page
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="ADE Heures",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.title("Analyse des heures d'enseignement — ADE")
st.caption("Importez un fichier `.ics` exporté depuis ADE pour obtenir le détail et le récapitulatif de vos heures.")

st.error("🔴 **Version bêta** --- Cet outil est en cours de beta test. Les résultats n'ont pas encore été validés et une vérification de votre part est nécessaire. Merci de signaler toute anomalie sur [le lien suivant](https://docs.google.com/document/d/1QvYGU6BJAivPvYUNZ4nP_qpJm5ZgR8SAdQZUg8VvDrY/edit?usp=sharing).")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def records_to_df(records):
    """Convert list of record dicts to a clean DataFrame."""
    rows = []
    for r in records:
        rows.append({
            "Nom":        r["nom"],
            "Filière":    r["filiere"],
            "Promo":      r["promo"],
            "Date":       r["dtstart"].strftime("%Y-%m-%d"),
            "Début":      r["dtstart"].strftime("%H:%M"),
            "Fin":        r["dtend"].strftime("%H:%M"),
            "Durée (h)":  round(r["duration_h"], 2),
            "HETD (h)":   round(hetd(r["duration_h"], r["modality"]), 2),
            "HETP (h)":   round(hetp(r["duration_h"], r["modality"]), 2),
            "Lieu":       r["location"],
            "Modalité":   r["modality"],
            # keep raw dtstart for sorting
            "_dtstart":   r["dtstart"],
        })
    return pd.DataFrame(rows)


def build_modality_summary(df):
    """Build summary table grouped by modality."""
    grp = (
        df.groupby("Modalité", sort=False)
        .agg(
            Séances=("Nom", "count"),
            Heures=("Durée (h)", "sum"),
            HETD=("HETD (h)", "sum"),
            HETP=("HETP (h)", "sum"),
        )
        .reset_index()
    )
    grp["Coeff HETD"] = grp["Modalité"].apply(lambda m: HETD_COEFFICIENTS_ESIEE.get(m.replace('_Trou_ADE', ''), 0)).round(2)
    grp["Coeff HETP"] = grp["Modalité"].apply(lambda m: HETP_COEFFICIENTS.get(m.replace('_Trou_ADE', ''), 0)).round(2)
    grp["Heures"] = grp["Heures"].round(2)
    grp["HETD"]   = grp["HETD"].round(2)
    grp["HETP"]   = grp["HETP"].round(2)
    order = {m: i for i, m in enumerate(MODALITY_ORDER)}
    grp["_order"] = grp["Modalité"].map(order).fillna(999)
    grp = grp.sort_values("_order").drop(columns="_order").reset_index(drop=True)
    total = pd.DataFrame([{
        "Modalité":   "TOTAL",
        "Coeff HETD": "",
        "Coeff HETP": "",
        "Séances":    int(grp["Séances"].sum()),
        "Heures":     round(grp["Heures"].sum(), 2),
        "HETD":       round(grp["HETD"].sum(), 2),
        "HETP":       round(grp["HETP"].sum(), 2),
    }])
    return pd.concat([grp, total], ignore_index=True)[
        ["Modalité", "Coeff HETD", "Coeff HETP", "Séances", "Heures", "HETD", "HETP"]
    ]


def get_unique_filieres(df):
    """Return sorted list of unique filière codes (splitting multi-filière cells)."""
    filieres = set()
    for val in df["Filière"]:
        if val and val != "—":
            for f in val.split(" / "):
                f = f.strip()
                if f and f != "—":
                    filieres.add(f)
    return sorted(filieres)


def get_unique_promos(df):
    """Return sorted list of unique promo codes (splitting multi-promo cells)."""
    promos = set()
    for val in df["Promo"]:
        if val and val != "—":
            for p in val.split(" / "):
                p = p.strip()
                if p and p != "—":
                    promos.add(p)
    return sorted(promos)


def build_filiere_summary(df):
    """Build summary grouped by filière, exploding multi-filière events."""
    rows = []
    for _, row in df.iterrows():
        val = row["Filière"]
        filieres = [f.strip() for f in val.split(" / ")] if val and val != "—" else ["—"]
        for f in filieres:
            rows.append({**row.to_dict(), "_filiere_key": f})
    df_exp = pd.DataFrame(rows)
    def collect_promos(s):
        promos = set()
        for val in s:
            if val and val != "—":
                for p in val.split(" / "):
                    p = p.strip()
                    if p and p != "—":
                        promos.add(p)
        return " / ".join(sorted(promos)) if promos else "—"

    grp = (
        df_exp.groupby("_filiere_key", sort=False)
        .agg(
            Séances=("Nom", "count"),
            Heures=("Durée (h)", "sum"),
            HETD=("HETD (h)", "sum"),
            HETP=("HETP (h)", "sum"),
            Cours=("Nom", lambda s: len(s.unique())),
            Promos=("Promo", collect_promos),
        )
        .reset_index()
        .rename(columns={"_filiere_key": "Filière"})
    )
    grp["Heures"] = grp["Heures"].round(2)
    grp["HETD"]   = grp["HETD"].round(2)
    grp["HETP"]   = grp["HETP"].round(2)
    return grp.sort_values("Heures", ascending=False).reset_index(drop=True)


_COURSE_SUFFIX_RE = re.compile(r'\s+(TP|TDR?|C(OURS)?)\s*\d+$', re.IGNORECASE)

def normalize_course_name(name):
    """Supprime les suffixes de groupe en fin de nom (TP1, TP2, C1, C2, TD1…)."""
    return _COURSE_SUFFIX_RE.sub('', name).strip()


def build_course_summary(df):
    """Build summary table grouped by normalized course name."""
    df = df.copy()
    df["_nom_base"] = df["Nom"].apply(normalize_course_name)
    grp = (
        df.groupby("_nom_base", sort=False)
        .agg(
            Séances=("Nom", "count"),
            Heures=("Durée (h)", "sum"),
            HETD=("HETD (h)", "sum"),
            HETP=("HETP (h)", "sum"),
            Modalités=("Modalité", lambda s: ", ".join(sorted(s.unique()))),
        )
        .reset_index()
        .rename(columns={"_nom_base": "Nom"})
        .sort_values("Heures", ascending=False)
        .reset_index(drop=True)
    )
    grp["Heures"] = grp["Heures"].round(2)
    grp["HETD"]   = grp["HETD"].round(2)
    grp["HETP"]   = grp["HETP"].round(2)
    return grp


def make_excel_bytes(records):
    """Generate Excel file in memory and return bytes."""
    buf = io.BytesIO()
    generate_excel(records, buf)
    buf.seek(0)
    return buf.read()


# Column config partagé : forcer 2 décimales sur toutes les colonnes numériques
NUM_COL_CONFIG = {
    "Durée (h)":  st.column_config.NumberColumn(format="%.2f"),
    "HETD (h)":   st.column_config.NumberColumn(format="%.2f"),
    "HETP (h)":   st.column_config.NumberColumn(format="%.2f"),
    "Heures":     st.column_config.NumberColumn(format="%.2f"),
    "HETD":       st.column_config.NumberColumn(format="%.2f"),
    "HETP":       st.column_config.NumberColumn(format="%.2f"),
    "Coeff HETD": st.column_config.NumberColumn(format="%.2f"),
    "Coeff HETP": st.column_config.NumberColumn(format="%.2f"),
}

# Modality → hex color for pandas Styler (strip FF alpha prefix)
def _hex(rrggbbaa):
    return "#" + rrggbbaa[2:]

MOD_CSS = {m: _hex(c) for m, c in MODALITY_COLORS.items()}


def style_modality(df_display):
    """Apply row background based on Modalité column."""
    def row_style(row):
        color = MOD_CSS.get(row.get("Modalité", ""), "#ffffff")
        return [f"background-color: {color}" for _ in row]
    return df_display.style.apply(row_style, axis=1)


# ---------------------------------------------------------------------------
# File upload
# ---------------------------------------------------------------------------

uploaded = st.file_uploader(
    "Choisissez un fichier ADE (.ics)",
    type=["ics"],
    help="Exportez votre emploi du temps depuis ADE au format iCalendar (.ics)",
)

if uploaded is None:
    st.info("Importez un fichier `.ics` pour commencer.")
    st.stop()

# Parse — save to temp file so parse_ics can open it normally
with tempfile.NamedTemporaryFile(suffix=".ics", delete=False) as tmp:
    tmp.write(uploaded.read())
    tmp_path = tmp.name

try:
    raw_events = parse_ics(tmp_path)
    records    = process_events(raw_events)
finally:
    os.unlink(tmp_path)

if not records:
    st.error("Aucun événement valide trouvé dans ce fichier.")
    st.stop()

df = records_to_df(records)

# Réinitialiser les filtres si un nouveau fichier est uploadé
_file_id = uploaded.name + str(uploaded.size)
if st.session_state.get("_last_file_id") != _file_id:
    for key in ["filter_mod", "filter_fil_mod", "filter_promo_mod", "filter_date_mod", "select_course", "select_filiere"]:
        st.session_state.pop(key, None)
    st.session_state["_last_file_id"] = _file_id

# ---------------------------------------------------------------------------
# Top metrics
# ---------------------------------------------------------------------------

st.markdown("---")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Séances totales",  len(records))
c2.metric("Heures totales",   f"{df['Durée (h)'].sum():.2f} h")
c3.metric("HETD totales",     f"{df['HETD (h)'].sum():.2f}")
c4.metric("HETP totales",     f"{df['HETP (h)'].sum():.2f}")
c5.metric("Cours distincts",  df["Nom"].nunique())

st.markdown("---")

# ---------------------------------------------------------------------------
# Section 1 — Récapitulatif par modalité
# ---------------------------------------------------------------------------

st.header("📊 Récapitulatif par modalité")

summary_df = build_modality_summary(df)

col_table, col_chart = st.columns([1, 1], gap="large")

with col_table:
    st.dataframe(
        style_modality(summary_df.iloc[:-1]),
        use_container_width=True,
        hide_index=True,
        column_config=NUM_COL_CONFIG,
    )
    # Ligne TOTAL en gras sous le tableau
    total_row = summary_df.iloc[-1]
    st.markdown(
        f"**TOTAL — {int(total_row['Séances'])} séances | "
        f"{total_row['Heures']:.2f} h | "
        f"{total_row['HETD']:.2f} HETD | "
        f"{total_row['HETP']:.2f} HETP**"
    )

with col_chart:
    chart_data = summary_df.iloc[:-1].set_index("Modalité")[["HETD", "HETP"]]
    st.bar_chart(chart_data, use_container_width=True, stack=False)

# ---------------------------------------------------------------------------
# Section 2 — Téléchargement Excel
# ---------------------------------------------------------------------------

st.markdown("---")
st.header("⬇️ Télécharger le fichier Excel")

excel_bytes = make_excel_bytes(records)
default_name = uploaded.name.replace(".ics", "_heures.xlsx")

st.download_button(
    label="📥 Télécharger le fichier Excel complet",
    data=excel_bytes,
    file_name=default_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

# ---------------------------------------------------------------------------
# Section 3 — Exploration
# ---------------------------------------------------------------------------

st.markdown("---")
st.header("🔍 Explorer les séances")

tab_mod, tab_cours, tab_filiere = st.tabs(["Par modalité", "Par nom de cours", "Par filière"])

# ---- Tab : Par modalité ----
with tab_mod:
    all_mods = [m for m in MODALITY_ORDER if m in df["Modalité"].values]
    for m in df["Modalité"].unique():
        if m not in all_mods:
            all_mods.append(m)

    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        selected_mods = st.multiselect(
            "Filtrer par modalité",
            options=all_mods,
            default=all_mods,
            key="filter_mod",
        )
    with col_f2:
        all_fil = get_unique_filieres(df)
        selected_fil = st.multiselect(
            "Filtrer par filière",
            options=all_fil,
            default=[],
            placeholder="Toutes les filières",
            key="filter_fil_mod",
        )
    with col_f3:
        all_promos = get_unique_promos(df)
        selected_promos = st.multiselect(
            "Filtrer par promo",
            options=all_promos,
            default=[],
            placeholder="Toutes les promos",
            key="filter_promo_mod",
        )
    with col_f4:
        date_min = df["_dtstart"].min().date()
        date_max = df["_dtstart"].max().date()
        date_range = st.date_input(
            "Filtrer par période",
            value=(date_min, date_max),
            min_value=date_min,
            max_value=date_max,
            key="filter_date_mod",
        )

    df_mod = df[df["Modalité"].isin(selected_mods)].copy()
    if selected_fil:
        mask = df_mod["Filière"].apply(
            lambda v: any(f in (v or "").split(" / ") for f in selected_fil)
        )
        df_mod = df_mod[mask]
    if selected_promos:
        mask = df_mod["Promo"].apply(
            lambda v: any(p in (v or "").split(" / ") for p in selected_promos)
        )
        df_mod = df_mod[mask]
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        d_start, d_end = date_range
        df_mod = df_mod[
            (df_mod["_dtstart"].dt.date >= d_start) &
            (df_mod["_dtstart"].dt.date <= d_end)
        ]

    df_mod = df_mod.sort_values("_dtstart").drop(columns="_dtstart").reset_index(drop=True)

    st.caption(f"{len(df_mod)} séance(s) — {df_mod['Durée (h)'].sum():.2f} h — {df_mod['HETD (h)'].sum():.2f} HETD — {df_mod['HETP (h)'].sum():.2f} HETP")
    st.dataframe(
        style_modality(df_mod),
        use_container_width=True,
        hide_index=True,
        column_config=NUM_COL_CONFIG,
    )

# ---- Tab : Par nom de cours ----
with tab_cours:
    course_summary = build_course_summary(df)

    st.subheader("Résumé par cours")
    st.dataframe(course_summary, use_container_width=True, hide_index=True, column_config=NUM_COL_CONFIG)

    st.subheader("Détail d'un cours")
    course_names = course_summary["Nom"].tolist()
    selected_course = st.selectbox("Sélectionner un cours", options=course_names, key="select_course")

    if selected_course:
        df_course = (
            df[df["Nom"].apply(normalize_course_name) == selected_course]
            .sort_values("_dtstart")
            .drop(columns="_dtstart")
            .reset_index(drop=True)
        )
        n = len(df_course)
        h = df_course["Durée (h)"].sum()
        e = df_course["HETD (h)"].sum()
        p = df_course["HETP (h)"].sum()
        mods = ", ".join(sorted(df_course["Modalité"].unique()))

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Séances",   n)
        c2.metric("Heures",    f"{h:.2f} h")
        c3.metric("HETD",      f"{e:.2f}")
        c4.metric("HETP",      f"{p:.2f}")
        c5.metric("Modalités", mods)

        st.dataframe(
            style_modality(df_course),
            use_container_width=True,
            hide_index=True,
            column_config=NUM_COL_CONFIG,
        )

# ---- Tab : Par filière ----
with tab_filiere:
    fil_summary = build_filiere_summary(df)

    col_fs, col_fc = st.columns([1, 1], gap="large")

    with col_fs:
        st.subheader("Récapitulatif par filière")
        st.dataframe(fil_summary, use_container_width=True, hide_index=True, column_config=NUM_COL_CONFIG)

    with col_fc:
        chart_fil = fil_summary[fil_summary["Filière"] != "—"].set_index("Filière")[["HETD", "HETP"]]
        st.bar_chart(chart_fil, use_container_width=True, stack=False)

    st.subheader("Détail d'une filière")
    all_fil_options = fil_summary["Filière"].tolist()
    selected_filiere = st.selectbox("Sélectionner une filière", options=all_fil_options, key="select_filiere")

    if selected_filiere:
        if selected_filiere == "—":
            mask = df["Filière"] == "—"
        else:
            mask = df["Filière"].apply(lambda v: selected_filiere in (v or "").split(" / "))

        df_fil = df[mask].sort_values("_dtstart").drop(columns="_dtstart").reset_index(drop=True)

        cf1, cf2, cf3, cf4, cf5 = st.columns(5)
        cf1.metric("Séances",         len(df_fil))
        cf2.metric("Heures",          f"{df_fil['Durée (h)'].sum():.2f} h")
        cf3.metric("HETD",            f"{df_fil['HETD (h)'].sum():.2f}")
        cf4.metric("HETP",            f"{df_fil['HETP (h)'].sum():.2f}")
        cf5.metric("Cours distincts", df_fil["Nom"].nunique())

        # Sub-breakdown by promo if multiple promos present
        promos_in_fil = sorted(df_fil["Promo"].unique())
        if len(promos_in_fil) > 1:
            with st.expander("Répartition par promo", expanded=False):
                promo_grp = (
                    df_fil.groupby("Promo")
                    .agg(Séances=("Nom", "count"), Heures=("Durée (h)", "sum"), HETD=("HETD (h)", "sum"), HETP=("HETP (h)", "sum"))
                    .reset_index()
                    .sort_values("Heures", ascending=False)
                )
                promo_grp["Heures"] = promo_grp["Heures"].round(2)
                promo_grp["HETD"]   = promo_grp["HETD"].round(2)
                promo_grp["HETP"]   = promo_grp["HETP"].round(2)
                st.dataframe(promo_grp, use_container_width=True, hide_index=True, column_config=NUM_COL_CONFIG)

        st.dataframe(
            style_modality(df_fil),
            use_container_width=True,
            hide_index=True,
            column_config=NUM_COL_CONFIG,
        )
