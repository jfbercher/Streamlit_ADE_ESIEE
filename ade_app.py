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
import requests
import time
from io import StringIO
from collections import defaultdict

import pandas as pd
import streamlit as st
from streamlit_local_storage import LocalStorage
storage = LocalStorage()


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
    get_teacher_name,
    compute_pdc_rea,
)

ANNEES = {'2025-2026':1, '2026-2027':2}
CURRENT_YEAR = "2025-2026"
statuts = {'80-20':400, '100-0':500, '60-40':300, 'MC UGE':288}
activites_non_planifiees = {"Décharge (HETP)":1, "Suivis stages E3/E4":2, "Suivis stages E5":5, "Suivis apprentis E3/E4FD":12,
                                        "Suivis apprentis E5":9, "Projet interne E4":35, "Projet interne E3":1,
                                        "Tremplin recherche":15, "Dépassement de forfait":1, "Autre (somme en HETP)":1}

if "edutime_csv_data" not in st.session_state:
    st.session_state.edutime_csv_data = None

# ---------------------------------------------------------------------------
# Session restore
# ---------------------------------------------------------------------------

old_setItem = storage.setItem
#local_storage.setItem = lambda key, value: old_setItem(key, json.dumps(value), key=key)
storage.setItem = lambda key, value: old_setItem(key, value, key=key+'_'+str(time.time()))
old_deleteItem = storage.deleteItem
storage.deleteItem = lambda key: old_deleteItem(key, key=key+'_'+str(time.time()))

if "loaded_ressource" not in st.session_state:
    st.session_state.loaded_ressource = 0

if "restored" not in st.session_state:

    st.session_state.RESSOURCE = storage.getItem("stored_RESSOURCE")
    if st.session_state.RESSOURCE is None:
        st.session_state.RESSOURCE = ""
    st.session_state.total_dech = storage.getItem("stored_total_dech")
    if st.session_state.total_dech is None:
        st.session_state.total_dech = 100.0
    st.session_state.select_totalhd =storage.getItem("stored_select_totalhd") 
    if st.session_state.select_totalhd is None:
        st.session_state.select_totalhd = "80-20"
    st.session_state.activities_to_remove = storage.getItem("stored_activities_to_remove")
    if st.session_state.activities_to_remove is None:
        st.session_state.activities_to_remove = []
    st.session_state.stored_selected_year = storage.getItem("stored_selected_year")
    if st.session_state.stored_selected_year is None:
        st.session_state.stored_selected_year = CURRENT_YEAR

    df_json = storage.getItem("stored_df_non_planifie") 
    if df_json: 
        st.session_state.df_non_planifie = pd.read_json(StringIO(df_json), orient="split")

    st.session_state.restored = True

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

st.caption("Utilisez votre numéro de ressource ADE ou importez un fichier `.ics` exporté depuis ADE pour obtenir le détail et le récapitulatif de vos heures.")

st.error("🔴 **Version bêta** --- Cet outil est en cours de beta test. Les résultats n'ont pas encore été validés et une vérification de votre part est nécessaire. Merci de signaler toute anomalie sur [le lien suivant](https://docs.google.com/document/d/1QvYGU6BJAivPvYUNZ4nP_qpJm5ZgR8SAdQZUg8VvDrY/edit?usp=sharing).")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

@st.cache_data()
def get_ressource_from_ADE(RESSOURCE, current_year=None):

    if current_year is None:
        current_year = CURRENT_YEAR

    ProjectId = ANNEES[current_year]
    StartYear, EndYear = current_year.split("-")

    ical_URL = f"https://edt-consult.univ-eiffel.fr/jsp/custom/modules/plannings/anonymous_cal.jsp?resources={RESSOURCE}&projectId={ProjectId}&calType=ical&startDay=01&startMonth=6&startYear={StartYear}&endDay=31&endMonth=08&endYear={EndYear}"
    
    response = requests.get(ical_URL)

    if response.status_code != 200:
        st.error("Erreur lors du téléchargement de l'agenda ADE : Erreur HTTP " + str(response.status_code) )
        if response.status_code == 500:
            st.error("Internal Server Error - Erreur interne du serveur ADE")
        if response.status_code == 404:
            st.error("Ressource introuvable")
        st.stop()
    else:
        st.session_state.loaded_ressource += 1
        #st.write("Ical chargé", st.session_state.loaded_ressource)
    
    return response.content

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


_COURSE_SUFFIX_RE = re.compile(r'\s+(TP|TDR?|TDRm?|TDm?|C(OURS)?)\s*\d+$', re.IGNORECASE)
_EP_RE = re.compile(r'\s*\(EP[^)]*\)', re.IGNORECASE)
_PROJET_INTERNE_RE = re.compile(r'Projet\s+interne\s+(E[34])', re.IGNORECASE)


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


# Column config partagé 
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

col_ressource, col_ressource_help, col_annee,col_ou, col_upload = st.columns([0.8, 0.4, 0.4, 0.2, 1], vertical_alignment="top")


with col_upload:
    uploaded = st.file_uploader(
        "Choisissez un fichier ADE (.ics)",
        type=["ics"],
        help="Exportez votre emploi du temps depuis ADE au format iCalendar (.ics)",
    )
    if uploaded: 
        st.session_state.RESSOURCE = ""

with col_annee:
    annee_courante = st.session_state.annee if "annee" in st.session_state else CURRENT_YEAR
    options = list(ANNEES.keys())
    st.selectbox("Choisissez une année scolaire", options=options, 
                 index=list(options).index(annee_courante), 
                 key="selected_year")

    storage.setItem( "stored_selected_year", st.session_state.selected_year) 

with col_ou:
    st.markdown("<div style='text-align:center'> <strong><br><br>OU</strong> </div>", unsafe_allow_html=True) 

with col_ressource_help:
    st.text('')
    st.text('')
    with st.popover("❓ Comment le trouver ?"):
        st.markdown("""
    Le numéro de ressource ADE est situé dans l'URL de votre emploi du temps exporté depuis ADE.
                    
    1. Connectez-vous sur [ADE](https://edt-consult.univ-eiffel.fr/direct/).
    2. Recherchez votre nom.
    4. Cliquez sur **Export Agenda**.
    4. Générez l'URL iCalendar.
    5. Copiez la valeur après `resources=`.
    """)

with col_ressource:
    value = storage.getItem("stored_RESSOURCE")
    RESSOURCE = st.text_input(
        "Entrez votre numéro de ressource ADE", # (voir le numéro dans l'URL générée via export agenda dans ADE)
        value = value if value else "",
    )
    if RESSOURCE:
        st.session_state.RESSOURCE = RESSOURCE
        storage.setItem("stored_RESSOURCE", RESSOURCE)

ical = None
# Déclenchement
if uploaded is None and not RESSOURCE:
    st.info("Entrez un numéro de ressource ADE **ou** importez un fichier `.ics`.")
    st.stop()

# si fichier uploadé
if uploaded is not None:
    ical = uploaded.getvalue()

# si téléchargement depuis ADE
elif RESSOURCE:

    ical = get_ressource_from_ADE(RESSOURCE, current_year=st.session_state.selected_year)

    uploaded = io.BytesIO(ical)
    uploaded.name = "edt.ics"
    uploaded.size = len(ical)


# Parse — save to temp file so parse_ics can open it normally
with tempfile.NamedTemporaryFile(suffix=".ics", delete=False) as tmp:
    tmp.write(uploaded.read())
    tmp_path = tmp.name

try:
    raw_events = parse_ics(tmp_path)
    records    = process_events(raw_events)
    teacher_name = get_teacher_name(records)
    st.session_state.teacher_name = teacher_name
    if len(teacher_name) > 0:
        st.markdown(
            f"<p style='margin-top:-50px;'>Nom de l'enseignant : {st.session_state.teacher_name.title()}</p>",
            unsafe_allow_html=True,
        )
        #st.write(f"Nom de l'enseignant : {st.session_state.teacher_name}")

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
        width='stretch',
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
    st.bar_chart(chart_data, width='stretch', stack=False)

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
    width='stretch',
)

# ---------------------------------------------------------------------------
# Section 3 — Exploration
# ---------------------------------------------------------------------------

st.markdown("---")
st.header("🔍 Explorer")

st.markdown("""
<style>
button[data-baseweb="tab"] {
    font-size: 1.4rem !important;
    font-weight: 700 !important;
    padding: 16px 32px !important;
    border-radius: 8px 8px 0 0 !important;
    background-color: #f0f2f6 !important;
    color: #444 !important;
    margin-right: 4px !important;
    border: 1px solid #d0d3da !important;
    border-bottom: none !important;
    line-height: 1.2 !important;
}
button[data-baseweb="tab"] p,
button[data-baseweb="tab"] span {
    font-size: 1.4rem !important;
    font-weight: 700 !important;
}
button[data-baseweb="tab"]:hover {
    background-color: #dde1ea !important;
    color: #111 !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    background-color: #ffffff !important;
    color: #1f77b4 !important;
    border-top: 3px solid #1f77b4 !important;
    border-left: 1px solid #d0d3da !important;
    border-right: 1px solid #d0d3da !important;
}
/* Sous-onglets (HETP/HETD...) : annuler les styles du 1er niveau */
div[role="tabpanel"] button[data-baseweb="tab"] p,
div[role="tabpanel"] button[data-baseweb="tab"] span {
    font-size: 0.9rem !important;
    font-weight: 500 !important;
}
div[role="tabpanel"] button[data-baseweb="tab"] {
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    padding: 6px 12px !important;
    border-radius: 0 !important;
    background-color: transparent !important;
    color: inherit !important;
    margin-right: 0 !important;
    border: none !important;
    border-bottom: 2px solid transparent !important;
}
div[role="tabpanel"] button[data-baseweb="tab"]:hover {
    background-color: transparent !important;
    color: inherit !important;
}
div[role="tabpanel"] button[data-baseweb="tab"][aria-selected="true"] {
    background-color: transparent !important;
    color: #1f77b4 !important;
    border-top: none !important;
    border-left: none !important;
    border-right: none !important;
    border-bottom: 2px solid #1f77b4 !important;
}
</style>
""", unsafe_allow_html=True)

tab_mod, tab_cours, tab_filiere, tab_pdc, tab_edutime = st.tabs(["Par modalité", "Par nom de cours", "Par filière", "Votre PdC réalisé", "Edutime"])

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
        width='stretch',
        hide_index=True,
        column_config=NUM_COL_CONFIG,
    )

# ---- Tab : Par nom de cours ----
with tab_cours:
    course_summary = build_course_summary(df)

    st.subheader("Résumé par cours")
    st.dataframe(course_summary, width='stretch', hide_index=True, column_config=NUM_COL_CONFIG)

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
            width='stretch',
            hide_index=True,
            column_config=NUM_COL_CONFIG,
        )

# ---- Tab : Par filière ----
with tab_filiere:
    fil_summary = build_filiere_summary(df)

    col_fs, col_fc = st.columns([1, 1], gap="large")

    with col_fs:
        st.subheader("Récapitulatif par filière")
        st.dataframe(fil_summary, width='stretch', hide_index=True, column_config=NUM_COL_CONFIG)

    with col_fc:
        chart_fil = fil_summary[fil_summary["Filière"] != "—"].set_index("Filière")[["HETD", "HETP"]]
        st.bar_chart(chart_fil, width='stretch', stack=False)

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
                st.dataframe(promo_grp, width='stretch', hide_index=True, column_config=NUM_COL_CONFIG)

        st.dataframe(
            style_modality(df_fil),
            width='stretch',
            hide_index=True,
            column_config=NUM_COL_CONFIG,
        )

# ---- Tab : Par nom de cours ----
with tab_pdc:
    df_pdc = df.copy()
    df_pdc["Nom"] = df_pdc["Nom"].apply(normalize_course_name)
    if "activities_to_remove" in st.session_state:
        activities_to_remove = list(set(st.session_state.activities_to_remove))
        for activity in activities_to_remove:
            if len(df_pdc["Nom"].str.contains(activity, na=False, regex=False)) == 0:
                st.warning(f"L'activité ''{activity}' n'est pas présent dans les activités importées d'ADE.")
            df_pdc = df_pdc[~df_pdc["Nom"].str.contains(activity, na=False, regex=False)]

    #st.dataframe(df_pdc, width='stretch', hide_index=True, column_config=NUM_COL_CONFIG)
    df_pdc = df_pdc.rename(columns={
    "HETP (h)": "HETP",
    "HETD (h)": "HETD",
    "Nom": "Cours",
    "Modalité": "Activité"
    })

    # Plan de charge attendu
    st.subheader("1) Plan de charge attendu")
    col1, col2, col3 = st.columns(3)
    with col1:
        statut_courant = st.session_state.select_totalhd

        statut = st.selectbox(
           "Total des heures dues au statut (en HETP)", 
           options=statuts, 
           index=list(statuts.keys()).index(statut_courant),
           key="select_totalhd")
        
        total_hd = float(statuts[st.session_state.select_totalhd])
        storage.setItem( "stored_select_totalhd", st.session_state.select_totalhd) 
        
    with col2:
        total_dech = st.number_input(
            "Total des décharges (en HETP)",
            value=float(st.session_state.total_dech),
            step=0.5,
            format="%.2f",
            #key="total_dech"
        )
        storage.setItem( "stored_total_dech", total_dech)
        st.session_state.total_dech = total_dech
        if "df_non_planifie" in st.session_state:
            st.session_state.df_non_planifie.loc["Décharge (HETP)", "Quantité"] = total_dech

    total_htodo = total_hd - total_dech
    with st.container(border=True):
        st.markdown("""- Total des heures attendues au statut (en HETP) : {total_hd}
- Total des décharges (en HETP) : {total_dech}
- Total à effectuer après décharges (en HETP) : {total_htodo}""".format(total_hd=total_hd, total_dech=total_dech, total_htodo=total_htodo), unsafe_allow_html=True)

    st.markdown("---")


    # ---- Activités non planifiées ----
    st.subheader("2) Activités non planifiées")
    # -----------
    st.markdown("""Si vous le souhaitez, entrez ici vos activités non planifiées, non récupérables à partir de ADE.
                    
⚠️ La colonne Quantité attend un nombre (nombre de suivis...) ; certaines activités peuvent être fractionnées (par exemple Projet E4 avec plusieurs suiveurs.)""")
  
    with st.expander("Activités non planifiées", expanded=False):
        if "df_non_planifie" not in st.session_state:
            st.session_state.df_non_planifie = pd.DataFrame( index=list(activites_non_planifiees.keys()), 
                                            columns=["Quantité", "Tarif/unité", "HETP"])
            st.session_state.df_non_planifie["Quantité"] = 0.0
            st.session_state.df_non_planifie["Tarif/unité"] = activites_non_planifiees.values()
            st.session_state.df_non_planifie["HETP"] = 0

        #st.session_state.df_non_planifie["Quantité"] = st.session_state.df_non_planifie["Quantité"].astype(float)
        edited = st.data_editor(
            st.session_state.df_non_planifie,
            key="non_planifie",
            column_config={
                "Quantité": st.column_config.NumberColumn(
                                min_value=0.0,
                                format="%.2f"),
                "Tarif/unité": st.column_config.NumberColumn(disabled=True),
                "HETP": st.column_config.NumberColumn(disabled=True)
            }
        )

        st.markdown(f"**Total des heures non planifiées (en HETP)** : {edited['HETP'].sum():.2f}")

        st.info("**Décharges**: Entrez directement la somme des HETP correspondantes. \n\n**Projets E3**: Entrez directement la somme des HETP correspondantes, qui dépend du nombre de projets suivis et dans chacun du nombre d'élèves. Si plusieurs suiveurs, ajustez en fonction des prorata de suivis. \n\n" \
        "👉🏼 Formule: Par suivi N_HETP = 8 + 0.5\*NbreSemaines\*NbreEtudiants (en 2025-26, NbreSemaines=7)")

        edited["HETP"] = edited["Quantité"]*edited["Tarif/unité"]

        # Reset
        if st.button("Reset", help="Effacer le tableau des heures non planifiées"):
            del st.session_state.df_non_planifie
            st.rerun()

        # Sauvegarde et rerun
        if not edited.equals(st.session_state.df_non_planifie):
            st.session_state.df_non_planifie = edited
            #storage.deleteItem( "stored_df_non_planifie" )
            storage.setItem( "stored_df_non_planifie", st.session_state.df_non_planifie.to_json( orient="split" ), )
            time.sleep(1.5)
            st.rerun()

    # -----------

    st.subheader("3) Plans de charge réalisés HETP ou HETD")
    tab_hetp, tab_hetd = st.tabs(["HETP", "HETD"])
    df_hetp = None

    with tab_hetp:
        st.subheader("PdC HETP")
        try:
            df_hetp = compute_pdc_rea(df_pdc, "HETP")
        except KeyError:
            st.error("La colonne 'HETP' est introuvable dans le fichier fourni.")
        df_hetp['Quantité'] = 0
        df_hetp['Total (HETP)'] = df_hetp.sum(axis=1)
        idx = df_hetp.index.tolist()
        for index, row in st.session_state.df_non_planifie.iterrows():
            if row['HETP'] != 0:
                L = len(df_hetp)
                df_hetp.loc[L, 'Quantité'] = row['Quantité']
                df_hetp.loc[L, 'Total (HETP)'] = row['HETP']
                idx.append(index)
        df_hetp.index = idx
        df_hetp = df_hetp.fillna(0)  
        ligne_total = df_hetp.sum(axis=0)
        df_hetp.loc[len(df_hetp)] = ligne_total
        idx.append('Total')
        df_hetp.index = idx
        #st.dataframe(df_hetp, width='stretch', height='content')
        event = st.dataframe(
            df_hetp,
            selection_mode="multi-row",
            on_select="rerun",
            width='stretch', 
            height='content',
            #hide_index=True,
        )
        if event.selection.rows:
            activities_to_remove = list(df_hetp.index[event.selection.rows]) 
            st.session_state.activities_to_remove.extend(activities_to_remove)
            st.write(f"Activities to remove: {activities_to_remove}")
            for activity in activities_to_remove:
                if sum(df_pdc["Cours"].str.contains(activity, na=False, regex=False)) == 0:
                    st.warning(f"⚠️ L'activité ''{activity}' n'est pas présente dans les activités importées d'ADE (non planifiée ?)")

        b1, b2, b3 = st.columns([1,1,4])
        with b1:
            if st.button("Supprimer les lignes sélectionnées", help="""Certaines activités peuvent devoir être neutralisées. 
Par exemple des actvités réalisées dans l'UGE mais Hors-ESIEE, et qui ne concernent donc pas le PdC réalisé ESIEE."""):
                st.rerun()

        with b2:
            if st.button("Restaurer toutes les lignes"):
                st.session_state.activities_to_remove = []
                st.rerun()

    if df_hetp is not None:
        df_hetd = df_hetp.apply(lambda x: x*2/3).rename(index={'Total (HETP)': 'Total (HETD)'}, 
                                                    columns={'Total (HETP)': 'Total (HETD)'})
    else: # Ceci ne devrait pas arriver, mais juste au cas où... 
        st.error("Visualiser le PdC HETP avant de visualiser le PdC HETD.")
            
    with tab_hetd:
        st.subheader("PdC HETD")
        try:
            #df_hetd = compute_pdc_rea(df_pdc, "HETD")
            st.dataframe(df_hetd, width='stretch', height='content')
        except KeyError:
            st.error("La colonne 'HETD' est introuvable dans le fichier fourni.")


    total_hreal_hetp = df_hetp.loc['Total', 'Total (HETP)']
    total_hreal_hetd = df_hetd.loc['Total', 'Total (HETD)']
    st.write("Total des heures réalisées : ", round(total_hreal_hetp,2), ' HETP, soit ', round(total_hreal_hetd,2), 'HETD.')

    # Calcul des heures complémentaires
    total_hcomp_hetp =  total_hreal_hetp - total_hd
    total_hcomp_hetd = total_hcomp_hetp * 2 / 3


    if total_hcomp_hetp > 0:
        # Rémunération attendue
        TARIF_HETP = 39.01
        TARIF_HETD_UNIV = 43.50
        TARIF_HETP_UNIV = TARIF_HETD_UNIV / 1.5
        remu = min(200, total_hcomp_hetp) * TARIF_HETP + \
                    max(0, total_hcomp_hetp - 200) * TARIF_HETP_UNIV


            
    # --- Extraction des totaux ---
    total_hreal_hetp = df_hetp.loc['Total', 'Total (HETP)']
    total_hreal_hetd = df_hetd.loc['Total', 'Total (HETD)'] 

    total_hcomp_hetp = total_hreal_hetp - total_hd
    total_hcomp_hetd = total_hcomp_hetp * 2 / 3

    st.markdown("#### Synthèse des Heures")

    st.markdown(
"""
<style>
/* Taille du titre (Label) */
[data-testid="stMetricLabel"] {
    font-size: 14px !important;
}

/* Taille de la valeur principale */
[data-testid="stMetricValue"] {
    font-size: 24px !important; /* Par défaut ~40px */
}

/* Taille du delta (sous-valeur en HETD) */
[data-testid="stMetricDelta"] {
    font-size: 18px !important;
}
</style>
""",
unsafe_allow_html=True
)

    # Affichage des métriques 
    col_real, col_comp = st.columns(2)

    with col_real:
        st.metric(
            label="⏳ Heures Réalisées", 
            value=f"{total_hreal_hetp:.2f} HETP", 
            delta=f"{total_hreal_hetd:.2f} HETD", 
            delta_color="off" # "off" pour juste afficher la valeur HETD en gris sans flèche
        )

    with col_comp:
        # On colore la métrique en vert si positif, rouge/gris si négatif
        color_inverse = "normal" if total_hcomp_hetp >= 0 else "inverse"
        st.metric(
            label="➕ Heures Complémentaires", 
            value=f"{total_hcomp_hetp:.2f} HETP", 
            delta=f"{total_hcomp_hetd:.2f} HETD",
            delta_color=color_inverse
        )

    # --- Section Rémunération ---
    if total_hcomp_hetp > 0:
        TARIF_HETP = 39.01
        TARIF_HETD_UNIV = 43.50
        TARIF_HETP_UNIV = TARIF_HETD_UNIV / 1.5
        
        remu = (min(200, total_hcomp_hetp) * TARIF_HETP + 
                max(0, total_hcomp_hetp - 200) * TARIF_HETP_UNIV)
        
        remu_HETD = remu / TARIF_HETD_UNIV

        st.markdown("#### Rémunération")
        
        # Un encadré vert (st.success) pour valoriser le gain
        st.success(f"""
        **Rémunération attendue  : {remu:.2f} €**, *soit l'équivalent de **{remu_HETD:.2f} HETD UGE** (sur la base de {TARIF_HETD_UNIV:.2f}€ / heure de TD universitaire).*
        """)

# ---- Tab : Edutime ----
with tab_edutime:
    st.info('💡 Dans Edutime : "Mes Services" → "Réalisé" → "Cette année universitaire" → "Appliquer les filtres" → "Exporter". Utiliser le fichier CSV résultant.')

    uploaded_edu = st.file_uploader(
        "Choisissez votre fichier CSV Edutime",
        type=["csv"],
        key="edutime_csv",
    )
    if uploaded_edu is not None:
        st.session_state["edutime_csv_data"] = uploaded_edu.getvalue()
        st.session_state["edutime_csv_name"] = uploaded_edu.name

    if st.session_state.edutime_csv_data is not None:
        try:
            df_edu = pd.read_csv(io.BytesIO(st.session_state.edutime_csv_data), sep=";", decimal=',')

            # Nettoyer Cours : supprimer (EP...) 
            # Sauvegarder quelles lignes étaient des vrais cours planifiés (avaient EP..)
            has_ep_mask = df_edu['Cours'].apply(
                lambda v: bool(_EP_RE.search(str(v))) if pd.notna(v) else False
            )
            df_edu['Cours'] = df_edu['Cours'].apply(
                lambda v: _EP_RE.sub('', str(v)).strip() if pd.notna(v) and str(v).strip() != '' else None
            )
            mask_empty = df_edu['Cours'].isna() | (df_edu['Cours'] == '')
            df_edu.loc[mask_empty, 'Cours'] = df_edu.loc[mask_empty, 'Activité']

            # Statut / décharges — pré-remplis depuis session_state (PdC réalisé)
            statut_courant = st.session_state.get('select_totalhd', '80-20')
            col_edu1, col_edu2, col_edu3 = st.columns(3)
            with col_edu1:
                statut_edu = st.selectbox(
                    "Total des heures dues au statut (en HETP)",
                    options=statuts,
                    index=list(statuts.keys()).index(statut_courant),
                    key="select_totalhd_edu",
                )
                total_hd_edu = float(statuts[statut_edu])

            
            with col_edu2:
                st.metric("Heures attendues au statut (HETP)", f"{total_hd_edu:.1f}")

            tab_hetp_edu, tab_hetd_edu = st.tabs(["HETP", "HETD"])

            with tab_hetp_edu:
                st.subheader("PdC HETP (Edutime)")
                try:
                    df_hetp_edu = compute_pdc_rea(df_edu, "HETP")
                except KeyError:
                    st.error("La colonne 'HETP' est introuvable dans le fichier fourni.")
                
                df_hetp_edu['Total (HETP)'] = df_hetp_edu.sum(axis=1)
                df_hetp_edu = df_hetp_edu.fillna(0)
                idx_edu = df_hetp_edu.index.tolist()
                if "Décharge (HETP)" in df_hetp_edu.columns:  # Suppression colonne disgracieuse Total conservé
                    df_hetp_edu.drop(columns=['Décharge (HETP)'], inplace=True)
                ligne_total_edu = df_hetp_edu.sum(axis=0)
                df_hetp_edu.loc[len(df_hetp_edu)] = ligne_total_edu
                idx_edu.append('Total')
                df_hetp_edu.index = idx_edu
                st.dataframe(df_hetp_edu, width='stretch')


            df_hetd_edu = df_hetp_edu.apply(lambda x: x * 2 / 3).rename(
                columns={'Total (HETP)': 'Total (HETD)'}
            )

            with tab_hetd_edu:
                st.subheader("PdC HETD (Edutime)")
                st.dataframe(df_hetd_edu, width='stretch')

            total_hreal_hetp_edu = df_hetp_edu.loc['Total', 'Total (HETP)']
            total_hreal_hetd_edu = df_hetd_edu.loc['Total', 'Total (HETD)']
            total_hcomp_hetp_edu = total_hreal_hetp_edu - total_hd_edu
            total_hcomp_hetd_edu = total_hcomp_hetp_edu * 2 / 3

            st.write("Total des heures réalisées :", round(total_hreal_hetp_edu, 2), "HETP, soit", round(total_hreal_hetd_edu, 2), "HETD.")

            st.markdown("#### Synthèse des Heures")

            col_real_edu, col_comp_edu = st.columns(2)
            with col_real_edu:
                st.metric(
                    label="⏳ Heures Réalisées",
                    value=f"{total_hreal_hetp_edu:.2f} HETP",
                    delta=f"{total_hreal_hetd_edu:.2f} HETD",
                    delta_color="off",
                )
            with col_comp_edu:
                color_inv_edu = "normal" if total_hcomp_hetp_edu >= 0 else "inverse"
                st.metric(
                    label="➕ Heures Complémentaires",
                    value=f"{total_hcomp_hetp_edu:.2f} HETP",
                    delta=f"{total_hcomp_hetd_edu:.2f} HETD",
                    delta_color=color_inv_edu,
                )

            if total_hcomp_hetp_edu > 0:
                TARIF_HETP = 39.01
                TARIF_HETD_UNIV = 43.50
                TARIF_HETP_UNIV = TARIF_HETD_UNIV / 1.5
                remu_edu = (min(200, total_hcomp_hetp_edu) * TARIF_HETP +
                            max(0, total_hcomp_hetp_edu - 200) * TARIF_HETP_UNIV)
                remu_HETD_edu = remu_edu / TARIF_HETD_UNIV
                st.markdown("#### Rémunération")
                st.success(f"""
                **Rémunération attendue : {remu_edu:.2f} €**, *soit l'équivalent de **{remu_HETD_edu:.2f} HETD UGE** (sur la base de {TARIF_HETD_UNIV:.2f}€ / heure de TD universitaire).*
                """)

            # ---- Comparaison ADE ↔ Edutime ----
            st.markdown("---")
            st.subheader("🔍 Comparaison (en HETP) ADE ↔ Edutime")

            ade_by_cours = (
                df.assign(Cours=df['Nom'].apply(normalize_course_name))
                .groupby('Cours')['HETP (h)'].sum()
                .round(2)
            )
            ade_ics_cours = set(ade_by_cours.index)  # noms cours ADE ICS avant ajout non-planifiés
            for index, row in st.session_state.df_non_planifie.iterrows():
                if row['HETP'] != 0:
                    ade_by_cours.loc[index] = row['HETP']

            if "activities_to_remove" in st.session_state:
                activities_to_remove = list(set(st.session_state.activities_to_remove))
                for activity in activities_to_remove:
                    if activity not in ade_by_cours.index:
                        st.warning(f"L'activité ''{activity}' n'est pas présent dans les activités importées d'ADE.")
                    ade_by_cours = ade_by_cours.drop(activity, errors="ignore")

            # Edutime : cours EP + Projet E3/E4 normalisés → "Projet interne E3/E4"
            _PROJET_COURT_RE = re.compile(r'^Projet\s+(E[34])$', re.IGNORECASE)
            is_projet_mask = df_edu['Cours'].str.contains(r'Projet.*E[34]', case=False, na=False, regex=True)
            # noms de cours planifiés (EP) pour le coloriage
            cours_planifies = ade_ics_cours | set(df_edu[has_ep_mask]['Cours'].unique())
            edu_by_cours = (
                df_edu[has_ep_mask | is_projet_mask | mask_empty]
                .assign(Cours=lambda d: d['Cours'].apply(
                    lambda v: _PROJET_COURT_RE.sub(r'Projet interne \1', str(v))
                ))
                .groupby('Cours')['HETP'].sum()
                .round(2)
            )

            comp = pd.concat(
                {"ADE": ade_by_cours, "Edutime": edu_by_cours}, axis=1
            ).fillna(0)
            comp['Δ (ADE−Edu)'] = (comp['ADE'] - comp['Edutime']).round(2)
            comp = comp[(comp['ADE'] != 0) | (comp['Edutime'] != 0)]
            comp = comp.sort_index()

            comp_cours = comp[comp.index.isin(cours_planifies)]
            comp_autres = comp[~comp.index.isin(cours_planifies)]

            COL_CFG = {
                'ADE':          st.column_config.NumberColumn(format="%.2f"),
                'Edutime':      st.column_config.NumberColumn(format="%.2f"),
                'Δ (ADE−Edu)': st.column_config.NumberColumn(format="%.2f"),
            }

            def _style_section(df_s, bg):
                styles = pd.DataFrame(f'background-color: {bg}', index=df_s.index, columns=df_s.columns)
                mask_delta = df_s['Δ (ADE−Edu)'].abs() > 0.01
                styles.loc[mask_delta, 'Δ (ADE−Edu)'] = f'background-color: {bg}; color: #c0392b; font-weight: bold'
                total_mask = df_s.index == 'Total'
                if total_mask.any():
                    styles.loc[total_mask] = 'background-color: #f0f0f0; font-weight: bold'
                return styles

            def _add_total(df_s):
                total = pd.DataFrame({
                    'ADE': [round(df_s['ADE'].sum(), 2)],
                    'Edutime': [round(df_s['Edutime'].sum(), 2)],
                    'Δ (ADE−Edu)': [round(df_s['Δ (ADE−Edu)'].sum(), 2)],
                }, index=['Total'])
                return pd.concat([df_s, total])

            st.markdown("**📘 Cours planifiés**")
            df_c = _add_total(comp_cours)
            st.dataframe(
                df_c.style.apply(lambda d: _style_section(d, '#e8f4fd'), axis=None),
                width='stretch', height='content',
                column_config=COL_CFG,
            )

            st.markdown("**📙 Activités non planifiées**")
            df_a = _add_total(comp_autres)
            st.dataframe(
                df_a.style.apply(lambda d: _style_section(d, '#fff8e1'), axis=None),
                width='stretch', height='content',
                column_config=COL_CFG,
            )

            # Total global
            total_global = pd.DataFrame({
                'ADE': [round(comp['ADE'].sum(), 2)],
                'Edutime': [round(comp['Edutime'].sum(), 2)],
                'Δ (ADE−Edu)': [round(comp['Δ (ADE−Edu)'].sum(), 2)],
            }, index=['Total général'])
            st.dataframe(
                total_global.style.apply(lambda d: _style_section(d, '#f0f0f0'), axis=None),
                width='stretch', height='content',
                column_config=COL_CFG,
            )

        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier : {e}")

