import streamlit as st
import io
import openpyxl
import pandas as pd
import time
import hashlib

from suppliers.fournisseur_abc import run_transform as run_abc

st.set_page_config(page_title="Générateur Shopify – Fichiers fournisseurs", layout="wide")

# --- CSS bouton (normal + hover) ---
st.markdown(
    """
    <style>
    div[data-testid="stButton"] > button {
        background: #ffffff !important;
        border: 1px solid #d6d6d9 !important;
        color: #2f5f8f !important;
        border-radius: 10px !important;
        padding: 0.55rem 1.1rem !important;
        font-weight: 500 !important;
        box-shadow: none !important;
    }
    div[data-testid="stButton"] > button:hover {
        background: #f0f2f6 !important;
        border: 1px solid #d6d6d9 !important;
        color: #2f5f8f !important;
    }
    div[data-testid="stButton"] > button:focus {
        box-shadow: none !important;
        outline: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Générateur de fichier Shopify")

SUPPLIERS = {
    "Balmoral": run_abc,
    "Bandit": run_abc,
    "Café du Cycliste": run_abc,
    "Ciele": run_abc,
    "District Vision": run_abc,
    "Fingerscrossed": run_abc,
    "MAAP": run_abc,
    "Pas Normal Studios": run_abc,
    "Tracksmith": run_abc,
}

st.markdown("### 1️⃣ Sélection du fournisseur")
supplier_name = st.selectbox("Choisir le fournisseur", list(SUPPLIERS.keys()))

st.markdown("### 2️⃣ Upload des fichiers")
supplier_file = st.file_uploader("Fichier fournisseur (.xlsx)", type=["xlsx"])
help_file = st.file_uploader("Help data (.xlsx)", type=["xlsx"])


# =========================
# 3️⃣ Tags
# =========================
event_promo_tag = ""
style_season_map = {}

def _first_existing_col(cols, candidates):
    cols_l = [c.lower() for c in cols]
    for c in candidates:
        if c.lower() in cols_l:
            return cols[cols_l.index(c.lower())]
    return None

def _extract_unique_style_rows(xlsx_bytes):
    """Extract unique styles for Seasonality tagging (per style).
    Displays (when available) in this order: Style Name, Style Number.
    """
    bio = io.BytesIO(xlsx_bytes)
    xls = pd.ExcelFile(bio)

    style_number_candidates = [
        "Style Number", "Style Num", "Style #", "style number", "style #", "Style",
    ]
    style_name_candidates = [
        "Style Name", "style name", "Product Name", "Name",
    ]

    rows = []
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
        except Exception:
            continue
        if df is None or df.empty:
            continue

        num_col = _first_existing_col(list(df.columns), style_number_candidates)
        name_col = _first_existing_col(list(df.columns), style_name_candidates)

        if not num_col and not name_col:
            continue

        data = {}
        if name_col:
            data["Style Name"] = df[name_col].astype(str).fillna("").map(lambda s: " ".join(str(s).strip().split()))
        if num_col:
            data["Style Number"] = df[num_col].astype(str).fillna("").map(lambda s: " ".join(str(s).strip().split()))

        tmp = pd.DataFrame(data)
        for c in tmp.columns:
            tmp = tmp[tmp[c].astype(str).str.strip().ne("").fillna(False)]

        if not tmp.empty:
            rows.append(tmp)

    if not rows:
        return None

    out = pd.concat(rows, ignore_index=True).drop_duplicates()

    cols = []
    if "Style Name" in out.columns:
        cols.append("Style Name")
    if "Style Number" in out.columns:
        cols.append("Style Number")

    return out[cols].reset_index(drop=True)

if supplier_file is not None:
    st.markdown("### 3️⃣ Tags")

    # Event/Promotion Related (applies to entire file)
    event_promo_tag = st.selectbox(
        "Event/Promotion Related",
        options=["", "spring-summer", "fall-winter"],
        index=0,
    )

    # Seasonality per style (table with free-text)
    style_rows_df = _extract_unique_style_rows(supplier_file.getvalue())
    if style_rows_df is not None and not style_rows_df.empty:
        st.markdown("#### Seasonality")

        key_col = "Style Number" if "Style Number" in style_rows_df.columns else "Style Name"
        # Stabilize ordering so the fingerprint doesn't change between reruns
        style_rows_df = style_rows_df.sort_values(by=key_col).reset_index(drop=True)

        supplier_fp = hashlib.md5(supplier_file.getvalue()).hexdigest()
        styles_fp = hashlib.md5("|".join(style_rows_df[key_col].astype(str).tolist()).encode("utf-8")).hexdigest()
        fp = f"{supplier_fp}:{styles_fp}:{key_col}"
        editor_key = f"seasonality_editor_{fp}"

        # Initialize ONLY when file/styles change (so typing doesn't get wiped)
        if st.session_state.get("seasonality_fp") != fp:
            st.session_state["seasonality_fp"] = fp

            prev = st.session_state.get("seasonality_df")
            prev_map = {}
            if prev is not None and key_col in prev.columns and "Seasonality Tags" in prev.columns:
                prev_map = {
                    str(k).strip(): str(v).strip()
                    for k, v in zip(prev[key_col].astype(str), prev["Seasonality Tags"].astype(str))
                    if str(k).strip()
                }

            init_df = style_rows_df.copy()
            init_df["Seasonality Tags"] = init_df[key_col].astype(str).map(lambda k: prev_map.get(str(k).strip(), ""))
            st.session_state["seasonality_df"] = init_df

        # IMPORTANT: always pass the same object from session_state to data_editor
        edited_df = st.data_editor(
            st.session_state["seasonality_df"],
            key=editor_key,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "Style Name": st.column_config.TextColumn(disabled=True),
                "Style Number": st.column_config.TextColumn(disabled=True),
                "Seasonality Tags": st.column_config.TextColumn(
                    help="Champ libre (ex: spring-summer, fall-winter, core, etc.)",
                    required=False,
                ),
            },
        )

        # Persist immediately so it sticks on the first write
        st.session_state["seasonality_df"] = edited_df

        style_season_map = {}
        for _, r in edited_df.iterrows():
            k = str(r.get(key_col, "")).strip()
            v = str(r.get("Seasonality Tags", "")).strip()
            if k and v:
                style_season_map[k] = v
    else:
        st.info("Aucun champ 'Style Name' ou 'Style Number' détecté dans le fichier. Seasonality ignorée.")

# 🔹 Projet pilote : pas de sélection de marque
brand_choice = ""

generate = st.button(
    "Générer le fichier Shopify",
    type="secondary",
    disabled=not (supplier_file and help_file),
)

if generate:
    st.markdown("### Génération en cours")
    status = st.empty()
    progress = st.progress(0)

    try:
        transform_fn = SUPPLIERS[supplier_name]

        status.info("Préparation des fichiers…")
        progress.progress(10)
        time.sleep(0.15)

        status.info("Lecture du fichier fournisseur…")
        progress.progress(25)
        time.sleep(0.15)

        status.info("Lecture du help data…")
        progress.progress(40)
        time.sleep(0.15)

        with st.spinner("Traitement en cours…"):
            output_bytes, warnings_df = transform_fn(
                supplier_xlsx_bytes=supplier_file.getvalue(),
                help_xlsx_bytes=help_file.getvalue(),
                vendor_name=supplier_name,
                brand_choice=brand_choice,  # toujours vide pour le pilote
                event_promo_tag=event_promo_tag,
                style_season_map=style_season_map,
            )

        status.info("Finalisation du fichier Shopify…")
        progress.progress(85)
        time.sleep(0.15)

        progress.progress(100)
        status.success("Fichier généré avec succès ✅")

        if warnings_df is not None and not warnings_df.empty:
            with st.expander("⚠️ Warnings détectés"):
                st.dataframe(warnings_df, use_container_width=True)

        st.download_button(
            label="⬇️ Télécharger output.xlsx",
            data=output_bytes,
            file_name=f"output_{supplier_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        progress.empty()
        status.error(f"Erreur lors de la génération : {e}")
