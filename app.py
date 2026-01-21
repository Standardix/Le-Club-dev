import streamlit as st
import io
import openpyxl
import time
import pandas as pd

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

st.markdown("### 3️⃣ Saisonality (par style)")

SEASONALITY_OPTIONS = [
    "",  # allow blank
    "spring-summer",
    "fall-winter",
    "all-season",
    "core",
    "seasonal",
]


def _first_existing_col(cols: list[str], candidates: list[str]) -> str | None:
    cols_l = {c.lower(): c for c in cols}
    for c in candidates:
        if c.lower() in cols_l:
            return cols_l[c.lower()]
    return None




def _extract_unique_style_rows(xlsx_bytes: bytes) -> pd.DataFrame | None:
    """Extract unique styles for Seasonality table.

    Prefers to display columns in this order:
      1) Style Name
      2) Style Number

    Key logic:
    - If both columns exist, returns both.
    - If only one exists, returns the one found.
    - If none exist, returns None.
    """
    bio = io.BytesIO(xlsx_bytes)
    xls = pd.ExcelFile(bio)

    style_number_candidates = [
        "Style Number", "Style Num", "Style", "style", "style number", "style #", "Style #",
    ]
    style_name_candidates = [
        "Style Name", "style name", "Name", "Product Name", "Product", "Style",
    ]

    rows: list[pd.DataFrame] = []

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

        out_cols = {}
        if name_col:
            out_cols["Style Name"] = (
                df[name_col].astype(str).fillna("").map(lambda s: " ".join(str(s).strip().split()))
            )
        if num_col:
            out_cols["Style Number"] = (
                df[num_col].astype(str).fillna("").map(lambda s: " ".join(str(s).strip().split()))
            )

        tmp = pd.DataFrame(out_cols)

        # drop blanks
        for c in list(tmp.columns):
            tmp = tmp[tmp[c].astype(str).str.strip().ne("").fillna(False)]

        rows.append(tmp)

    if not rows:
        return None

    out = pd.concat(rows, ignore_index=True).drop_duplicates()

    # enforce display order
    cols = []
    if "Style Name" in out.columns:
        cols.append("Style Name")
    if "Style Number" in out.columns:
        cols.append("Style Number")

    out = out[cols].copy()
    return out


def _extract_unique_styles(xlsx_bytes: bytes) -> list[str]:
    """Return sorted unique style identifiers from the supplier file (all valid sheets)."""
    bio = io.BytesIO(xlsx_bytes)
    xls = pd.ExcelFile(bio)

    style_candidates = [
        "Style Number",
        "Style",
        "style",
        "Style Num",
        "Style Name",
        "style name",
        "Style Number ",
    ]

    out: set[str] = set()
    for sn in xls.sheet_names:
        df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sn, dtype=str)
        if df is None or df.empty:
            continue
        df = df.dropna(how="all")
        if df.empty:
            continue

        scol = _first_existing_col(df.columns.tolist(), style_candidates)
        if scol is None:
            continue

        vals = (
            df[scol]
            .astype(str)
            .fillna("")
            .map(lambda s: " ".join(str(s).strip().split()))
        )
        for v in vals.tolist():
            if v and v.lower() != "nan":
                out.add(v)

    return sorted(out)


style_season_map: dict[str, str] = {}

if supplier_file is not None:
    try:
        style_rows_df = _extract_unique_style_rows(supplier_file.getvalue())

        if style_rows_df is not None and not style_rows_df.empty:
            st.caption("Le programme ajoutera ce tag directement dans la colonne **Tags** pour toutes les lignes ayant le même style.")

            # Use Style Number as key when available; fallback to Style Name
            key_col = "Style Number" if "Style Number" in style_rows_df.columns else "Style Name"

            # Initialize / refresh session table while preserving existing inputs when possible
            if "seasonality_df" not in st.session_state:
                seasonality_df_init = style_rows_df.copy()
                seasonality_df_init["Saisonality tag"] = ""
                st.session_state["seasonality_df"] = seasonality_df_init
            else:
                existing = st.session_state["seasonality_df"]
                existing_map = {}
                if key_col in existing.columns and "Saisonality tag" in existing.columns:
                    existing_map = {
                        str(k).strip(): str(v).strip()
                        for k, v in zip(existing[key_col].astype(str), existing["Saisonality tag"].astype(str))
                        if str(k).strip()
                    }

                seasonality_df_init = style_rows_df.copy()
                seasonality_df_init["Saisonality tag"] = seasonality_df_init[key_col].astype(str).map(
                    lambda k: existing_map.get(str(k).strip(), "")
                )
                st.session_state["seasonality_df"] = seasonality_df_init

            # Editable table (free-text)
            seasonality_df = st.data_editor(
                st.session_state["seasonality_df"],
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                column_config={
                    "Style Name": st.column_config.TextColumn(disabled=True),
                    "Style Number": st.column_config.TextColumn(disabled=True),
                    "Saisonality tag": st.column_config.TextColumn(
                        help="Champ libre : ex. spring-summer, fall, core, etc.",
                        required=False,
                    ),
                },
            )
            st.session_state["seasonality_df"] = seasonality_df

            # Build map {style_key: seasonality_tag}
            style_season_map = {}
            for _, r in seasonality_df.iterrows():
                k = str(r.get(key_col, "")).strip()
                v = str(r.get("Saisonality tag", "")).strip()
                if k and v:
                    style_season_map[k] = v
        else:
            st.info("Aucun champ 'Style Name' ou 'Style Number' détecté dans le fichier. La saisonalité par style sera ignorée.")
            style_season_map = {}

    except Exception as e:
        st.warning(f"Impossible d'extraire les styles pour la saisonalité : {e}")

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
