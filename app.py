import streamlit as st
import io
import time
import hashlib
import openpyxl
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

st.markdown(
    """
1) Téléverse ton **fichier fournisseur** (Excel)  
2) Téléverse ton fichier **Help Data** (Excel)  
3) Configure tes **Tags**  
4) Clique sur **Générer le fichier Shopify**
"""
)

# -----------------------
# Uploads
# -----------------------
supplier_file = st.file_uploader("📄 Fichier fournisseur (Excel)", type=["xlsx"])
help_file = st.file_uploader("🧩 Help Data (Excel)", type=["xlsx"])


# -----------------------
# Suppliers
# -----------------------
SUPPLIERS = {
    "Balmoral": run_abc,
    "Bandit": run_abc,
    "Black Sheep": run_abc,
    "Ciele": run_abc,
    "Clif Bar": run_abc,
    "Cody": run_abc,
    "Duer": run_abc,
    "Hoka": run_abc,
    "Maurten": run_abc,
    "Norda": run_abc,
    "Oakley": run_abc,
    "Pas Normal Studios": run_abc,
    "Tracksmith": run_abc,
}

supplier_name = st.selectbox("Choisir le fournisseur", list(SUPPLIERS.keys()))

# -----------------------
# Helpers
# -----------------------
def _first_existing_col(cols, candidates):
    cols_l = [c.lower() for c in cols]
    for c in candidates:
        if c.lower() in cols_l:
            return cols[cols_l.index(c.lower())]
    return None


def _extract_unique_style_rows(xlsx_bytes):
    """Return a dataframe of unique styles for the seasonality table.

    Display order:
      1) Style Name
      2) Style Number

    Returns None if neither column is found.
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


# -----------------------
# 3) Tags
# -----------------------
style_season_map = {}
event_promo_tag = ""

if supplier_file:
    st.markdown("### 3) Tags")

    event_promo_tag = st.selectbox(
        "Event/Promotion Related",
        options=["", "spring-summer", "fall-winter"],
        index=0,
        help="S'applique à l'ensemble des pièces du fichier (optionnel).",
    )

    st.markdown("---")

    style_rows_df = _extract_unique_style_rows(supplier_file.getvalue())

    if style_rows_df is not None and not style_rows_df.empty:
        st.caption("Saisonality (par style) — **champ libre**. Le tag sera ajouté dans **Tags** pour toutes les lignes du même style.")

        key_col = "Style Number" if "Style Number" in style_rows_df.columns else "Style Name"

        # fingerprint for widget key (prevents 'type twice' issue)
        supplier_fp = hashlib.md5(supplier_file.getvalue()).hexdigest()
        styles_fp = hashlib.md5("|".join(style_rows_df[key_col].astype(str).tolist()).encode("utf-8")).hexdigest()
        fp = f"{supplier_fp}:{styles_fp}:{key_col}"

        # init/refresh only when file/styles change
        if st.session_state.get("seasonality_fp") != fp:
            st.session_state["seasonality_fp"] = fp

            # preserve existing values (if any)
            existing_map = {}
            existing = st.session_state.get("seasonality_df")
            if existing is not None and key_col in existing.columns and "Saisonality tag" in existing.columns:
                existing_map = {
                    str(k).strip(): str(v).strip()
                    for k, v in zip(existing[key_col].astype(str), existing["Saisonality tag"].astype(str))
                    if str(k).strip()
                }

            init_df = style_rows_df.copy()
            init_df["Saisonality tag"] = init_df[key_col].astype(str).map(lambda k: existing_map.get(str(k).strip(), ""))
            st.session_state["seasonality_df"] = init_df

        base_df = st.session_state.get("seasonality_df", style_rows_df.copy())
        if "Saisonality tag" not in base_df.columns:
            base_df = base_df.copy()
            base_df["Saisonality tag"] = ""

        # IMPORTANT: key is tied to fp to stabilize widget state and avoid first entry disappearing
        editor_key = f"seasonality_editor_{fp}"

        edited_df = st.data_editor(
            base_df,
            key=editor_key,
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

        st.session_state["seasonality_df"] = edited_df

        style_season_map = {}
        for _, r in edited_df.iterrows():
            k = str(r.get(key_col, "")).strip()
            v = str(r.get("Saisonality tag", "")).strip()
            if k and v:
                style_season_map[k] = v
    else:
        st.info("Aucun champ 'Style Name' ou 'Style Number' détecté dans le fichier. La saisonalité par style sera ignorée.")
        style_season_map = {}

# -----------------------
# 4) Génération
# -----------------------
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
        time.sleep(0.2)
        progress.progress(20)

        help_wb = openpyxl.load_workbook(io.BytesIO(help_file.getvalue()), data_only=True)

        status.info("Transformation en cours…")
        time.sleep(0.2)
        progress.progress(60)

        output_bytes = transform_fn(
            supplier_xlsx_bytes=supplier_file.getvalue(),
            help_xlsx_bytes=help_file.getvalue(),
            vendor_name=supplier_name,
            brand_choice="",
            style_season_map=style_season_map,
            event_promo_tag=event_promo_tag,
        )

        status.success("Fichier Shopify généré ✅")
        progress.progress(100)

        st.download_button(
            "⬇️ Télécharger le fichier Shopify",
            data=output_bytes,
            file_name="shopify_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        status.error(f"Erreur lors de la génération : {e}")
        progress.empty()
