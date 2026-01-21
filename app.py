import streamlit as st
import io
import time
import hashlib
import openpyxl
import pandas as pd

from suppliers.fournisseur_abc import run_transform as run_abc

st.set_page_config(page_title="Générateur Shopify", layout="wide")

st.title("Générateur de fichier Shopify")

# =========================
# Upload des fichiers
# =========================
supplier_file = st.file_uploader("📄 Fichier fournisseur (Excel)", type=["xlsx"])
help_file = st.file_uploader("🧩 Help Data (Excel)", type=["xlsx"])


# =========================
# Helpers
# =========================
def _first_existing_col(cols, candidates):
    cols_l = [c.lower() for c in cols]
    for c in candidates:
        if c.lower() in cols_l:
            return cols[cols_l.index(c.lower())]
    return None


def _extract_unique_style_rows(xlsx_bytes):
    bio = io.BytesIO(xlsx_bytes)
    xls = pd.ExcelFile(bio)

    style_number_candidates = [
        "Style Number", "Style Num", "Style", "style", "style number", "Style #"
    ]
    style_name_candidates = [
        "Style Name", "style name", "Product Name", "Name"
    ]

    rows = []

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
        except Exception:
            continue

        if df.empty:
            continue

        num_col = _first_existing_col(df.columns, style_number_candidates)
        name_col = _first_existing_col(df.columns, style_name_candidates)

        if not num_col and not name_col:
            continue

        data = {}
        if name_col:
            data["Style Name"] = df[name_col].astype(str).str.strip()
        if num_col:
            data["Style Number"] = df[num_col].astype(str).str.strip()

        tmp = pd.DataFrame(data)

        for c in tmp.columns:
            tmp = tmp[tmp[c] != ""]

        rows.append(tmp)

    if not rows:
        return None

    out = pd.concat(rows).drop_duplicates()

    cols = []
    if "Style Name" in out.columns:
        cols.append("Style Name")
    if "Style Number" in out.columns:
        cols.append("Style Number")

    return out[cols].reset_index(drop=True)


# =========================
# Saisonality par style
# =========================
style_season_map = {}

if supplier_file:
    style_rows_df = _extract_unique_style_rows(supplier_file.getvalue())

    if style_rows_df is not None and not style_rows_df.empty:

        st.caption(
            "Le **Saisonality tag** est appliqué à toutes les lignes partageant le même style."
        )

        key_col = "Style Number" if "Style Number" in style_rows_df.columns else "Style Name"

        # 🔒 Fingerprint pour éviter les resets pendant la saisie
        supplier_fp = hashlib.md5(supplier_file.getvalue()).hexdigest()
        styles_fp = hashlib.md5(
            "|".join(style_rows_df[key_col].astype(str).tolist()).encode()
        ).hexdigest()
        fp = f"{supplier_fp}:{styles_fp}:{key_col}"

        if st.session_state.get("seasonality_fp") != fp:
            st.session_state["seasonality_fp"] = fp

            # récupérer anciennes valeurs si présentes
            existing_map = {}
            existing = st.session_state.get("seasonality_df")
            if existing is not None and "Saisonality tag" in existing.columns:
                existing_map = {
                    str(k).strip(): str(v).strip()
                    for k, v in zip(existing[key_col], existing["Saisonality tag"])
                    if str(k).strip()
                }

            init_df = style_rows_df.copy()
            init_df["Saisonality tag"] = init_df[key_col].map(
                lambda k: existing_map.get(str(k).strip(), "")
            )

            st.session_state["seasonality_df"] = init_df

        # 📝 Tableau éditable (champ libre)
        seasonality_df = st.data_editor(
            st.session_state["seasonality_df"],
            key="seasonality_editor",
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "Style Name": st.column_config.TextColumn(disabled=True),
                "Style Number": st.column_config.TextColumn(disabled=True),
                "Saisonality tag": st.column_config.TextColumn(
                    help="Champ libre (ex. spring-summer, fall, core)",
                ),
            },
        )

        st.session_state["seasonality_df"] = seasonality_df

        # Construire le mapping final
        for _, r in seasonality_df.iterrows():
            k = str(r.get(key_col, "")).strip()
            v = str(r.get("Saisonality tag", "")).strip()
            if k and v:
                style_season_map[k] = v

    else:
        st.info("Aucun Style détecté — la saisonalité par style sera ignorée.")


# =========================
# Génération Shopify
# =========================
generate = st.button(
    "Générer le fichier Shopify",
    disabled=not (supplier_file and help_file),
)

if generate:
    status = st.empty()
    progress = st.progress(0)

    try:
        status.info("Lecture des fichiers…")
        progress.progress(20)
        time.sleep(0.3)

        help_wb = openpyxl.load_workbook(
            io.BytesIO(help_file.getvalue()), data_only=True
        )

        status.info("Transformation en cours…")
        progress.progress(60)
        time.sleep(0.3)

        output_bytes = run_abc(
            supplier_file.getvalue(),
            help_wb,
            style_season_map=style_season_map,
        )

        progress.progress(100)
        status.success("Fichier Shopify généré ✅")

        st.download_button(
            "⬇️ Télécharger le fichier Shopify",
            data=output_bytes,
            file_name="shopify_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        status.error(f"Erreur : {e}")
        progress.empty()
