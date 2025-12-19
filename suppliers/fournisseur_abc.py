import io
import re
import math
import pandas as pd
from slugify import slugify
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ---------------------------------------------------------
# 45 colonnes EXACTES + ordre EXACT (comme ton message)
# ---------------------------------------------------------
SHOPIFY_OUTPUT_COLUMNS = [
    "Handle",
    "Command",
    "Title",
    "Body (HTML)",
    "Vendor",
    "Custom Product Type",
    "Tags",
    "Published",
    "Published Scope",
    "Option1 Name",
    "Option1 Value",
    "Variant SKU",
    "Variant Barcode",
    "Variant Country of Origin",
    "Variant HS Code",
    "Variant Grams",
    "Variant Inventory Tracker",
    "Variant Inventory Policy",
    "Variant Fulfillment Service",
    "Variant Price",
    "Variant Requires Shipping",
    "Variant Taxable",
    "SEO Title",
    "SEO Description",
    "Variant Weight Unit",
    "Cost per item",
    "Status",
    "Metafield: my_fields.product_use_case [multi_line_text_field]",
    "Metafield: my_fields.product_features [multi_line_text_field]",
    "Metafield: my_fields.behind_the_brand [multi_line_text_field]",
    "Metafield: my_fields.size_comment [single_line_text_field]",
    "Metafield: my_fields.gender [single_line_text_field]",
    "Metafield: my_fields.colour [single_line_text_field]",
    "Metafield: mm-google-shopping.color",
    "Variant Metafield: mm-google-shopping.size",
    "Metafield: mm-google-shopping.size_system",
    "Metafield: mm-google-shopping.condition",
    "Metafield: mm-google-shopping.google_product_category",
    "Metafield: mm-google-shopping.gender",
    "Variant Metafield: mm-google-shopping.mpn",
    "Variant Metafield: mm-google-shopping.gtin",
    "Metafield: theme.siblings [single_line_text_field]",
    "Category: ID",
    "Inventory Available: Boutique",
    "Inventory Available: Le Club",
]

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


# ----------------------------
# Helpers
# ----------------------------
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())


def _words(s: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]+", str(s).lower()))


def _first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in cols:
            return cols[c.lower()]
    return None


# ----------------------------
# Read 2-column sheets from help data
# ----------------------------
def _read_help_2cols(help_bytes: bytes, sheet_name: str) -> pd.DataFrame | None:
    """
    Reads first 2 columns of a sheet as strings.
    Returns df with columns: a, b
    """
    try:
        df = pd.read_excel(io.BytesIO(help_bytes), sheet_name=sheet_name, dtype=str)
        if df is None or df.empty or df.shape[1] < 2:
            return None
        out = df.iloc[:, :2].copy()
        out.columns = ["a", "b"]
        out["a"] = out["a"].astype(str).str.strip()
        out["b"] = out["b"].astype(str).str.strip()
        out = out[(out["a"] != "") & (out["a"].str.lower() != "nan")]
        return out
    except Exception:
        return None


def _build_standardization_map(help_bytes: bytes, sheet_name: str) -> dict[str, str]:
    """
    Standardization sheets:
      Col A = raw, Col B = standard
    """
    df = _read_help_2cols(help_bytes, sheet_name)
    if df is None:
        return {}
    return {str(k).strip().lower(): str(v).strip() for k, v in zip(df["a"], df["b"])}


def _standardize(val: str, mapping: dict[str, str]) -> str:
    s = _norm(val)
    if not s or s.lower() == "nan":
        return ""
    return mapping.get(s.lower(), s)


def _read_product_types(help_bytes: bytes) -> list[str]:
    """
    Product Types sheet is a single column list (any header ok).
    """
    try:
        df = pd.read_excel(io.BytesIO(help_bytes), sheet_name="Product Types", dtype=str)
        if df is None or df.empty:
            return []
        col = df.columns[0]
        vals = [str(v).strip() for v in df[col].astype(str).tolist() if str(v).strip().lower() != "nan" and str(v).strip() != ""]
        return vals
    except Exception:
        return []


def _read_size_reco(help_bytes: bytes) -> pd.DataFrame | None:
    """
    Size Recommandation: expects columns Garment, Comment
    """
    try:
        df = pd.read_excel(io.BytesIO(help_bytes), sheet_name="Size Recommandation", dtype=str)
        if df is None or df.empty:
            return None
        return df
    except Exception:
        return None


# ----------------------------
# Fallback parse Color/Size from Description
# ----------------------------
def _extract_color_size_from_description(desc: str) -> tuple[str, str]:
    """
    Simple fallback:
    - looks for ending patterns "... COLOR - SIZE" / "... COLOR, SIZE" / "... COLOR / SIZE"
    """
    text = _norm(desc)
    if not text:
        return "", ""

    parts = re.split(r"\s*[-,/]\s*|\s*,\s*", text)
    parts = [p.strip() for p in parts if p and p.strip()]
    if len(parts) < 2:
        return "", ""

    last = parts[-1]
    if re.fullmatch(r"(X{0,3}S|X{0,3}L|S|M|L|XL|XXL|XXXL|\d{1,2}([./-]\d{1,2})?)", last, flags=re.IGNORECASE):
        return parts[-2], last

    return parts[-1], ""


# ----------------------------
# Exact-match: category IDs + product type
# ----------------------------
def _best_match_id(description: str, cat_df: pd.DataFrame | None) -> str:
    """
    cat_df: a=name/keyword, b=id
    Exact match: all words in 'a' must exist in description.
    Returns 'b' (ID).
    """
    if cat_df is None:
        return ""

    desc_words = _words(description)
    best_id = ""
    best_len = 0

    for _, row in cat_df.iterrows():
        name = str(row["a"]).strip()
        cid = str(row["b"]).strip()
        name_words = _words(name)
        if not name_words:
            continue
        if name_words.issubset(desc_words):
            if len(name_words) > best_len:
                best_len = len(name_words)
                best_id = cid

    # avoid "123.0"
    best_id = re.sub(r"\.0$", "", best_id) if best_id else ""
    return best_id


def _best_match_product_type(description: str, product_types: list[str]) -> str:
    desc_words = _words(description)
    best = ""
    best_len = 0
    for pt in product_types:
        w = _words(pt)
        if w and w.issubset(desc_words) and len(w) > best_len:
            best = pt
            best_len = len(w)
    return best


# ----------------------------
# Price: round to nearest x9.99
# ----------------------------
def _round_to_nearest_9_99(price) -> float:
    if price is None or (isinstance(price, float) and math.isnan(price)):
        return float("nan")
    p = float(price)
    nearest10 = math.floor(p / 10.0 + 0.5) * 10.0
    return round(nearest10 - 0.01, 2)


# ----------------------------
# Barcode: keep leading zeros
# ----------------------------
def _barcode_keep_zeros(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if re.fullmatch(r"\d+", s):
        return s.zfill(12) if len(s) <= 12 else s
    return s


# ----------------------------
# Apply yellow fill to empty cells in selected columns
# ----------------------------
def _apply_yellow_for_empty(writer_buffer: io.BytesIO, sheet_name: str, columns_to_yellow_if_empty: list[str]) -> io.BytesIO:
    writer_buffer.seek(0)
    wb = load_workbook(writer_buffer)
    ws = wb[sheet_name]

    # Map header names to column index (1-based)
    header = [cell.value for cell in ws[1]]
    col_index = {name: i + 1 for i, name in enumerate(header) if name}

    for col_name in columns_to_yellow_if_empty:
        if col_name not in col_index:
            continue
        c = col_index[col_name]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val is None or (isinstance(val, str) and val.strip() == ""):
                cell.fill = YELLOW_FILL

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ----------------------------
# Main transform for Fournisseur ABC
# ----------------------------
def run_transform(supplier_xlsx_bytes: bytes, help_xlsx_bytes: bytes, vendor_name: str):
    sup = pd.read_excel(io.BytesIO(supplier_xlsx_bytes), sheet_name=0, dtype=str).copy()
    warnings: list[dict] = []

    # Help data
    color_map = _build_standardization_map(help_xlsx_bytes, "Color Standardization")
    size_map = _build_standardization_map(help_xlsx_bytes, "Size Standardization")
    country_map = _build_standardization_map(help_xlsx_bytes, "Country Abbreviations")
    gender_map = _build_standardization_map(help_xlsx_bytes, "Gender Standardization")

    shopify_cat = _read_help_2cols(help_xlsx_bytes, "Shopify Product Category")
    google_cat = _read_help_2cols(help_xlsx_bytes, "Google Product Category")
    product_types = _read_product_types(help_xlsx_bytes)
    size_reco = _read_size_reco(help_xlsx_bytes)

    # Supplier columns (flexibles)
    desc_col = _first_existing_col(sup, ["Description", "Product Name", "Title"])
    product_col = _first_existing_col(sup, ["Product", "Product Code", "SKU"])
    color_col = _first_existing_col(sup, ["Color", "Colour"])
    size_col = _first_existing_col(sup, ["Size"])
    upc_col = _first_existing_col(sup, ["UPC", "UPC Code"])
    origin_col = _first_existing_col(sup, ["Country Code", "Origin", "Manufacturing Country"])
    hs_col = _first_existing_col(sup, ["HS Code", "HTS Code"])
    extid_col = _first_existing_col(sup, ["External ID", "ExternalID"])
    msrp_col = _first_existing_col(sup, ["Cad MSRP", "MSRP"])
    landed_col = _first_existing_col(sup, ["Landed"])
    grams_col = _first_existing_col(sup, ["Grams", "Weight (g)", "Weight"])

    gender_col = _first_existing_col(sup, ["Gender"])

    if desc_col is None:
        raise ValueError('Colonne Description / Product Name / Title introuvable dans le fichier fournisseur.')
    if msrp_col is None:
        raise ValueError('Colonne Cad MSRP / MSRP introuvable dans le fichier fournisseur.')

    # Base fields
    sup["_vendor"] = vendor_name
    sup["_description"] = sup[desc_col].astype(str).fillna("").map(_norm)

    # Color / Size with fallback
    sup["_color_raw"] = sup[color_col].astype(str).fillna("").map(_norm) if color_col else ""
    sup["_size_raw"] = sup[size_col].astype(str).fillna("").map(_norm) if size_col else ""

    parsed = sup["_description"].apply(_extract_color_size_from_description)
    sup["_color_fb"] = parsed.map(lambda t: t[0])
    sup["_size_fb"] = parsed.map(lambda t: t[1])

    sup["_color"] = sup["_color_raw"]
    sup.loc[sup["_color"].eq(""), "_color"] = sup["_color_fb"]

    sup["_size"] = sup["_size_raw"]
    sup.loc[sup["_size"].eq(""), "_size"] = sup["_size_fb"]

    # Standardize
    sup["_color_std"] = sup["_color"].apply(lambda x: _standardize(x, color_map))
    sup["_size_std"] = sup["_size"].apply(lambda x: _standardize(x, size_map))

    # Gender (supplier OR infer via mapping on text)
    if gender_col:
        sup["_gender_raw"] = sup[gender_col].astype(str).fillna("").map(_norm)
    else:
        sup["_gender_raw"] = ""

    sup["_gender_std"] = sup["_gender_raw"].apply(lambda x: _standardize(x, gender_map)) if gender_map else sup["_gender_raw"]

    # Title = Description + Color (real/standardized)
    sup["_title"] = (sup["_description"] + " " + sup["_color_std"]).str.strip()

    # Handle = Vendor + Gender + Description + Color, with hyphens (slugify)
    def _make_handle(r):
        parts = [r["_vendor"], r["_gender_std"], r["_description"], r["_color_std"]]
        parts = [p for p in parts if p and str(p).strip()]
        return slugify(" ".join(parts))

    sup["_handle"] = sup.apply(_make_handle, axis=1)

    # SEO Title = Handle with real spaces (not hyphens)
    sup["_seo_title"] = sup["_handle"].astype(str).str.replace("-", " ").str.strip()

    # SEO Description (exact wording from rules file)
    sup["_seo_desc"] = sup.apply(
        lambda r: f"Shop the {r['_seo_title']} with free worldwide shipping, and 30-day returns on leclub.cc. Discover {r['_vendor']} products.",
        axis=1
    )

    # Custom Product Type = exact-match from help data list
    sup["_product_type"] = sup["_description"].apply(lambda d: _best_match_product_type(d, product_types))

    # Tags = Vendor + standardized color + gender + "_badge_new" + product type
    # + if gender includes both -> add Men and Women
    def _make_tags(r):
        tags = []
        if r["_vendor"]:
            tags.append(r["_vendor"])
        if r["_color_std"]:
            tags.append(r["_color_std"])
        if r["_gender_std"]:
            tags.append(r["_gender_std"])
        tags.append("_badge_new")
        if r["_product_type"]:
            tags.append(r["_product_type"])

        g = str(r["_gender_std"]).lower()
        if "men" in g and "women" in g:
            tags.extend(["Men", "Women"])

        return ", ".join([t for t in tags if t])

    sup["_tags"] = sup.apply(_make_tags, axis=1)

    # SKU rule: External ID else Product code else "SKU-[Size]-[Color]"
    sup["_external_id"] = sup[extid_col].astype(str).fillna("").map(_norm) if extid_col else ""
    sup["_product_code"] = sup[product_col].astype(str).fillna("").map(_norm) if product_col else ""

    def _make_sku(r):
        if r["_external_id"]:
            return r["_external_id"]
        if r["_product_code"]:
            return r["_product_code"]
        base = r["_product_code"] or "SKU"
        return f"{base}-{r['_size_std']}-{r['_color_std']}".strip("-")

    sup["_variant_sku"] = sup.apply(_make_sku, axis=1)

    # Barcode = UPC (keep leading zeros)
    sup["_barcode"] = sup[upc_col].apply(_barcode_keep_zeros) if upc_col else ""

    # Country of origin = standardized via Country Abbreviations
    sup["_origin_raw"] = sup[origin_col].astype(str).fillna("").map(_norm) if origin_col else ""
    sup["_origin_std"] = sup["_origin_raw"].apply(lambda x: _standardize(x, country_map))

    # HS Code
    sup["_hs"] = sup[hs_col].astype(str).fillna("").map(_norm) if hs_col else ""

    # Grams (if present)
    sup["_grams"] = sup[grams_col].astype(str).fillna("").map(_norm) if grams_col else ""

    # Price: Cad MSRP rounded to x9.99
    msrp_num = pd.to_numeric(
        sup[msrp_col].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False),
        errors="coerce",
    )
    sup["_price"] = msrp_num.apply(_round_to_nearest_9_99)

    # Cost per item = Landed (as string)
    sup["_cost"] = sup[landed_col].astype(str).fillna("").map(_norm) if landed_col else ""

    # Size comment: from Size Recommandation sheet (Garment == vendor)
    def _size_comment(vendor: str) -> str:
        if size_reco is None or size_reco.empty:
            return ""
        v = vendor.strip().lower()
        gcol = None
        ccol = None
        # try to find columns "Garment" and "Comment" robustly
        for c in size_reco.columns:
            if c.strip().lower() == "garment":
                gcol = c
            if c.strip().lower() == "comment":
                ccol = c
        if not gcol or not ccol:
            return ""

        hits = size_reco[size_reco[gcol].astype(str).str.strip().str.lower() == v]
        if not hits.empty:
            return _norm(hits.iloc[0][ccol])

        hits = size_reco[size_reco[gcol].astype(str).str.strip().str.lower().apply(lambda g: g in v or v in g)]
        if not hits.empty:
            return _norm(hits.iloc[0][ccol])

        return ""

    sup["_size_comment"] = sup["_vendor"].apply(_size_comment)

    # Categories: IDs via exact-match in description
    sup["_shopify_cat_id"] = sup["_description"].apply(lambda d: _best_match_id(d, shopify_cat))
    sup["_google_cat_id"] = sup["_description"].apply(lambda d: _best_match_id(d, google_cat))

    # theme.siblings = Vendor + Description (slugify)
    sup["_siblings"] = sup.apply(lambda r: slugify(f"{r['_vendor']} {r['_description']}"), axis=1)

    # ----------------------------
    # Build output
    # ----------------------------
    out = pd.DataFrame(columns=SHOPIFY_OUTPUT_COLUMNS)

    out["Handle"] = sup["_handle"]
    out["Command"] = "NEW"
    out["Title"] = sup["_title"]
    out["Body (HTML)"] = ""  # keep empty (yellow if empty)
    out["Vendor"] = sup["_vendor"]

    out["Custom Product Type"] = sup["_product_type"]
    out["Tags"] = sup["_tags"]

    out["Published"] = False
    out["Published Scope"] = "global"

    out["Option1 Name"] = "Size"
    out["Option1 Value"] = sup["_size_std"]

    out["Variant SKU"] = sup["_variant_sku"]
    out["Variant Barcode"] = sup["_barcode"]
    out["Variant Country of Origin"] = sup["_origin_std"]
    out["Variant HS Code"] = sup["_hs"]
    out["Variant Grams"] = sup["_grams"]

    out["Variant Inventory Tracker"] = "shopify"
    out["Variant Inventory Policy"] = "deny"
    out["Variant Fulfillment Service"] = "manual"
    out["Variant Price"] = sup["_price"]

    out["Variant Requires Shipping"] = True
    out["Variant Taxable"] = True

    out["SEO Title"] = sup["_seo_title"]
    out["SEO Description"] = sup["_seo_desc"]

    out["Variant Weight Unit"] = "g"
    out["Cost per item"] = sup["_cost"]
    out["Status"] = "draft"

    out["Metafield: my_fields.product_use_case [multi_line_text_field]"] = ""  # keep empty
    out["Metafield: my_fields.product_features [multi_line_text_field]"] = ""  # no clear rule -> empty
    out["Metafield: my_fields.behind_the_brand [multi_line_text_field]"] = ""  # no clear rule -> empty
    out["Metafield: my_fields.size_comment [single_line_text_field]"] = sup["_size_comment"]
    out["Metafield: my_fields.gender [single_line_text_field]"] = sup["_gender_std"]
    out["Metafield: my_fields.colour [single_line_text_field]"] = sup["_color_std"]

    out["Metafield: mm-google-shopping.color"] = sup["_color_std"]
    out["Variant Metafield: mm-google-shopping.size"] = sup["_size_std"]
    out["Metafield: mm-google-shopping.size_system"] = "US"  # confirmed
    out["Metafield: mm-google-shopping.condition"] = "new"
    out["Metafield: mm-google-shopping.google_product_category"] = sup["_google_cat_id"]  # ID per rules file
    out["Metafield: mm-google-shopping.gender"] = sup["_gender_std"]
    out["Variant Metafield: mm-google-shopping.mpn"] = sup["_variant_sku"]
    out["Variant Metafield: mm-google-shopping.gtin"] = sup["_barcode"]

    out["Metafield: theme.siblings [single_line_text_field]"] = sup["_siblings"]

    out["Category: ID"] = sup["_shopify_cat_id"]  # ID per rules file

    out["Inventory Available: Boutique"] = 0
    out["Inventory Available: Le Club"] = 0

    # Ensure exact order
    out = out.reindex(columns=SHOPIFY_OUTPUT_COLUMNS)

    # ----------------------------
    # Yellow columns when empty (per rules file)
    # ----------------------------
    yellow_if_empty_cols = [
        "Handle",
        "Title",
        "Body (HTML)",
        "Custom Product Type",
        "Option1 Name",
        "Option1 Value",
        "Variant Price",
        "SEO Title",
        "SEO Description",
        "Metafield: my_fields.size_comment [single_line_text_field]",
        "Metafield: my_fields.gender [single_line_text_field]",
        "Metafield: my_fields.colour [single_line_text_field]",
        "Metafield: mm-google-shopping.color",
        "Variant Metafield: mm-google-shopping.size",
    ]

    # Export to Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="shopify_import")
        pd.DataFrame(warnings).to_excel(writer, index=False, sheet_name="warnings")

    buffer = _apply_yellow_for_empty(buffer, "shopify_import", yellow_if_empty_cols)

    return buffer.getvalue(), pd.DataFrame(warnings)
