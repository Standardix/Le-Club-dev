from __future__ import annotations


def _norm_upc(v) -> str:
    """Normalize UPC/Barcode: keep digits only, drop trailing .0 from numeric."""
    if v is None:
        return ""
    s = str(v).strip()
    # drop .0 for floats represented as '123.0'
    s = re.sub(r"\.0$", "", s)
    # keep digits only
    s = re.sub(r"\D", "", s)
    return s



def _colkey(c: str) -> str:
    """Normalize column name: lower, remove spaces/punct/parentheses for robust matching."""
    s = str(c or "").strip().lower()
    s = re.sub(r"[\s\-_/]+", "", s)
    s = s.replace("(", "").replace(")", "")
    return s

def _find_col(df_cols, candidates):
    """Return first column in df_cols matching any normalized candidate."""
    norm_map = {_colkey(c): c for c in df_cols}
    for cand in candidates:
        k = _colkey(cand)
        if k in norm_map:
            return norm_map[k]
    # also allow partial contains on normalized
    cols_norm = [(_colkey(c), c) for c in df_cols]
    for cand in candidates:
        ck = _colkey(cand)
        for cn, orig in cols_norm:
            if ck and ck in cn:
                return orig
    return None


def _header_has_cad(col_name: str) -> bool:
    return "cad" in str(col_name or "").lower()



import io
import re
import math

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# Optional dependency: python-slugify
try:
    from slugify import slugify  # type: ignore
except Exception:
    def slugify(value: str) -> str:
        s = str(value or "").strip().lower()
        s = re.sub(r"[^a-z0-9]+", "-", s)
        s = re.sub(r"-{2,}", "-", s).strip("-")
        return s

def _build_existing_shopify_index(existing_shopify_xlsx_bytes: bytes | None):
    """Build matching indexes from an existing Shopify product export/list.

    Keys priority (as requested):
      1) brand + SKU + UPC
      2) brand + UPC
      3) brand + SKU
      4) SKU + UPC
      5) UPC

    Also returns a set of existing handles (normalized).
    """
    handles_set: set[str] = set()
    key_sets = {
        "brand_sku_upc": set(),
        "brand_upc": set(),
        "brand_sku": set(),
        "sku_upc": set(),
        "upc": set(),
    }
    if not existing_shopify_xlsx_bytes:
        return handles_set, key_sets

    try:
        bio = io.BytesIO(existing_shopify_xlsx_bytes)
        df = pd.read_excel(bio)  # first sheet by default
    except Exception:
        return handles_set, key_sets

    cols_l = {str(c).strip().lower(): c for c in df.columns}
    handle_col = cols_l.get("handle")
    vendor_col = cols_l.get("vendor") or cols_l.get("brand")
    sku_col = cols_l.get("variant sku") or cols_l.get("sku")
    upc_col = cols_l.get("variant barcode") or cols_l.get("barcode") or cols_l.get("upc")

    for _, r in df.iterrows():
        h = _norm_handle(r.get(handle_col, "")) if handle_col else ""
        if h:
            handles_set.add(h)

        brand = _norm(r.get(vendor_col, "")) if vendor_col else ""
        sku = _norm(r.get(sku_col, "")) if sku_col else ""
        upc = _norm_upc(r.get(upc_col, "")) if upc_col else ""

        if brand and sku and upc:
            key_sets["brand_sku_upc"].add((brand, sku, upc))
        if brand and upc:
            key_sets["brand_upc"].add((brand, upc))
        if brand and sku:
            key_sets["brand_sku"].add((brand, sku))
        if sku and upc:
            key_sets["sku_upc"].add((sku, upc))
        if upc:
            key_sets["upc"].add((upc,))

    return handles_set, key_sets


def _row_is_existing(brand: str, sku: str, upc: str, key_sets) -> bool:
    b = _norm(brand)
    s = _norm(sku)
    u = _norm_upc(upc)
    if b and s and u and (b, s, u) in key_sets["brand_sku_upc"]:
        return True
    if b and u and (b, u) in key_sets["brand_upc"]:
        return True
    if b and s and (b, s) in key_sets["brand_sku"]:
        return True
    if s and u and (s, u) in key_sets["sku_upc"]:
        return True
    if u and (u,) in key_sets["upc"]:
        return True
    return False


def _apply_red_font_for_handle(buffer: io.BytesIO, sheet_name: str, rows_to_color: list[int]) -> io.BytesIO:
    """Color the Handle cell red for the given 0-based row indexes (dataframe rows)."""
    wb = load_workbook(buffer)
    ws = wb[sheet_name]

    # locate Handle column
    headers = [str(c.value or "") for c in ws[1]]
    try:
        handle_col_idx = headers.index("Handle") + 1
    except ValueError:
        # nothing to do
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out

    red_font = Font(color="FF0000")
    for df_i in rows_to_color:
        excel_row = df_i + 2  # header +1, df row offset
        cell = ws.cell(row=excel_row, column=handle_col_idx)
        cell.font = red_font

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
import io
import re
import math
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# ---------------------------------------------------------
# ORDRE FINAL DES COLONNES (strict)
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

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())


def _norm_handle(v) -> str:
    s = str(v or "").strip().lower()
    # collapse whitespace
    s = re.sub(r"\s+", "", s)
    return s


def _strip_gender_prefix_size(v: str) -> str:
    s = _norm(v)
    if not s:
        return ""
    if re.match(r"^[WwMm]\s*\d", s):
        return s[1:].strip()
    return s


def _strip_gender_tokens(text: str) -> str:
    """Remove embedded gender markers like -w-, - W -, -m-, etc from a string."""
    s = str(text or "")
    # remove patterns like -w- , - W - , /w/ etc surrounded by dashes/spaces
    s = re.sub(r"(?i)(\s*-\s*[wm]\s*-\s*)", " ", s)
    s = re.sub(r"(?i)(\b[wm]\b)", lambda m: "" if m.group(0).lower() in ("w","m") else m.group(0), s)
    s = re.sub(r"\s+", " ", s).strip()
    return s




def _clean_style_key(v) -> str:
    s = _norm(v)
    # if Excel treated numeric as float: 123.0 -> 123
    s = re.sub(r"^(\d+)\.0+$", r"\1", s)
    return s

def _strip_reg_for_handle(s: str) -> str:
    """Handle only: remove ® and (r)/[r] to keep URL safe."""
    t = _norm(s)
    t = t.replace("®", "")
    t = re.sub(r"[\(\[\{]\s*r\s*[\)\]\}]", "", t, flags=re.IGNORECASE)
    return _norm(t)


def _convert_r_to_registered(s: str) -> str:
    """Display/SEO: convert (r)/[r] to ®."""
    t = _norm(s)
    t = re.sub(r"[\(\[\{]\s*r\s*[\)\]\}]", "®", t, flags=re.IGNORECASE)
    return t


def _title_case_preserve_registered(text: str) -> str:
    """
    Strict Title Case while preserving ®.
    - Title-cases each space-separated token
    - Also title-cases sub-tokens split by "/" and "-"
    - Keeps tokens containing digits as-is
    """
    text = _norm(text)
    if not text:
        return ""

    def _tc_token(tok: str) -> str:
        if not tok:
            return tok
        if any(ch.isdigit() for ch in tok):
            return tok

        # preserve ® inside token
        if "®" in tok:
            sub = tok.split("®")
            sub = [(_tc_token(s) if s else "") for s in sub]
            return "®".join(sub)

        # separators inside token
        for sep in ["/", "-"]:
            if sep in tok:
                parts = tok.split(sep)
                parts = [(_tc_token(p) if p else "") for p in parts]
                return sep.join(parts)

        return tok[:1].upper() + tok[1:].lower()

    return " ".join(_tc_token(w) for w in text.split(" "))


def _normalize_match_text(s: str) -> str:
    """
    Normalization used ONLY for matching category/product type:
    - tee -> tshirt
    - t-shirt / t shirt -> tshirt
    - long-sleeve / long sleeve -> long sleeve
    """
    t = str(s or "")
    t = t.replace("®", "")
    t = t.lower()

    t = re.sub(r"\bt\s*[- ]\s*shirt\b", "tshirt", t)
    t = re.sub(r"\btshirt\b", "tshirt", t)

    t = re.sub(r"\btee\b", "tshirt", t)
    t = re.sub(r"\btees\b", "tshirt", t)

    t = re.sub(r"\blong\s*[- ]\s*sleeve\b", "long sleeve", t)
    return t


def _words(s: str) -> list[str]:
    return re.findall(r"[a-z0-9]+", _normalize_match_text(s))


def _singularize_token(tok: str) -> str:
    if tok.endswith("s") and len(tok) >= 4:
        return tok[:-1]
    return tok


def _wordset_loose(s: str) -> set[str]:
    return set(_singularize_token(t) for t in _words(s))


def _first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in cols:
            return cols[c.lower()]
    return None


# ---------------------------------------------------------
# Help data readers (openpyxl)
# ---------------------------------------------------------
def _load_help_wb(help_bytes: bytes):
    return openpyxl.load_workbook(io.BytesIO(help_bytes), data_only=True)


def _read_2col_map(wb, sheet_candidates: list[str]) -> dict[str, str]:
    """Col A raw -> Col B standard"""
    sheet = None
    for name in sheet_candidates:
        if name in wb.sheetnames:
            sheet = wb[name]
            break
    if sheet is None:
        return {}

    m: dict[str, str] = {}
    for r in range(2, sheet.max_row + 1):
        a = sheet.cell(row=r, column=1).value
        b = sheet.cell(row=r, column=2).value
        if a is None or b is None:
            continue
        ra = str(a).strip()
        rb = str(b).strip()
        if not ra or ra.lower() == "nan":
            continue
        m[ra.lower()] = rb
    return m


def _standardize(val: str, mapping: dict[str, str]) -> str:
    s = _norm(val)
    if not s or s.lower() == "nan":
        return ""
    return mapping.get(s.lower(), s)


def _read_list_column(wb, sheet_name: str) -> list[str]:
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    out = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        s = str(v).strip()
        if s and s.lower() != "nan":
            out.append(s)
    return out



def _read_variant_weight_map(wb, sheet_name: str = "Variant Weight (Grams)") -> dict[str, str]:
    """
    Map Custom Product Type -> Variant Weight (Grams)
    from Help Data sheet "Variant Weight (Grams)".
    """
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]
    m: dict[str, str] = {}
    for r in range(2, ws.max_row + 1):
        k = ws.cell(row=r, column=1).value
        v = ws.cell(row=r, column=2).value
        if k is None or v is None:
            continue
        ks = str(k).strip()
        if not ks or ks.lower() == "nan":
            continue
        # keep grams as string (Shopify expects a number; empty string means "unknown")
        if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
            vs = str(int(v)) if float(v).is_integer() else str(v)
        else:
            vs = str(v).strip()
        if vs.lower() == "nan":
            continue
        m[ks.lower()] = vs
    return m

def _read_category_rows(wb, sheet_name: str):
    """returns list[(name_keywords, id)] from columns A,B. Handles sheets with or without headers."""
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]

    # Detect header row (light heuristic)
    a1 = ws.cell(row=1, column=1).value
    b1 = ws.cell(row=1, column=2).value
    start_row = 1
    if isinstance(a1, str) and a1.strip().lower() in {"name", "keyword", "category", "product category"}:
        start_row = 2
    if isinstance(b1, str) and b1.strip().lower() in {"id", "category id"}:
        start_row = 2

    rows = []
    for r in range(start_row, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if a is None:
            continue
        aa = str(a).strip()
        bb = "" if b is None else str(b).strip()
        if aa and aa.lower() != "nan":
            rows.append((aa, bb))
    return rows or None


def _read_brand_line_map(wb, sheet_name: str) -> dict[str, str]:
    """Col A = brand, Col B+ concatenated text parts"""
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]
    m = {}
    for r in range(2, ws.max_row + 1):
        brand = ws.cell(row=r, column=1).value
        if brand is None:
            continue
        b = str(brand).strip()
        if not b or b.lower() == "nan":
            continue

        parts = []
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            s = str(v).strip()
            if s and s.lower() != "nan":
                parts.append(s)

        if parts:
            # IMPORTANT: keep exact spacing; caller decides punctuation
            m[b.lower()] = " ".join(parts).strip()
    return m


def _read_size_reco_map(wb) -> dict[str, str]:
    """Garment -> Comment"""
    if "Size Recommandation" not in wb.sheetnames:
        return {}
    ws = wb["Size Recommandation"]

    headers = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if h is None:
            continue
        headers[str(h).strip().lower()] = c

    gcol = headers.get("garment")
    ccol = headers.get("comment")
    if not gcol or not ccol:
        return {}

    m = {}
    for r in range(2, ws.max_row + 1):
        g = ws.cell(row=r, column=gcol).value
        c = ws.cell(row=r, column=ccol).value
        if g is None or c is None:
            continue
        gs = str(g).strip()
        cs = str(c).strip()
        if gs and cs:
            m[gs.lower()] = cs
    return m


# ---------------------------------------------------------
# Matching functions
# ---------------------------------------------------------
def _best_match_id(text: str, cat_rows) -> str:
    """
    Exact-match (loose singular/plural): all words in name must be in text.
    Returns ID (col B).

    Special rule:
    - If text contains "long sleeve" but no specific garment match is found,
      ALWAYS try "long sleeve jersey" (never tshirt).
    """
    if not cat_rows:
        return ""

    def _match(t: str) -> str:
        tset = _wordset_loose(t)
        best_id = ""
        best_len = 0
        for name, cid in cat_rows:
            nset = _wordset_loose(name)
            if nset and nset.issubset(tset):
                if len(nset) > best_len:
                    best_len = len(nset)
                    best_id = str(cid or "").strip()
        best_id = re.sub(r"\.0$", "", best_id) if best_id else ""
        return best_id

    # 1) normal match
    got = _match(text)
    if got:
        return got

    # 2) LONG SLEEVE fallback -> ALWAYS Jersey
    w = _wordset_loose(text)
    if {"long", "sleeve"}.issubset(w):
        got = _match(f"{text} jersey")
        if got:
            return got

    return ""

    def _match(t: str) -> str:
        tset = _wordset_loose(t)
        best_id = ""
        best_len = 0
        for name, cid in cat_rows:
            nset = _wordset_loose(name)
            if nset and nset.issubset(tset):
                if len(nset) > best_len:
                    best_len = len(nset)
                    best_id = str(cid or "").strip()
        best_id = re.sub(r"\.0$", "", best_id) if best_id else ""
        return best_id

    # 1) normal match
    got = _match(text)
    if got:
        return got

    # 2) LONG SLEEVE fallback: if only "long sleeve" appears, we try common garment types
    w = _wordset_loose(text)
    if {"long", "sleeve"}.issubset(w):
        # Prefer T-Shirt when no further hint is present
        got = _match(f"{text} tshirt")
        if got:
            return got
        got = _match(f"{text} jersey")
        if got:
            return got

    return ""


def _best_match_product_type(text: str, product_types: list[str]) -> str:
    """
    Match product type by word-subset (loose singular/plural).

    Special rule:
    - If text contains "long sleeve" but no specific garment match is found,
      ALWAYS try "long sleeve jersey" (never tshirt).
    """
    def _match(t: str) -> str:
        tset = _wordset_loose(t)
        best = ""
        best_len = 0
        for pt in product_types:
            pset = _wordset_loose(pt)
            if pset and pset.issubset(tset):
                if len(pset) > best_len:
                    best_len = len(pset)
                    best = pt
        return best

    # 1) normal match
    got = _match(text)
    if got:
        return got

    # 2) LONG SLEEVE fallback -> ALWAYS Jersey
    w = _wordset_loose(text)
    if {"long", "sleeve"}.issubset(w):
        got = _match(f"{text} jersey")
        if got:
            return got

    return ""


# ---------------------------------------------------------
# Parsing & formatting
# ---------------------------------------------------------
def _extract_color_size_from_description(desc: str) -> tuple[str, str]:
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


def _round_to_nearest_9_99(price) -> float:
    if price is None or (isinstance(price, float) and math.isnan(price)):
        return float("nan")
    p = float(price)
    nearest10 = math.floor(p / 10.0 + 0.5) * 10.0
    return round(nearest10 - 0.01, 2)


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


def _hs_code_clean(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""
    return re.sub(r"\.0$", "", s)


# ---------------------------------------------------------
# Excel highlighting
# ---------------------------------------------------------
def _apply_yellow_for_empty(buffer: io.BytesIO, sheet_name: str, cols_to_yellow: list[str]) -> io.BytesIO:
    buffer.seek(0)
    wb = load_workbook(buffer)
    ws = wb[sheet_name]

    header = [cell.value for cell in ws[1]]
    col_index = {name: i + 1 for i, name in enumerate(header) if name}

    for col_name in cols_to_yellow:
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


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
def run_transform(
    supplier_xlsx_bytes: bytes,
    help_xlsx_bytes: bytes,
    vendor_name: str,
    brand_choice: str = "",
    event_promo_tag: str = "",
    style_season_map: dict[str, str] | None = None,
    existing_shopify_xlsx_bytes: bytes | None = None,
):
    # Defensive defaults (avoid NameError when price columns absent)
    detected_cost_col = None
    detected_price_col = None
    warnings: list[dict] = []

    style_season_map = style_season_map or {}
    style_season_map = { _clean_style_key(k): v for k, v in style_season_map.items() }

    # -----------------------------------------------------
    # Supplier reader (multi-sheet capable)
    # -----------------------------------------------------
    def _read_supplier_multi_sheet(xlsx_bytes: bytes) -> pd.DataFrame:
        """
        Reads supplier XLSX.
        - If there are multiple sheets, keep only sheets that contain the minimum required columns
          (Description-like), then concatenate.
        - If there is a single valid sheet, behaves like the previous implementation.
        """
        bio = io.BytesIO(xlsx_bytes)
        xls = pd.ExcelFile(bio)

        # Column candidates duplicated from the main logic (kept local to avoid refactors).
        desc_candidates = [
            "description", "Description", "Product Name", "product name",
            "Title", "title", "Style", "style", "Style Name", "style name",
            "Display Name", "display name", "Online Display Name", "online display name",
            "Technical Specifications", "technical specifications",
        ]
        msrp_candidates = [
            "Cad MSRP", "MSRP", "Retail Price (CAD)", "retail price (CAD)", "retail price (cad)",
        ]

        dfs: list[pd.DataFrame] = []
        for sn in xls.sheet_names:
            df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sn, dtype=str)
            # --- Price columns (robust) ---
            cost_col = _find_col(df.columns, [
                "Wholesale CAD", "Wholesale (CAD)", "CAD Wholesale", "WholesaleCAD", "wholesale cad"
            ])
            price_col = _find_col(df.columns, [
                "Retail CAD", "Retail (CAD)", "CAD Retail", "RetailCAD", "retail cad"
            ])

            # Legacy MSRP-like columns (optional)
            msrp_col = _find_col(df.columns, [
                "Retail Price (CAD)", "Cad MSRP", "MSRP", "msrp"
            ])

            # Prefer explicit Norda CAD columns
            detected_cost_col = cost_col if cost_col else None
            if price_col:
                detected_price_col = price_col
            elif msrp_col:
                detected_price_col = msrp_col
            else:
                detected_price_col = None

            # Drop fully empty rows early
            if df is None or df.empty:
                warnings.append({
                    "type": "sheet_skipped",
                    "sheet": sn,
                    "reason": "empty",
                })
                continue
            df = df.dropna(how="all")
            if df.empty:
                warnings.append({
                    "type": "sheet_skipped",
                    "sheet": sn,
                    "reason": "empty",
                })
                continue

            # Validate minimum required columns
            has_desc = _first_existing_col(df, desc_candidates) is not None
            has_msrp = _first_existing_col(df, msrp_candidates) is not None
            if not has_desc:
                warnings.append({
                    "type": "sheet_skipped",
                    "sheet": sn,
                    "reason": "missing_required_columns",
                    "has_desc": has_desc,
                    "has_msrp": has_msrp,
                })
                continue

            df["_source_sheet"] = sn
            dfs.append(df)

        if not dfs:
            raise ValueError(
                "Aucun onglet valide détecté dans le fichier fournisseur (colonne Description requise)."
            )
        return pd.concat(dfs, ignore_index=True, sort=False)

    sup = _read_supplier_multi_sheet(supplier_xlsx_bytes).copy()


    # -----------------------------------------------------
    # Detect price columns on the concatenated supplier dataframe
    # (fix: detected_* set inside sheet loop are not propagated)
    # -----------------------------------------------------
    detected_cost_col = _find_col(sup.columns, [
        "Wholesale CAD", "Wholesale (CAD)", "CAD Wholesale", "WholesaleCAD", "wholesale cad",
        "Landed", "landed", "Wholesale Price (CAD)", "wholesale price (cad)", "Wholesale Price", "wholesale price"
    ])
    detected_price_col = _find_col(sup.columns, [
        "Retail CAD", "Retail (CAD)", "CAD Retail", "RetailCAD", "retail cad",
        "Retail Price (CAD)", "Cad MSRP", "MSRP", "msrp"
    ])

    wb = _load_help_wb(help_xlsx_bytes)

    # Standardization
    color_map = _read_2col_map(wb, ["Color Standardization", "Color Variable"])
    size_map = _read_2col_map(wb, ["Size Standardization", "Size Variante"])
    country_map = _read_2col_map(wb, ["Country Abbreviations", "Country of Origin"])
    gender_map = _read_2col_map(wb, ["Gender Standardization", "Gender"])

    # Categories & Product types
    shopify_cat_rows = _read_category_rows(wb, "Shopify Product Category")
    google_cat_rows = _read_category_rows(wb, "Google Product Category")
    product_types = _read_list_column(wb, "Product Types")
    variant_weight_map = _read_variant_weight_map(wb)


    # Brand maps
    brand_desc_map = _read_brand_line_map(wb, "SEO Description Brand Part")
    brand_lines_map = _read_brand_line_map(wb, "Brand lines")

    # Size reco
    size_comment_map = _read_size_reco_map(wb)

    # Supplier columns
    desc_col = _first_existing_col(
        sup,
        [
            "Description", "description",
            "Product Details", "product details",
            "Technical Specifications", "technical specifications",
            "Product Name", "product name",
            "Title", "title", "Style", "style", "Style Name", "style name",
            "Display Name", "display name", "Online Display Name", "online display name",
        ],
    )

    # If we picked Technical Specifications but it is mostly empty, fallback to Description when available.
    desc_col_fallback = _first_existing_col(sup, ["Description", "description"])
    if desc_col and _colkey(desc_col) in ("technicalspecifications", "technicalspecification") and desc_col_fallback:
        non_empty_ratio = sup[desc_col].astype(str).fillna("").str.strip().ne("").mean() if len(sup) else 0
        if non_empty_ratio < 0.2:
            desc_col = desc_col_fallback

    product_col = _first_existing_col(sup, ["Product", "Product Code", "SKU", "sku"])
    color_col = _first_existing_col(sup, ["Vendor Color", "vendor color", "Color", "color", "Colour", "colour", "Color Code", "color code"])
    size_col = _first_existing_col(sup, ["Size 1","Size1","Size", "size", "Vendor Size1", "vendor size1"])
    upc_col = _first_existing_col(sup, ["UPC", "UPC Code", "UPC Code 1", "UPC Code1", "UPC1", "Variant Barcode", "Barcode", "bar code", "upc", "upc code"])
    origin_col = _first_existing_col(sup, ["Country Code", "Origin", "Manufacturing Country", "COO", "country code", "origin", "manufacturing country", "coo"])
    hs_col = _first_existing_col(sup, ["HS Code", "HTS Code", "hs code", "hts code"])
    extid_col = _first_existing_col(sup, ["External ID", "ExternalID"])
    msrp_col = _first_existing_col(sup, ["Cad MSRP", "MSRP", "Retail Price (CAD)", "retail price (CAD)", "retail price (cad)"])
    landed_col = _first_existing_col(sup, ["Landed", "landed", "Wholesale Price", "wholesale price", "Wholesale Price (CAD)", "wholesale price (cad)"])
    grams_col = _first_existing_col(sup, ["Grams", "Weight (g)", "Weight"])
    gender_col = _first_existing_col(sup, ["Gender", "gender", "Genre", "genre", "Sex", "sex", "Sexe", "sexe"])


    # -----------------------------------------------------
    # Gender inference: detect "-w-" / "- W -" / "-m-" / "- M -" in Name or SKU
    # -----------------------------------------------------
    name_hint_col = _first_existing_col(sup, ["Style Name", "Name", "Product Name", "Title", "Style"])
    sku_hint_col = extid_col or product_col

    def _infer_gender_from_texts(name_val: str, sku_val: str) -> str:
        t = f"{_norm(name_val)} {_norm(sku_val)}".lower()
        if re.search(r"-\s*w\s*-", t):
            return "Women"
        if re.search(r"-\s*m\s*-", t):
            return "Men"
        return ""

    if desc_col is None:
        raise ValueError(
            "Colonne Description introuvable. Colonnes acceptées: Description, Style, Style Name, Product Name, Title, Display Name, Online Display Name."
        )
    if msrp_col is None:
        msrp_col = None  # MSRP not found; leave prices blank per rules

    # -----------------------------------------------------
    # De-duplicate across sheets (SKU and/or UPC)
    # -----------------------------------------------------
    def _clean_sku(x) -> str:
        s = _norm(x)
        return re.sub(r"\.0$", "", s)

    def _make_dedupe_key(r) -> str:
        sku = _clean_sku(r.get(extid_col, "")) if extid_col else ""
        if not sku:
            sku = _clean_sku(r.get(product_col, "")) if product_col else ""
        upc = _barcode_keep_zeros(r.get(upc_col, "")) if upc_col else ""

        if sku and upc:
            return f"{sku}|{upc}"
        if sku:
            return sku
        if upc:
            return upc
        # Fallback (rare): keep row unique if neither exists
        return ""

    sup["_dedupe_key"] = sup.apply(_make_dedupe_key, axis=1)

    before = len(sup)
    # Only dedupe rows where we have at least one identifier; keep all others.
    has_key = sup["_dedupe_key"].astype(str).str.strip().ne("")
    sup_keyed = sup.loc[has_key].drop_duplicates(subset=["_dedupe_key"], keep="first")
    sup_unkeyed = sup.loc[~has_key]
    sup = pd.concat([sup_keyed, sup_unkeyed], ignore_index=True)
    after = len(sup)
    if after < before:
        warnings.append({
            "type": "dedupe_applied",
            "dedupe_by": "SKU and/or UPC",
            "rows_before": before,
            "rows_after": after,
            "rows_removed": before - after,
        })

    
    # Base description (keep both a normalized version and the original source text)
    sup["_desc_source"] = sup[desc_col].astype(str).fillna("")  # preserve original (length, punctuation, line breaks)
    sup["_desc_raw"] = sup["_desc_source"].map(_norm)
    sup["_desc_seo"] = sup["_desc_raw"].apply(_convert_r_to_registered)
    sup["_desc_handle"] = sup.apply(lambda r: _strip_reg_for_handle(r["_title_name_raw"]) if r.get("_desc_is_long") and r.get("_title_name_raw") else _strip_reg_for_handle(r["_desc_raw"]), axis=1)

    # -----------------------------------------------------
    # Long description rule:
    # If the SOURCE description text is > 200 chars, move it to Body (HTML)
    # and build Title from Style Name / Name instead of the long description.
    # -----------------------------------------------------
    title_name_col = _first_existing_col(sup, ["Style Name", "Name", "Product Name", "Title", "Style"])
    sup["_title_name_raw"] = sup[title_name_col].astype(str).fillna("").map(_norm) if title_name_col else ""

    sup["_desc_is_long"] = sup["_desc_source"].apply(lambda x: len(str(x)) > 200)

    # Put the original description in Body (HTML) when long (not the normalized one)
    sup["_body_html"] = sup.apply(lambda r: str(r["_desc_source"]).strip() if r["_desc_is_long"] else "", axis=1)
    # Color / Size input
    sup["_color_raw"] = sup[color_col].astype(str).fillna("").map(_norm) if color_col else ""
    sup["_size_raw"] = sup[size_col].astype(str).fillna("").map(_norm) if size_col else ""

    # Fallback parse from description if missing
    parsed = sup["_desc_raw"].apply(_extract_color_size_from_description)
    sup["_color_fb"] = parsed.map(lambda t: t[0])
    sup["_size_fb"] = parsed.map(lambda t: t[1])

    sup["_color_in"] = sup["_color_raw"]
    sup.loc[sup["_color_in"].eq(""), "_color_in"] = sup["_color_fb"]

    sup["_size_in"] = sup["_size_raw"]
    sup.loc[sup["_size_in"].eq(""), "_size_in"] = sup["_size_fb"]

    # Standardize
    sup["_color_std"] = sup["_color_in"].apply(lambda x: _standardize(x, color_map))
    sup["_size_std"] = sup["_size_in"].apply(lambda x: _standardize(x, size_map))

    # Gender (standardize if possible)
    sup["_gender_raw"] = sup[gender_col].astype(str).fillna("").map(_norm) if gender_col else ""

    sup["_gender_inferred"] = sup.apply(
        lambda r: _infer_gender_from_texts(
            r.get(name_hint_col, "") if name_hint_col else "",
            r.get(sku_hint_col, "") if sku_hint_col else "",
        ),
        axis=1,
    )
    sup.loc[sup["_gender_raw"].astype(str).str.strip().eq(""), "_gender_raw"] = sup.loc[
        sup["_gender_raw"].astype(str).str.strip().eq(""),
        "_gender_inferred",
    ]

    sup["_gender_std"] = sup["_gender_raw"].apply(lambda x: _standardize(x, gender_map)) if gender_map else sup["_gender_raw"]

    # Vendor / Brand
    sup["_vendor"] = vendor_name
    sup["_brand_choice"] = _norm(brand_choice)

    # Title: Gender('s) + Description - Color (NON-standardized, Title Cased)
    def _gender_for_title(g: str) -> str:
        gg = _norm(g)
        if gg.lower() in ("men", "women"):
            gg = f"{gg}'s"
        return _title_case_preserve_registered(gg)

    sup["_gender_title"] = sup["_gender_std"].astype(str).fillna("").map(_gender_for_title)
    sup["_desc_title"] = sup.apply(lambda r: _title_case_preserve_registered(r["_title_name_raw"]) if r.get("_desc_is_long") and r["_title_name_raw"] else _title_case_preserve_registered(r["_desc_seo"]), axis=1)
    sup["_color_title"] = sup["_color_in"].astype(str).fillna("").map(_title_case_preserve_registered)

    sup["_title"] = (sup["_gender_title"].str.strip() + " " + sup["_desc_title"].str.strip()).str.strip()
    sup.loc[sup["_color_title"].str.strip().ne(""), "_title"] = sup.apply(
        lambda r: r["_title"]
        if _norm(r["_color_title"]) and _norm(r["_color_title"]) in _norm(r["_title"])
        else (r["_title"].strip() + " - " + r["_color_title"].strip()).strip(),
        axis=1,
    )

    # Handle: Vendor + Gender + Description + Color (color NON-standardized)
    def _make_handle(r):
        # When description is long and moved to Body (HTML), build handle from Style Name/Name (same rule as Title)
        base_text = r.get("_title_name_raw") if r.get("_desc_is_long") and r.get("_title_name_raw") else r.get("_desc_handle")
        base_text = _strip_gender_tokens(base_text)
        desc_for_handle = _strip_reg_for_handle(base_text)
        color_for_handle = _strip_reg_for_handle(r.get("_color_in", ""))
        # Avoid duplicating color if it's already present in the base text
        if color_for_handle and _norm(color_for_handle).lower() in _norm(desc_for_handle).lower():
            color_for_handle = ""
        parts = [
            _strip_reg_for_handle(r.get("_vendor")),
            _strip_reg_for_handle(r.get("_gender_std")),
            desc_for_handle,
            color_for_handle,
        ]
        parts = [p for p in parts if p and str(p).strip()]
        return slugify(" ".join(parts))
    sup["_handle"] = sup.apply(_make_handle, axis=1)

    # Custom Product Type: match using DESCRIPTION (to catch TEE / LONG SLEEVE etc.)
    sup["_product_type"] = sup["_desc_raw"].apply(lambda t: _best_match_product_type(t, product_types))

    # Tags (keep standardized color/gender tags)
    # -----------------------------------------------------
    # Seasonality key (to apply Seasonality Tags per style)
    # -----------------------------------------------------
    style_num_col = _first_existing_col(sup, ["Style Number", "Style Num", "Style #", "style number", "style #", "Style"])
    style_name_col = _first_existing_col(sup, ["Style Name", "style name", "Product Name", "Name"])
    sup["_seasonality_key"] = ""
    if style_num_col is not None:
        sup["_seasonality_key"] = sup[style_num_col].astype(str).fillna("").map(_clean_style_key)
    elif style_name_col is not None:
        sup["_seasonality_key"] = sup[style_name_col].astype(str).fillna("").map(_clean_style_key)

    def _make_tags(r):
        tags = []
        if r["_vendor"]:
            tags.append(r["_vendor"])
        if r["_color_std"]:
            tags.append(r["_color_std"])
        # Colour-based tags
        if _norm(r["_color_std"]).lower() == "black":
            tags.append("Core")
        else:
            tags.append("Seasonal")

        if r["_gender_std"]:
            tags.append(r["_gender_std"])
        tags.append("_badge_new")
        if r["_product_type"]:
            tags.append(r["_product_type"])
        # Event/Promotion Related (applies to entire file)
        if event_promo_tag:
            tags.append(event_promo_tag)

        # Seasonality tag (per style)
        stg = style_season_map.get(_clean_style_key(r.get("_seasonality_key", "")))
        if stg:
            tags.append(stg)

        return ", ".join([t for t in tags if t])

    sup["_tags"] = sup.apply(_make_tags, axis=1)

    # SKU
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

    # Barcode
    sup["_barcode"] = sup[upc_col].apply(_barcode_keep_zeros) if upc_col else ""

    # Country (standardize)
    sup["_origin_raw"] = sup[origin_col].astype(str).fillna("").map(_norm) if origin_col else ""
    sup["_origin_std"] = sup["_origin_raw"].apply(lambda x: _standardize(x, country_map))

    # HS Code
    sup["_hs"] = sup[hs_col].apply(_hs_code_clean) if hs_col else ""

    # Grams
    if grams_col:
        sup["_grams"] = sup[grams_col].astype(str).fillna("").map(_norm)
    else:
        # Fallback: use Help Data -> "Variant Weight (Grams)" mapped by Custom Product Type
        sup["_grams"] = sup["_product_type"].apply(lambda pt: variant_weight_map.get(str(pt).strip().lower(), "") if pt else "")

    # Price
    if detected_price_col is not None and _header_has_cad(detected_price_col):
        price_num = pd.to_numeric(
            sup[detected_price_col].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce",
        )
        sup["_price"] = price_num.apply(_round_to_nearest_9_99)
    else:
        sup["_price"] = ""

    # Cost (leave blank unless CAD column detected per rules)
    if detected_cost_col is not None and _header_has_cad(detected_cost_col):
        sup["_cost"] = sup[detected_cost_col].astype(str).fillna("").map(_norm)
    else:
        sup["_cost"] = ""

    # Size comment
    def _size_comment(r):
        key = (r["_brand_choice"] or r["_vendor"]).strip().lower()
        return size_comment_map.get(key, "")

    sup["_size_comment"] = sup.apply(_size_comment, axis=1)

    # Categories: match using DESCRIPTION (to catch LONG SLEEVE, TEE → tshirt)
    sup["_shopify_cat_id"] = sup["_desc_raw"].apply(lambda t: _best_match_id(t, shopify_cat_rows))
    sup["_google_cat_id"] = sup["_desc_raw"].apply(lambda t: _best_match_id(t, google_cat_rows))

    # Siblings
    sup["_siblings"] = sup["_handle"]

    # SEO Title (adds 's for Men/Women, Title Case)
    def _seo_title(r):
        g = _norm(r["_gender_std"])
        if g.lower() in ("men", "women"):
            g = f"{g}'s"
        main = f"{r['_vendor']} {g} {r['_desc_seo']}".strip()
        main = _title_case_preserve_registered(main)
        color = _title_case_preserve_registered(r["_color_std"])
        return f"{main} - {color}".strip() if color else main

    sup["_seo_title"] = sup.apply(_seo_title, axis=1)

    # SEO Description rules
    def _seo_desc(r):
        prefix = f"Shop the {r['_seo_title']} with free worldwide shipping, and 30-day returns on leclub.cc. "
        brand_name = _norm(r["_brand_choice"] or r["_vendor"])
        brand_disp = _title_case_preserve_registered(brand_name)

        bkey = brand_name.strip().lower()
        if bkey and bkey in brand_desc_map:
            part = _norm(brand_desc_map[bkey]).rstrip().rstrip(".")
            return f"{prefix}Discover {brand_disp} {part}."
        return f"{prefix}Discover {brand_disp} products."

    sup["_seo_desc"] = sup.apply(_seo_desc, axis=1)

    # behind the brand
    def _behind_brand(r):
        bkey = (r["_brand_choice"] or "").strip().lower()
        return brand_lines_map.get(bkey, "") if bkey else ""

    sup["_behind_the_brand"] = sup.apply(_behind_brand, axis=1)

    # ---------------------------------------------------------
    # Build output (strict order)
    # ---------------------------------------------------------
    out = pd.DataFrame(columns=SHOPIFY_OUTPUT_COLUMNS)

    out["Handle"] = sup["_handle"]
    out["Command"] = "NEW"
    out["Title"] = sup["_title"]
    out["Body (HTML)"] = sup["_body_html"]
    out["Vendor"] = sup["_vendor"]
    out["Custom Product Type"] = sup["_product_type"]
    out["Tags"] = sup["_tags"]

    out["Published"] = False
    out["Published Scope"] = "global"

    out["Option1 Name"] = "Size"
    out["Option1 Value"] = sup["_size_std"].map(_strip_gender_prefix_size)

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

    out["Metafield: my_fields.product_use_case [multi_line_text_field]"] = ""
    out["Metafield: my_fields.product_features [multi_line_text_field]"] = ""
    out["Metafield: my_fields.behind_the_brand [multi_line_text_field]"] = sup["_behind_the_brand"]
    out["Metafield: my_fields.size_comment [single_line_text_field]"] = sup["_size_comment"]
    out["Metafield: my_fields.gender [single_line_text_field]"] = sup["_gender_std"]

    out["Metafield: my_fields.colour [single_line_text_field]"] = sup["_color_std"]
    out["Metafield: mm-google-shopping.color"] = sup["_color_std"]
    out["Variant Metafield: mm-google-shopping.size"] = sup["_size_std"].map(_strip_gender_prefix_size)

    out["Metafield: mm-google-shopping.size_system"] = "US"
    out["Metafield: mm-google-shopping.condition"] = "new"
    out["Metafield: mm-google-shopping.google_product_category"] = sup["_google_cat_id"]
    out["Metafield: mm-google-shopping.gender"] = sup["_gender_std"]

    out["Variant Metafield: mm-google-shopping.mpn"] = sup["_variant_sku"]
    out["Variant Metafield: mm-google-shopping.gtin"] = sup["_barcode"]

    out["Metafield: theme.siblings [single_line_text_field]"] = sup["_siblings"]
    out["Category: ID"] = sup["_shopify_cat_id"]

    out["Inventory Available: Boutique"] = 0
    out["Inventory Available: Le Club"] = 0

    out = out.reindex(columns=SHOPIFY_OUTPUT_COLUMNS)

    # Yellow rules
    yellow_if_empty_cols = [
        "Handle",
        "Title",
        "Body (HTML)",
        "Custom Product Type",
        "Option1 Name",
        "Option1 Value",
        "Variant Price",
        "Variant Grams",
        "Variant Country of Origin",
        "Variant HS Code",
        "SEO Title",
        "SEO Description",
        "Metafield: my_fields.size_comment [single_line_text_field]",
        "Metafield: my_fields.gender [single_line_text_field]",
        "Metafield: my_fields.colour [single_line_text_field]",
        "Metafield: mm-google-shopping.color",
        "Variant Metafield: mm-google-shopping.size",
        "Metafield: mm-google-shopping.google_product_category",
        "Category: ID",
    ]

    def _apply_red_font_for_tags(buffer: io.BytesIO, sheet_name: str, rows_to_color_red: list[int]) -> io.BytesIO:
        """Apply red font to the 'Tags' cell for given 0-based dataframe row indexes."""
        buffer.seek(0)
        wb = openpyxl.load_workbook(buffer)
        if sheet_name not in wb.sheetnames:
            return buffer
        ws = wb[sheet_name]

        # Find Tags column index from header row (row 1)
        tags_col_idx = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if str(v).strip() == "Tags":
                tags_col_idx = c
                break
        if tags_col_idx is None:
            return buffer

        red_font = openpyxl.styles.Font(color="FFFF0000")

        # Data rows start at Excel row 2
        if not rows_to_color_red:
            rows_to_color_red = []
            for excel_row in range(2, ws.max_row + 1):
                v = ws.cell(row=excel_row, column=tags_col_idx).value
                if v and "seasonal" in str(v).lower():
                    rows_to_color_red.append(excel_row - 2)

        for df_i in rows_to_color_red:
            excel_row = df_i + 2
            cell = ws.cell(row=excel_row, column=tags_col_idx)
            cell.font = red_font

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        # Split into "products" and "do not import" based on existing Shopify file (if provided)
        existing_handles_set, existing_key_sets = _build_existing_shopify_index(existing_shopify_xlsx_bytes)

        # columns expected in output
        vendor_col = "Vendor" if "Vendor" in out.columns else None
        sku_col = "Variant SKU" if "Variant SKU" in out.columns else ("SKU" if "SKU" in out.columns else None)
        upc_col = "Variant Barcode" if "Variant Barcode" in out.columns else ("Barcode" if "Barcode" in out.columns else ("UPC" if "UPC" in out.columns else None))

        def _getcol(r, c):
            return r.get(c, "") if c else ""

        mask_existing = []
        for _, r in out.iterrows():
            brand = _getcol(r, vendor_col) or vendor_name
            sku = _getcol(r, sku_col)
            upc = _getcol(r, upc_col)
            mask_existing.append(_row_is_existing(str(brand), str(sku), str(upc), existing_key_sets))

        mask_existing = pd.Series(mask_existing, index=out.index)

        products_df = out.loc[~mask_existing].copy()
        do_not_import_df = out.loc[mask_existing].copy()

        products_df.to_excel(writer, index=False, sheet_name="products")
        do_not_import_df.to_excel(writer, index=False, sheet_name="do not import")
        pd.DataFrame(warnings).to_excel(writer, index=False, sheet_name="warnings")

    
    # Red font for Tags when colour is NOT Black (i.e., Seasonal) — apply on both sheets
    existing_handles_set, existing_key_sets = _build_existing_shopify_index(existing_shopify_xlsx_bytes)

    def _rows_to_color_for_df(df_slice: pd.DataFrame) -> list[int]:
        if "_color_std" not in df_slice.columns:
            return []
        return [
            i
            for i, c in enumerate(df_slice["_color_std"].astype(str).tolist())
            if _norm(c) != "" and _norm(c).lower() != "black"
        ]

    # For handle red: when output handle already exists in Shopify
    def _rows_handle_conflict(df_slice: pd.DataFrame) -> list[int]:
        """Rows where Handle conflicts with an existing Shopify handle."""
        if "Handle" not in df_slice.columns:
            return []
        handles_norm = df_slice["Handle"].apply(_norm_handle)
        mask = handles_norm.isin(existing_handles_set) & handles_norm.ne("")
        return [i for i, v in enumerate(mask.tolist()) if v]
        return [i for i, h in enumerate(df_slice["Handle"].astype(str).tolist()) if _norm(h) in existing_handles_set and _norm(h) != ""]

    # Reload workbook buffer as BytesIO for styling helpers
    buffer.seek(0)

    # Apply tag red and yellow empty on each sheet
    buffer = _apply_red_font_for_tags(buffer, "products", _rows_to_color_for_df(products_df))
    buffer = _apply_red_font_for_tags(buffer, "do not import", _rows_to_color_for_df(do_not_import_df))

    buffer = _apply_yellow_for_empty(buffer, "products", yellow_if_empty_cols)
    buffer = _apply_yellow_for_empty(buffer, "do not import", yellow_if_empty_cols)

    # Apply red font for handle conflicts (only the cell in Handle column)
    buffer = _apply_red_font_for_handle(buffer, "products", _rows_handle_conflict(products_df))
    buffer = _apply_red_font_for_handle(buffer, "do not import", _rows_handle_conflict(do_not_import_df))

    return buffer.getvalue(), pd.DataFrame(warnings)
