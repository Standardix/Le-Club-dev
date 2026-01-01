# fournisseur_abc.py
# Clean, indentation-safe final version

import io
import re
import math
import pandas as pd
import openpyxl
from slugify import slugify
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

SHOPIFY_OUTPUT_COLUMNS = [
    "Handle","Command","Title","Body (HTML)","Vendor","Custom Product Type","Tags",
    "Published","Published Scope","Option1 Name","Option1 Value","Variant SKU",
    "Variant Barcode","Variant Country of Origin","Variant HS Code","Variant Grams",
    "Variant Inventory Tracker","Variant Inventory Policy","Variant Fulfillment Service",
    "Variant Price","Variant Requires Shipping","Variant Taxable","SEO Title",
    "SEO Description","Variant Weight Unit","Cost per item","Status",
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
    "Category: ID","Inventory Available: Boutique","Inventory Available: Le Club"
]

def _norm(s):
    return re.sub(r"\s+", " ", str(s or "").strip())

def _title_case(text):
    text = _norm(text)
    if not text:
        return ""
    def tc(tok):
        if any(c.isdigit() for c in tok):
            return tok
        for sep in ["/","-"]:
            if sep in tok:
                return sep.join(tc(p) for p in tok.split(sep))
        return tok[:1].upper() + tok[1:].lower()
    return " ".join(tc(w) for w in text.split(" "))

def _normalize_match_text(s):
    t = str(s or "").lower()
    t = re.sub(r"\btee(s)?\b", "tshirt", t)
    t = re.sub(r"\bt\s*[- ]\s*shirt\b", "tshirt", t)
    t = re.sub(r"\blong\s*[- ]\s*sleeve\b", "long sleeve", t)
    return t

def _words(s):
    return re.findall(r"[a-z0-9]+", _normalize_match_text(s))

def _wordset(s):
    return set(_words(s))

def _best_match_id(text, rows):
    if not rows:
        return ""
    tset = _wordset(text)
    best = ("",0)
    for name,cid in rows:
        nset = _wordset(name)
        if nset and nset.issubset(tset) and len(nset) > best[1]:
            best = (cid,len(nset))
    return str(best[0]).replace(".0","") if best[0] else ""

def _read_category_rows(wb, sheet):
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    rows=[]
    for r in range(1, ws.max_row+1):
        a = ws.cell(r,1).value
        b = ws.cell(r,2).value
        if a:
            rows.append((str(a),str(b or "")))
    return rows

def run_transform(supplier_xlsx_bytes, help_xlsx_bytes, vendor_name, brand_choice=""):
    sup = pd.read_excel(io.BytesIO(supplier_xlsx_bytes), dtype=str).fillna("")
    wb = openpyxl.load_workbook(io.BytesIO(help_xlsx_bytes), data_only=True)

    shopify_cat = _read_category_rows(wb,"Shopify Product Category")
    google_cat = _read_category_rows(wb,"Google Product Category")

    desc_col = next(c for c in sup.columns if c.lower() in ["description","product name","title","display name"])
    color_col = next((c for c in sup.columns if "color" in c.lower()), None)
    gender_col = next((c for c in sup.columns if c.lower()=="gender"), None)

    sup["_desc"] = sup[desc_col]
    sup["_color_raw"] = sup[color_col] if color_col else ""
    sup["_gender"] = sup[gender_col] if gender_col else ""

    def gender_disp(g):
        g=_norm(g)
        if g.lower() in ("men","women"):
            return f"{g}'s"
        return g

    sup["_title"] = (
        sup["_gender"].map(gender_disp).map(_title_case).str.strip() + " " +
        sup["_desc"].map(_title_case).str.strip()
    ).str.strip()

    sup.loc[sup["_color_raw"].str.strip()!="","_title"] = (
        sup["_title"] + " - " + sup["_color_raw"].map(_title_case)
    )

    sup["_handle"] = sup.apply(
        lambda r: slugify(" ".join([vendor_name,r["_gender"],r["_desc"],r["_color_raw"]])),
        axis=1
    )

    sup["_shopify_cat_id"] = sup["_desc"].apply(lambda t: _best_match_id(t, shopify_cat))
    sup["_google_cat_id"] = sup["_desc"].apply(lambda t: _best_match_id(t, google_cat))

    out = pd.DataFrame()
    out["Handle"] = sup["_handle"]
    out["Title"] = sup["_title"]
    out["Vendor"] = vendor_name
    out["Custom Product Type"] = sup["_shopify_cat_id"]
    out["Category: ID"] = sup["_shopify_cat_id"]
    out["Metafield: mm-google-shopping.google_product_category"] = sup["_google_cat_id"]

    for col in SHOPIFY_OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""

    out = out[SHOPIFY_OUTPUT_COLUMNS]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="shopify_import")

    return buf.getvalue(), pd.DataFrame()
