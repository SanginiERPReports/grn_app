# app.py — PR–PO–GRN–GIN Linker with header validation + styled Excel output
import streamlit as st

# Display your logo image (update path if in subfolder, e.g. 'images/logo.png')
st.image("SANGINI_LOGOG.JPEG", width=150)

st.title("📘 PR–PO–GRN–GIN Linker")


import io, re, math
import streamlit as st
import pandas as pd
import xlsxwriter

st.set_page_config(page_title="PR–PO–GRN–GIN Linker", layout="wide")
st.title("📘 PR–PO–GRN–GIN Linker")

# ----------------- helpers -----------------
def clean_key(v):
    if v is None: return ""
    s = str(v).strip().upper()
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"\s+", " ", s)
    return s

def to_str(v):
    if v is None: return ""
    try:
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return ""
    except Exception:
        pass
    return v

def read_export_with_header_row7(uploaded_file):
    """Read first sheet with header row = 7 (index=6)."""
    return pd.read_excel(uploaded_file, header=6)

def safe_lookup_dict(df, key_col, value_col):
    if key_col not in df.columns or value_col not in df.columns:
        return {}
    tmp = df[[key_col, value_col]].copy()
    tmp[key_col] = tmp[key_col].map(clean_key)
    tmp = tmp[tmp[key_col] != ""]
    tmp = tmp.drop_duplicates(subset=[key_col], keep="last")
    return dict(zip(tmp[key_col], tmp[value_col]))

def check_columns(df, required, title):
    """Return (ok, missing_list). Shows a nice table in Streamlit."""
    have = list(df.columns)
    missing = [c for c in required if c not in df.columns]
    status = "✅ OK" if not missing else "❌ Missing"
    st.caption(f"**{title}** – {status}")
    left, right = st.columns([2,3])
    with left:
        st.write("Required:")
        st.code("\n".join(required))
    with right:
        st.write("Found:")
        st.code("\n".join(have[:30] + (["…"] if len(have) > 30 else [])))
    if missing:
        st.error(f"Missing columns in **{title}**: {missing}")
    return (len(missing) == 0, missing)

# ----------------- UI -----------------
c1, c2 = st.columns(2)
with c1:
    stock_file = st.file_uploader("Upload **Stock Ledger (main)**", type=["xlsx"])
    po_file    = st.file_uploader("Upload **PO Export** (header on row 7)", type=["xlsx"])
with c2:
    grn_file   = st.file_uploader("Upload **GRN Export** (header on row 7)", type=["xlsx"])
    gin_file   = st.file_uploader("Upload **GIN Export** (header on row 7)", type=["xlsx"])

if st.button("🚀 Validate & Run"):
    if not (stock_file and po_file and grn_file and gin_file):
        st.error("Please upload all **4 files**.")
        st.stop()

    # 1) Read files
    stock = pd.read_excel(stock_file)
    po    = read_export_with_header_row7(po_file)
    grn   = read_export_with_header_row7(grn_file)
    gin   = read_export_with_header_row7(gin_file)

    # 2) Validate headers
    ok_all = True

    stock_required = [
        "Date","Name of Company","Project Name","Sub Project",
        "Level","Activity Code","Activity Name","Godown Name",
        "P.O. No","G.R. No","Voucher No","From Voucher",
        "Contractor / Service Provider Name",
        "Item Group","Item Desc","Received Qty","Received Amt","Issued Qty","Issued Amt"
    ]
    ok1,_ = check_columns(stock, stock_required, "Stock Ledger")
    ok_all = ok_all and ok1

    po_required  = ["P.O. No.", "P.O. Date", "Remarks"]
    ok2,_ = check_columns(po, po_required, "PO Export (row 7 header)")
    ok_all = ok_all and ok2

    grn_required = ["G.R. No", "GRN Date", "Remarks"]
    ok3,_ = check_columns(grn, grn_required, "GRN Export (row 7 header)")
    ok_all = ok_all and ok3

    gin_required = ["G.I.N. Sr No@S/GIN/S/Y/S/5", "Issue Date", "Remarks"]
    ok4,_ = check_columns(gin, gin_required, "GIN Export (row 7 header)")
    ok_all = ok_all and ok4

    if not ok_all:
        st.stop()

    st.success("✅ All headers look good. Processing…")

    # 3) Build lookups (drop dups)
    po_date_map  = safe_lookup_dict(po,  "P.O. No.", "P.O. Date")
    po_rem_map   = safe_lookup_dict(po,  "P.O. No.", "Remarks")
    grn_date_map = safe_lookup_dict(grn, "G.R. No",  "GRN Date")
    grn_rem_map  = safe_lookup_dict(grn, "G.R. No",  "Remarks")
    gin_key_col  = "G.I.N. Sr No@S/GIN/S/Y/S/5"
    gin_date_map = safe_lookup_dict(gin, gin_key_col, "Issue Date")
    gin_rem_map  = safe_lookup_dict(gin, gin_key_col, "Remarks")

    # 4) Normalized keys in stock
    stock["_PO_KEY"]  = stock.get("P.O. No", stock.get("P.O. No.", "")).map(clean_key)
    stock["_GRN_KEY"] = stock.get("G.R. No", "").map(clean_key)
    stock["_GIN_KEY"] = stock.get("Voucher No", "").map(clean_key)

    # 5) Enrich with dates/remarks
    stock["P.O. Date"]   = stock["_PO_KEY"].map(po_date_map).fillna("")
    stock["PO_Remarks"]  = stock["_PO_KEY"].map(po_rem_map).fillna("")
    stock["GRN Date"]    = stock["_GRN_KEY"].map(grn_date_map).fillna("")
    stock["GRN_Remarks"] = stock["_GRN_KEY"].map(grn_rem_map).fillna("")
    stock["Issue Date"]  = stock["_GIN_KEY"].map(gin_date_map).fillna("")
    stock["GIN_Remarks"] = stock["_GIN_KEY"].map(gin_rem_map).fillna("")

    # Rename Voucher No → GIN No (display)
    if "Voucher No" in stock.columns:
        stock.rename(columns={"Voucher No": "GIN No"}, inplace=True)

    # 6) Compute GRN_Issued_Qty / GRN_Issued_Amt from ledger
    for c in ["Issued Qty", "Issued Amt"]:
        if c in stock.columns:
            stock[c] = pd.to_numeric(stock[c], errors="coerce").fillna(0.0)
        else:
            stock[c] = 0.0

    grn_issue_sum = (
        stock.groupby("_GRN_KEY", dropna=False)[["Issued Qty", "Issued Amt"]]
        .sum().reset_index()
        .rename(columns={"Issued Qty": "GRN_Issued_Qty", "Issued Amt": "GRN_Issued_Amt"})
    )
    grn_issue_qty_map = dict(zip(grn_issue_sum["_GRN_KEY"], grn_issue_sum["GRN_Issued_Qty"]))
    grn_issue_amt_map = dict(zip(grn_issue_sum["_GRN_KEY"], grn_issue_sum["GRN_Issued_Amt"]))
    stock["GRN_Issued_Qty"] = stock["_GRN_KEY"].map(grn_issue_qty_map).fillna(0.0)
    stock["GRN_Issued_Amt"] = stock["_GRN_KEY"].map(grn_issue_amt_map).fillna(0.0)

    # 7) OUTPUT: Summary + RAW_Data
    headers = [
        "Date","Name of Company","Project Name","Sub Project",
        "Level","Activity Code","Activity Name","Godown Name",
        "P.O. No","P.O. Date","PO_Remarks",
        "G.R. No","GRN Date","GRN_Remarks",
        "Item Group","Item Desc",
        "GIN No","Issue Date","GIN_Remarks",
        "Received Qty","Received Amt","Issued Qty","Issued Amt",
        "GRN_Issued_Qty","GRN_Issued_Amt",
        "From Voucher","Contractor / Service Provider Name"
    ]
    for col in headers:
        if col not in stock.columns:
            stock[col] = ""

    summary = stock[headers].copy()

    out_buf = io.BytesIO()
    wb = xlsxwriter.Workbook(out_buf, {'in_memory': True, 'nan_inf_to_errors': True})
    ws_sum = wb.add_worksheet("Summary")
    ws_raw = wb.add_worksheet("RAW_Data")

    header_fmt = wb.add_format({
        'bold': True, 'font_size': 14, 'bg_color': '#FCE4D6',
        'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    ws_sum.set_default_row(20)
    ws_sum.set_row(0, 75)
    ws_sum.freeze_panes(1, 0)
    ws_sum.autofilter(0, 0, 0, len(headers)-1)

    # Write headers
    for j, h in enumerate(headers):
        ws_sum.write(0, j, h, header_fmt)
        ws_raw.write(0, j, h)

    # Write rows
    for i, row in summary.iterrows():
        for j, h in enumerate(headers):
            ws_sum.write(i+1, j, to_str(row[h]))

    for i, row in stock.iterrows():
        for j, h in enumerate(headers):
            ws_raw.write(i+1, j, to_str(row[h] if h in stock.columns else ""))

    wb.close()
    out_buf.seek(0)

    st.success("✅ Done. Download your Excel below.")
    st.download_button(
        "⬇️ Download (Summary + RAW_Data)",
        data=out_buf.getvalue(),
        file_name="Stock_Ledger_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
