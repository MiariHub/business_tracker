# app.py
import os
from io import BytesIO
from typing import List, Tuple, Dict
from datetime import datetime
import urllib.parse
import urllib.request

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import altair as alt

from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
)
from reportlab.lib.utils import ImageReader




# ---------------------------- Constants ----------------------------
SAVED_EDITS_FILE = "saved_edits.csv"
DATA_FILE = "data.csv"
EXCEL_EXPORT_FILE = "edited_report.xlsx"
REQUIRED_COLUMNS = [
    "Month", "New Leads", "Repeat Leads", "Revenue", "Net Profit", "Customers", "Conversion Rate",
]

# PDF layout caps
FRAME_MAX_W = 468   # 612 - 2*72
CHART_MAX_H = 420
LOGO_MAX_H = 120    # compact logo

# ---------------------------- Utilities ----------------------------
def normalize_month(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce").dt.to_period("M").astype(str)

def month_to_timestamp(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce").dt.to_period("M").dt.to_timestamp()

def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            if col == "Month":
                if "Date" in df.columns:
                    df[col] = normalize_month(df["Date"])
                else:
                    n = len(df) or 12
                    df[col] = pd.period_range("2025-01", periods=n, freq="M").astype(str)
            elif col == "Conversion Rate":
                df[col] = 0.0
            else:
                df[col] = 0
    df["Month"] = normalize_month(df["Month"])
    return df

def load_csv_if_exists(path: str) -> pd.DataFrame:
    if os.path.exists(path) and os.path.getsize(path) > 0:
        try:
            df = pd.read_csv(path)
            if "Month" in df.columns:
                df["Month"] = normalize_month(df["Month"])
            return ensure_columns(df)
        except Exception:
            return pd.DataFrame(columns=REQUIRED_COLUMNS)
    return pd.DataFrame(columns=REQUIRED_COLUMNS)

def load_saved_edits(path: str) -> pd.DataFrame:
    return load_csv_if_exists(path)

def save_edits(df: pd.DataFrame, path: str) -> None:
    df.copy()[REQUIRED_COLUMNS].to_csv(path, index=False)

def _try_excel_writer(df: pd.DataFrame, engine: str) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return buf.getvalue()

def bytes_excel_or_csv(df: pd.DataFrame, preferred_xlsx_name: str = "data.xlsx"):
    try:
        return _try_excel_writer(df, "openpyxl"), preferred_xlsx_name, \
               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "excel"
    except ImportError:
        pass
    try:
        return _try_excel_writer(df, "xlsxwriter"), preferred_xlsx_name, \
               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "excel"
    except ImportError:
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        csv_name = preferred_xlsx_name.replace(".xlsx", ".csv")
        return csv_bytes, csv_name, "text/csv", "csv"

def get_customer_type_toggle() -> str:
    return st.radio("Customer Type", ["All", "New", "Repeat"], index=0, horizontal=True)

# ---------- Query param compatibility ----------
def set_query_params_compat(**kwargs) -> None:
    if hasattr(st, "query_params"):
        try:
            st.query_params.clear()
            for k, v in kwargs.items():
                st.query_params[k] = v
            return
        except Exception:
            pass
    st.experimental_set_query_params(**kwargs)

# ---------------------------- KPI + helpers ----------------------------
@st.cache_data(show_spinner=False)
def calculate_ytd_metrics(df: pd.DataFrame) -> dict:
    d = df.copy()
    for col in ["Revenue", "Net Profit", "Customers"]:
        d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0)
    return {
        "YTD Revenue": float(d["Revenue"].sum()),
        "YTD Net Profit": float(d["Net Profit"].sum()),
        "YTD Customers": float(d["Customers"].sum()),
    }

def kpi_vs_last(df: pd.DataFrame) -> dict:
    d = df.copy()
    d["_ts_key"] = normalize_month(d["Month"])
    d.sort_values("_ts_key", inplace=True)
    def delta(series):
        s = pd.to_numeric(series, errors="coerce").fillna(0)
        return (s.iloc[-1] - s.iloc[-2]) if len(s) >= 2 else 0
    return {"rev": delta(d["Revenue"]), "np": delta(d["Net Profit"]), "cust": delta(d["Customers"])}

def fmt_delta(x: float) -> str:
    return f"{float(x):+,.0f}"

# ---------------------------- CSV Uploader & Validation ----------------------------
def schema_template_df() -> pd.DataFrame:
    return pd.DataFrame([{
        "Month": "2025-01",
        "New Leads": 10,
        "Repeat Leads": 5,
        "Revenue": 20000,
        "Net Profit": 6000,
        "Customers": 18,
        "Conversion Rate": 0.12,
    }])

def validate_and_coerce_upload(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str], List[str]]:
    notes: Dict[str, str] = {}
    warns: List[str] = []
    incoming_cols = set(df.columns)
    missing = [c for c in REQUIRED_COLUMNS if c not in incoming_cols]
    if missing:
        warns.append(f"Missing columns were added with defaults: {', '.join(missing)}.")
    df = ensure_columns(df)

    num_cols = ["New Leads","Repeat Leads","Revenue","Net Profit","Customers","Conversion Rate"]
    for col in num_cols:
        before_bad = df[col].isna().sum()
        df[col] = (
            df[col].astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce")
        after_bad = df[col].isna().sum()
        if after_bad > before_bad:
            notes[col] = "Some values were non-numeric and set to NaN → filled to 0."
        df[col] = df[col].fillna(0)

    df["Month"] = normalize_month(df["Month"])
    dupes = df["Month"].duplicated(keep="first").sum()
    if dupes:
        warns.append(f"Found {dupes} duplicate Month rows; keeping first occurrence.")
        df = df.drop_duplicates(subset=["Month"], keep="first")
    return df[REQUIRED_COLUMNS], notes, warns

# ---------------------------- Charts (Altair for app) ----------------------------
def alt_month_axis():
    return alt.X("Month:T", axis=alt.Axis(format="%b %Y", labelAngle=-45))

def chart_revenue_alt(df):
    d = df.copy(); d["Month"] = month_to_timestamp(d["Month"])
    return alt.Chart(d).mark_bar().encode(
        alt_month_axis(), y=alt.Y("Revenue:Q"),
        tooltip=["Month:T","Revenue:Q","Net Profit:Q","Customers:Q"]
    ).interactive()

def chart_netprofit_alt(df):
    d = df.copy(); d["Month"] = month_to_timestamp(d["Month"])
    return alt.Chart(d).mark_line(point=True).encode(
        alt_month_axis(), y=alt.Y("Net Profit:Q"),
        tooltip=["Month:T","Net Profit:Q"]
    ).interactive()

def chart_conversion_alt(df):
    d = df.copy(); d["Month"] = month_to_timestamp(d["Month"])
    d["Conversion %"] = pd.to_numeric(d["Conversion Rate"], errors="coerce").fillna(0)*100
    return alt.Chart(d).mark_line(point=True).encode(
        alt_month_axis(), y=alt.Y("Conversion %:Q"),
        tooltip=["Month:T","Conversion %:Q"]
    ).interactive()

# ---------------------------- Charts (Matplotlib for PDF) ----------------------------
def _safe_sort(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    out = df.copy()
    out["_ts"] = month_to_timestamp(out["Month"])
    out.sort_values("_ts", inplace=True)
    return out

def plot_revenue_comparison(df: pd.DataFrame):
    fig, ax = plt.subplots()
    if df.empty:
        ax.set_title("Revenue by Month (no data)"); return fig
    d = _safe_sort(df)
    ax.bar(d["_ts"], pd.to_numeric(d["Revenue"], errors="coerce").fillna(0))
    ax.set_title("Revenue by Month"); ax.set_xlabel("Month"); ax.set_ylabel("Revenue")
    fig.autofmt_xdate(); return fig

def plot_net_profit_trend(df: pd.DataFrame):
    fig, ax = plt.subplots()
    if df.empty:
        ax.set_title("Net Profit Trend (no data)"); return fig
    d = _safe_sort(df)
    ax.plot(d["_ts"], pd.to_numeric(d["Net Profit"], errors="coerce").fillna(0), marker="o")
    ax.set_title("Net Profit Trend"); ax.set_xlabel("Month"); ax.set_ylabel("Net Profit")
    fig.autofmt_xdate(); return fig

def plot_conversion_rates(df: pd.DataFrame):
    fig, ax = plt.subplots()
    if df.empty:
        ax.set_title("Conversion Rate Trend (no data)"); return fig
    d = _safe_sort(df)
    y = pd.to_numeric(d["Conversion Rate"], errors="coerce").fillna(0)
    ax.plot(d["_ts"], y, marker="o")
    ax.set_title("Conversion Rate Trend"); ax.set_xlabel("Month"); ax.set_ylabel("Conversion Rate")
    ax.set_ylim(0, max(0.21, float(y.max()) * 1.1))
    fig.autofmt_xdate(); return fig

# ---------------------------- Quarterly aggregation + charts ----------------------------
def quarterly_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Quarter","Revenue","Net Profit","Customers","Conversion Rate","_qt"])
    tmp = df.copy()
    tmp["Quarter"] = pd.to_datetime(tmp["Month"], errors="coerce").dt.to_period("Q").astype(str)
    rev = pd.to_numeric(tmp["Revenue"], errors="coerce").fillna(0)
    npf = pd.to_numeric(tmp["Net Profit"], errors="coerce").fillna(0)
    cust = pd.to_numeric(tmp["Customers"], errors="coerce").fillna(0)
    conv = pd.to_numeric(tmp["Conversion Rate"], errors="coerce").fillna(0)
    tmp["__rev"], tmp["__np"], tmp["__cust"], tmp["__conv_w"] = rev, npf, cust, conv*cust
    agg = tmp.groupby("Quarter", as_index=False).agg(
        Revenue=("__rev","sum"), Net_Profit=("__np","sum"), Customers=("__cust","sum"), Conv_W=("__conv_w","sum")
    )
    agg["Conversion Rate"] = (agg["Conv_W"] / agg["Customers"]).fillna(0)
    agg.rename(columns={"Net_Profit":"Net Profit"}, inplace=True)
    qidx = pd.PeriodIndex(agg["Quarter"], freq="Q")
    agg["_qt"] = qidx.to_timestamp()
    agg.sort_values("_qt", inplace=True)
    return agg[["Quarter","Revenue","Net Profit","Customers","Conversion Rate","_qt"]]

def plot_quarterly_revenue(df: pd.DataFrame):
    q = quarterly_summary_df(df); fig, ax = plt.subplots()
    if q.empty: ax.set_title("Revenue by Quarter (no data)"); return fig
    ax.bar(q["_qt"], q["Revenue"]); ax.set_title("Revenue by Quarter"); ax.set_xlabel("Quarter"); ax.set_ylabel("Revenue")
    fig.autofmt_xdate(); return fig

def plot_quarterly_net_profit(df: pd.DataFrame):
    q = quarterly_summary_df(df); fig, ax = plt.subplots()
    if q.empty: ax.set_title("Net Profit by Quarter (no data)"); return fig
    ax.plot(q["_qt"], q["Net Profit"], marker="o"); ax.set_title("Net Profit by Quarter"); ax.set_xlabel("Quarter"); ax.set_ylabel("Net Profit")
    fig.autofmt_xdate(); return fig

def plot_quarterly_conversion(df: pd.DataFrame):
    q = quarterly_summary_df(df); fig, ax = plt.subplots()
    if q.empty: ax.set_title("Conversion Rate by Quarter (no data)"); return fig
    ax.plot(q["_qt"], q["Conversion Rate"], marker="o"); ax.set_title("Conversion Rate by Quarter"); ax.set_xlabel("Quarter"); ax.set_ylabel("Conversion Rate")
    ax.set_ylim(0, max(0.21, float(q["Conversion Rate"].max())*1.1))
    fig.autofmt_xdate(); return fig

# ---------------------------- PDF helpers ----------------------------
def fig_to_png_bytes(fig) -> bytes:
    buf = BytesIO(); fig.savefig(buf, format="png", dpi=150, bbox_inches="tight"); plt.close(fig); return buf.getvalue()

def scaled_rl_image(data_or_bytesio, max_w=FRAME_MAX_W, max_h=CHART_MAX_H, hAlign="CENTER") -> RLImage:
    if isinstance(data_or_bytesio, (bytes, bytearray)):
        bio = BytesIO(data_or_bytesio)
    else:
        bio = data_or_bytesio
    rdr = ImageReader(bio)
    iw, ih = rdr.getSize()
    scale = min(max_w / float(iw), max_h / float(ih), 1.0)
    bio.seek(0)
    img = RLImage(bio, width=iw * scale, height=ih * scale)
    img.hAlign = hAlign
    return img

# -------- MONTHLY TABLE (centered + aggregates) ----------
def classy_monthly_table_centered(df: pd.DataFrame, max_rows: int = 10):
    d = _safe_sort(df)
    if d.empty:
        t = Table([["No data available"]]); t.hAlign = "CENTER"; return t, TableStyle([])

    # Coerce numerics
    d["Revenue"] = pd.to_numeric(d["Revenue"], errors="coerce").fillna(0.0)
    d["Net Profit"] = pd.to_numeric(d["Net Profit"], errors="coerce").fillna(0.0)
    d["Customers"] = pd.to_numeric(d["Customers"], errors="coerce").fillna(0.0)
    d["Conversion Rate"] = pd.to_numeric(d["Conversion Rate"], errors="coerce").fillna(0.0)

    # Limit to visible rows
    body = d.head(max_rows).copy()

    # Aggregates
    rev_sum = float(body["Revenue"].sum())
    np_sum  = float(body["Net Profit"].sum())
    cust_sum = float(body["Customers"].sum())
    if cust_sum > 0:
        conv_wavg = float((body["Conversion Rate"] * body["Customers"]).sum() / cust_sum)
    else:
        conv_wavg = float(body["Conversion Rate"].mean() if len(body) else 0.0)

    # Build rows
    rows = [["Month","Revenue","Net Profit","Customers","Conv. Rate"]]
    for _, r in body.iterrows():
        rows.append([
            str(r["Month"]),
            f"${r['Revenue']:,.0f}",
            f"${r['Net Profit']:,.0f}",
            f"{r['Customers']:,.0f}",
            f"{r['Conversion Rate']:.1%}",
        ])
    # Totals row
    rows.append([
        "Aggregates",
        f"${rev_sum:,.0f}",
        f"${np_sum:,.0f}",
        f"{cust_sum:,.0f}",
        f"{conv_wavg:.1%}",
    ])

    col_widths = [
        FRAME_MAX_W*0.24,
        FRAME_MAX_W*0.18,
        FRAME_MAX_W*0.18,
        FRAME_MAX_W*0.20,
        FRAME_MAX_W*0.20,
    ]
    tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    tbl.hAlign = "CENTER"

    navy = colors.HexColor("#173a5e")
    zebra1 = colors.whitesmoke
    zebra2 = colors.HexColor("#f7f9fb")
    rule = colors.HexColor("#d8e0ea")

    style_cmds = [
        ("FONTNAME",(0,0),(-1,0),"Times-Bold"),
        ("FONTSIZE",(0,0),(-1,0),10),
        ("BACKGROUND",(0,0),(-1,0),navy),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),

        # Center EVERYTHING
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),

        ("FONTNAME",(0,1),(-1,-2),"Times-Roman"),
        ("FONTSIZE",(0,1),(-1,-2),9),

        # Body zebra (excluding the aggregate row)
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[zebra1, zebra2]),

        # Rules
        ("LINEBELOW",(0,0),(-1,0),1, navy),       # header underline
        ("LINEABOVE",(0,-1),(-1,-1),1, navy),     # line above aggregates

        # Aggregates styling (last row)
        ("FONTNAME",(0,-1),(-1,-1),"Times-Bold"),
        ("FONTSIZE",(0,-1),(-1,-1),10),

        # Padding
        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("RIGHTPADDING",(0,0),(-1,-1),6),
        ("TOPPADDING",(0,0),(-1,-1),6),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]
    return tbl, TableStyle(style_cmds)

# -------- QUARTERLY TABLE (centered + aggregates) ----------
def quarterly_table_centered(df: pd.DataFrame, max_rows: int = 10):
    if df.empty:
        t = Table([["No data available"]]); t.hAlign = "CENTER"; return t, TableStyle([])
    agg = quarterly_summary_df(df).copy()

    # Limit to visible rows
    view = agg.head(max_rows).copy()

    # Aggregates
    rev_sum = float(view["Revenue"].sum())
    np_sum  = float(view["Net Profit"].sum())
    cust_sum = float(view["Customers"].sum())
    if cust_sum > 0:
        conv_wavg = float((view["Conversion Rate"] * view["Customers"]).sum() / cust_sum)
    else:
        conv_wavg = float(view["Conversion Rate"].mean() if len(view) else 0.0)

    rows = [["Quarter","Revenue","Net Profit","Customers","Conv. Rate"]]
    for _, r in view.iterrows():
        rows.append([
            r["Quarter"],
            f"${r['Revenue']:,.0f}",
            f"${r['Net Profit']:,.0f}",
            f"{r['Customers']:,.0f}",
            f"{r['Conversion Rate']:.1%}",
        ])
    rows.append([
        "Aggregates",
        f"${rev_sum:,.0f}",
        f"${np_sum:,.0f}",
        f"{cust_sum:,.0f}",
        f"{conv_wavg:.1%}",
    ])

    tbl = Table(rows, repeatRows=1)
    tbl.hAlign = "CENTER"
    navy = colors.HexColor("#173a5e")
    zebra1 = colors.whitesmoke
    zebra2 = colors.HexColor("#f7f9fb")

    style = TableStyle([
        ("FONTNAME",(0,0),(-1,0),"Times-Bold"),
        ("FONTSIZE",(0,0),(-1,0),10),
        ("BACKGROUND",(0,0),(-1,0),navy),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),

        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),

        ("FONTNAME",(0,1),(-1,-2),"Times-Roman"),
        ("FONTSIZE",(0,1),(-1,-2),9),

        ("ROWBACKGROUNDS",(0,1),(-1,-2),[zebra1, zebra2]),

        ("LINEBELOW",(0,0),(-1,0),1, navy),
        ("LINEABOVE",(0,-1),(-1,-1),1, navy),

        ("FONTNAME",(0,-1),(-1,-1),"Times-Bold"),
        ("FONTSIZE",(0,-1),(-1,-1),10),

        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("RIGHTPADDING",(0,0),(-1,-1),6),
        ("TOPPADDING",(0,0),(-1,-1),6),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ])
    return tbl, style

def make_page_decorator(timestamp_text: str):
    def _decor(c, doc):
        c.saveState()
        c.setFont("Times-Italic", 9)
        c.drawString(40, 30, f"Generated on: {timestamp_text}")
        c.drawRightString(LETTER[0]-36, 30, f"Page {c.getPageNumber()}")
        c.restoreState()
    return _decor

# ---------- Logo helpers ----------
def github_blob_to_raw(url: str) -> str:
    try:
        parsed = urllib.parse.urlparse(url)
        if parsed.netloc.lower() == "github.com" and "/blob/" in parsed.path:
            parts = parsed.path.strip("/").split("/")
            if len(parts) >= 5 and parts[2] == "blob":
                user, repo, _, branch = parts[:4]
                rest = "/".join(parts[4:])
                return f"https://raw.githubusercontent.com/{user}/{repo}/{branch}/{rest}"
    except Exception:
        pass
    return url

def fetch_url_bytes(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=15) as resp:
        return resp.read()

def load_logo_bytes(logo_file, logo_url: str | None) -> BytesIO | None:
    if logo_file is not None:
        return BytesIO(logo_file.getvalue())
    if logo_url:
        url = github_blob_to_raw(logo_url.strip())
        try:
            data = fetch_url_bytes(url)
            return BytesIO(data)
        except Exception as e:
            st.warning(f"Could not fetch logo from URL: {e}")
    return None

# ---------------------------- Sample Data ----------------------------
def make_sample_data() -> pd.DataFrame:
    months = pd.period_range("2025-01", periods=12, freq="M").astype(str)
    return pd.DataFrame({
        "Month": months,
        "New Leads": [20,18,22,25,27,24,26,30,28,29,31,35],
        "Repeat Leads": [10,12,11,13,12,14,15,16,18,17,19,20],
        "Revenue": [20000,22000,21000,25000,27000,26000,30000,32000,31000,33000,35000,38000],
        "Net Profit": [6000,6500,6200,7000,7400,7300,8200,8700,8500,9000,9500,10500],
        "Customers": [25,27,28,30,32,31,35,38,37,39,42,45],
        "Conversion Rate": [0.12,0.13,0.12,0.14,0.15,0.14,0.16,0.17,0.16,0.17,0.18,0.19],
    })

# ---------------------------- Filtering ----------------------------
def apply_filters(df: pd.DataFrame, months_selected: List[str], customer_type: str) -> pd.DataFrame:
    df_f = df[df["Month"].isin(months_selected)].copy()
    if customer_type == "New":
        df_f = df_f[df_f["New Leads"] > 0]
    elif customer_type == "Repeat":
        df_f = df_f[df_f["Repeat Leads"] > 0]
    return df_f

# ---------------------------- Main Renderer ----------------------------
def render_dashboard(df: pd.DataFrame, role: str = "owner"):
    df = ensure_columns(df)

    uploaded_df = load_csv_if_exists(DATA_FILE)
    if not uploaded_df.empty:
        df = uploaded_df

    saved = load_saved_edits(SAVED_EDITS_FILE)
    if not saved.empty:
        m = df.merge(saved[REQUIRED_COLUMNS], on="Month", how="left", suffixes=("", "_saved"))
        for col in REQUIRED_COLUMNS:
            if col != "Month" and f"{col}_saved" in m.columns:
                m[col] = m[f"{col}_saved"].combine_first(m[col])
                m.drop(columns=[f"{col}_saved"], inplace=True)
        df = m

    st.title("📊 Business Performance Tracker")

    with st.sidebar:
        st.header("Data")
        up = st.file_uploader("Upload CSV (with Month, Revenue, etc.)", type=["csv"])
        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬇️ Download CSV Template"):
                tpl = schema_template_df()
                st.download_button("Download now", data=tpl.to_csv(index=False).encode("utf-8"),
                                   file_name="template.csv", mime="text/csv", type="primary")
        with c2:
            if st.button("↩️ Reset to sample"):
                if os.path.exists(DATA_FILE): os.remove(DATA_FILE)
                if os.path.exists(SAVED_EDITS_FILE): os.remove(SAVED_EDITS_FILE)
                st.toast("Reset complete. Reloading…", icon="♻️")
                st.experimental_rerun()

        if up is not None:
            try:
                raw = pd.read_csv(up)
                clean, notes, warns = validate_and_coerce_upload(raw)
                clean.to_csv(DATA_FILE, index=False)
                st.success(f"Uploaded {len(clean)} rows → saved to {DATA_FILE}")
                for w in warns: st.warning(w)
                if notes: st.caption("Coercions: " + "; ".join([f"{k}: {v}" for k,v in notes.items()]))
                st.toast("Data uploaded", icon="✅")
                st.experimental_rerun()
            except Exception as e:
                st.error(f"Upload failed: {e}")

        st.header("Filters")
        preset = st.selectbox("Quick range", ["All","YTD","QTD","MTD"], index=0)
        df["_MonthTS"] = month_to_timestamp(df["Month"])
        today = pd.to_datetime("today").normalize()
        if preset == "YTD":
            start = pd.Timestamp(today.year, 1, 1)
            default_months = sorted(df.loc[df["_MonthTS"]>=start, "Month"].unique())
        elif preset == "QTD":
            start = pd.Timestamp(today.year, ((today.month-1)//3)*3 + 1, 1)
            default_months = sorted(df.loc[df["_MonthTS"]>=start, "Month"].unique())
        elif preset == "MTD":
            start = pd.Timestamp(today.year, today.month, 1)
            default_months = sorted(df.loc[df["_MonthTS"]>=start, "Month"].unique())
        else:
            default_months = sorted(df["Month"].unique().tolist())

        months_selected = st.multiselect("Months", sorted(df["Month"].unique().tolist()), default=default_months)
        customer_type = get_customer_type_toggle()
        kpi_scope = st.toggle("Apply filters to KPIs", value=False)

        if st.button("🔗 Copy link with current filters"):
            set_query_params_compat(months=months_selected, customer=customer_type, kpi=int(kpi_scope), preset=preset)
            st.toast("Link updated with current filters", icon="🔗")

        with st.popover("ℹ️ Data entry tips"):
            st.markdown("- Month must be unique (YYYY-MM)\n- Conversion is 0–1\n- Numbers accept commas")

    df_filtered = apply_filters(df, months_selected, customer_type)
    if df_filtered.empty:
        st.info("No rows match your filters. Try widening the date range or switching Customer Type.")

    kpi_df = df_filtered if kpi_scope else df
    ytd = calculate_ytd_metrics(kpi_df)
    deltas = kpi_vs_last(kpi_df)

    c1, c2, c3 = st.columns(3)
    c1.metric("YTD Revenue", f"${ytd['YTD Revenue']:,.0f}", fmt_delta(deltas["rev"]))
    c2.metric("YTD Net Profit", f"${ytd['YTD Net Profit']:,.0f}", fmt_delta(deltas["np"]))
    c3.metric("YTD Customers", f"{ytd['YTD Customers']:,.0f}", fmt_delta(deltas["cust"]))

    with st.expander("🎯 Goals & Progress", expanded=True):
        colg1, colg2 = st.columns(2)
        with colg1:
            rev_goal = st.number_input("Revenue Goal (YTD)", min_value=0, value=int((ytd["YTD Revenue"] or 1)*1.1))
            st.progress(min(1.0, ytd["YTD Revenue"]/max(1, rev_goal)), text=f"Revenue progress ({ytd['YTD Revenue']:,.0f}/{rev_goal:,.0f})")
        with colg2:
            np_goal  = st.number_input("Net Profit Goal (YTD)", min_value=0, value=int((ytd["YTD Net Profit"] or 1)*1.1))
            st.progress(min(1.0, ytd["YTD Net Profit"]/max(1, np_goal)), text=f"Net Profit progress ({ytd['YTD Net Profit']:,.0f}/{np_goal:,.0f})")

    tab1, tab2, tab3, tab4 = st.tabs(["📋 Monthly Data", "📈 Charts", "💰 Cashflow Planner", "🧾 Export"])

    with tab1:
        st.subheader("📝 Edit Monthly Business Data")
        df_flag = df_filtered.copy()
        if not df_flag.empty:
            x = pd.to_numeric(df_flag["Revenue"], errors="coerce").fillna(0)
            if x.std(ddof=0) != 0:
                df_flag["Rev Outlier"] = ((x - x.mean()).abs() > 2.0 * x.std(ddof=0))
                df_flag["Revenue ⚠️"] = df_flag.apply(
                    lambda r: f"${float(r['Revenue']):,.0f}" + (" ⚠️" if r.get("Rev Outlier", False) else ""), axis=1
                )
        st.dataframe(
            (df_flag[["Month","Revenue ⚠️","Net Profit","Customers","Conversion Rate"]] if not df_flag.empty else df_filtered),
            use_container_width=True
        )

        edit_cols = ["Month","New Leads","Repeat Leads","Revenue","Net Profit","Customers","Conversion Rate"]
        editable_df = df_filtered[edit_cols].copy()
        edited_df = st.data_editor(
            editable_df,
            use_container_width=True,
            num_rows="dynamic",
            key="editor",
            column_config={
                "Month": st.column_config.TextColumn("Month (YYYY-MM)", help="Must be unique per row."),
                "New Leads": st.column_config.NumberColumn(min_value=0, step=1),
                "Repeat Leads": st.column_config.NumberColumn(min_value=0, step=1),
                "Revenue": st.column_config.NumberColumn(min_value=0, format="%.0f"),
                "Net Profit": st.column_config.NumberColumn(min_value=0, format="%.0f"),
                "Customers": st.column_config.NumberColumn(min_value=0, step=1),
                "Conversion Rate": st.column_config.NumberColumn(min_value=0.0, max_value=1.0, format="%.3f"),
            },
        )
        st.session_state["edited_df"] = edited_df

        col_add, col_save, col_dl = st.columns([1,1,1])
        with col_add:
            if st.button("➕ Add Month"):
                with st.form("add_month"):
                    m = st.text_input("Month (YYYY-MM)")
                    rev = st.number_input("Revenue", min_value=0)
                    npf = st.number_input("Net Profit", min_value=0)
                    cust = st.number_input("Customers", min_value=0, step=1)
                    conv = st.number_input("Conversion Rate (0–1)", min_value=0.0, max_value=1.0, step=0.001)
                    newl = st.number_input("New Leads", min_value=0, step=1)
                    repl = st.number_input("Repeat Leads", min_value=0, step=1)
                    submitted = st.form_submit_button("Add")
                if submitted:
                    try:
                        m_norm = pd.to_datetime(m).to_period("M").strftime("%Y-%m")
                        if m_norm in st.session_state["edited_df"]["Month"].values:
                            st.warning("Month already exists; updating existing row.")
                            mask = st.session_state["edited_df"]["Month"] == m_norm
                            st.session_state["edited_df"].loc[mask, ["Revenue","Net Profit","Customers","Conversion Rate","New Leads","Repeat Leads"]] = [rev,npf,cust,conv,newl,repl]
                        else:
                            st.session_state["edited_df"] = pd.concat([
                                st.session_state["edited_df"],
                                pd.DataFrame([{
                                    "Month": m_norm, "Revenue": rev, "Net Profit": npf, "Customers": cust,
                                    "Conversion Rate": conv, "New Leads": newl, "Repeat Leads": repl
                                }])
                            ], ignore_index=True)
                        st.toast("Month added/updated", icon="✅")
                    except Exception:
                        st.error("Invalid month format (use YYYY-MM).")

        with col_dl:
            data_bytes, fname, mime, used = bytes_excel_or_csv(edited_df, preferred_xlsx_name=EXCEL_EXPORT_FILE)
            if used == "csv":
                st.info("Excel engines missing; exported CSV instead. Install 'openpyxl' or 'xlsxwriter' for .xlsx.")
            st.download_button("⬇️ Download current view", data=data_bytes, file_name=fname, mime=mime)

        with col_save:
            if st.button("💾 Save Changes"):
                with st.status("Saving changes…", expanded=False) as s:
                    if edited_df["Month"].isna().any() or (edited_df["Month"].astype(str).str.len() == 0).any():
                        st.warning("Month cannot be empty."); st.stop()
                    try:
                        m_norm = pd.to_datetime(edited_df["Month"]).dt.to_period("M").astype(str)
                    except Exception:
                        st.warning("Invalid Month values. Use format like 2025-01."); st.stop()
                    if len(m_norm.unique()) != len(m_norm):
                        st.warning("Duplicate Month values detected."); st.stop()

                    num_cols = ["New Leads","Repeat Leads","Revenue","Net Profit","Customers","Conversion Rate"]
                    coerced = edited_df.copy()
                    for col in num_cols:
                        coerced[col] = (
                            coerced[col].astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False)
                        )
                        coerced[col] = pd.to_numeric(coerced[col], errors="coerce")
                    if coerced[num_cols].isna().any().any():
                        st.warning("Some numeric fields contain invalid values."); st.stop()
                    if (coerced["Conversion Rate"] < 0).any() or (coerced["Conversion Rate"] > 1).any():
                        st.warning("Conversion Rate must be between 0 and 1."); st.stop()

                    coerced["Month"] = m_norm
                    base = df.copy()
                    merged = base.merge(coerced, on="Month", how="left", suffixes=("", "_new"))
                    for col in num_cols:
                        merged[col] = merged[f"{col}_new"].combine_first(merged[col])
                        merged.drop(columns=[f"{col}_new"], inplace=True)

                    save_edits(merged, SAVED_EDITS_FILE)
                    calculate_ytd_metrics.clear()
                    s.update(label="Saved!", state="complete", expanded=False)
                    st.toast("✅ Changes saved")
                    st.rerun()

    with tab2:
        st.subheader("📈 Visual Charts (hover/zoom/save)")
        st.altair_chart(chart_revenue_alt(df_filtered), use_container_width=True)
        st.altair_chart(chart_netprofit_alt(df_filtered), use_container_width=True)
        st.altair_chart(chart_conversion_alt(df_filtered), use_container_width=True)

    with tab3:
        st.subheader("📅 Weekly Cashflow Planner + ⚙️ What-if")
        cashflow_data = {
            "Week Starting": pd.date_range("2025-01-06", periods=12, freq="W-MON"),
            "Cash In": [10000 + i * 500 for i in range(12)],
            "Cash Out": [8000 + i * 400 for i in range(12)],
        }
        cashflow_df = pd.DataFrame(cashflow_data)
        cashflow_df["Net Cashflow"] = cashflow_df["Cash In"] - cashflow_df["Cash Out"]
        st.dataframe(cashflow_df, use_container_width=True)

        colg, colc, colp = st.columns(3)
        with colg:
            growth = st.slider("Revenue growth from next week (%)", 0, 50, 10, 1)
        with colc:
            ap_lag = st.slider("AP lag weeks", 0, 8, 2, 1, help="Delay cash out")
        with colp:
            ar_lag = st.slider("AR lag weeks", 0, 8, 1, 1, help="Delay cash in")

        sim = cashflow_df.copy()
        sim["Cash In"] = (sim["Cash In"] * (1 + growth/100)).shift(ar_lag, fill_value=0)
        sim["Cash Out"] = sim["Cash Out"].shift(ap_lag, fill_value=0)
        sim["Net Cashflow"] = sim["Cash In"] - sim["Cash Out"]

        base_line = cashflow_df[["Week Starting","Net Cashflow"]].copy(); base_line["Label"] = "Base"
        sim_line = sim[["Week Starting","Net Cashflow"]].copy(); sim_line["Label"] = "Simulated"
        plot_df = pd.concat([base_line, sim_line], ignore_index=True)
        st.altair_chart(
            alt.Chart(plot_df).mark_line(point=True).encode(
                x="Week Starting:T", y="Net Cashflow:Q", color="Label:N",
                tooltip=["Week Starting:T","Net Cashflow:Q","Label:N"]
            ).interactive(),
            use_container_width=True
        )
        st.caption("Solid lines reflect base vs simulated net cashflow under current sliders.")

    with tab4:
        st.subheader("🧾 Export Report")
        logo_file = st.file_uploader("Optional logo (PNG/JPG)", type=["png","jpg","jpeg"])
        logo_url = st.text_input(
            "Or paste logo URL",
            value="https://raw.githubusercontent.com/MiariHub/cement_prediction/main/electro-pi.png",
            help="Direct image links recommended. GitHub blob links are auto-converted."
        )
        logo_width = st.slider("Logo width (px)", 80, 260, 140, 10)
        subtitle = st.text_input("Subtitle (e.g., Region / Team)", value="Consolidated YTD")
        report_date = st.date_input("Report date", pd.to_datetime("today"))

        st.markdown("**Include charts:**")
        col_a, col_b, col_c = st.columns(3)
        with col_a: include_rev = st.checkbox("Revenue", value=True)
        with col_b: include_np = st.checkbox("Net Profit", value=True)
        with col_c: include_conv = st.checkbox("Conversion", value=False)
        aggregate_charts_by_quarter = st.checkbox("Aggregate charts by quarter", value=False)

        st.markdown("**Summary table options:**")
        col1, col2 = st.columns(2)
        with col1:
            max_rows = st.number_input("Max rows per table", min_value=5, max_value=100, value=10, step=1)
        with col2:
            group_by_quarter = st.checkbox("Group by quarter (single table)", value=False)
        include_both_tables = st.checkbox("Include both monthly and quarterly tables in PDF", value=False)

        if group_by_quarter or include_both_tables:
            qdf = quarterly_summary_df(df_filtered).drop(columns=["_qt"]).copy()
            qcol1, qcol2 = st.columns(2)
            with qcol1:
                st.download_button("📄 Download Quarterly Summary (CSV)",
                                   data=qdf.to_csv(index=False).encode("utf-8"),
                                   file_name="quarterly_summary.csv", mime="text/csv")
            with qcol2:
                data_bytes, fname, mime, used = bytes_excel_or_csv(qdf, preferred_xlsx_name="quarterly_summary.xlsx")
                if used == "csv":
                    st.info("Excel engines missing; exported CSV instead. Install 'openpyxl' or 'xlsxwriter' for .xlsx.")
                st.download_button("📊 Download Quarterly Summary (Excel)", data=data_bytes, file_name=fname, mime=mime)
        else:
            st.caption("Turn on a quarterly table option to enable CSV/Excel quarterly exports.")

        def build_chart_blocks():
            blocks = []
            if include_rev:
                fig = plot_quarterly_revenue(df_filtered) if aggregate_charts_by_quarter else plot_revenue_comparison(df_filtered)
                blocks.append(("Revenue by Quarter" if aggregate_charts_by_quarter else "Revenue by Month", fig))
            if include_np:
                fig = plot_quarterly_net_profit(df_filtered) if aggregate_charts_by_quarter else plot_net_profit_trend(df_filtered)
                blocks.append(("Net Profit by Quarter" if aggregate_charts_by_quarter else "Net Profit Trend", fig))
            if include_conv:
                fig = plot_quarterly_conversion(df_filtered) if aggregate_charts_by_quarter else plot_conversion_rates(df_filtered)
                blocks.append(("Conversion Rate by Quarter" if aggregate_charts_by_quarter else "Conversion Rate Trend", fig))
            return blocks

        if st.button("📤 Generate PDF"):
            buffer = BytesIO()
            doc = SimpleDocTemplate(
                buffer, pagesize=LETTER,
                leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36
            )
            styles = getSampleStyleSheet()
            styles["Title"].fontName = "Times-Bold"
            styles["Italic"].fontName = "Times-Italic"
            styles["Normal"].fontName = "Times-Roman"
            styles["Heading3"].fontName = "Times-Bold"
            styles["Title"].textColor = colors.HexColor("#1f4e79")
            styles["Title"].leading = 22

            elements = []

            logo_bytes = load_logo_bytes(logo_file, logo_url)
            if logo_bytes is not None:
                try:
                    img_logo = scaled_rl_image(
                        logo_bytes,
                        max_w=min(float(logo_width), FRAME_MAX_W),
                        max_h=LOGO_MAX_H,
                        hAlign="CENTER"
                    )
                    elements.append(img_logo)
                    elements.append(Spacer(1, 6))
                except Exception:
                    st.warning("Logo could not be read. Make sure the URL points to an image file (PNG/JPG).")

            elements.append(Paragraph("Business Performance Report", styles["Title"]))
            elements.append(Paragraph(subtitle, styles["Italic"]))
            elements.append(Paragraph(pd.to_datetime(report_date).strftime("%B %d, %Y"), styles["Normal"]))
            elements.append(Spacer(1, 10))

            toc_items = []
            if include_rev: toc_items.append("• Revenue")
            if include_np:  toc_items.append("• Net Profit")
            if include_conv: toc_items.append("• Conversion")
            toc_items.append("• Tables")
            elements.append(Paragraph("Contents", styles["Heading3"]))
            elements.append(Paragraph("<br/>".join(toc_items), styles["Normal"]))
            elements.append(Spacer(1, 12))

            kpi_df_for_pdf = kpi_df if 'kpi_scope' in st.session_state and st.session_state['kpi_scope'] else df
            kpis = calculate_ytd_metrics(kpi_df_for_pdf)
            kpi_text = f"YTD Revenue: ${kpis['YTD Revenue']:,.0f} &nbsp;&nbsp; YTD Net Profit: ${kpis['YTD Net Profit']:,.0f} &nbsp;&nbsp; YTD Customers: {kpis['YTD Customers']:,.0f}"
            elements.append(Paragraph(kpi_text, styles["Normal"]))
            elements.append(Spacer(1, 12))

            for title, fig in build_chart_blocks():
                png = fig_to_png_bytes(fig)
                elements.append(Paragraph(title, styles["Heading3"]))
                elements.append(scaled_rl_image(png, max_w=FRAME_MAX_W, max_h=CHART_MAX_H, hAlign="CENTER"))
                elements.append(Spacer(1, 12))

            if include_both_tables:
                tbl_m, style_m = classy_monthly_table_centered(df_filtered, max_rows=int(max_rows))
                tbl_q, style_q = quarterly_table_centered(df_filtered, max_rows=int(max_rows))
                tbl_m.setStyle(style_m); tbl_q.setStyle(style_q)
                elements.append(Paragraph("Summary by Month", styles["Heading3"])); elements.append(tbl_m); elements.append(Spacer(1, 12))
                elements.append(Paragraph("Summary by Quarter", styles["Heading3"])); elements.append(tbl_q)
            else:
                if group_by_quarter:
                    tbl, style = quarterly_table_centered(df_filtered, max_rows=int(max_rows))
                    tbl.setStyle(style)
                    elements.append(Paragraph("Summary by Quarter", styles["Heading3"])); elements.append(tbl)
                else:
                    tbl, style = classy_monthly_table_centered(df_filtered, max_rows=int(max_rows))
                    tbl.setStyle(style)
                    elements.append(Paragraph("Summary by Month", styles["Heading3"])); elements.append(tbl)

            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            decor = make_page_decorator(ts)
            doc.build(elements, onFirstPage=decor, onLaterPages=decor)

            st.download_button(
                label="📥 Download PDF",
                data=buffer.getvalue(),
                file_name="business_report.pdf",
                mime="application/pdf"
            )

# ---------------------------- Entrypoint ----------------------------
def main():
    st.set_page_config(page_title="Business Performance Tracker", layout="wide")
    if "base_df" not in st.session_state:
        st.session_state["base_df"] = make_sample_data()
    render_dashboard(st.session_state["base_df"], role="owner")

if __name__ == "__main__":
    main()
