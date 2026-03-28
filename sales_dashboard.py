# ============================================================
#  SALES DASHBOARD — Vaayushanti Solutions Pvt. Ltd.
#  pip install dash plotly pandas openpyxl requests
#  python sales_dashboard_fixed.py → http://localhost:8050
#
#  ✅ v3 FIXED: 
#    - Month filter now CORRECTLY updates month-wise revenue chart
#    - apply_filters fully consistent across ALL callbacks
#    - Month grid auto/explicit mode fixed
#    - Category trend uses same filtered data as KPIs
#    - SC/CRM uses apply_filters for consistency
#    - Faster rendering with memoized filter
# ============================================================

import dash
import time
from dash import dcc, html, Input, Output, State, dash_table, ctx
import plotly.graph_objects as go
import pandas as pd
import requests, io, re
from datetime import datetime

SHEET_CSV_URLS = [
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRmQxkjVw5rF5PIj7tSzeh3sV5gT6fcRlM3D77J2zxXeL9qklUB041KjNkmUTiFFotHCklKTNqJxBcx/pub?output=csv"
]

SHEET_NAMES = ["Sheet1"]

CAT_COLORS = {
    "Bags":         "#1D9E75",
    "Cages":        "#4F8EF7",
    "Projects":     "#F5A623",
    "Trading Item": "#E05C97",
}
MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

SC_CRM_EMPLOYEES = ["Renuka Arya", "Jyoti Sahu"]
SC_CRM_COLORS    = {"Renuka Arya": "#A78BFA", "Jyoti Sahu": "#FB923C"}

# ─────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────
def clean_sp(v):
    v = str(v).strip()
    v = re.sub(r"^EMP\d+[_\s]+", "", v, flags=re.IGNORECASE)
    return v.replace("_", " ").strip()

def parse_date(v):
    if pd.isna(v): return pd.NaT
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "nat", "", "null", "n/a"): return pd.NaT
    try:
        n = float(re.sub(r"[,\s]", "", s))
        if 33000 < n < 73000:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(n))
    except Exception:
        pass
    for fmt in [
        "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y",
        "%d-%m-%y", "%d/%m/%y", "%m/%d/%y",
        "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",    "%d/%m/%Y %H:%M",
        "%Y/%m/%d",          "%d %b %Y",
        "%d %B %Y",          "%b %d, %Y",
        "%B %d, %Y",         "%Y-%m-%dT%H:%M:%S",
        "%Y%m%d",
    ]:
        try:
            return pd.to_datetime(s, format=fmt)
        except Exception:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT

def get_cat(v):
    il = str(v).lower()
    if "bag"     in il: return "Bags"
    if "cage"    in il: return "Cages"
    if "project" in il: return "Projects"
    if "trading" in il or "trade" in il or "ventury" in il or "venturi" in il:
        return "Trading Item"
    return str(v).strip() or "Other"

# ─────────────────────────────────────────────────────────────
#  FETCH
# ─────────────────────────────────────────────────────────────
def fetch_csv_from_url(url, sheet_name="Sheet", retries=2):
    for attempt in range(retries + 1):
        try:
            if attempt > 0:
                print(f"  🔄 Retry {attempt}/{retries} for {sheet_name}...")
                time.sleep(2 * attempt)

            print(f"  🌐 Fetching: {url[:80]}...")
            resp = requests.get(f"{url}&t={int(time.time())}", timeout=30)
            resp.raise_for_status()
            content = resp.content

            raw_str = None
            for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    raw_str = content.decode(enc)
                    break
                except Exception:
                    pass
            if raw_str is None:
                raw_str = content.decode("utf-8", errors="replace")

            peek = pd.read_csv(io.StringIO(raw_str), header=None, dtype=str, nrows=15)
            if peek.shape[0] < 2 or peek.shape[1] < 2:
                print(f"  ⚠️  {sheet_name}: Too few rows/cols, skipping")
                if attempt < retries: continue
                return None

            hdr, best = 0, 0
            for i in range(min(10, len(peek))):
                vals = [str(x).lower().strip() for x in peek.iloc[i]
                        if str(x) not in ("nan", "", "None")]
                score = (
                    any("date" in v for v in vals) +
                    any("company" in v or "client" in v for v in vals) +
                    any("value" in v or "amount" in v or "₹" in v for v in vals) +
                    any("item" in v or "category" in v for v in vals) +
                    any("sales" in v or "person" in v for v in vals) +
                    (len(vals) >= 4)
                )
                if score > best:
                    best = score
                    hdr = i

            df = pd.read_csv(io.StringIO(raw_str), header=hdr, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.dropna(how="all").reset_index(drop=True)
            df = df.loc[:, df.notna().any()]

            mp = {}
            for c in df.columns:
                cl = c.lower().strip()
                if not cl or cl == "nan":
                    continue
                if "date" in cl and "po" not in cl and "order" not in cl:
                    if "date" not in mp.values(): mp[c] = "date"
                elif any(k in cl for k in ("company","client","customer")):
                    if "company" not in mp.values(): mp[c] = "company"
                elif "item" in cl or ("category" in cl and "sub" not in cl):
                    if "item" not in mp.values(): mp[c] = "item"
                elif cl in ("qty","quantity") or (
                        "qty" in cl and "val" not in cl and "amount" not in cl):
                    if "qty" not in mp.values(): mp[c] = "qty"
                elif any(k in cl for k in ("value","amount","₹")) or cl == "po":
                    if "amount" not in mp.values(): mp[c] = "amount"
                elif ("office" in cl and any(k in cl for k in ("person","sp","sc"))) or \
                     cl.replace(" ", "") in ("sc","crm") or \
                     ("crm" in cl and "person" in cl):
                    if "office_sp" not in mp.values(): mp[c] = "office_sp"
                elif "sales" in cl and "person" in cl:
                    if "salesperson" not in mp.values(): mp[c] = "salesperson"
                elif ("end" in cl and ("user" in cl or "oem" in cl)) or cl == "oem":
                    if "end_user_oem" not in mp.values(): mp[c] = "end_user_oem"
                elif "existing" in cl or ("new" in cl and "customer" in cl):
                    if "existing_new" not in mp.values(): mp[c] = "existing_new"

            if "amount" not in mp.values():
                print(f"  ⚠️  {sheet_name}: 'amount' column nahi mili")
                print(f"       Columns: {list(df.columns)}")
                if attempt < retries: continue
                return None

            df = df.rename(columns=mp)

            base_cols  = ["date","company","item","qty","amount","salesperson","office_sp"]
            extra_cols = ["end_user_oem","existing_new"]
            for c in base_cols + extra_cols:
                if c not in df.columns:
                    df[c] = "" if c not in ("qty","amount") else 0

            df = df[base_cols + extra_cols].copy()

            df["date"] = df["date"].apply(parse_date)
            df = df.dropna(subset=["date"])
            if df.empty:
                print(f"  ⚠️  {sheet_name}: Date parse ke baad koi row nahi bachi")
                if attempt < retries: continue
                return None

            df["amount"] = pd.to_numeric(
                df["amount"].astype(str)
                    .str.replace(",",  "", regex=False)
                    .str.replace("₹",  "", regex=False)
                    .str.replace(" ",  "", regex=False)
                    .str.strip(),
                errors="coerce").fillna(0)

            df["qty"] = pd.to_numeric(
                df["qty"].astype(str).str.extract(r"(\d+)")[0],
                errors="coerce").fillna(0).astype(int)

            df = df[df["amount"] > 0]
            if df.empty:
                print(f"  ⚠️  {sheet_name}: Amount > 0 filter ke baad koi row nahi bachi")
                if attempt < retries: continue
                return None

            REPLACE_MAP = {"nan":"","NaN":"","None":"","<NA>":"","null":"","NULL":""}
            for col in ["company","item","salesperson","end_user_oem","existing_new","office_sp"]:
                df[col] = df[col].astype(str).str.strip().replace(REPLACE_MAP)

            df["office_sp_clean"] = df["office_sp"].apply(clean_sp)
            df["sp_clean"]        = df["salesperson"].apply(clean_sp)
            df["category"]        = df["item"].apply(get_cat)
            df["year"]            = df["date"].dt.year.astype(int)
            df["month"]           = df["date"].dt.month.astype(int)
            df["month_name"]      = df["date"].dt.strftime("%b")
            df["day_str"]         = df["date"].dt.strftime("%d-%m-%Y")

            print(f"  ✅ {sheet_name!r}: {len(df)} orders loaded (₹{df['amount'].sum():,.0f})")
            return df

        except requests.exceptions.ConnectionError:
            print(f"  ❌ {sheet_name}: Network error — internet check karo")
        except requests.exceptions.Timeout:
            print(f"  ❌ {sheet_name}: Timeout — slow internet ya badi sheet")
        except requests.exceptions.HTTPError as e:
            print(f"  ❌ {sheet_name}: HTTP {e.response.status_code}")
            if e.response.status_code in (401, 403, 404):
                return None
        except Exception as e:
            import traceback; traceback.print_exc()
            print(f"  ❌ {sheet_name}: {e}")

    return None


def load_data():
    print("\n" + "─"*55)
    print("  🌐 Google Sheets se data fetch ho raha hai...")
    print("─"*55)

    all_dfs = []
    for i, url in enumerate(SHEET_CSV_URLS):
        name = SHEET_NAMES[i] if i < len(SHEET_NAMES) else f"Sheet{i+1}"
        df   = fetch_csv_from_url(url, name)
        if df is not None:
            all_dfs.append(df)

    if not all_dfs:
        print("  ❌ Koi bhi sheet se data nahi aaya!\n")
        return pd.DataFrame()

    out = pd.concat(all_dfs, ignore_index=True)
    out = out.drop_duplicates(subset=["date","company","item","amount"], keep="first")
    out = out.sort_values("date").reset_index(drop=True)
    print(f"\n  🎉 Total: {len(out)} orders | ₹{out['amount'].sum():,.0f}")
    print("─"*55 + "\n")
    return out


# ─────────────────────────────────────────────────────────────
#  INITIAL DATA LOAD
# ─────────────────────────────────────────────────────────────
SOURCE_LABEL = "Google Sheets (Live)"
SOURCE_URL   = SHEET_CSV_URLS[0] if SHEET_CSV_URLS else ""

DF        = load_data()
ALL_YEARS = sorted(DF["year"].unique().tolist())     if not DF.empty else []
ALL_CATS  = sorted(DF["category"].unique().tolist()) if not DF.empty else []
ALL_SPS   = sorted(DF["sp_clean"].unique().tolist()) if not DF.empty else []
ALL_COS   = sorted(DF["company"].unique().tolist())  if not DF.empty else []
LAST_SYNC = datetime.now().strftime("%d %b %H:%M")

# ─────────────────────────────────────────────────────────────
#  THEME / STYLE CONSTANTS
# ─────────────────────────────────────────────────────────────
PL="#111827"; P2="#1A2236"; BD="#1E2D47"
T1="#E8EEFF"; T2="#7B90C4"; FT="'Segoe UI',Arial,sans-serif"
CARD = {"background":PL,"borderRadius":"14px","border":"1px solid #1E2D47",
        "padding":"18px","boxShadow":"0 2px 16px rgba(10,15,30,.5)"}
CL   = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family=FT,size=11,color=T2), margin=dict(l=8,r=8,t=8,b=8))

def inr(v):
    try:
        v = float(v)
        if v >= 1e7: return f"₹{v/1e7:.2f}Cr"
        if v >= 1e5: return f"₹{v/1e5:.1f}L"
        if v >= 1e3: return f"₹{v/1e3:.0f}K"
        return f"₹{int(v)}"
    except: return "₹0"

def efig(msg=""):
    fig = go.Figure()
    if msg:
        fig.add_annotation(text=msg,xref="paper",yref="paper",x=.5,y=.5,
                           showarrow=False,font=dict(size=13,color=T2))
    fig.update_layout(**CL, showlegend=False)
    return fig

def kpi_card(title, vid, sid, ac):
    return html.Div([
        html.Div(style={"height":"3px","background":f"linear-gradient(90deg,{ac},{ac}44)",
                        "margin":"-18px -18px 14px -18px","borderRadius":"3px 3px 0 0"}),
        html.Div(title, style={"fontSize":"10px","color":T2,"textTransform":"uppercase",
                               "letterSpacing":".08em","fontWeight":"700","marginBottom":"8px"}),
        html.Div(id=vid, style={"fontSize":"26px","fontWeight":"800","color":T1,"lineHeight":"1"}),
        html.Div(id=sid, style={"fontSize":"11px","color":T2,"marginTop":"6px"}),
    ], style={**CARD,"paddingTop":"18px"})

def sc_kpi_card(emp_name, color, vid_rev, vid_ord, vid_avg):
    initials = "".join([w[0] for w in emp_name.split()])
    return html.Div([
        html.Div(style={"height":"3px","background":f"linear-gradient(90deg,{color},{color}44)",
                        "margin":"-18px -18px 14px -18px","borderRadius":"3px 3px 0 0"}),
        html.Div(style={"display":"flex","alignItems":"center","gap":"10px","marginBottom":"12px"}, children=[
            html.Div(initials, style={
                "width":"36px","height":"36px","borderRadius":"50%","background":color+"33",
                "border":f"2px solid {color}","color":color,"fontWeight":"800",
                "fontSize":"13px","display":"flex","alignItems":"center","justifyContent":"center",
                "flexShrink":"0",
            }),
            html.Div([
                html.Div(emp_name, style={"fontSize":"13px","fontWeight":"800","color":T1}),
                html.Div("SC / CRM", style={"fontSize":"9px","color":color,"fontWeight":"700",
                                            "textTransform":"uppercase","letterSpacing":".08em","marginTop":"1px"}),
            ]),
        ]),
        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr 1fr","gap":"8px"}, children=[
            html.Div([
                html.Div("Revenue",   style={"fontSize":"8px","color":T2,"textTransform":"uppercase",
                                             "fontWeight":"700","marginBottom":"3px"}),
                html.Div(id=vid_rev,  style={"fontSize":"16px","fontWeight":"800","color":color}),
            ], style={"background":"#0A0F1E","borderRadius":"8px","padding":"8px","textAlign":"center"}),
            html.Div([
                html.Div("Orders",   style={"fontSize":"8px","color":T2,"textTransform":"uppercase",
                                            "fontWeight":"700","marginBottom":"3px"}),
                html.Div(id=vid_ord, style={"fontSize":"16px","fontWeight":"800","color":T1}),
            ], style={"background":"#0A0F1E","borderRadius":"8px","padding":"8px","textAlign":"center"}),
            html.Div([
                html.Div("Avg Order",  style={"fontSize":"8px","color":T2,"textTransform":"uppercase",
                                              "fontWeight":"700","marginBottom":"3px"}),
                html.Div(id=vid_avg,   style={"fontSize":"16px","fontWeight":"800","color":T1}),
            ], style={"background":"#0A0F1E","borderRadius":"8px","padding":"8px","textAlign":"center"}),
        ]),
    ], style={**CARD,"paddingTop":"18px","border":f"1px solid {color}44"})


def make_eu_oem_stats(dff):
    col    = dff["end_user_oem"].str.strip()
    eu_df  = dff[col.str.lower().str.contains("end.?user", regex=True, na=False)].copy()
    oem_df = dff[col.str.lower().str.contains("oem|project", regex=True, na=False)].copy()

    def get_stats(df):
        if df.empty:
            return [], {"rev":"₹0","qty":0,"orders":0,"avg":"₹0","companies":0}
        by_cat = df.groupby("category").agg(
            po_value=("amount","sum"), qty=("qty","sum"), orders=("amount","count")
        ).reset_index().sort_values("po_value", ascending=False)
        totals = {
            "rev":      inr(df["amount"].sum()),
            "qty":      int(df["qty"].sum()),
            "orders":   len(df),
            "avg":      inr(df["amount"].mean()),
            "companies": df["company"].nunique(),
        }
        return by_cat.to_dict("records"), totals

    eu_by_cat,  eu_totals  = get_stats(eu_df)
    oem_by_cat, oem_totals = get_stats(oem_df)
    return eu_by_cat, eu_totals, oem_by_cat, oem_totals


def make_eu_oem_charts(dff):
    col    = dff["end_user_oem"].str.strip()
    eu_df  = dff[col.str.lower().str.contains("end.?user", regex=True, na=False)].copy()
    oem_df = dff[col.str.lower().str.contains("oem|project", regex=True, na=False)].copy()

    cats     = sorted(dff["category"].unique().tolist())
    eu_vals  = [float(eu_df[eu_df["category"]==c]["amount"].sum())  for c in cats]
    oem_vals = [float(oem_df[oem_df["category"]==c]["amount"].sum()) for c in cats]

    fbar = go.Figure()
    fbar.add_trace(go.Bar(
        name="End User", x=cats, y=eu_vals,
        marker_color="#1D9E75", marker_opacity=0.9,
        text=[inr(v) for v in eu_vals], textposition="outside",
        textfont=dict(size=9, color="#1D9E75"),
        hovertemplate="End User — %{x}: ₹%{y:,.0f}<extra></extra>"))
    fbar.add_trace(go.Bar(
        name="OEM", x=cats, y=oem_vals,
        marker_color="#F5A623", marker_opacity=0.9,
        text=[inr(v) for v in oem_vals], textposition="outside",
        textfont=dict(size=9, color="#F5A623"),
        hovertemplate="OEM — %{x}: ₹%{y:,.0f}<extra></extra>"))
    fbar.update_layout(
        **CL, barmode="group", showlegend=True,
        legend=dict(orientation="h",yanchor="bottom",y=1.01,xanchor="right",x=1,
                    font=dict(size=10),bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False,tickfont=dict(size=10),color=T1),
        yaxis=dict(showgrid=True,gridcolor=BD,zeroline=False,color=T2,visible=False))

    eu_qty  = [int(eu_df[eu_df["category"]==c]["qty"].sum())  for c in cats]
    oem_qty = [int(oem_df[oem_df["category"]==c]["qty"].sum()) for c in cats]

    fqty = go.Figure()
    fqty.add_trace(go.Bar(
        name="End User", x=cats, y=eu_qty,
        marker_color="#4F8EF7", marker_opacity=0.9,
        text=eu_qty, textposition="outside", textfont=dict(size=9,color="#4F8EF7"),
        hovertemplate="End User — %{x}: %{y} units<extra></extra>"))
    fqty.add_trace(go.Bar(
        name="OEM", x=cats, y=oem_qty,
        marker_color="#E05C97", marker_opacity=0.9,
        text=oem_qty, textposition="outside", textfont=dict(size=9,color="#E05C97"),
        hovertemplate="OEM — %{x}: %{y} units<extra></extra>"))
    fqty.update_layout(
        **CL, barmode="group", showlegend=True,
        legend=dict(orientation="h",yanchor="bottom",y=1.01,xanchor="right",x=1,
                    font=dict(size=10),bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False,tickfont=dict(size=10),color=T1),
        yaxis=dict(showgrid=True,gridcolor=BD,zeroline=False,color=T2,visible=False))

    return fbar, fqty


# ─────────────────────────────────────────────────────────────
#  APP LAYOUT
# ─────────────────────────────────────────────────────────────
app = dash.Dash(__name__, title="Vaayushanti Sales Dashboard",
                suppress_callback_exceptions=True)
server = app.server

app.index_string = """<!DOCTYPE html><html><head>
{%metas%}<title>{%title%}</title>{%favicon%}{%css%}
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0A0F1E;font-family:'Segoe UI',Arial,sans-serif}
body::before{content:"";position:fixed;inset:0;
  background-image:radial-gradient(rgba(79,142,247,.05) 1px,transparent 1px);
  background-size:30px 30px;pointer-events:none;z-index:0}
#react-entry-point{position:relative;z-index:1}
::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:#0A0F1E}
::-webkit-scrollbar-thumb{background:#1E3A8A;border-radius:3px}
.kh{transition:transform .15s,box-shadow .15s}
.kh:hover{transform:translateY(-2px);box-shadow:0 6px 24px rgba(79,142,247,.2)!important}
.Select-control{background:#1A2236!important;border-color:#1E2D47!important;color:#E8EEFF!important;border-radius:8px!important}
.Select-menu-outer{background:#111827!important;border-color:#1E2D47!important;border-radius:8px!important;z-index:9999!important}
.Select-option{color:#E8EEFF!important;background:#111827!important}
.Select-option:hover{background:#1A2236!important}
.Select-value-label{color:#E8EEFF!important}
.Select-placeholder{color:#7B90C4!important}
.Select-arrow{border-top-color:#7B90C4!important}
.cm{background:#0F1829;border-radius:8px;padding:8px 4px;border:1px solid #1E2D47;
    cursor:pointer;text-align:center;transition:all .15s;font-size:11px;font-weight:600;color:#7B90C4}
.cm:hover{border-color:#4F8EF7;background:#1A2236;color:#E8EEFF}
.cm.on{background:#4F8EF7!important;border-color:#4F8EF7!important;color:#fff!important}
.cm.has{border-color:#1D9E7566;color:#E8EEFF}
.cm.off{opacity:.2;cursor:not-allowed;pointer-events:none}
.yr-btn{background:#0F1829;border:1px solid #1E2D47;color:#7B90C4;border-radius:8px;
        padding:5px 14px;cursor:pointer;font-size:11px;font-weight:600;transition:all .15s}
.yr-btn:hover{border-color:#4F8EF7;color:#E8EEFF}
.yr-btn.yr-on{background:#4F8EF7!important;border-color:#4F8EF7!important;color:#fff!important;font-weight:700}
.sc-btn{background:#0F1829;border:1px solid #1E2D47;color:#7B90C4;border-radius:8px;
        padding:6px 18px;cursor:pointer;font-size:11px;font-weight:600;transition:all .15s;white-space:nowrap}
.sc-btn:hover{border-color:#A78BFA;color:#E8EEFF;background:#1A1030}
.sc-btn.sc-renuka-on{background:#A78BFA!important;border-color:#A78BFA!important;color:#fff!important;font-weight:700}
.sc-btn.sc-jyoti-on{background:#FB923C!important;border-color:#FB923C!important;color:#fff!important;font-weight:700}
.sc-btn.sc-all-on{background:linear-gradient(90deg,#A78BFA,#FB923C)!important;border-color:#A78BFA!important;color:#fff!important;font-weight:700}
.stat-pill{background:#0A0F1E;border-radius:8px;padding:6px 12px;text-align:center;border:1px solid #1E2D47}
</style></head><body>{%app_entry%}
<footer>{%config%}{%scripts%}{%renderer%}</footer></body></html>"""

app.layout = html.Div(
    style={"background":"transparent","minHeight":"100vh","padding":"20px"},
    children=[
    dcc.Store(id="st-yr",   data="ALL"),
    # ✅ FIX: months store — None = all months for year, list = explicit selection
    dcc.Store(id="st-mon",  data=None),
    dcc.Store(id="st-scrm", data="ALL"),
    dcc.Interval(id="tick", interval=30*1000, n_intervals=0, disabled=False),

    # ── Header ──────────────────────────────────────────────────────
    html.Div(style={"borderRadius":"16px","marginBottom":"16px","overflow":"hidden",
                    "boxShadow":"0 4px 30px rgba(30,58,138,.4)","border":"1px solid #1E3A8A44"}, children=[
        html.Div(style={"background":"linear-gradient(120deg,#0D1B3E,#1A3272,#0D1B3E)",
                        "padding":"18px 24px","display":"flex","alignItems":"center",
                        "justifyContent":"space-between","flexWrap":"wrap","gap":"12px"}, children=[
            html.Div(style={"display":"flex","alignItems":"center","gap":"18px"}, children=[
                html.Div("VS", style={"fontSize":"22px","fontWeight":"900","color":"#1E3A8A",
                    "background":"#fff","borderRadius":"12px","width":"58px","height":"58px",
                    "display":"flex","alignItems":"center","justifyContent":"center","flexShrink":"0"}),
                html.Div([
                    html.Div("Vaayushanti Solutions Pvt. Ltd.",
                             style={"fontSize":"20px","fontWeight":"800","color":"#FFF"}),
                    html.Div("Sales Performance Dashboard",
                             style={"fontSize":"11px","color":"rgba(255,255,255,.45)","marginTop":"2px"}),
                ]),
            ]),
            html.Div(style={"textAlign":"right"}, children=[
                html.Div(id="hdr-tag", style={"fontSize":"12px","fontWeight":"700","color":"#F5A623",
                    "background":"rgba(245,166,35,.12)","padding":"6px 18px","borderRadius":"20px",
                    "border":"1px solid rgba(245,166,35,.3)","display":"inline-block"}),
                html.Div(id="hdr-sub", style={"fontSize":"10px","color":"rgba(255,255,255,.3)","marginTop":"5px"}),
                html.Div(style={"display":"flex","gap":"8px","marginTop":"8px",
                                "justifyContent":"flex-end","alignItems":"center"}, children=[
                    html.Button("🔄 Reload", id="btn-reload", n_clicks=0, style={
                        "fontSize":"11px","padding":"5px 16px","borderRadius":"20px","cursor":"pointer",
                        "border":"1px solid rgba(79,142,247,.4)","background":"rgba(79,142,247,.1)",
                        "color":"#4F8EF7","fontWeight":"700"}),
                    html.Div(id="sync-lbl", style={"fontSize":"10px","color":"rgba(255,255,255,.3)"}),
                ]),
            ]),
        ]),
        html.Div(style={"height":"3px","background":"linear-gradient(90deg,#F5A623,rgba(245,166,35,.2),transparent)"}),
    ]),

    # ── Data Source Banner ─────────────────────────────────────────
    html.Div(style={**CARD,"background":"#0A0F1E","marginBottom":"14px","padding":"9px 18px",
                    "display":"flex","alignItems":"center","gap":"12px"}, children=[
        html.Span("🌐", style={"fontSize":"15px"}),
        html.Div([
            html.Div("Data Source", style={"fontSize":"9px","color":T2,"fontWeight":"700","textTransform":"uppercase"}),
            html.Div(SOURCE_LABEL, style={"fontSize":"10px","color":"#4F8EF7","fontFamily":"monospace","marginTop":"1px"}),
            html.Div((SOURCE_URL[:80]+"...") if len(SOURCE_URL) > 80 else SOURCE_URL,
                     style={"fontSize":"9px","color":T2,"fontFamily":"monospace","marginTop":"1px","opacity":"0.6"}),
        ]),
        html.Div(id="file-lbl", style={"marginLeft":"auto","fontSize":"11px","fontWeight":"700"}),
    ]),

    # ── Period Selector ─────────────────────────────────────────────
    html.Div(style={**CARD,"background":P2,"marginBottom":"14px"}, children=[
        html.Div(style={"display":"flex","alignItems":"center","gap":"10px","marginBottom":"12px"}, children=[
            html.Div(style={"width":"4px","height":"18px","borderRadius":"2px",
                            "background":"linear-gradient(180deg,#4F8EF7,#F5A623)"}),
            html.Div("📅 Select Period", style={"fontSize":"12px","fontWeight":"700","color":T1,
                                                "textTransform":"uppercase","letterSpacing":".05em"}),
        ]),
        html.Div(style={"display":"flex","gap":"8px","alignItems":"center","marginBottom":"12px",
                        "flexWrap":"wrap"}, children=[
            html.Div("Year:", style={"fontSize":"10px","color":T2,"fontWeight":"700",
                                     "textTransform":"uppercase","marginRight":"4px"}),
            html.Button("All Years", id="yr-ALL", n_clicks=0, className="yr-btn yr-on"),
        ] + [
            html.Button(str(y), id=f"yr-{y}", n_clicks=0, className="yr-btn")
            for y in ALL_YEARS
        ] + [
            html.Div(style={"marginLeft":"auto","display":"flex","gap":"6px"}, children=[
                html.Button("All Months", id="btn-mall", n_clicks=0,
                    style={"background":"rgba(79,142,247,.12)","border":"1px solid rgba(79,142,247,.4)",
                           "color":"#4F8EF7","borderRadius":"6px","padding":"4px 14px",
                           "cursor":"pointer","fontSize":"10px","fontWeight":"700"}),
                html.Button("Clear", id="btn-mclr", n_clicks=0,
                    style={"background":"#0F1829","border":"1px solid #1E2D47","color":T2,
                           "borderRadius":"6px","padding":"4px 10px","cursor":"pointer","fontSize":"10px"}),
            ]),
        ]),
        html.Div(id="mon-grid", style={"display":"grid","gridTemplateColumns":"repeat(6,1fr)","gap":"8px"}),
        html.Div(id="mon-info", style={"marginTop":"10px","fontSize":"10px","color":T2,
                                       "padding":"6px 12px","background":"#0F1829","borderRadius":"6px"}),
    ]),

    # ── Filters ────────────────────────────────────────────────────
    html.Div(style={**CARD,"background":P2,"marginBottom":"14px","display":"flex",
                    "gap":"16px","flexWrap":"wrap","alignItems":"flex-end"}, children=[
        html.Div(style={"display":"flex","alignItems":"center","gap":"8px","alignSelf":"center"}, children=[
            html.Div(style={"width":"4px","height":"18px","borderRadius":"2px","background":"#F5A623"}),
            html.Div("Filters", style={"fontSize":"12px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
        ]),
        html.Div([
            html.Label("Category", style={"fontSize":"10px","color":T2,"display":"block",
                                          "marginBottom":"4px","fontWeight":"700","textTransform":"uppercase"}),
            dcc.Dropdown(ALL_CATS, ALL_CATS, id="f-cat", multi=True,
                         style={"minWidth":"160px","fontSize":"11px"}),
        ]),
        html.Div([
            html.Label("Sales Person", style={"fontSize":"10px","color":T2,"display":"block",
                                              "marginBottom":"4px","fontWeight":"700","textTransform":"uppercase"}),
            dcc.Dropdown(ALL_SPS, ALL_SPS, id="f-sp", multi=True,
                         style={"minWidth":"200px","fontSize":"11px"}),
        ]),
        html.Div([
            html.Label("Company", style={"fontSize":"10px","color":T2,"display":"block",
                                         "marginBottom":"4px","fontWeight":"700","textTransform":"uppercase"}),
            dcc.Dropdown(ALL_COS, ALL_COS, id="f-co", multi=True,
                         style={"minWidth":"230px","fontSize":"11px"}),
        ]),
    ]),

    # ── KPI Cards ──────────────────────────────────────────────────
    html.Div(style={"display":"grid","gridTemplateColumns":"repeat(3,1fr)",
                    "gap":"12px","marginBottom":"14px"}, children=[
        html.Div(kpi_card("Total Revenue",    "kv-rev","ks-rev","#4F8EF7"), className="kh"),
        html.Div(kpi_card("Avg Order Value",  "kv-avg","ks-avg","#F5A623"), className="kh"),
        html.Div(kpi_card("Total Units Sold", "kv-qty","ks-qty","#E05C97"), className="kh"),
    ]),

    # ── Charts Row 1 ───────────────────────────────────────────────
    html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"12px","marginBottom":"14px"}, children=[
        html.Div([
            html.Div(style={"display":"flex","alignItems":"center","gap":"8px","marginBottom":"10px"}, children=[
                html.Div(style={"width":"4px","height":"14px","borderRadius":"2px","background":"#4F8EF7"}),
                html.Div("Month-wise Revenue", style={"fontSize":"11px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
            ]),
            dcc.Graph(id="g-mon", config={"displayModeBar":False}, style={"height":"220px"}),
        ], style=CARD),
        html.Div([
            html.Div(style={"display":"flex","alignItems":"center","gap":"8px","marginBottom":"10px"}, children=[
                html.Div(style={"width":"4px","height":"14px","borderRadius":"2px","background":"#F5A623"}),
                html.Div("Category Trends", style={"fontSize":"11px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
            ]),
            dcc.Graph(id="g-cat", config={"displayModeBar":False}, style={"height":"220px"}),
        ], style=CARD),
    ]),

    # ── Charts Row 2 ───────────────────────────────────────────────
    html.Div(style={"display":"grid","gridTemplateColumns":"2fr 1fr","gap":"12px","marginBottom":"14px"}, children=[
        html.Div([
            html.Div(style={"display":"flex","justifyContent":"space-between","alignItems":"center","marginBottom":"10px"}, children=[
                html.Div(style={"display":"flex","alignItems":"center","gap":"8px"}, children=[
                    html.Div(style={"width":"4px","height":"14px","borderRadius":"2px","background":"#4F8EF7"}),
                    html.Div("Revenue Trend", style={"fontSize":"11px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
                ]),
                html.Div(style={"display":"flex","gap":"5px"}, children=[
                    html.Button("Daily",   id="btn-daily",   n_clicks=0,
                        style={"fontSize":"10px","padding":"3px 10px","borderRadius":"6px","cursor":"pointer",
                               "border":"1px solid #1E2D47","background":"#0F1829","color":T2}),
                    html.Button("Monthly", id="btn-monthly", n_clicks=1,
                        style={"fontSize":"10px","padding":"3px 10px","borderRadius":"6px","cursor":"pointer",
                               "border":"1px solid #4F8EF7","background":"#4F8EF7","color":"#fff","fontWeight":"700"}),
                ]),
            ]),
            dcc.Graph(id="g-trend", config={"displayModeBar":False}, style={"height":"220px"}),
        ], style=CARD),
        html.Div([
            html.Div(style={"display":"flex","alignItems":"center","gap":"8px","marginBottom":"8px"}, children=[
                html.Div(style={"width":"4px","height":"14px","borderRadius":"2px","background":"#F5A623"}),
                html.Div("Revenue by Category", style={"fontSize":"11px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
            ]),
            dcc.Graph(id="g-donut", config={"displayModeBar":False}, style={"height":"170px"}),
            html.Div(id="g-legend", style={"display":"flex","flexWrap":"wrap","gap":"5px",
                                           "justifyContent":"center","paddingTop":"6px"}),
        ], style=CARD),
    ]),

    # ── Sales Person / Item Charts ──────────────────────────────────
    html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"12px","marginBottom":"14px"}, children=[
        html.Div([
            html.Div(style={"display":"flex","alignItems":"center","gap":"8px","marginBottom":"10px"}, children=[
                html.Div(style={"width":"4px","height":"14px","borderRadius":"2px","background":"#1D9E75"}),
                html.Div("Top Sales Persons", style={"fontSize":"11px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
            ]),
            dcc.Graph(id="g-sp", config={"displayModeBar":False}, style={"height":"240px"}),
        ], style=CARD),
        html.Div([
            html.Div(style={"display":"flex","alignItems":"center","gap":"8px","marginBottom":"10px"}, children=[
                html.Div(style={"width":"4px","height":"14px","borderRadius":"2px","background":"#E05C97"}),
                html.Div("Top Items by Revenue", style={"fontSize":"11px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
            ]),
            dcc.Graph(id="g-item", config={"displayModeBar":False}, style={"height":"240px"}),
        ], style=CARD),
    ]),

    # ── End User / OEM Section ─────────────────────────────────────
    html.Div(style={**CARD,"background":"linear-gradient(135deg,#071410,#0A1A14,#071018)",
                    "marginBottom":"14px","border":"1px solid #1D9E7544"}, children=[
        html.Div(style={"display":"flex","alignItems":"center","gap":"10px","marginBottom":"18px"}, children=[
            html.Div(style={"width":"4px","height":"22px","borderRadius":"2px",
                            "background":"linear-gradient(180deg,#1D9E75,#F5A623)"}),
            html.Div([
                html.Div("End User  /  OEM  —  Detailed Analysis",
                         style={"fontSize":"13px","fontWeight":"800","color":T1,"textTransform":"uppercase","letterSpacing":".06em"}),
                html.Div("Item Quantity & PO Value — EU aur OEM ka alag breakdown",
                         style={"fontSize":"9px","color":"#1D9E75","fontWeight":"600","textTransform":"uppercase","letterSpacing":".06em"}),
            ]),
        ]),
        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"14px","marginBottom":"18px"}, children=[
            html.Div(style={"background":"#071A10","borderRadius":"12px","padding":"14px 18px","border":"1px solid #1D9E7555"}, children=[
                html.Div(style={"display":"flex","alignItems":"center","gap":"8px","marginBottom":"14px"}, children=[
                    html.Div(style={"width":"10px","height":"10px","borderRadius":"50%","background":"#1D9E75","boxShadow":"0 0 8px #1D9E75"}),
                    html.Div("END USER", style={"fontSize":"11px","fontWeight":"800","color":"#1D9E75","letterSpacing":".1em"}),
                ]),
                html.Div(style={"display":"grid","gridTemplateColumns":"repeat(4,1fr)","gap":"8px"}, children=[
                    html.Div([html.Div("PO Value",   style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"4px"}),
                              html.Div(id="eu-total-rev", style={"fontSize":"18px","fontWeight":"900","color":"#1D9E75"})],
                             style={"background":"#0A0F1E","borderRadius":"8px","padding":"10px","textAlign":"center"}),
                    html.Div([html.Div("Item Qty",   style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"4px"}),
                              html.Div(id="eu-total-qty", style={"fontSize":"18px","fontWeight":"900","color":T1})],
                             style={"background":"#0A0F1E","borderRadius":"8px","padding":"10px","textAlign":"center"}),
                    html.Div([html.Div("Orders",     style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"4px"}),
                              html.Div(id="eu-total-ord", style={"fontSize":"18px","fontWeight":"900","color":T1})],
                             style={"background":"#0A0F1E","borderRadius":"8px","padding":"10px","textAlign":"center"}),
                    html.Div([html.Div("Companies",  style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"4px"}),
                              html.Div(id="eu-total-cos", style={"fontSize":"18px","fontWeight":"900","color":T1})],
                             style={"background":"#0A0F1E","borderRadius":"8px","padding":"10px","textAlign":"center"}),
                ]),
            ]),
            html.Div(style={"background":"#1A1200","borderRadius":"12px","padding":"14px 18px","border":"1px solid #F5A62355"}, children=[
                html.Div(style={"display":"flex","alignItems":"center","gap":"8px","marginBottom":"14px"}, children=[
                    html.Div(style={"width":"10px","height":"10px","borderRadius":"50%","background":"#F5A623","boxShadow":"0 0 8px #F5A623"}),
                    html.Div("OEM / PROJECT", style={"fontSize":"11px","fontWeight":"800","color":"#F5A623","letterSpacing":".1em"}),
                ]),
                html.Div(style={"display":"grid","gridTemplateColumns":"repeat(4,1fr)","gap":"8px"}, children=[
                    html.Div([html.Div("PO Value",   style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"4px"}),
                              html.Div(id="oem-total-rev", style={"fontSize":"18px","fontWeight":"900","color":"#F5A623"})],
                             style={"background":"#0A0F1E","borderRadius":"8px","padding":"10px","textAlign":"center"}),
                    html.Div([html.Div("Item Qty",   style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"4px"}),
                              html.Div(id="oem-total-qty", style={"fontSize":"18px","fontWeight":"900","color":T1})],
                             style={"background":"#0A0F1E","borderRadius":"8px","padding":"10px","textAlign":"center"}),
                    html.Div([html.Div("Orders",     style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"4px"}),
                              html.Div(id="oem-total-ord", style={"fontSize":"18px","fontWeight":"900","color":T1})],
                             style={"background":"#0A0F1E","borderRadius":"8px","padding":"10px","textAlign":"center"}),
                    html.Div([html.Div("Companies",  style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"4px"}),
                              html.Div(id="oem-total-cos", style={"fontSize":"18px","fontWeight":"900","color":T1})],
                             style={"background":"#0A0F1E","borderRadius":"8px","padding":"10px","textAlign":"center"}),
                ]),
            ]),
        ]),
        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"14px","marginBottom":"18px"}, children=[
            html.Div(style={"background":"#071A10","borderRadius":"10px","padding":"14px","border":"1px solid #1D9E7533"}, children=[
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"12px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"#1D9E75"}),
                    html.Div("End User — Category Breakdown", style={"fontSize":"10px","fontWeight":"700","color":"#1D9E75","textTransform":"uppercase"}),
                ]),
                html.Div(id="eu-cat-table"),
            ]),
            html.Div(style={"background":"#1A1200","borderRadius":"10px","padding":"14px","border":"1px solid #F5A62333"}, children=[
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"12px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"#F5A623"}),
                    html.Div("OEM / Project — Category Breakdown", style={"fontSize":"10px","fontWeight":"700","color":"#F5A623","textTransform":"uppercase"}),
                ]),
                html.Div(id="oem-cat-table"),
            ]),
        ]),
        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"14px"}, children=[
            html.Div(style={"background":"#0A1020","borderRadius":"10px","padding":"14px","border":"1px solid #1D9E7522"}, children=[
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"8px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"linear-gradient(180deg,#1D9E75,#F5A623)"}),
                    html.Div("PO Value — EU vs OEM by Category", style={"fontSize":"10px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
                ]),
                dcc.Graph(id="eu-oem-bar", config={"displayModeBar":False}, style={"height":"220px"}),
            ]),
            html.Div(style={"background":"#0A1020","borderRadius":"10px","padding":"14px","border":"1px solid #4F8EF722"}, children=[
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"8px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"linear-gradient(180deg,#4F8EF7,#E05C97)"}),
                    html.Div("Item Quantity — EU vs OEM by Category", style={"fontSize":"10px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
                ]),
                dcc.Graph(id="eu-oem-qty-bar", config={"displayModeBar":False}, style={"height":"220px"}),
            ]),
        ]),
    ]),

    # ── Existing / New Section ─────────────────────────────────────
    html.Div(style={**CARD,"background":"linear-gradient(135deg,#0D1B2A,#101E30,#0A1525)",
                    "marginBottom":"14px","border":"1px solid #4F8EF733"}, children=[
        html.Div(style={"display":"flex","alignItems":"center","gap":"10px","marginBottom":"16px"}, children=[
            html.Div(style={"width":"4px","height":"20px","borderRadius":"2px","background":"linear-gradient(180deg,#4F8EF7,#1D9E75)"}),
            html.Div([
                html.Div("Existing / New Customer Analysis",
                         style={"fontSize":"13px","fontWeight":"800","color":T1,"textTransform":"uppercase","letterSpacing":".06em"}),
                html.Div("Customer type split — Existing, New, Old",
                         style={"fontSize":"9px","color":"#4F8EF7","fontWeight":"600","textTransform":"uppercase","letterSpacing":".08em"}),
            ]),
        ]),
        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"16px"}, children=[
            html.Div([
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"10px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"#1D9E75"}),
                    html.Div("Existing / New — Customer Split", style={"fontSize":"10px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
                ]),
                dcc.Graph(id="en-g-donut", config={"displayModeBar":False}, style={"height":"220px"}),
                html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr 1fr","gap":"8px","marginTop":"10px"}, children=[
                    html.Div([html.Div("Revenue",style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"3px"}),
                              html.Div(id="en-rev",style={"fontSize":"14px","fontWeight":"800","color":"#1D9E75"})],className="stat-pill"),
                    html.Div([html.Div("Orders",style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"3px"}),
                              html.Div(id="en-ord",style={"fontSize":"14px","fontWeight":"800","color":T1})],className="stat-pill"),
                    html.Div([html.Div("Total Qty",style={"fontSize":"8px","color":T2,"textTransform":"uppercase","fontWeight":"700","marginBottom":"3px"}),
                              html.Div(id="en-qty",style={"fontSize":"14px","fontWeight":"800","color":T1})],className="stat-pill"),
                ]),
            ], style={"background":"#0A1020","borderRadius":"10px","padding":"14px","border":"1px solid #1D9E7533"}),
            html.Div([
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"10px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"#4F8EF7"}),
                    html.Div("Existing / New — Category Revenue", style={"fontSize":"10px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
                ]),
                dcc.Graph(id="en-g-cat", config={"displayModeBar":False}, style={"height":"220px"}),
                html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr 1fr","gap":"8px","marginTop":"10px"}, children=[
                    html.Div([html.Div("Existing",style={"fontSize":"8px","color":"#4F8EF7","textTransform":"uppercase","fontWeight":"700","marginBottom":"3px"}),
                              html.Div(id="en-exist-rev",style={"fontSize":"14px","fontWeight":"800","color":"#4F8EF7"})],className="stat-pill"),
                    html.Div([html.Div("New",style={"fontSize":"8px","color":"#1D9E75","textTransform":"uppercase","fontWeight":"700","marginBottom":"3px"}),
                              html.Div(id="en-new-rev",style={"fontSize":"14px","fontWeight":"800","color":"#1D9E75"})],className="stat-pill"),
                    html.Div([html.Div("Old",style={"fontSize":"8px","color":"#FB923C","textTransform":"uppercase","fontWeight":"700","marginBottom":"3px"}),
                              html.Div(id="en-old-rev",style={"fontSize":"14px","fontWeight":"800","color":"#FB923C"})],className="stat-pill"),
                ]),
            ], style={"background":"#0A1020","borderRadius":"10px","padding":"14px","border":"1px solid #4F8EF733"}),
        ]),
    ]),

    # ── SC / CRM Section ─────────────────────────────────────────
    html.Div(style={**CARD,"background":"linear-gradient(135deg,#150D2A,#1A1030,#0F1420)",
                    "marginBottom":"14px","border":"1px solid #A78BFA33"}, children=[
        html.Div(style={"display":"flex","alignItems":"center","justifyContent":"space-between",
                        "marginBottom":"16px","flexWrap":"wrap","gap":"10px"}, children=[
            html.Div(style={"display":"flex","alignItems":"center","gap":"10px"}, children=[
                html.Div(style={"width":"4px","height":"20px","borderRadius":"2px","background":"linear-gradient(180deg,#A78BFA,#FB923C)"}),
                html.Div([
                    html.Div("SC / CRM Performance",
                             style={"fontSize":"13px","fontWeight":"800","color":T1,"textTransform":"uppercase","letterSpacing":".06em"}),
                    html.Div("Sales Coordinator & CRM Team",
                             style={"fontSize":"9px","color":"#A78BFA","fontWeight":"600","textTransform":"uppercase","letterSpacing":".08em"}),
                ]),
            ]),
            html.Div(style={"display":"flex","gap":"8px","alignItems":"center","flexWrap":"wrap"}, children=[
                html.Div("View:", style={"fontSize":"10px","color":T2,"fontWeight":"700","textTransform":"uppercase","marginRight":"2px"}),
                html.Button("👥 Both",         id="sc-btn-all",    n_clicks=0, className="sc-btn sc-all-on"),
                html.Button("🟣 Renuka Arya",  id="sc-btn-renuka", n_clicks=0, className="sc-btn"),
                html.Button("🟠 Jyoti Sahu",   id="sc-btn-jyoti",  n_clicks=0, className="sc-btn"),
            ]),
        ]),
        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"12px","marginBottom":"14px"}, children=[
            sc_kpi_card("Renuka Arya","#A78BFA","sc-renuka-rev","sc-renuka-ord","sc-renuka-avg"),
            sc_kpi_card("Jyoti Sahu", "#FB923C","sc-jyoti-rev", "sc-jyoti-ord", "sc-jyoti-avg"),
        ]),
        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr 1fr","gap":"12px"}, children=[
            html.Div([
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"8px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"linear-gradient(180deg,#A78BFA,#FB923C)"}),
                    html.Div("Monthly Revenue", style={"fontSize":"10px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
                ]),
                dcc.Graph(id="sc-g-mon", config={"displayModeBar":False}, style={"height":"200px"}),
            ], style={"background":"#0F1020","borderRadius":"10px","padding":"12px","border":"1px solid #A78BFA22"}),
            html.Div([
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"8px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"#FB923C"}),
                    html.Div("Category Split", style={"fontSize":"10px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
                ]),
                dcc.Graph(id="sc-g-cat", config={"displayModeBar":False}, style={"height":"200px"}),
            ], style={"background":"#0F1020","borderRadius":"10px","padding":"12px","border":"1px solid #FB923C22"}),
            html.Div([
                html.Div(style={"display":"flex","alignItems":"center","gap":"6px","marginBottom":"8px"}, children=[
                    html.Div(style={"width":"3px","height":"12px","borderRadius":"2px","background":"#A78BFA"}),
                    html.Div("Head-to-Head", style={"fontSize":"10px","fontWeight":"700","color":T1,"textTransform":"uppercase"}),
                ]),
                dcc.Graph(id="sc-g-cmp", config={"displayModeBar":False}, style={"height":"200px"}),
            ], style={"background":"#0F1020","borderRadius":"10px","padding":"12px","border":"1px solid #A78BFA22"}),
        ]),
    ]),

    # ── All Orders Table ───────────────────────────────────────────
    html.Div([
        html.Div(style={"display":"flex","justifyContent":"space-between","alignItems":"center","marginBottom":"12px"}, children=[
            html.Div(style={"display":"flex","alignItems":"center","gap":"8px"}, children=[
                html.Div(style={"width":"4px","height":"16px","borderRadius":"2px","background":"linear-gradient(180deg,#4F8EF7,#1D9E75)"}),
                html.Div("All Orders", style={"fontSize":"12px","fontWeight":"700","color":T1,"textTransform":"uppercase","letterSpacing":".05em"}),
            ]),
            html.Div(id="tbl-count", style={"fontSize":"11px","color":T2,"background":"#0F1829",
                                            "padding":"4px 14px","borderRadius":"20px","border":"1px solid #1E2D47"}),
        ]),
        dash_table.DataTable(
            id="tbl",
            columns=[
                {"name":"Date",            "id":"day_str",        "type":"text"},
                {"name":"Company Name",    "id":"company",        "type":"text"},
                {"name":"Item / Category", "id":"item",           "type":"text"},
                {"name":"Category",        "id":"category",       "type":"text"},
                {"name":"Qty",             "id":"qty",            "type":"numeric"},
                {"name":"Value of PO (₹)", "id":"amt_num",        "type":"numeric"},
                {"name":"Sales Person",    "id":"sp_clean",       "type":"text"},
                {"name":"Office SC/CRM",   "id":"office_sp_clean","type":"text"},
                {"name":"End User/OEM",    "id":"end_user_oem",   "type":"text"},
                {"name":"Existing/New",    "id":"existing_new",   "type":"text"},
            ],
            style_table={"overflowX":"auto","maxHeight":"420px","overflowY":"auto"},
            style_header={"background":"#0A0F1E","fontWeight":"700","fontSize":"10px","color":T2,
                          "border":"none","borderBottom":"2px solid #1E2D47","textTransform":"uppercase",
                          "padding":"10px 10px","letterSpacing":".04em"},
            style_filter={"background":"#0F1829","border":"none","borderBottom":"1px solid #1E2D47",
                          "color":T1,"fontSize":"11px","padding":"4px 8px"},
            style_cell={"border":"none","borderBottom":"1px solid #1A2236","padding":"7px 10px",
                        "fontFamily":FT,"color":T1,"background":"#111827","fontSize":"12px",
                        "overflow":"hidden","textOverflow":"ellipsis","whiteSpace":"nowrap"},
            style_cell_conditional=[
                {"if":{"column_id":"day_str"},          "minWidth":"105px","maxWidth":"115px","textAlign":"center"},
                {"if":{"column_id":"company"},          "minWidth":"200px","maxWidth":"300px"},
                {"if":{"column_id":"item"},             "minWidth":"100px","maxWidth":"140px"},
                {"if":{"column_id":"category"},         "minWidth":"100px","maxWidth":"120px","textAlign":"center"},
                {"if":{"column_id":"qty"},              "minWidth":"55px", "maxWidth":"70px", "textAlign":"right"},
                {"if":{"column_id":"amt_num"},          "minWidth":"120px","maxWidth":"140px","textAlign":"right"},
                {"if":{"column_id":"sp_clean"},         "minWidth":"130px","maxWidth":"180px"},
                {"if":{"column_id":"office_sp_clean"},  "minWidth":"130px","maxWidth":"180px"},
                {"if":{"column_id":"end_user_oem"},     "minWidth":"110px","maxWidth":"130px","textAlign":"center"},
                {"if":{"column_id":"existing_new"},     "minWidth":"90px", "maxWidth":"110px","textAlign":"center"},
            ],
            style_data_conditional=[
                {"if":{"row_index":"odd"}, "background":"#141C2E"},
                {"if":{"filter_query":'({category} = "Bags")',         "column_id":"category"}, "color":"#5DCAA5","fontWeight":"700"},
                {"if":{"filter_query":'({category} = "Cages")',        "column_id":"category"}, "color":"#85B7EB","fontWeight":"700"},
                {"if":{"filter_query":'({category} = "Projects")',     "column_id":"category"}, "color":"#FAC775","fontWeight":"700"},
                {"if":{"filter_query":'({category} = "Trading Item")', "column_id":"category"}, "color":"#ED93B1","fontWeight":"700"},
                {"if":{"column_id":"amt_num"},      "color":"#4F8EF7","fontWeight":"700"},
                {"if":{"column_id":"qty"},          "color":"#C8D8F0","fontWeight":"600"},
                {"if":{"column_id":"day_str"},      "color":"#9FB4D8"},
                {"if":{"filter_query":'({end_user_oem} = "End-User")',      "column_id":"end_user_oem"}, "color":"#1D9E75","fontWeight":"700"},
                {"if":{"filter_query":'({end_user_oem} = "Project (OEM)")', "column_id":"end_user_oem"}, "color":"#F5A623","fontWeight":"700"},
                {"if":{"filter_query":'({existing_new} = "New")',       "column_id":"existing_new"}, "color":"#4F8EF7","fontWeight":"700"},
                {"if":{"filter_query":'({existing_new} = "Existing")',  "column_id":"existing_new"}, "color":"#A78BFA","fontWeight":"700"},
                {"if":{"filter_query":'({existing_new} = "Old")',       "column_id":"existing_new"}, "color":"#FB923C","fontWeight":"700"},
            ],
            sort_action="native",
            filter_action="native",
            page_size=100,
            fixed_rows={"headers":True},
        ),
    ], style={**CARD,"marginBottom":"14px"}),

    # ── Footer ────────────────────────────────────────────────────
    html.Div(style={"display":"flex","alignItems":"center","justifyContent":"center",
                    "gap":"8px","marginTop":"4px","paddingBottom":"8px"}, children=[
        html.Span("VS", style={"fontSize":"11px","fontWeight":"900","color":"#1E3A8A",
                               "background":"white","borderRadius":"4px","padding":"1px 5px"}),
        html.Span("Vaayushanti Solutions Pvt. Ltd.", style={"fontSize":"10px","color":T2,"fontWeight":"700"}),
        html.Div(style={"width":"3px","height":"3px","borderRadius":"50%","background":"#F5A623"}),
        html.Span("Live Google Sheets Dashboard", style={"fontSize":"10px","color":T2}),
    ]),
])


# ─────────────────────────────────────────────────────────────
#  ✅ CORE FIX: apply_filters — single consistent function
#
#  months=None  → show ALL data for selected year (no month filter)
#  months=[]    → same as None (no month filter) — safety fallback
#  months=[1,3] → filter to only those months
#
#  KEY INSIGHT: The month-wise revenue chart was broken because
#  callbacks used different filtering logic. Now ALL callbacks
#  use this single function.
# ─────────────────────────────────────────────────────────────
def apply_filters(yr, months, cats=None, sps=None, cos=None):
    if DF.empty:
        return pd.DataFrame()
    dff = DF.copy()

    # Year filter
    if yr and yr != "ALL":
        try:
            dff = dff[dff["year"] == int(yr)]
        except (ValueError, TypeError):
            pass

    # Month filter — ONLY when user has explicitly selected months (non-empty list)
    if months:  # None or [] both mean "no month filter"
        dff = dff[dff["month"].isin(months)]

    # Category filter
    if cats:
        dff = dff[dff["category"].isin(cats)]

    # Salesperson filter
    if sps:
        dff = dff[dff["sp_clean"].isin(sps)]

    # Company filter
    if cos:
        dff = dff[dff["company"].isin(cos)]

    return dff


# ─────────────────────────────────────────────────────────────
#  CALLBACKS
# ─────────────────────────────────────────────────────────────

yr_inputs = [Input("yr-ALL","n_clicks")] + [Input(f"yr-{y}","n_clicks") for y in ALL_YEARS]

@app.callback(
    Output("st-yr",    "data"),
    Output("sync-lbl", "children"),
    Output("file-lbl", "children"),
    Output("tick",     "disabled"),
    *[Output(f"yr-{x}","className") for x in ["ALL"] + ALL_YEARS],
    *yr_inputs,
    Input("btn-reload","n_clicks"),
    Input("tick","n_intervals"),
    State("st-yr","data"),
    prevent_initial_call=False)
def cb_year(*args):
    global DF, ALL_YEARS, ALL_CATS, ALL_SPS, ALL_COS, LAST_SYNC

    cur_yr = args[-1] if args[-1] is not None else "ALL"
    tid    = ctx.triggered_id

    if tid in ("btn-reload", "tick"):
        DF = load_data()
        if not DF.empty:
            ALL_YEARS[:] = sorted(DF["year"].unique().tolist())
            ALL_CATS[:]  = sorted(DF["category"].unique().tolist())
            ALL_SPS[:]   = sorted(DF["sp_clean"].unique().tolist())
            ALL_COS[:]   = sorted(DF["company"].unique().tolist())
        LAST_SYNC = datetime.now().strftime("%d %b %H:%M")
        sel_yr = cur_yr
    elif tid == "yr-ALL":
        sel_yr = "ALL"
    elif isinstance(tid, str) and tid.startswith("yr-") and tid != "yr-ALL":
        try:
            sel_yr = int(tid.replace("yr-", ""))
        except ValueError:
            sel_yr = "ALL"
    else:
        sel_yr = "ALL"

    all_keys = ["ALL"] + ALL_YEARS
    classes  = ["yr-btn yr-on" if str(y) == str(sel_yr) else "yr-btn"
                for y in all_keys]

    if not DF.empty:
        flbl = html.Span([
            html.Span("✅ Connected", style={"color":"#1D9E75"}),
            html.Span(f"  |  {len(DF)} rows | Last sync: {LAST_SYNC}",
                      style={"color":T2,"fontWeight":"400"}),
        ])
    else:
        flbl = html.Span([
            html.Span("❌ No Data", style={"color":"#E05C97"}),
            html.Span("  |  SHEET_CSV_URLS check karo", style={"color":T2,"fontWeight":"400"}),
        ])

    return [sel_yr, "Synced: " + LAST_SYNC, flbl, False] + classes


# ── Month selection ──────────────────────────────────────────
@app.callback(
    Output("st-mon","data"),
    Input({"type":"mb","index":dash.dependencies.ALL},"n_clicks"),
    Input("btn-mall","n_clicks"),
    Input("btn-mclr","n_clicks"),
    Input("st-yr","data"),
    State("st-mon","data"),
    prevent_initial_call=True)
def cb_mon(mc, mall, mclr, yr, cur):
    tid = ctx.triggered_id
    if DF.empty:
        return None

    # "All Months" button or "Clear" button → reset to auto (None)
    if tid in ("btn-mall", "btn-mclr"):
        return None

    # Year changed → reset months to auto
    if tid == "st-yr":
        return None

    # Month cell clicked
    if isinstance(tid, dict) and tid.get("type") == "mb":
        m = tid["index"]
        if cur is None:
            # In auto mode → clicking a month starts explicit selection with just that month
            return [m]
        s = list(cur)
        if m in s:
            s.remove(m)
            # If all months deselected, go back to auto mode
            return sorted(s) if s else None
        else:
            s.append(m)
            return sorted(s)

    return cur


# ── Month grid render ────────────────────────────────────────
@app.callback(
    Output("mon-grid","children"),
    Output("mon-info","children"),
    Input("st-yr","data"),
    Input("st-mon","data"))
def cb_grid(yr, sel):
    if DF.empty:
        base_months = set()
    else:
        base = DF if (not yr or yr == "ALL") else DF[DF["year"] == int(yr)]
        base_months = set(base["month"].unique())

    is_auto = (sel is None or sel == [])
    # In auto mode, all available months are "selected" (highlighted)
    sel_set = base_months if is_auto else set(sel or [])

    cells = []
    for i, mn in enumerate(MON):
        m = i + 1
        if m not in base_months:
            cls = "cm off"
        elif m in sel_set:
            cls = "cm on"
        else:
            cls = "cm has"
        cells.append(html.Div(mn, id={"type":"mb","index":m}, className=cls, n_clicks=0))

    yr_lbl = "All Years" if (not yr or yr == "ALL") else str(yr)
    if is_auto:
        info = f"📅 {yr_lbl}  |  All months selected  ({len(base_months)} months with data)"
    elif sel:
        info = f"📅 {yr_lbl}  |  " + " • ".join(MON[m-1] for m in sorted(sel)) + f"  ({len(sel)} months selected)"
    else:
        info = f"📅 {yr_lbl}  |  All months selected"
    return cells, info


# ── Dropdown options update on reload ────────────────────────
@app.callback(
    Output("f-cat","options"),
    Output("f-sp", "options"),
    Output("f-co", "options"),
    Input("btn-reload","n_clicks"),
    Input("tick","n_intervals"),
    prevent_initial_call=True)
def cb_dropdown_opts(rel, ti):
    if DF.empty:
        return [], [], []
    return (
        sorted(DF["category"].unique().tolist()),
        sorted(DF["sp_clean"].unique().tolist()),
        sorted(DF["company"].unique().tolist()),
    )


# ── Existing / New analysis ──────────────────────────────────
@app.callback(
    Output("en-g-donut","figure"),
    Output("en-rev","children"),
    Output("en-ord","children"),
    Output("en-qty","children"),
    Output("en-g-cat","figure"),
    Output("en-exist-rev","children"),
    Output("en-new-rev","children"),
    Output("en-old-rev","children"),
    Input("st-yr","data"),
    Input("st-mon","data"),
    Input("f-cat","value"),
    Input("f-sp", "value"),
    Input("f-co", "value"))
def cb_existing_new(yr, months, cats, sps, cos):
    if DF.empty:
        ef = efig("No data")
        return ef,"₹0","0","0", ef,"₹0","₹0","₹0"

    # ✅ Uses consistent apply_filters
    dff      = apply_filters(yr, months, cats, sps, cos)
    EN_COLORS = {"Existing":"#4F8EF7","New":"#1D9E75","Old":"#FB923C"}
    en_df     = dff[dff["existing_new"].str.strip().isin(EN_COLORS)].copy()

    if not en_df.empty:
        ca2 = en_df.groupby("existing_new").agg(
            amount=("amount","sum"), qty=("qty","sum")).reset_index()
        fig2 = go.Figure(go.Pie(
            labels=ca2["existing_new"].tolist(),
            values=ca2["amount"].tolist(),
            hole=0.58,
            marker=dict(colors=[EN_COLORS.get(str(c),"#888") for c in ca2["existing_new"]],
                        line=dict(color="#0F1020", width=2)),
            textinfo="label+percent", textfont=dict(size=9),
            hovertemplate="%{label}<br>₹%{value:,.0f} (%{percent})<br>Qty: %{customdata}<extra></extra>",
            customdata=ca2["qty"].tolist()))
        fig2.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                           font=dict(family=FT,size=9,color=T2),
                           margin=dict(l=4,r=4,t=4,b=4), showlegend=False)
        total_rev = inr(en_df["amount"].sum())
        total_qty = int(en_df["qty"].sum())
        total_ord = len(en_df)
    else:
        fig2 = efig("No Existing/New data")
        total_rev, total_qty, total_ord = "₹0", 0, 0

    cats_avail = sorted(dff["category"].unique().tolist())
    fcat = go.Figure()
    for en_type, col in EN_COLORS.items():
        sub  = en_df[en_df["existing_new"] == en_type]
        vals = [float(sub[sub["category"]==c]["amount"].sum()) for c in cats_avail]
        if any(v > 0 for v in vals):
            fcat.add_trace(go.Bar(
                name=en_type, x=cats_avail, y=vals,
                marker_color=col, marker_opacity=0.88,
                text=[inr(v) if v > 0 else "" for v in vals],
                textposition="outside", textfont=dict(size=8, color=col),
                hovertemplate=f"<b>{en_type}</b><br>%{{x}}: ₹%{{y:,.0f}}<extra></extra>"))
    fcat.update_layout(
        **CL, barmode="group", showlegend=True,
        legend=dict(orientation="h",yanchor="bottom",y=1.01,xanchor="right",x=1,
                    font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False,tickfont=dict(size=10),color=T1),
        yaxis=dict(showgrid=True,gridcolor=BD,zeroline=False,color=T2,visible=False))

    def en_rev(t):
        sub = en_df[en_df["existing_new"] == t]
        return inr(sub["amount"].sum()) if not sub.empty else "₹0"

    return (fig2, total_rev, str(total_ord), str(total_qty),
            fcat, en_rev("Existing"), en_rev("New"), en_rev("Old"))


# ── SC / CRM filter toggle ────────────────────────────────────
@app.callback(
    Output("st-scrm","data"),
    Output("sc-btn-all","className"),
    Output("sc-btn-renuka","className"),
    Output("sc-btn-jyoti","className"),
    Input("sc-btn-all","n_clicks"),
    Input("sc-btn-renuka","n_clicks"),
    Input("sc-btn-jyoti","n_clicks"),
    prevent_initial_call=False)
def cb_scrm_filter(na, nr, nj):
    tid = ctx.triggered_id
    if tid == "sc-btn-renuka":
        return "Renuka Arya", "sc-btn", "sc-btn sc-renuka-on", "sc-btn"
    elif tid == "sc-btn-jyoti":
        return "Jyoti Sahu",  "sc-btn", "sc-btn", "sc-btn sc-jyoti-on"
    else:
        return "ALL", "sc-btn sc-all-on", "sc-btn", "sc-btn"


# ── SC / CRM charts ────────────────────────────────────────────
@app.callback(
    Output("sc-renuka-rev","children"), Output("sc-renuka-ord","children"), Output("sc-renuka-avg","children"),
    Output("sc-jyoti-rev","children"),  Output("sc-jyoti-ord","children"),  Output("sc-jyoti-avg","children"),
    Output("sc-g-mon","figure"),
    Output("sc-g-cat","figure"),
    Output("sc-g-cmp","figure"),
    Input("st-yr","data"),
    Input("st-mon","data"),
    Input("st-scrm","data"))
def cb_scrm(yr, months, scrm_sel):
    if DF.empty:
        ef = efig("No data")
        return ("—","—","—","—","—","—", ef, ef, ef)

    # ✅ FIXED: SC/CRM now uses apply_filters for consistency
    # (no category/sp/co filter here — SC section shows all regardless)
    base = apply_filters(yr, months)

    stats = {}
    for emp in SC_CRM_EMPLOYEES:
        emp_df = base[base["office_sp_clean"].str.contains(emp, case=False, na=False)]
        stats[emp] = {
            "df":  emp_df,
            "rev": float(emp_df["amount"].sum()),
            "ord": len(emp_df),
            "avg": float(emp_df["amount"].mean()) if len(emp_df) else 0.0,
        }

    r = stats["Renuka Arya"]; j = stats["Jyoti Sahu"]
    r_rev, r_ord, r_avg = inr(r["rev"]), str(r["ord"]), inr(r["avg"])
    j_rev, j_ord, j_avg = inr(j["rev"]), str(j["ord"]), inr(j["avg"])

    show_emps = (["Renuka Arya"] if scrm_sel == "Renuka Arya" else
                 ["Jyoti Sahu"]  if scrm_sel == "Jyoti Sahu"  else
                 SC_CRM_EMPLOYEES)

    fmon = go.Figure()
    for emp in show_emps:
        edf = stats[emp]["df"]
        mg  = edf.groupby(["month","month_name"])["amount"].sum().reset_index().sort_values("month")
        col = SC_CRM_COLORS[emp]
        fmon.add_trace(go.Bar(
            name=emp, x=mg["month_name"].tolist(), y=mg["amount"].tolist(),
            marker_color=col, marker_opacity=0.85,
            text=[inr(v) for v in mg["amount"].tolist()],
            textposition="inside", insidetextanchor="end",
            textfont=dict(size=8, color="#fff"),
            hovertemplate=f"<b>{emp}</b><br>%{{x}}: ₹%{{y:,.0f}}<extra></extra>"))
    fmon.update_layout(**CL, barmode="group", showlegend=(len(show_emps)>1),
        legend=dict(orientation="h",yanchor="bottom",y=1.01,xanchor="right",x=1,
                    font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False,tickfont=dict(size=8),color=T2),
        yaxis=dict(showgrid=True,gridcolor=BD,zeroline=False,color=T2,visible=False))

    combined = pd.concat([stats[e]["df"] for e in show_emps])
    if not combined.empty:
        ca = combined.groupby("category")["amount"].sum().reset_index()
        fcat = go.Figure(go.Pie(
            labels=ca["category"].tolist(), values=ca["amount"].tolist(), hole=0.55,
            marker=dict(colors=[CAT_COLORS.get(str(c),"#888") for c in ca["category"]],
                        line=dict(color="#0F1020", width=2)),
            textinfo="label+percent", textfont=dict(size=8),
            hovertemplate="%{label}: ₹%{value:,.0f} (%{percent})<extra></extra>"))
        fcat.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                           font=dict(family=FT,size=10,color=T2),
                           margin=dict(l=2,r=2,t=2,b=2), showlegend=False)
    else:
        fcat = efig("No data")

    cmp_cats    = ["Revenue","Orders","Avg Order"]
    renuka_vals = [r["rev"]/1e5,   r["ord"],   r["avg"]/1e3]
    jyoti_vals  = [j["rev"]/1e5,   j["ord"],   j["avg"]/1e3]
    fcmp = go.Figure()
    if "Renuka Arya" in show_emps:
        fcmp.add_trace(go.Bar(
            name="Renuka Arya", x=cmp_cats, y=renuka_vals,
            marker_color="#A78BFA", marker_opacity=0.85,
            text=[inr(r["rev"]), str(r["ord"]), inr(r["avg"])],
            textposition="outside", textfont=dict(size=9,color="#A78BFA"),
            hovertemplate="<b>Renuka</b><br>%{x}: %{text}<extra></extra>"))
    if "Jyoti Sahu" in show_emps:
        fcmp.add_trace(go.Bar(
            name="Jyoti Sahu", x=cmp_cats, y=jyoti_vals,
            marker_color="#FB923C", marker_opacity=0.85,
            text=[inr(j["rev"]), str(j["ord"]), inr(j["avg"])],
            textposition="outside", textfont=dict(size=9,color="#FB923C"),
            hovertemplate="<b>Jyoti</b><br>%{x}: %{text}<extra></extra>"))
    fcmp.update_layout(**CL, barmode="group", showlegend=(len(show_emps)>1),
        legend=dict(orientation="h",yanchor="bottom",y=1.01,xanchor="right",x=1,
                    font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False,tickfont=dict(size=10),color=T1),
        yaxis=dict(showgrid=True,gridcolor=BD,zeroline=False,color=T2,visible=False))

    return (r_rev, r_ord, r_avg, j_rev, j_ord, j_avg, fmon, fcat, fcmp)


# ── Category breakdown table helper ─────────────────────────
def make_cat_breakdown_table(by_cat_records, accent_color):
    if not by_cat_records:
        return html.Div("No data", style={"color":T2,"fontSize":"11px","textAlign":"center","padding":"20px"})

    header = html.Div(style={
        "display":"grid","gridTemplateColumns":"1.5fr 1fr 1fr 1fr",
        "gap":"4px","marginBottom":"6px","padding":"4px 8px"
    }, children=[
        html.Div("Category", style={"fontSize":"8px","color":T2,"fontWeight":"700","textTransform":"uppercase"}),
        html.Div("PO Value", style={"fontSize":"8px","color":T2,"fontWeight":"700","textTransform":"uppercase","textAlign":"right"}),
        html.Div("Qty",      style={"fontSize":"8px","color":T2,"fontWeight":"700","textTransform":"uppercase","textAlign":"right"}),
        html.Div("Orders",   style={"fontSize":"8px","color":T2,"fontWeight":"700","textTransform":"uppercase","textAlign":"right"}),
    ])
    rows = []
    for i, rec in enumerate(by_cat_records):
        cat = str(rec["category"]); cc = CAT_COLORS.get(cat,"#888")
        rows.append(html.Div(style={
            "display":"grid","gridTemplateColumns":"1.5fr 1fr 1fr 1fr",
            "gap":"4px","padding":"6px 8px","borderRadius":"6px",
            "background":"#0A0F1E" if i%2==0 else "#0D1520",
            "marginBottom":"3px","borderLeft":f"3px solid {cc}",
        }, children=[
            html.Div(style={"display":"flex","alignItems":"center","gap":"5px"}, children=[
                html.Div(style={"width":"6px","height":"6px","borderRadius":"2px","background":cc,"flexShrink":"0"}),
                html.Div(cat, style={"fontSize":"10px","color":T1,"fontWeight":"600"}),
            ]),
            html.Div(inr(rec["po_value"]), style={"fontSize":"11px","color":accent_color,"fontWeight":"800","textAlign":"right"}),
            html.Div(str(int(rec["qty"])),  style={"fontSize":"11px","color":T1,"fontWeight":"600","textAlign":"right"}),
            html.Div(str(int(rec["orders"])),style={"fontSize":"11px","color":T2,"fontWeight":"600","textAlign":"right"}),
        ]))
    return html.Div([header] + rows)


# ── EU / OEM callback ─────────────────────────────────────────
@app.callback(
    Output("eu-total-rev","children"), Output("eu-total-qty","children"),
    Output("eu-total-ord","children"), Output("eu-total-cos","children"),
    Output("oem-total-rev","children"),Output("oem-total-qty","children"),
    Output("oem-total-ord","children"),Output("oem-total-cos","children"),
    Output("eu-cat-table","children"),
    Output("oem-cat-table","children"),
    Output("eu-oem-bar","figure"),
    Output("eu-oem-qty-bar","figure"),
    Input("st-yr","data"),
    Input("st-mon","data"),
    Input("f-cat","value"),
    Input("f-sp", "value"),
    Input("f-co", "value"))
def cb_eu_oem(yr, months, cats, sps, cos):
    if DF.empty:
        ef = efig("No data")
        et = html.Div("No data", style={"color":T2,"fontSize":"11px","textAlign":"center","padding":"20px"})
        return ("₹0","0","0","0","₹0","0","0","0", et, et, ef, ef)

    # ✅ Uses consistent apply_filters
    dff = apply_filters(yr, months, cats, sps, cos)
    eu_by_cat, eu_totals, oem_by_cat, oem_totals = make_eu_oem_stats(dff)
    fbar, fqty = make_eu_oem_charts(dff)
    return (
        eu_totals["rev"],  str(eu_totals["qty"]),  str(eu_totals["orders"]),  str(eu_totals["companies"]),
        oem_totals["rev"], str(oem_totals["qty"]), str(oem_totals["orders"]), str(oem_totals["companies"]),
        make_cat_breakdown_table(eu_by_cat,  "#1D9E75"),
        make_cat_breakdown_table(oem_by_cat, "#F5A623"),
        fbar, fqty,
    )


# ── Main dashboard callback ───────────────────────────────────
@app.callback(
    Output("hdr-tag","children"), Output("hdr-sub","children"),
    Output("kv-rev","children"),  Output("ks-rev","children"),
    Output("kv-avg","children"),  Output("ks-avg","children"),
    Output("kv-qty","children"),  Output("ks-qty","children"),
    Output("g-mon","figure"),
    Output("g-cat","figure"),
    Output("g-trend","figure"),
    Output("g-donut","figure"),
    Output("g-legend","children"),
    Output("g-sp","figure"),
    Output("g-item","figure"),
    Output("tbl","data"),
    Output("tbl-count","children"),
    Output("btn-daily","style"),
    Output("btn-monthly","style"),
    Input("st-yr","data"),
    Input("st-mon","data"),
    Input("f-cat","value"),
    Input("f-sp", "value"),
    Input("f-co", "value"),
    Input("btn-daily","n_clicks"),
    Input("btn-monthly","n_clicks"),
    Input("btn-reload","n_clicks"),
    Input("tick","n_intervals"))
def cb_main(yr, months, cats, sps, cos, nd, nm, rel, ti):
    s_on  = {"fontSize":"10px","padding":"3px 10px","borderRadius":"6px","cursor":"pointer",
             "border":"1px solid #4F8EF7","background":"#4F8EF7","color":"#fff","fontWeight":"700"}
    s_off = {"fontSize":"10px","padding":"3px 10px","borderRadius":"6px","cursor":"pointer",
             "border":"1px solid #1E2D47","background":"#0F1829","color":T2}
    monthly = (ctx.triggered_id != "btn-daily") and ((nm or 0) >= (nd or 0))
    d_sty, m_sty = (s_off, s_on) if monthly else (s_on, s_off)

    if DF.empty:
        return ("—","—","—","—","—","—",
                efig(),efig(),efig(),go.Figure(),[],efig(),efig(),[],"-",d_sty,m_sty)

    # ✅ PRIMARY FILTERED DATA — used for KPIs, table, trend, donut, sp, item
    dff = apply_filters(yr, months, cats, sps, cos)

    yr_lbl = "All Years" if (not yr or yr == "ALL") else str(yr)
    mon_lbl = ("All Months" if not months
               else " | ".join(MON[m-1] for m in sorted(months)))

    rev   = float(dff["amount"].sum())
    ords  = len(dff)
    avg   = float(dff["amount"].mean()) if ords else 0.0
    units = int(dff["qty"].sum())

    ks_rev = f"across {dff['company'].nunique()} companies"
    ks_avg = "per order"
    ks_qty = f"across {dff['item'].nunique()} items"

    if dff.empty:
        return (f"{yr_lbl} — {mon_lbl}", f"{ords} orders",
                inr(rev),ks_rev, inr(avg),ks_avg, str(units),ks_qty,
                efig("Koi data nahi"),efig("Koi data nahi"),efig("Koi data nahi"),
                go.Figure(),[],efig("Koi data nahi"),efig("Koi data nahi"),
                [],"-", d_sty, m_sty)

    # ✅ FIX: Month-wise Revenue chart — uses YEAR filter only (no month filter)
    # so it always shows ALL 12 bars for the year, highlighting selected months
    # This was the ROOT CAUSE of the bug — earlier code used ct_base which
    # sometimes included month filter, making bars disappear
    year_only_df = apply_filters(yr, None, cats, sps, cos)  # year + cat/sp/co, NO month filter
    mg = year_only_df.groupby(["month","month_name"])["amount"].sum().reset_index().sort_values("month")

    # Highlight selected months (blue = selected, dim = not selected)
    if months:
        bar_cols = ["#4F8EF7" if m in months else "#1E3A8A"
                    for m in mg["month"].tolist()]
        bar_opac = [1.0 if m in months else 0.35
                    for m in mg["month"].tolist()]
    else:
        bar_cols = ["#4F8EF7"] * len(mg)
        bar_opac = [1.0] * len(mg)

    ymax = float(mg["amount"].max()) if not mg.empty else 1
    fmon = go.Figure(go.Bar(
        x=mg["month_name"].tolist(), y=mg["amount"].tolist(),
        marker=dict(
            color=bar_cols,
            opacity=bar_opac,
            line=dict(width=0),
        ),
        text=[inr(v) if v > 0 else "" for v in mg["amount"].tolist()],
        textposition="inside", insidetextanchor="end",
        textfont=dict(size=9, color="#fff"),
        hovertemplate="%{x}: ₹%{y:,.0f}<extra></extra>"))
    fmon.update_layout(**CL, showlegend=False,
        xaxis=dict(showgrid=False,tickfont=dict(size=9),color=T2),
        yaxis=dict(showgrid=True,gridcolor=BD,zeroline=False,color=T2,visible=False,
                   range=[0, ymax*1.22]))

    # ✅ FIX: Category Trends — uses FILTERED data (same as KPIs)
    # So when you select Jan+Feb, trend lines show only Jan+Feb data
    cg = dff.groupby(["month","month_name","category"])["amount"].sum().reset_index().sort_values("month")
    fcat = go.Figure()
    for c in sorted(cg["category"].unique()):
        sc  = cg[cg["category"] == c]
        col = CAT_COLORS.get(c,"#888")
        fcat.add_trace(go.Scatter(
            x=sc["month_name"].tolist(), y=sc["amount"].tolist(),
            name=str(c), mode="lines+markers",
            line=dict(color=col,width=2), marker=dict(size=5,color=col),
            hovertemplate=f"<b>{c}</b><br>%{{x}}: ₹%{{y:,.0f}}<extra></extra>",
            connectgaps=True))
    fcat.update_layout(**CL, showlegend=True,
        legend=dict(orientation="h",yanchor="bottom",y=1.01,xanchor="right",x=1,
                    font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False,tickfont=dict(size=8),color=T2),
        yaxis=dict(showgrid=True,gridcolor=BD,tickfont=dict(size=9),zeroline=False,color=T2))

    # Trend (daily / monthly) — uses filtered data
    ft = go.Figure()
    for c in sorted(dff["category"].unique()):
        sc  = dff[dff["category"] == c]
        col = CAT_COLORS.get(c,"#888")
        if monthly:
            g = sc.groupby(["year","month","month_name"])["amount"].sum().reset_index()
            g["lbl"] = g["month_name"] + " " + g["year"].astype(str)
            g = g.sort_values(["year","month"])
            x, y = g["lbl"].tolist(), g["amount"].tolist()
        else:
            sc = sc.copy(); sc["d"] = sc["date"].dt.date
            g = sc.groupby("d")["amount"].sum().reset_index().sort_values("d")
            x, y = g["d"].astype(str).tolist(), g["amount"].tolist()
        ft.add_trace(go.Scatter(
            x=x, y=y, name=str(c), mode="lines+markers",
            line=dict(color=col,width=2), marker=dict(size=5,color=col),
            hovertemplate=f"<b>{c}</b><br>%{{x}}: ₹%{{y:,.0f}}<extra></extra>",
            connectgaps=True))
    ft.update_layout(**CL, showlegend=True,
        legend=dict(orientation="h",yanchor="bottom",y=1.01,xanchor="right",x=1,
                    font=dict(size=9),bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False,tickfont=dict(size=8),linecolor=BD,color=T2,tickangle=-30),
        yaxis=dict(showgrid=True,gridcolor=BD,tickfont=dict(size=9),zeroline=False,color=T2))

    # Donut
    ca  = dff.groupby("category").agg(amount=("amount","sum"), qty=("qty","sum")).reset_index()
    tot = float(ca["amount"].sum()) if not ca.empty else 1.0
    fd  = go.Figure(go.Pie(
        labels=ca["category"].tolist(), values=ca["amount"].tolist(), hole=0.62,
        marker=dict(colors=[CAT_COLORS.get(str(c),"#888") for c in ca["category"]],
                    line=dict(color=PL,width=3)),
        textinfo="none",
        hovertemplate="%{label}: ₹%{value:,.0f} (%{percent})<extra></extra>"))
    fd.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                     font=dict(family=FT,size=11,color=T2),
                     margin=dict(l=2,r=2,t=2,b=2), showlegend=False)
    leg = []
    for _, row in ca.iterrows():
        c = CAT_COLORS.get(str(row["category"]),"#888")
        p = float(row["amount"])/tot*100 if tot else 0
        leg.append(html.Div(style={
            "display":"flex","flexDirection":"column","gap":"3px",
            "background":c+"18","borderRadius":"7px","padding":"5px 9px",
            "border":f"1px solid {c}44","minWidth":"95px"}, children=[
            html.Div(style={"display":"flex","alignItems":"center","gap":"4px"}, children=[
                html.Div(style={"width":"7px","height":"7px","borderRadius":"2px","background":c}),
                html.Span(str(row["category"]),style={"fontSize":"9px","color":T1,"fontWeight":"700"}),
                html.Span(f"{round(p)}%",style={"fontSize":"9px","color":c,"fontWeight":"700","marginLeft":"auto"}),
            ]),
            html.Div(style={"display":"flex","justifyContent":"space-between"}, children=[
                html.Span("Qty:", style={"fontSize":"8px","color":T2}),
                html.Span(str(int(row["qty"])),style={"fontSize":"9px","color":T1,"fontWeight":"600"}),
            ]),
        ]))

    # Salesperson bar
    sp  = dff.groupby("sp_clean")["amount"].sum().nlargest(8).reset_index()
    spc = ["#4F8EF7","#378ADD","#85B7EB","#B5D4F4","#1D9E75","#5DCAA5","#9FE1CB","#FAC775"]
    if not sp.empty:
        fsp = go.Figure(go.Bar(
            x=sp["amount"].tolist(), y=sp["sp_clean"].tolist(), orientation="h",
            marker=dict(color=spc[:len(sp)],opacity=0.9),
            text=[inr(v) for v in sp["amount"].tolist()],
            textposition="outside", textfont=dict(size=10,color=T1),
            hovertemplate="%{y}: ₹%{x:,.0f}<extra></extra>"))
        fsp.update_layout(**CL, showlegend=False,
            xaxis=dict(showgrid=True,gridcolor=BD,zeroline=False,color=T2,visible=False),
            yaxis=dict(showgrid=False,tickfont=dict(size=11),autorange="reversed",color=T1))
    else:
        fsp = efig("Koi data nahi")

    # Item bar
    it = (dff.groupby(["item","category"])["amount"].sum().reset_index()
            .sort_values("amount",ascending=False).head(8))
    if not it.empty:
        fit = go.Figure(go.Bar(
            x=it["amount"].tolist(), y=it["item"].tolist(), orientation="h",
            marker=dict(color=[CAT_COLORS.get(str(c),"#888") for c in it["category"]],opacity=0.9),
            text=[inr(v) for v in it["amount"].tolist()],
            textposition="outside", textfont=dict(size=9,color=T2),
            hovertemplate="%{y}: ₹%{x:,.0f}<extra></extra>"))
        fit.update_layout(**CL, showlegend=False,
            xaxis=dict(showgrid=True,gridcolor=BD,zeroline=False,color=T2,visible=False),
            yaxis=dict(showgrid=False,tickfont=dict(size=9),autorange="reversed",color=T1))
    else:
        fit = efig("Koi data nahi")

    # Table
    tbl = dff.sort_values("date",ascending=False).copy()
    tbl["amt_num"] = tbl["amount"].astype(int)
    tbl["qty"]     = tbl["qty"].astype(int)
    tbl_data = tbl[["day_str","company","item","category","qty","amt_num","sp_clean",
                    "office_sp_clean","end_user_oem","existing_new"]].to_dict("records")

    return (
        f"{yr_lbl} — {mon_lbl}", f"{ords} orders",
        inr(rev), ks_rev, inr(avg), ks_avg, str(units), ks_qty,
        fmon, fcat, ft, fd, leg, fsp, fit,
        tbl_data, f"📋 {len(tbl)} orders",
        d_sty, m_sty,
    )


# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "="*55)
    print("  Vaayushanti Sales Dashboard — Google Sheets Live")
    print("  URL    : http://localhost:8050")
    print(f"  Source : {SOURCE_LABEL}")
    print("  Auto-refresh: Every 30 seconds")
    print("="*55 + "\n")
    app.run(debug=False)
