import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
import glob
import os
import re

# ═══════════════════════════════════════════════════════════
# 설정
# ═══════════════════════════════════════════════════════════
DATA_DIR          = Path(__file__).parent / "data"
BRAND_TARGET_FILE = DATA_DIR / "brand_target.xlsx"

BRAND_KEYWORDS = {
    "Betadine":   ["베타딘"],
    "Lamisil":    ["라미실"],
    "Nicotinell": ["니코틴엘"],
    "Cyclogest":  ["사이클로제스트"],
    "Cialis":     ["시알리스"],
    "Physiogel":  ["피지오겔"],
    "LG":         ["피지오겔"],
    "Handok":     ["케토톱", "훼스탈", "클리어틴", "로푸록스", "캄비손", "페스내추럴",
                   "하이비날", "싸이타민", "하이렉스", "이치레스큐", "클리어티앤"],
    "RB":         ["스트렙실", "듀렉스", "개비스콘", "데톨"],
    "Haleon":     ["센트룸", "센티렉스", "칼트레이트", "애드빌"],
}

MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun",
               "Jul","Aug","Sep","Oct","Nov","Dec"]

COL_DATE     = "Billing Date"
COL_PRODUCT  = "Material Description"
COL_INV_QTY  = "Invoice Quantity"
COL_VALUE    = "Total Sales Value"
COL_CUSTOMER = "Sold To Customer Code"
COL_TEAM     = "Team Name"

WS_TEAM_KW   = "key account management team"  # 이 팀 = WS, 나머지 = PHA

INLICENSE_BRANDS = ["Lamisil", "Betadine", "Nicotinell", "Cyclogest", "Cialis", "LG"]

# (business_label, [(brand, split_ws_pha), ...])
BUSINESS_GROUPS = [
    ("In-license OTC & ETC", [
        ("Lamisil",    True),
        ("Betadine",   True),
        ("Nicotinell", True),
        ("Cyclogest",  False),
        ("Cialis",     False),
    ]),
    ("Agency",       [("LG",     False)]),
    ("CSO Business", [("Handok", False), ("RB", False), ("Haleon", False)]),
]

SKU_NORMALIZE: dict[str, list[tuple[str, list[str]]]] = {
    "Lamisil": [
        ("Lamisil Once 4g",        ["원스"]),
        ("Lamisil DermGel 1% 15g", ["덤겔"]),
        ("Lamisil Cream 1% 15g",   ["크림", "15"]),
        ("Lamisil Cream 1% 30g",   ["크림", "30"]),
        ("Lamisil Spray 1% 30ml",    ["외용액"]),
    ],
    "Betadine": [
        ("Betadine Throat Spray 0.45% 25ml",          ["인후스프레이", "25"]),
        ("Betadine Throat Spray 0.45% 50ml",          ["인후스프레이", "50"]),
        ("Dry Powder Spray 2.5% 5g",                  ["드라이파우더"]),
        ("Gynobetadine Pessaries 10T",                ["질좌제", "10t"]),
        ("Gynobetadine Pessaries 50T",                ["질좌제", "50t"]),
        ("Gynobetadine V.Douche 180ML",               ["질세정", "180"]),
        ("Gynobetadine V.Douche 360Ml",               ["질세정", "360"]),
        ("Vetadine Antis. Sol 10% 1G/L",              ["농후액"]),
        ("Betadine Antiseptic Solution 10% 1G/L",     ["소독액", "gl"]),
        ("Betadine Antiseptic Solution 10% 1L(2014)", ["소독액", "2014"]),
        ("Betadine Antiseptic Solution 10% 1L",       ["소독액"]),
        ("Betadine Skincleanser7.5% 1G/L",            ["세정액", "gl"]),
        ("Betadine Skincleanser7.5% 1L",              ["세정액"]),
    ],
    "Nicotinell": [
        ("Nicotinell TTS10 7P",      ["tts10"]),
        ("Nicotinell TTS20 7P",      ["tts20"]),
        ("Nicotinell TTS30 7P",      ["tts30"]),
        ("Nicotinell Gum 2mg 24P",   ["껌"]),
        ("Nicotinell Lozenge 1mg 36P", ["로젠즈"]),
    ],
    "Cialis": [
        ("Cialis 10mg 4T",  ["10mg"]),
        ("Cialis 20mg 4T",  ["20mg"]),
        ("Cialis 5mg 14T",  ["5mg", "14"]),
        ("Cialis 5mg 28T",  ["5mg", "28"]),
    ],
    "Cyclogest": [
        ("Cyclogest 200mg", ["200"]),
        ("Cyclogest 400mg", ["400"]),
    ],
}

# SKU-level whitelist: only rows whose product name contains one of these substrings
# (case-insensitive, literal match — not regex)
BRAND_SKU_FILTER: dict[str, list[str]] = {
    "Cialis": ["시알리스정(ptp)"],
}

# Team-level whitelist: only rows whose Team Name is in this list are counted.
# (exact match after lowercase + strip)
BRAND_TEAM_FILTER: dict[str, list[str]] = {
    "Lamisil": [
        "zpt p team",
        "zpt o team",
        "zpt t team",
        "zpt z team",
        "key account management team",
    ],
    "Handok": [
        "zpt p team",
        "zpt o team",
        "zpt t team",
        "zpt z team",
    ],
    "RB": [
        "zpt p team",
        "zpt o team",
        "zpt t team",
        "zpt z team",
        "rb odt team",
        "zpt chc-r team",
    ],
    "Haleon": [
        "zpt p team",
        "zpt o team",
        "zpt t team",
        "zpt z team",
        "zanovex chc-g team",
        "zpt chc-h team",
    ],
}
# ═══════════════════════════════════════════════════════════

st.set_page_config(page_title="DSR Dashboard", page_icon="📊", layout="wide")
st.title("📊 DSR Sales Dashboard")


CACHE_PARQUET = DATA_DIR / "_actuals_cache.parquet"


@st.cache_data(ttl=300)
def load_actuals():
    exclude = {
        os.path.abspath(BRAND_TARGET_FILE),
        os.path.abspath(DATA_DIR / "budget.xlsx"),
    }
    xlsx_files = [
        f for f in glob.glob(str(DATA_DIR / "*.xlsx"))
        if os.path.abspath(f) not in exclude
        and not Path(f).name.startswith("~$")
    ]
    if not xlsx_files:
        return None, []

    # xlsx 중 가장 최근 수정 시각
    xlsx_mtime = max(os.path.getmtime(f) for f in xlsx_files)

    # parquet 캐시가 있고 xlsx보다 최신이면 빠르게 로드
    if CACHE_PARQUET.exists() and os.path.getmtime(CACHE_PARQUET) >= xlsx_mtime:
        try:
            return pd.read_parquet(CACHE_PARQUET), xlsx_files
        except Exception:
            pass  # 캐시 깨진 경우 xlsx에서 재로드

    # xlsx 읽기 (최초 또는 파일 변경 시)
    frames = []
    for f in xlsx_files:
        try:
            frames.append(pd.read_excel(f))
        except Exception:
            pass
    if not frames:
        return None, []

    df = pd.concat(frames, ignore_index=True).drop_duplicates()
    df[COL_DATE] = pd.to_datetime(
        df[COL_DATE].astype(str).str[:8], format="%Y%m%d", errors="coerce"
    )

    # Convert all non-numeric columns to proper string type for parquet/pandas 3.x compat.
    # Some object columns contain mixed float+string values from xlsx concat; we normalize
    # them to StringDtype (None → pd.NA, everything else → str(x)).
    for col in df.columns:
        if df[col].dtype == object or pd.api.types.is_string_dtype(df[col]):
            if df[col].dtype == bool:
                continue
            df[col] = (
                df[col]
                .apply(lambda x: None if pd.isna(x) else str(x))
                .astype("string")   # force StringDtype so pyarrow can serialise cleanly
            )

    # 다음 실행을 위해 parquet으로 저장
    try:
        df.to_parquet(CACHE_PARQUET, index=False)
    except Exception:
        pass

    return df, xlsx_files


@st.cache_data(ttl=3600)
def load_brand_targets() -> pd.DataFrame:
    if not BRAND_TARGET_FILE.exists():
        return pd.DataFrame()
    df = pd.read_excel(str(BRAND_TARGET_FILE))
    df.columns = [str(c).strip() for c in df.columns]
    return df


def get_ytd_target(targets_df: pd.DataFrame, brand: str, up_to_month_idx: int) -> float:
    if targets_df.empty or brand not in targets_df["Brand"].values:
        return 0.0
    rows = targets_df[targets_df["Brand"] == brand]
    return sum(
        float(row.get(MONTH_NAMES[i], 0) or 0)
        for _, row in rows.iterrows()
        for i in range(up_to_month_idx + 1)
    )


def get_mtd_target(targets_df: pd.DataFrame, brand: str, month_idx: int) -> float:
    if targets_df.empty or brand not in targets_df["Brand"].values:
        return 0.0
    rows = targets_df[targets_df["Brand"] == brand]
    return sum(float(row.get(MONTH_NAMES[month_idx], 0) or 0) for _, row in rows.iterrows())


def filter_by_brand(df: pd.DataFrame, brand: str) -> pd.DataFrame:
    keywords = BRAND_KEYWORDS.get(brand, [])
    if not keywords:
        return df.iloc[0:0]
    pattern = "|".join(k.lower() for k in keywords)
    mask = df[COL_PRODUCT].fillna("").str.lower().str.contains(pattern, na=False)
    result = df[mask]
    sku_filters = BRAND_SKU_FILTER.get(brand, [])
    if sku_filters:
        # Use str.contains (always returns bool dtype, even on empty Series).
        # apply(lambda) on an empty StringDtype Series returns object dtype in
        # pandas 3.x, which pandas then treats as column labels — dropping all columns.
        pattern_sku = "|".join(re.escape(s.lower()) for s in sku_filters)
        sku_mask = result[COL_PRODUCT].fillna("").str.lower().str.contains(
            pattern_sku, na=False, regex=True
        )
        result = result[sku_mask]
    team_filters = BRAND_TEAM_FILTER.get(brand, [])
    if team_filters and COL_TEAM in result.columns:
        allowed_teams = [t.lower().strip() for t in team_filters]
        team_mask = result[COL_TEAM].fillna("").str.lower().str.strip().isin(allowed_teams)
        result = result[team_mask]
    return result


def _strip_product_name(name: str) -> str:
    s = str(name).strip()
    s = re.sub(r"^\([A-Za-z]\)\s*", "", s)
    s = re.sub(r"^\*+\s*", "", s)
    s = re.sub(r"^[가-힣]+[-]\s*", "", s)
    s = re.sub(r"\s+\d+EA\s*$", "", s, re.I)
    return re.sub(r"\s+", " ", s).strip()


def normalize_sku(name: str, brand: str) -> str:
    key = re.sub(r"[^\w가-힣%]", "", str(name).lower())
    if brand in SKU_NORMALIZE:
        for canonical, kws in SKU_NORMALIZE[brand]:
            kws_clean = [re.sub(r"[^\w가-힣%]", "", k.lower()) for k in kws]
            if all(k in key for k in kws_clean):
                return canonical
    return _strip_product_name(name)


def fmt(v):
    return f"{v/1_000_000:,.0f}M" if pd.notna(v) and v != 0 else "—"


def fmt_pct(v):
    return f"{v:.1f}%" if pd.notna(v) and v is not None else "—"


def fmt_growth(v):
    return f"{v:+.1f}%" if pd.notna(v) and v is not None else "—"


# ─── 데이터 로드 ───────────────────────────────────────────
df, actuals_file = load_actuals()

if df is None:
    st.warning(f"No sales Excel files found in `{DATA_DIR}/`.")
    st.stop()

targets_df = load_brand_targets()
today      = pd.Timestamp.today()

# ─── 연도·월 선택기 (사이드바) ────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    available_years = sorted(
        df[COL_DATE].dropna().dt.year.astype(int).unique().tolist(),
        reverse=True,
    )
    if not available_years:
        available_years = [today.year]

    sel_year = st.selectbox(
        "Year",
        options=available_years,
        index=0,
        format_func=lambda y: str(y),
    )

    if sel_year == today.year:
        available_months = [m for m in range(1, 13) if m <= today.month]
    else:
        available_months = list(range(1, 13))

    sel_month = st.selectbox(
        "Month",
        options=available_months,
        index=len(available_months) - 1,
        format_func=lambda m: f"{MONTH_NAMES[m-1]} {sel_year}",
    )

sel_month_idx  = sel_month - 1
sel_month_name = MONTH_NAMES[sel_month_idx]
year           = sel_year
prev_year      = sel_year - 1

mtd_df    = df[(df[COL_DATE].dt.year == year) & (df[COL_DATE].dt.month == sel_month)]
ytd_df    = df[(df[COL_DATE].dt.year == year) & (df[COL_DATE].dt.month <= sel_month)]
py_ytd_df = df[(df[COL_DATE].dt.year == prev_year) & (df[COL_DATE].dt.month <= sel_month)]
latest_date = df[COL_DATE].dropna().dt.date.max()
today_df    = df[df[COL_DATE].dt.date == latest_date]

st.caption(
    f"{len(actuals_file)} file(s) loaded | "
    f"As of {sel_month_name} {year} | "
    f"MTD {len(mtd_df):,} / YTD {len(ytd_df):,} records"
)

all_brands = targets_df["Brand"].tolist() if not targets_df.empty else []


def calc_kpis(brand_list, mtd_data, ytd_data, py_ytd_data):
    mtd_act    = sum(filter_by_brand(mtd_data, b)[COL_VALUE].sum()
                     for b in brand_list if BRAND_KEYWORDS.get(b))
    ytd_act    = sum(filter_by_brand(ytd_data, b)[COL_VALUE].sum()
                     for b in brand_list if BRAND_KEYWORDS.get(b))
    py_ytd_act = sum(filter_by_brand(py_ytd_data, b)[COL_VALUE].sum()
                     for b in brand_list if BRAND_KEYWORDS.get(b))
    mtd_tgt = sum(get_mtd_target(targets_df, b, sel_month_idx) for b in brand_list)
    ytd_tgt = sum(get_ytd_target(targets_df, b, sel_month_idx) for b in brand_list)
    mtd_ach = (mtd_act / mtd_tgt * 100) if mtd_tgt > 0 else None
    ytd_ach = (ytd_act / ytd_tgt * 100) if ytd_tgt > 0 else None
    growth  = ((ytd_act - py_ytd_act) / py_ytd_act * 100) if py_ytd_act > 0 else None
    return mtd_act, mtd_tgt, mtd_ach, ytd_act, ytd_tgt, ytd_ach, py_ytd_act, growth




# ─── 브랜드 분석 탭 공통 렌더링 ────────────────────────────
def _style_total(s):
    name = s.name
    is_total = (
        name[1] == "Total" if isinstance(name, tuple) else "Total" in str(name)
    )
    return (
        ["background-color: #DAE8FC; font-weight: bold"] * len(s)
        if is_total else [""] * len(s)
    )


def _build_sku_pivot(bdf, ws_df, pha_df, value_col, month_range, show_pha=True):
    all_skus = bdf["_sku"].unique().tolist()
    sku_ytd  = {s: bdf[bdf["_sku"] == s][value_col].sum() for s in all_skus}
    all_skus = sorted(all_skus, key=lambda s: -sku_ytd[s])

    channels = [(ws_df, "WS"), (pha_df, "PHA"), (bdf, "Total")]
    if not show_pha:
        channels = [(ws_df, "WS"), (bdf, "Total")]

    index_tuples, data = [], []
    for sku in all_skus:
        for ch_df, ch_label in channels:
            sub = ch_df[ch_df["_sku"] == sku]
            row = {MONTH_NAMES[m - 1]: sub[sub["_month"] == m][value_col].sum()
                   for m in month_range}
            row["YTD Total"] = sub[value_col].sum()
            index_tuples.append((sku, ch_label))
            data.append(row)

    idx = pd.MultiIndex.from_tuples(index_tuples, names=["SKU", "Channel"])
    return pd.DataFrame(data, index=idx)


def render_brand_tab(brand_name: str, ytd_data: pd.DataFrame, py_ytd_data: pd.DataFrame,
                     c_actual="#2980B9", c_target="#AEB6BF", c_line="#E67E22",
                     show_pha: bool = True, show_customer_count: bool = True):
    bdf    = filter_by_brand(ytd_data, brand_name).copy()
    py_bdf = filter_by_brand(py_ytd_data, brand_name).copy()

    if bdf.empty:
        st.info(f"📭 No {brand_name} data through {sel_month_name} {year}.")
        return

    bdf["_month"] = bdf[COL_DATE].dt.month
    bdf["_sku"]   = bdf[COL_PRODUCT].apply(lambda x: normalize_sku(x, brand_name))
    month_range   = list(range(1, sel_month + 1))
    has_cust      = COL_CUSTOMER in bdf.columns
    has_py        = not py_bdf.empty

    if has_py:
        py_bdf["_month"] = py_bdf[COL_DATE].dt.month

    # WS / PHA 분리
    if COL_TEAM in bdf.columns:
        ws_mask = bdf[COL_TEAM].fillna("").str.lower().str.strip() == WS_TEAM_KW
        ws_df   = bdf[ws_mask]
        pha_df  = bdf[~ws_mask]
    else:
        ws_df  = bdf.iloc[0:0]
        pha_df = bdf

    # ── 월별 Actual vs Target 콤보 차트 ──────────────────────
    month_labels = [MONTH_NAMES[m - 1] for m in month_range]
    monthly_act  = [bdf[bdf["_month"] == m][COL_VALUE].sum() for m in month_range]

    if not targets_df.empty and brand_name in targets_df["Brand"].values:
        _tgt_rows = targets_df[targets_df["Brand"] == brand_name]
        monthly_tgt = [
            sum(float(r.get(MONTH_NAMES[m - 1], 0) or 0) for _, r in _tgt_rows.iterrows())
            for m in month_range
        ]
    else:
        monthly_tgt = [0] * len(month_range)

    monthly_ar = [(a / t * 100) if t > 0 else None for a, t in zip(monthly_act, monthly_tgt)]

    fig_combo = make_subplots(specs=[[{"secondary_y": True}]])
    fig_combo.add_trace(go.Bar(
        x=month_labels, y=monthly_tgt,
        name="Target", marker_color=c_target, opacity=0.7,
    ), secondary_y=False)
    fig_combo.add_trace(go.Bar(
        x=month_labels, y=monthly_act,
        name="Actual", marker_color=c_actual, opacity=0.92,
        text=[f"{v/1e6:.1f}M" for v in monthly_act],
        textposition="outside", textfont=dict(size=10, color=c_actual),
    ), secondary_y=False)
    fig_combo.add_trace(go.Scatter(
        x=month_labels, y=monthly_ar,
        name="A/R %", mode="lines+markers+text",
        line=dict(color=c_line, width=2.5),
        marker=dict(size=8, color=c_line),
        text=[f"{v:.1f}%" if v is not None else "" for v in monthly_ar],
        textposition="top center",
        textfont=dict(size=11, color=c_line),
    ), secondary_y=True)
    fig_combo.update_layout(
        barmode="group", height=420,
        margin=dict(l=0, r=10, t=45, b=0),
        legend=dict(orientation="h", y=1.1, x=0),
        plot_bgcolor="#FAFAFA", paper_bgcolor="white",
    )
    fig_combo.update_yaxes(title_text="Sales Value (KRW)", secondary_y=False, tickformat=",",
                           gridcolor="#EFEFEF")
    fig_combo.update_yaxes(title_text="A/R %", secondary_y=True, ticksuffix="%")
    st.plotly_chart(fig_combo, use_container_width=True)
    st.markdown("---")

    # ── Monthly Summary (WS / PHA / Total) ──────────────────────
    st.markdown("#### Monthly Summary")
    summary = {}

    _channels = [(ws_df, "WS"), (pha_df, "PHA"), (bdf, "Total")] if show_pha else [(ws_df, "WS"), (bdf, "Total")]
    for metric_label, val_col in [("Qty", COL_INV_QTY), ("Value", COL_VALUE)]:
        for ch_df, ch_label in _channels:
            row = {}
            for m in month_range:
                mdf = ch_df[ch_df["_month"] == m]
                col = MONTH_NAMES[m - 1]
                row[col] = (
                    f"{int(mdf[val_col].sum()):,}"
                    if val_col == COL_INV_QTY
                    else f"{mdf[val_col].sum():,.0f}"
                )
            row["YTD Total"] = (
                f"{int(ch_df[val_col].sum()):,}"
                if val_col == COL_INV_QTY
                else f"{ch_df[val_col].sum():,.0f}"
            )
            summary[f"{metric_label} ({ch_label})"] = row

    if has_py:
        py_val_row, yoy_row = {}, {}
        for m in month_range:
            col     = MONTH_NAMES[m - 1]
            cur_val = bdf[bdf["_month"] == m][COL_VALUE].sum()
            py_val  = py_bdf[py_bdf["_month"] == m][COL_VALUE].sum()
            py_val_row[col] = f"{py_val:,.0f}"
            yoy_row[col]    = f"{(cur_val - py_val) / py_val * 100:+.1f}%" if py_val > 0 else "—"
        py_ytd  = py_bdf[COL_VALUE].sum()
        cur_ytd = bdf[COL_VALUE].sum()
        py_val_row["YTD Total"] = f"{py_ytd:,.0f}"
        yoy_row["YTD Total"]    = (
            f"{(cur_ytd - py_ytd) / py_ytd * 100:+.1f}%" if py_ytd > 0 else "—"
        )
        summary[f"PY Value ({prev_year})"] = py_val_row
        summary["YoY GR%"]                 = yoy_row

    st.dataframe(
        pd.DataFrame(summary).T.style.apply(_style_total, axis=1),
        use_container_width=True,
    )
    st.markdown("---")

    st.markdown("#### SKU Qty (Invoice)")
    qty_piv  = _build_sku_pivot(bdf, ws_df, pha_df, COL_INV_QTY, month_range, show_pha)
    qty_disp = qty_piv.copy()
    for col in qty_disp.columns:
        qty_disp[col] = qty_disp[col].apply(lambda x: f"{int(x):,}")
    st.dataframe(
        qty_disp.style.apply(_style_total, axis=1),
        use_container_width=True,
    )
    st.markdown("---")

    st.markdown("#### SKU Value (KRW)")
    val_piv  = _build_sku_pivot(bdf, ws_df, pha_df, COL_VALUE, month_range, show_pha)
    val_disp = val_piv.copy()
    for col in val_disp.columns:
        val_disp[col] = val_disp[col].apply(lambda x: f"{x:,.0f}")
    st.dataframe(
        val_disp.style.apply(_style_total, axis=1),
        use_container_width=True,
    )

    if has_cust and show_customer_count:
        st.markdown("---")
        st.markdown("#### Monthly Customer Count")
        month_labels = [MONTH_NAMES[m - 1] for m in month_range]
        ws_counts    = [ws_df[ws_df["_month"] == m][COL_CUSTOMER].nunique()  for m in month_range]
        pha_counts   = [pha_df[pha_df["_month"] == m][COL_CUSTOMER].nunique() for m in month_range]

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=month_labels, y=ws_counts,
            name="WS", marker_color="#636EFA",
            text=ws_counts, textposition="outside",
        ))
        fig.add_trace(go.Bar(
            x=month_labels, y=pha_counts,
            name="PHA", marker_color="#EF553B",
            text=pha_counts, textposition="outside",
        ))
        fig.update_layout(
            barmode="group",
            height=350,
            margin=dict(l=0, r=0, t=30, b=0),
            yaxis_title="No. of Customers",
            legend=dict(orientation="h", y=1.08),
        )
        st.plotly_chart(fig, use_container_width=True)

        ytd_ws  = ws_df[COL_CUSTOMER].nunique()
        ytd_pha = pha_df[COL_CUSTOMER].nunique()
        ytd_tot = bdf[COL_CUSTOMER].nunique()
        c1, c2, c3 = st.columns(3)
        c1.metric("YTD Customers (WS)",    f"{ytd_ws:,}")
        c2.metric("YTD Customers (PHA)",   f"{ytd_pha:,}")
        c3.metric("YTD Customers (Total)", f"{ytd_tot:,}")


# ─── Tabs ───────────────────────────────────────────────────
tab_ov, tab_team, tab_lms, tab_btd, tab_ntl, tab_cyc, tab_cis, tab_ws = st.tabs([
    "🏷️ Brand Achievement",
    "👥 Team Achievement",
    "💊 Lamisil",
    "🩺 Betadine",
    "🚭 Nicotinell",
    "🔵 Cyclogest",
    "💙 Cialis",
    "🏪 Wholesaler",
])

# ══════════════════════════════════════════════
with tab_ov:
    st.subheader(f"Brand Performance — {sel_month_name} {year}")
    st.caption(f"Daily Actual as of: {latest_date} (latest date in data)")

    py_col_name = f"PY YTD ({prev_year})"

    def _ws_pha(df):
        if df.empty or COL_TEAM not in df.columns:
            return df.iloc[0:0], df
        mask = df[COL_TEAM].fillna("").str.lower().str.strip() == WS_TEAM_KW
        return df[mask], df[~mask]

    def _tgt(brand, kind, ch=None):
        if targets_df.empty or brand not in targets_df["Brand"].values:
            return None
        brand_rows = targets_df[targets_df["Brand"] == brand]
        if ch and "Channel" in brand_rows.columns:
            ch_rows = brand_rows[brand_rows["Channel"].str.upper() == ch.upper()]
            if ch_rows.empty:
                return None
            row = ch_rows.iloc[0]
            v = (float(row.get(MONTH_NAMES[sel_month_idx], 0) or 0) if kind == "mtd"
                 else sum(float(row.get(MONTH_NAMES[i], 0) or 0) for i in range(sel_month_idx + 1)))
        else:
            if kind == "mtd":
                v = sum(float(r.get(MONTH_NAMES[sel_month_idx], 0) or 0) for _, r in brand_rows.iterrows())
            else:
                v = sum(float(r.get(MONTH_NAMES[i], 0) or 0)
                        for _, r in brand_rows.iterrows()
                        for i in range(sel_month_idx + 1))
        return v if v else None

    def _row(business, product, channel, mtd_d, ytd_d, py_d,
             brand=None, row_type="data", daily_d=None):
        ma = mtd_d[COL_VALUE].sum(); ya = ytd_d[COL_VALUE].sum(); pa = py_d[COL_VALUE].sum()
        da = daily_d[COL_VALUE].sum() if daily_d is not None else 0
        if brand:
            tgt_ch = channel if channel in ("PHA", "WS") else None
            mt = _tgt(brand, "mtd", tgt_ch)
            yt = _tgt(brand, "ytd", tgt_ch)
        else:
            mt = None; yt = None
        return {
            "Business": business, "Product": product, "Channel": channel,
            "Daily Actual": da,
            "MTD Actual": ma, "MTD Target": mt,
            "MTD A/R %":   (ma / mt * 100) if mt else None,
            "YTD Actual": ya, "YTD Target": yt,
            "YTD A/R %":   (ya / yt * 100) if yt else None,
            py_col_name: pa,
            "YoY GR%": ((ya - pa) / pa * 100) if pa > 0 else None,
            "_rt": row_type,
        }

    rows = []
    il_ytd_f, il_mtd_f, il_py_f, il_daily_f = [], [], [], []

    for g_idx, (business, brand_list) in enumerate(BUSINESS_GROUPS):
        first = True
        for brand, split_ch in brand_list:
            if not BRAND_KEYWORDS.get(brand):
                continue
            b_ytd   = filter_by_brand(ytd_df,    brand)
            b_mtd   = filter_by_brand(mtd_df,    brand)
            b_py    = filter_by_brand(py_ytd_df, brand)
            b_daily = filter_by_brand(today_df,  brand)

            if g_idx < 2:
                il_ytd_f.append(b_ytd);   il_mtd_f.append(b_mtd)
                il_py_f.append(b_py);     il_daily_f.append(b_daily)

            biz = business if first else ""; first = False

            if split_ch:
                ws_ytd,   pha_ytd   = _ws_pha(b_ytd)
                ws_mtd,   pha_mtd   = _ws_pha(b_mtd)
                ws_py,    pha_py    = _ws_pha(b_py)
                ws_daily, pha_daily = _ws_pha(b_daily)
                rows.append(_row(biz, brand, "PHA",   pha_mtd, pha_ytd, pha_py, brand, "data",       pha_daily))
                rows.append(_row("",  brand, "WS",    ws_mtd,  ws_ytd,  ws_py,  brand, "data",       ws_daily))
                rows.append(_row("",  "",    "Total", b_mtd,   b_ytd,   b_py,   brand, "brand_total", b_daily))
            else:
                rows.append(_row(biz, brand, "Total", b_mtd, b_ytd, b_py, brand, "brand_total", b_daily))

        if business == "Agency" and il_ytd_f:
            iy  = pd.concat(il_ytd_f,   ignore_index=True)
            im  = pd.concat(il_mtd_f,   ignore_index=True)
            ip  = pd.concat(il_py_f,    ignore_index=True)
            id_ = pd.concat(il_daily_f, ignore_index=True) if il_daily_f else pd.DataFrame(columns=[COL_VALUE])
            il_bs = [b for bg, brs in BUSINESS_GROUPS[:2] for b, _ in brs if BRAND_KEYWORDS.get(b)]
            imt = sum((_tgt(b, "mtd") or 0) for b in il_bs)
            iyt = sum((_tgt(b, "ytd") or 0) for b in il_bs)
            ima = im[COL_VALUE].sum(); iya = iy[COL_VALUE].sum(); ipa = ip[COL_VALUE].sum()
            ida = id_[COL_VALUE].sum() if COL_VALUE in id_.columns else 0
            rows.append({
                "Business": "In-license & Agency Total", "Product": "", "Channel": "",
                "Daily Actual": ida,
                "MTD Actual": ima, "MTD Target": imt or None,
                "MTD A/R %":   (ima / imt * 100) if imt else None,
                "YTD Actual": iya, "YTD Target": iyt or None,
                "YTD A/R %":   (iya / iyt * 100) if iyt else None,
                py_col_name: ipa,
                "YoY GR%": ((iya - ipa) / ipa * 100) if ipa > 0 else None,
                "_rt": "subtotal",
            })

    all_bl = [b for _, brs in BUSINESS_GROUPS for b, _ in brs if BRAND_KEYWORDS.get(b)]
    gma = sum(filter_by_brand(mtd_df,    b)[COL_VALUE].sum() for b in all_bl)
    gya = sum(filter_by_brand(ytd_df,    b)[COL_VALUE].sum() for b in all_bl)
    gpa = sum(filter_by_brand(py_ytd_df, b)[COL_VALUE].sum() for b in all_bl)
    gda = sum(filter_by_brand(today_df,  b)[COL_VALUE].sum() for b in all_bl)
    gmt = sum((_tgt(b, "mtd") or 0) for b in all_bl)
    gyt = sum((_tgt(b, "ytd") or 0) for b in all_bl)
    rows.append({
        "Business": "TOTAL", "Product": "", "Channel": "",
        "Daily Actual": gda,
        "MTD Actual": gma, "MTD Target": gmt or None,
        "MTD A/R %":   (gma / gmt * 100) if gmt else None,
        "YTD Actual": gya, "YTD Target": gyt or None,
        "YTD A/R %":   (gya / gyt * 100) if gyt else None,
        py_col_name: gpa,
        "YoY GR%": ((gya - gpa) / gpa * 100) if gpa > 0 else None,
        "_rt": "grand_total",
    })

    ov_df   = pd.DataFrame(rows)
    rt_list = ov_df["_rt"].tolist()

    cols_show = ["Business", "Product", "Channel",
                 "Daily Actual",
                 "MTD Actual", "MTD Target", "MTD A/R %",
                 "YTD Actual", "YTD Target", "YTD A/R %",
                 py_col_name, "YoY GR%"]
    disp = ov_df[cols_show].copy()

    for c in ["Daily Actual", "MTD Actual", "MTD Target", "YTD Actual", "YTD Target", py_col_name]:
        disp[c] = ov_df[c].apply(lambda x: f"{x:,.0f}" if pd.notna(x) and x is not None else "")
    disp["MTD A/R %"] = ov_df["MTD A/R %"].apply(lambda x: f"{x:.1f}%"  if pd.notna(x) and x is not None else "")
    disp["YTD A/R %"] = ov_df["YTD A/R %"].apply(lambda x: f"{x:.1f}%"  if pd.notna(x) and x is not None else "")
    disp["YoY GR%"]   = ov_df["YoY GR%"].apply(  lambda x: f"{x:+.1f}%" if pd.notna(x) and x is not None else "")

    # ── HTML table: Business 열 rowspan 적용 ─────────────────
    def _biz_rowspan(raw_rows):
        result = {}; i = 0
        while i < len(raw_rows):
            biz = raw_rows[i].get("Business", "")
            if biz:
                span = 1; j = i + 1
                while j < len(raw_rows) and not raw_rows[j].get("Business", ""):
                    span += 1; j += 1
                result[i] = span
                for k in range(i + 1, i + span):
                    result[k] = 0
                i += span
            else:
                result[i] = 1; i += 1
        return result

    def _prod_rowspan(raw_rows):
        result = {}; i = 0
        while i < len(raw_rows):
            prod = raw_rows[i].get("Product", "")
            rt   = raw_rows[i].get("_rt", "data")
            if prod and rt not in ("subtotal", "grand_total"):
                span = 1; j = i + 1
                while j < len(raw_rows):
                    np  = raw_rows[j].get("Product", "")
                    nrt = raw_rows[j].get("_rt", "data")
                    if nrt in ("subtotal", "grand_total"):
                        break
                    if np == prod or np == "":
                        span += 1; j += 1
                    else:
                        break
                result[i] = span
                for k in range(i + 1, i + span):
                    result[k] = 0
                i += span
            else:
                result[i] = 1; i += 1
        return result

    biz_rs  = _biz_rowspan(rows)
    prod_rs = _prod_rowspan(rows)
    records = disp.to_dict("records")

    def _rt_css(rt):
        if rt == "grand_total": return "background-color:#1A252F;color:white;font-weight:bold;"
        if rt == "subtotal":    return "background-color:#4A4A4A;color:white;font-weight:bold;"
        if rt == "brand_total": return "background-color:#D6EAF8;font-weight:bold;"
        return ""

    tbl  = '<div style="overflow-x:auto;">'
    tbl += '<table style="width:100%;border-collapse:collapse;font-size:12px;font-family:sans-serif;">'
    tbl += '<thead><tr style="background-color:#2C3E50;color:white;">'
    for col in cols_show:
        tbl += f'<th style="padding:5px 8px;border:1px solid #555;white-space:nowrap;text-align:center;">{col}</th>'
    tbl += '</tr></thead><tbody>'

    for i, rec in enumerate(records):
        rt  = rt_list[i]
        rs  = _rt_css(rt)
        tbl += f'<tr style="{rs}">'
        for col in cols_show:
            val = rec.get(col, "")
            if col == "Business":
                span = biz_rs.get(i, 1)
                if span == 0:
                    continue
                biz_bg = "background-color:white;color:#1A1A1A;" if span > 1 else ""
                cell_s = f"text-align:center;padding:5px 8px;border:1px solid #ccc;vertical-align:middle;font-weight:bold;{biz_bg}"
                rs_attr = f' rowspan="{span}"' if span > 1 else ""
                tbl += f'<td{rs_attr} style="{cell_s}">{val}</td>'
            elif col in ("MTD A/R %", "YTD A/R %"):
                ar_s = ""
                if rt not in ("grand_total", "subtotal"):
                    try:
                        n = float(str(val).replace("%", ""))
                        if n >= 100:  ar_s = "background-color:#D5F5E3;font-weight:bold;"
                        elif n >= 80: ar_s = "background-color:#FDEBD0;font-weight:bold;"
                        else:         ar_s = "background-color:#FADBD8;font-weight:bold;"
                    except Exception:
                        pass
                tbl += f'<td style="text-align:right;padding:5px 8px;border:1px solid #ccc;{ar_s}">{val}</td>'
            elif col == "YoY GR%":
                gr_s = ""
                if rt not in ("grand_total", "subtotal"):
                    try:
                        n = float(str(val).replace("%", "").replace("+", ""))
                        if n < 0:   gr_s = "color:#C0392B;font-weight:bold;"
                        elif n > 0: gr_s = "color:#1E8449;font-weight:bold;"
                    except Exception:
                        pass
                tbl += f'<td style="text-align:right;padding:5px 8px;border:1px solid #ccc;{gr_s}">{val}</td>'
            elif col == "Product":
                span = prod_rs.get(i, 1)
                if span == 0:
                    continue
                prod_bg = "background-color:white;color:#1A1A1A;" if span > 1 else ""
                cell_s  = f"text-align:center;padding:5px 8px;border:1px solid #ccc;vertical-align:middle;font-weight:bold;{prod_bg}"
                rs_attr = f' rowspan="{span}"' if span > 1 else ""
                tbl += f'<td{rs_attr} style="{cell_s}">{val}</td>'
            elif col == "Channel":
                tbl += f'<td style="text-align:left;padding:5px 8px;border:1px solid #ccc;">{val}</td>'
            else:
                tbl += f'<td style="text-align:right;padding:5px 8px;border:1px solid #ccc;">{val}</td>'
        tbl += '</tr>'

    tbl += '</tbody></table></div>'
    st.html(tbl)

# ══════════════════════════════════════════════
with tab_team:
    st.subheader(f"Team Performance — {sel_month_name} {year}")
    st.caption(f"Daily Actual as of: {latest_date} (latest date in data)")

    if COL_TEAM not in ytd_df.columns:
        st.warning("No 'Team Name' column found in data.")
    else:
        IL_TM_BRANDS = ["Lamisil", "Betadine", "Nicotinell", "Cyclogest", "Cialis"]
        KAM_SHOW     = ["Lamisil", "Betadine", "Nicotinell"]

        TEAM_SRC = {
            "KEY ACCOUNT MANAGEMENT TEAM": ["key account management team"],
            "CHC-MM":      ["zpt z team", "zpt p team", "zpt t team", "zpt o team"],
            "CHC-H":       ["zanovex chc-g", "zpt chc-h team"],
            "CHC-R":       ["zpt chc-r team", "rb odt team"],
            "Clinic Team": ["zpt lg dermacare"],
        }
        EXCL_KW    = ["default", "한독"]
        EXCL_EXACT = {"hospital/clinic access", "zpt chc-s"}

        def _tdf_t(src, keys):
            if src.empty or COL_TEAM not in src.columns:
                return src.iloc[0:0]
            return src[src[COL_TEAM].fillna("").str.lower().str.strip().isin(set(keys))]

        def _il_split(src):
            if src.empty:
                return src.iloc[0:0], src.iloc[0:0]
            kws = [kw for b in IL_TM_BRANDS for kw in BRAND_KEYWORDS.get(b, [])]
            m = src[COL_PRODUCT].fillna("").apply(lambda p: any(kw in str(p) for kw in kws))
            return src[m], src[~m]

        def _vals(d, mt, yt, py):
            da = d[COL_VALUE].sum()  if not d.empty  else 0.0
            ma = mt[COL_VALUE].sum() if not mt.empty else 0.0
            ya = yt[COL_VALUE].sum() if not yt.empty else 0.0
            pa = py[COL_VALUE].sum() if not py.empty else 0.0
            gr = ((ya - pa) / pa * 100) if pa > 0 else None
            return da, ma, ya, pa, gr

        def _rd(team, product, rt, da, ma, ya, pa, gr, mt=None, yt=None):
            mtd_ar = (ma / mt * 100) if mt else None
            ytd_ar = (ya / yt * 100) if yt else None
            return {"_team": team, "_product": product, "_rt": rt,
                    "Daily Actual": da,
                    "MTD Actual": ma, "MTD Target": mt, "MTD A/R %": mtd_ar,
                    "YTD Actual": ya, "YTD Target": yt, "YTD A/R %": ytd_ar,
                    py_col_name: pa, "YoY GR%": gr}

        mapped_keys = {k for kl in TEAM_SRC.values() for k in kl}
        other_teams_list = []
        for rt_name in sorted(ytd_df[COL_TEAM].dropna().unique().tolist(), key=str):
            lk = str(rt_name).lower().strip()
            if any(ek in lk for ek in EXCL_KW):
                continue
            if lk in mapped_keys or lk in EXCL_EXACT:
                continue
            other_teams_list.append((str(rt_name), lk))

        t_rows = []

        for disp, keys in TEAM_SRC.items():
            d  = _tdf_t(today_df,  keys)
            m_ = _tdf_t(mtd_df,    keys)
            y  = _tdf_t(ytd_df,    keys)
            p  = _tdf_t(py_ytd_df, keys)

            if disp == "KEY ACCOUNT MANAGEMENT TEAM":
                for idx, brand in enumerate(KAM_SHOW):
                    bd, bm, by, bp = (filter_by_brand(x, brand) for x in (d, m_, y, p))
                    da, ma, ya, pa, gr = _vals(bd, bm, by, bp)
                    mt = _tgt(brand, "mtd", "WS")
                    yt = _tgt(brand, "ytd", "WS")
                    t_rows.append(_rd(disp if idx == 0 else "", brand, "data", da, ma, ya, pa, gr, mt, yt))
                s3 = t_rows[-3:]
                tot_da, tot_ma, tot_ya, tot_pa = (sum(r[c] for r in s3) for c in
                    ["Daily Actual", "MTD Actual", "YTD Actual", py_col_name])
                tot_gr = ((tot_ya - tot_pa) / tot_pa * 100) if tot_pa > 0 else None
                tot_mt = sum((_tgt(b, "mtd", "WS") or 0) for b in KAM_SHOW) or None
                tot_yt = sum((_tgt(b, "ytd", "WS") or 0) for b in KAM_SHOW) or None
                t_rows.append(_rd("", "Total", "team_total", tot_da, tot_ma, tot_ya, tot_pa, tot_gr, tot_mt, tot_yt))

            elif disp == "CHC-MM":
                d_il, d_oth  = _il_split(d)
                m_il, m_oth  = _il_split(m_)
                y_il, y_oth  = _il_split(y)
                p_il, p_oth  = _il_split(p)
                da1, ma1, ya1, pa1, gr1 = _vals(d_il,  m_il,  y_il,  p_il)
                da2, ma2, ya2, pa2, gr2 = _vals(d_oth, m_oth, y_oth, p_oth)
                da0, ma0, ya0, pa0, gr0 = _vals(d,     m_,    y,     p)
                il_pha = ["Lamisil", "Betadine", "Nicotinell"]
                il_tot = ["Cyclogest", "Cialis"]
                mt_il = (sum((_tgt(b, "mtd", "PHA") or 0) for b in il_pha) +
                         sum((_tgt(b, "mtd") or 0) for b in il_tot)) or None
                yt_il = (sum((_tgt(b, "ytd", "PHA") or 0) for b in il_pha) +
                         sum((_tgt(b, "ytd") or 0) for b in il_tot)) or None
                t_rows.append(_rd(disp, "In-license", "data",       da1, ma1, ya1, pa1, gr1, mt_il, yt_il))
                t_rows.append(_rd("",   "Others",     "data",       da2, ma2, ya2, pa2, gr2))
                t_rows.append(_rd("",   "Total",      "team_total", da0, ma0, ya0, pa0, gr0, mt_il, yt_il))

            elif disp == "Clinic Team":
                # Cyclogest: 팀 구분 없이 전체 데이터에서 (데이터상 KAM 태깅이지만 Clinic Team 소속)
                cyc = [filter_by_brand(x, "Cyclogest") for x in (today_df, mtd_df, ytd_df, py_ytd_df)]
                phy = [filter_by_brand(x, "Physiogel") for x in (d, m_, y, p)]
                da1, ma1, ya1, pa1, gr1 = _vals(*cyc)
                da2, ma2, ya2, pa2, gr2 = _vals(*phy)
                da0 = da1 + da2; ma0 = ma1 + ma2; ya0 = ya1 + ya2; pa0 = pa1 + pa2
                gr0 = ((ya0 - pa0) / pa0 * 100) if pa0 > 0 else None
                mt_cyc = _tgt("Cyclogest", "mtd")
                yt_cyc = _tgt("Cyclogest", "ytd")
                t_rows.append(_rd(disp, "Cyclogest", "data",       da1, ma1, ya1, pa1, gr1, mt_cyc, yt_cyc))
                t_rows.append(_rd("",   "Physiogel", "data",       da2, ma2, ya2, pa2, gr2))
                t_rows.append(_rd("",   "Total",     "team_total", da0, ma0, ya0, pa0, gr0, mt_cyc, yt_cyc))

            else:
                da, ma, ya, pa, gr = _vals(d, m_, y, p)
                t_rows.append(_rd(disp, "", "data", da, ma, ya, pa, gr))

        for raw_name, lk in other_teams_list:
            d  = _tdf_t(today_df,  [lk])
            m_ = _tdf_t(mtd_df,    [lk])
            y  = _tdf_t(ytd_df,    [lk])
            p  = _tdf_t(py_ytd_df, [lk])
            da, ma, ya, pa, gr = _vals(d, m_, y, p)
            t_rows.append(_rd(raw_name, "", "data", da, ma, ya, pa, gr))

        # Grand Total (actuals only — targets are partial so A/R % omitted)
        all_incl = list(mapped_keys) + [lk for _, lk in other_teams_list]
        gt_d, gt_m, gt_y, gt_p = (_tdf_t(x, all_incl) for x in (today_df, mtd_df, ytd_df, py_ytd_df))
        da_gt, ma_gt, ya_gt, pa_gt, gr_gt = _vals(gt_d, gt_m, gt_y, gt_p)
        t_rows.append(_rd("GRAND TOTAL", "", "grand_total", da_gt, ma_gt, ya_gt, pa_gt, gr_gt))

        # Team column rowspans
        team_rs = {}; ri = 0
        while ri < len(t_rows):
            tv = t_rows[ri]["_team"]
            if tv:
                span = 1; j = ri + 1
                while j < len(t_rows) and not t_rows[j]["_team"]:
                    span += 1; j += 1
                team_rs[ri] = span
                for k in range(ri + 1, ri + span):
                    team_rs[k] = 0
                ri += span
            else:
                team_rs[ri] = 1; ri += 1

        tcols = ["Team", "Product", "Daily Actual",
                 "MTD Actual", "MTD Target", "MTD A/R %",
                 "YTD Actual", "YTD Target", "YTD A/R %",
                 py_col_name, "YoY GR%"]

        ttbl  = '<div style="overflow-x:auto;">'
        ttbl += '<table style="width:100%;border-collapse:collapse;font-size:12px;font-family:sans-serif;">'
        ttbl += '<thead><tr style="background-color:#2C3E50;color:white;">'
        for col in tcols:
            ttbl += f'<th style="padding:5px 8px;border:1px solid #555;white-space:nowrap;text-align:center;">{col}</th>'
        ttbl += '</tr></thead><tbody>'

        for ri, r in enumerate(t_rows):
            rt_v = r["_rt"]
            if rt_v == "grand_total":
                row_s = "background-color:#1A252F;color:white;font-weight:bold;"
            elif rt_v == "team_total":
                row_s = "background-color:#D6EAF8;font-weight:bold;"
            else:
                row_s = ""
            ttbl += f'<tr style="{row_s}">'
            for col in tcols:
                if col == "Team":
                    span = team_rs.get(ri, 1)
                    if span == 0:
                        continue
                    cell_s = "text-align:center;padding:5px 8px;border:1px solid #ccc;vertical-align:middle;font-weight:bold;"
                    rs_attr = f' rowspan="{span}"' if span > 1 else ""
                    ttbl += f'<td{rs_attr} style="{cell_s}">{r["_team"]}</td>'
                elif col == "Product":
                    ttbl += f'<td style="text-align:left;padding:5px 8px;border:1px solid #ccc;">{r["_product"]}</td>'
                elif col in ("MTD A/R %", "YTD A/R %"):
                    raw_ar = r.get(col)
                    val = f"{raw_ar:.1f}%" if raw_ar is not None else ""
                    ar_s = ""
                    if rt_v != "grand_total" and raw_ar is not None:
                        if raw_ar >= 100:  ar_s = "background-color:#D5F5E3;font-weight:bold;"
                        elif raw_ar >= 80: ar_s = "background-color:#FDEBD0;font-weight:bold;"
                        else:              ar_s = "background-color:#FADBD8;font-weight:bold;"
                    ttbl += f'<td style="text-align:right;padding:5px 8px;border:1px solid #ccc;{ar_s}">{val}</td>'
                elif col == "YoY GR%":
                    raw_gr = r["YoY GR%"]
                    val = f"{raw_gr:+.1f}%" if raw_gr is not None else ""
                    gr_s = ""
                    if rt_v not in ("grand_total", "team_total") and raw_gr is not None:
                        if raw_gr < 0:   gr_s = "color:#C0392B;font-weight:bold;"
                        elif raw_gr > 0: gr_s = "color:#1E8449;font-weight:bold;"
                    ttbl += f'<td style="text-align:right;padding:5px 8px;border:1px solid #ccc;{gr_s}">{val}</td>'
                else:
                    raw_v = r.get(col)
                    val = f"{raw_v:,.0f}" if raw_v is not None else ""
                    ttbl += f'<td style="text-align:right;padding:5px 8px;border:1px solid #ccc;">{val}</td>'
            ttbl += '</tr>'

        ttbl += '</tbody></table></div>'
        st.html(ttbl)

# ══════════════════════════════════════════════
with tab_lms:
    st.subheader(f"Lamisil — {sel_month_name} {year}")
    render_brand_tab("Lamisil", ytd_df, py_ytd_df,
                     c_actual="#27AE60", c_target="#A9DFBF", c_line="#E74C3C")

with tab_btd:
    st.subheader(f"Betadine — {sel_month_name} {year}")
    render_brand_tab("Betadine", ytd_df, py_ytd_df,
                     c_actual="#CA6F1E", c_target="#F0B27A", c_line="#7D3C98")

with tab_ntl:
    st.subheader(f"Nicotinell — {sel_month_name} {year}")
    render_brand_tab("Nicotinell", ytd_df, py_ytd_df,
                     c_actual="#1F618D", c_target="#85C1E9", c_line="#E67E22")

with tab_cyc:
    st.subheader(f"Cyclogest — {sel_month_name} {year}")
    render_brand_tab("Cyclogest", ytd_df, py_ytd_df,
                     c_actual="#7D3C98", c_target="#C39BD3", c_line="#17A589",
                     show_pha=False, show_customer_count=False)

with tab_cis:
    st.subheader(f"Cialis — {sel_month_name} {year}")
    render_brand_tab("Cialis", ytd_df, py_ytd_df,
                     c_actual="#1A5276", c_target="#85C1E9", c_line="#F4A62A")

# ══════════════════════════════════════════════
with tab_ws:
    st.subheader(f"Wholesaler — {sel_month_name} {year}")

    _WS_BRANDS = ["Lamisil", "Betadine", "Nicotinell", "Cyclogest", "Cialis"]
    _mrange    = list(range(1, sel_month + 1))
    _mcols     = [MONTH_NAMES[m - 1] for m in _mrange]

    # Collect all in-license YTD rows, WS channel (KAM team) only
    _parts = [filter_by_brand(ytd_df, b) for b in _WS_BRANDS if BRAND_KEYWORDS.get(b)]
    _base  = pd.concat(_parts, ignore_index=True) if _parts else ytd_df.iloc[0:0]
    if COL_TEAM in _base.columns:
        _base = _base[
            _base[COL_TEAM].fillna("").str.lower().str.strip() == WS_TEAM_KW
        ].copy()
    else:
        _base = _base.copy()

    if _base.empty or COL_CUSTOMER not in _base.columns:
        st.info("No wholesaler data available for the selected period.")
    else:
        # Tag brand, SKU, month
        _base["_brand"] = None
        for _b in _WS_BRANDS:
            _kws = BRAND_KEYWORDS.get(_b, [])
            if _kws:
                _pat = "|".join(k.lower() for k in _kws)
                _msk = _base[COL_PRODUCT].fillna("").str.lower().str.contains(_pat, na=False)
                _base.loc[_msk, "_brand"] = _b
        _base = _base[_base["_brand"].notna()].copy()
        _base["_sku"]   = _base.apply(
            lambda r: normalize_sku(r[COL_PRODUCT], r["_brand"]), axis=1
        )
        _base["_month"] = _base[COL_DATE].dt.month

        # Optional customer name column
        _CNAME = next(
            (c for c in _base.columns
             if "customer name" in c.lower() and "sold" in c.lower()),
            None,
        )
        if _CNAME:
            _nm = (
                _base.drop_duplicates(COL_CUSTOMER)
                     .set_index(COL_CUSTOMER)[_CNAME].fillna("").to_dict()
            )
            _fcust = lambda c: f"{_nm.get(c, '')}  [{c}]" if _nm.get(c) else str(c)
        else:
            _fcust = lambda c: str(c)

        # Customers sorted by YTD sales desc
        _cust_order = (
            _base.groupby(COL_CUSTOMER)[COL_VALUE].sum()
            .sort_values(ascending=False)
            .index.tolist()
        )

        # ── Search box + dropdown ────────────────────────────────
        _search = st.text_input("Search Wholesaler", placeholder="Type name or code to filter…")
        _filtered_custs = (
            [c for c in _cust_order if _search.strip().lower() in _fcust(c).lower()]
            if _search.strip() else _cust_order
        )
        if not _filtered_custs:
            st.warning(f"No wholesaler matching '{_search}'.")
        else:
            _sel_cust = st.selectbox(
                f"Select ({len(_filtered_custs)} result{'s' if len(_filtered_custs) != 1 else ''})",
                options=_filtered_custs,
                format_func=_fcust,
            )

            _cdf = _base[_base[COL_CUSTOMER] == _sel_cust]

            # ── KPI cards ─────────────────────────────────────────
            _ytd_v = _cdf[COL_VALUE].sum()
            _ytd_q = _cdf[COL_INV_QTY].sum() if COL_INV_QTY in _cdf.columns else 0
            _n_sku = int(_cdf["_sku"].nunique())
            kc1, kc2, kc3 = st.columns(3)
            kc1.metric("YTD Sales Value (KRW)", f"{_ytd_v:,.0f}")
            kc2.metric("YTD Invoice Qty",       f"{int(_ytd_q):,}")
            kc3.metric("Active SKUs",           str(_n_sku))
            st.markdown("---")

            # ── Build (Brand, SKU) row order: brand order → YTD desc
            _sku_ytd = _cdf.groupby(["_brand", "_sku"])[COL_VALUE].sum().to_dict()
            _bidx    = {b: i for i, b in enumerate(_WS_BRANDS)}
            _sku_pairs = (
                _cdf[["_brand", "_sku"]].drop_duplicates()
                .assign(_bidx=lambda d: d["_brand"].map(_bidx).fillna(99))
                .assign(_ytd=lambda d: d.apply(
                    lambda r: _sku_ytd.get((r["_brand"], r["_sku"]), 0), axis=1
                ))
                .sort_values(["_bidx", "_ytd"], ascending=[True, False])
                [["_brand", "_sku"]].values.tolist()
            )

            _dcols = ["Brand", "SKU"] + _mcols + ["YTD Total"]
            _val_rows, _qty_rows = [], []
            for brand_, sku_ in _sku_pairs:
                sdf_ = _cdf[(_cdf["_brand"] == brand_) & (_cdf["_sku"] == sku_)]
                vr = {"Brand": brand_, "SKU": sku_}
                qr = {"Brand": brand_, "SKU": sku_}
                for m in _mrange:
                    mdf_ = sdf_[sdf_["_month"] == m]
                    vr[MONTH_NAMES[m - 1]] = mdf_[COL_VALUE].sum()
                    qr[MONTH_NAMES[m - 1]] = mdf_[COL_INV_QTY].sum() if COL_INV_QTY in mdf_.columns else 0
                vr["YTD Total"] = sdf_[COL_VALUE].sum()
                qr["YTD Total"] = sdf_[COL_INV_QTY].sum() if COL_INV_QTY in sdf_.columns else 0
                _val_rows.append(vr)
                _qty_rows.append(qr)

            def _mk_tbl(rows_, is_qty_=False):
                if not rows_:
                    return pd.DataFrame()
                df_ = pd.DataFrame(rows_)[_dcols].set_index(["Brand", "SKU"])
                for c in df_.columns:
                    df_[c] = df_[c].apply(
                        lambda x: (f"{int(x):,}" if is_qty_ else f"{x:,.0f}") if x else "—"
                    )
                return df_

            st.markdown("#### Sales Value (KRW) by SKU / Month")
            st.dataframe(_mk_tbl(_val_rows), use_container_width=True)

            st.markdown("#### Invoice Qty by SKU / Month")
            st.dataframe(_mk_tbl(_qty_rows, is_qty_=True), use_container_width=True)
            st.markdown("---")

            # ── All Wholesalers YTD summary (collapsible) ─────────
            with st.expander("📋 All Wholesalers — YTD Summary"):
                _brands_present = [b for b in _WS_BRANDS if b in _base["_brand"].unique()]
                _summ = (
                    _base.groupby([COL_CUSTOMER, "_brand"])[COL_VALUE]
                    .sum()
                    .unstack(fill_value=0)
                    .reindex(columns=_brands_present, fill_value=0)
                )
                _summ["Total"] = _summ.sum(axis=1)
                _summ = _summ.sort_values("Total", ascending=False)
                _summ.index = _summ.index.map(_fcust)
                _summ.index.name = "Wholesaler"
                _summ_disp = _summ.copy()
                for c in _summ_disp.columns:
                    _summ_disp[c] = _summ_disp[c].apply(lambda x: f"{x:,.0f}" if x else "—")
                st.dataframe(_summ_disp, use_container_width=True)

st.caption("Press F5 to refresh and load the latest data.")
