import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
    "LG":         ["멘소래담", "아크네스", "립아이스", "소프트립스", "메디케이티드립"],
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

INLICENSE_BRANDS = ["Lamisil", "Betadine", "Nicotinell", "Cyclogest", "Cialis", "LG"]

# 브랜드별 정규 SKU 매핑: (정규명, [매칭 키워드 ALL-match])
# 키워드가 모두 포함되면 해당 정규명으로 통합 (특수문자·접두/접미 무시)
# 순서 중요: 구체적 패턴을 먼저, 포괄적 패턴을 나중에
SKU_NORMALIZE: dict[str, list[tuple[str, list[str]]]] = {
    "Lamisil": [
        ("라미실원스외용액 4g",       ["원스"]),
        ("라미실덤겔 1% 15g",        ["덤겔"]),
        ("라미실크림 1% 15g",        ["크림", "15"]),
        ("라미실크림 1% 30g",        ["크림", "30"]),
        ("라미실외용액(병) 1% 30ml", ["외용액"]),
    ],
    "Betadine": [
        # 인후/파우더 스프레이
        ("Betadine Throat Spray 0.45% 25ml",          ["인후스프레이", "25"]),
        ("Betadine Throat Spray 0.45% 50ml",          ["인후스프레이", "50"]),
        ("Dry Powder Spray 2.5% 5g",                  ["드라이파우더"]),
        # Gynobetadine — 질좌제(pessaries) 먼저, 질세정(douche) 다음
        ("Gynobetadine Pessaries 10T",                ["질좌제", "10t"]),
        ("Gynobetadine Pessaries 50T",                ["질좌제", "50t"]),
        ("Gynobetadine V.Douche 180ML",               ["질세정", "180"]),
        ("Gynobetadine V.Douche 360Ml",               ["질세정", "360"]),
        # 농후액(Vetadine) — 소독액보다 먼저
        ("Vetadine Antis. Sol 10% 1G/L",              ["농후액"]),
        # Antiseptic 소독액: 1G/L → 2014 → 1L 순으로 구체적인 것 먼저
        ("Betadine Antiseptic Solution 10% 1G/L",     ["소독액", "gl"]),
        ("Betadine Antiseptic Solution 10% 1L(2014)", ["소독액", "2014"]),
        ("Betadine Antiseptic Solution 10% 1L",       ["소독액"]),
        # Skincleanser 세정액: 1G/L → 1L 순
        ("Betadine Skincleanser7.5% 1G/L",            ["세정액", "gl"]),
        ("Betadine Skincleanser7.5% 1L",              ["세정액"]),
    ],
    "Nicotinell": [
        ("니코틴엘 TTS10 7매",       ["tts10"]),
        ("니코틴엘 TTS20 7매",       ["tts20"]),
        ("니코틴엘 TTS30 7매",       ["tts30"]),
        ("니코틴엘 껌 2MG 24개",     ["껌"]),
        ("니코틴엘 로젠즈 1MG 36개", ["로젠즈"]),
    ],
}
# ═══════════════════════════════════════════════════════════

st.set_page_config(page_title="DSR Dashboard", page_icon="📊", layout="wide")
st.title("📊 DSR Sales Dashboard")


@st.cache_data(ttl=300)
def load_actuals():
    exclude = {
        os.path.abspath(BRAND_TARGET_FILE),
        os.path.abspath(DATA_DIR / "budget.xlsx"),
    }
    files = [
        f for f in glob.glob(str(DATA_DIR / "*.xlsx"))
        if os.path.abspath(f) not in exclude
        and not Path(f).name.startswith("~$")
    ]
    if not files:
        return None, []
    frames = []
    for f in files:
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
    return df, files


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
    row = targets_df[targets_df["Brand"] == brand].iloc[0]
    return sum(float(row.get(MONTH_NAMES[i], 0) or 0) for i in range(up_to_month_idx + 1))


def get_mtd_target(targets_df: pd.DataFrame, brand: str, month_idx: int) -> float:
    if targets_df.empty or brand not in targets_df["Brand"].values:
        return 0.0
    row = targets_df[targets_df["Brand"] == brand].iloc[0]
    return float(row.get(MONTH_NAMES[month_idx], 0) or 0)


def filter_by_brand(df: pd.DataFrame, brand: str) -> pd.DataFrame:
    keywords = BRAND_KEYWORDS.get(brand, [])
    if not keywords:
        return df.iloc[0:0]
    pattern = "|".join(k.lower() for k in keywords)
    mask = df[COL_PRODUCT].fillna("").str.lower().str.contains(pattern, na=False)
    return df[mask]


def _strip_product_name(name: str) -> str:
    """공통 접두/접미사를 제거해 표시용 이름을 정리."""
    s = str(name).strip()
    s = re.sub(r"^\([A-Za-z]\)\s*", "", s)       # (K) (G) 등 제거
    s = re.sub(r"^\*+\s*", "", s)                # 선행 * 제거
    s = re.sub(r"^[가-힣]+[-]\s*", "", s)        # 구반품- 등 한글 접두어+하이픈 제거
    s = re.sub(r"\s+\d+EA\s*$", "", s, re.I)    # 후행 1EA 2EA 등 제거
    return re.sub(r"\s+", " ", s).strip()


def normalize_sku(name: str, brand: str) -> str:
    """브랜드 정규 SKU 매핑 우선 적용, 없으면 일반 정제명 반환."""
    key = re.sub(r"[^\w가-힣%]", "", str(name).lower())  # 숫자·한글·영문·%만 남김
    if brand in SKU_NORMALIZE:
        for canonical, kws in SKU_NORMALIZE[brand]:
            kws_clean = [re.sub(r"[^\w가-힣%]", "", k.lower()) for k in kws]
            if all(k in key for k in kws_clean):
                return canonical
    return _strip_product_name(name)


def fmt(v):
    return f"{v/1_000_000:,.0f}M" if pd.notna(v) and v != 0 else "—"


def fmt_pct(v):
    return f"{v:.1f}%" if pd.notna(v) else "—"


# ─── 데이터 로드 ───────────────────────────────────────────
df, actuals_file = load_actuals()

if df is None:
    st.warning(f"`{DATA_DIR}/` 폴더에 실적 Excel 파일이 없습니다.")
    st.stop()

targets_df = load_brand_targets()
today      = pd.Timestamp.today()

# ─── 월 선택기 (사이드바) ──────────────────────────────────
with st.sidebar:
    st.header("⚙️ 조회 설정")
    available_months = [m for m in range(1, 13) if m <= today.month]
    sel_month = st.selectbox(
        "조회 월 선택",
        options=available_months,
        index=len(available_months) - 1,
        format_func=lambda m: f"{today.year}년 {m}월",
    )

sel_month_idx  = sel_month - 1
sel_month_name = MONTH_NAMES[sel_month_idx]
year           = today.year

mtd_df = df[(df[COL_DATE].dt.year == year) & (df[COL_DATE].dt.month == sel_month)]
ytd_df = df[(df[COL_DATE].dt.year == year) & (df[COL_DATE].dt.month <= sel_month)]

st.caption(
    f"실적 파일 {len(actuals_file)}개 로드 | "
    f"기준: {year}년 {sel_month}월 | "
    f"MTD {len(mtd_df):,}건 / YTD {len(ytd_df):,}건"
)

all_brands = targets_df["Brand"].tolist() if not targets_df.empty else []


def calc_kpis(brand_list, mtd_data, ytd_data):
    mtd_act = sum(filter_by_brand(mtd_data, b)[COL_VALUE].sum()
                  for b in brand_list if BRAND_KEYWORDS.get(b))
    ytd_act = sum(filter_by_brand(ytd_data, b)[COL_VALUE].sum()
                  for b in brand_list if BRAND_KEYWORDS.get(b))
    mtd_tgt = sum(get_mtd_target(targets_df, b, sel_month_idx) for b in brand_list)
    ytd_tgt = sum(get_ytd_target(targets_df, b, sel_month_idx) for b in brand_list)
    mtd_ach = (mtd_act / mtd_tgt * 100) if mtd_tgt > 0 else None
    ytd_ach = (ytd_act / ytd_tgt * 100) if ytd_tgt > 0 else None
    return mtd_act, mtd_tgt, mtd_ach, ytd_act, ytd_tgt, ytd_ach


# ─── Company Total ─────────────────────────────────────────
co_mtd_act, co_mtd_tgt, co_mtd_ach, co_ytd_act, co_ytd_tgt, co_ytd_ach = calc_kpis(all_brands, mtd_df, ytd_df)

st.markdown("### 🏢 Company Total")
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("MTD 실적 (M)", fmt(co_mtd_act))
c2.metric("MTD 목표 (M)", fmt(co_mtd_tgt))
c3.metric("MTD 달성률",   fmt_pct(co_mtd_ach),
          delta=f"{co_mtd_ach-100:.1f}%p" if co_mtd_ach else None)
c4.metric("YTD 실적 (M)", fmt(co_ytd_act))
c5.metric("YTD 목표 (M)", fmt(co_ytd_tgt))
c6.metric("YTD 달성률",   fmt_pct(co_ytd_ach),
          delta=f"{co_ytd_ach-100:.1f}%p" if co_ytd_ach else None)

st.divider()

# ─── Inlicense Total ───────────────────────────────────────
il_brands_valid = [b for b in INLICENSE_BRANDS if b in all_brands]
il_mtd_act, il_mtd_tgt, il_mtd_ach, il_ytd_act, il_ytd_tgt, il_ytd_ach = calc_kpis(il_brands_valid, mtd_df, ytd_df)

st.markdown("### 🔬 Inlicense Total")
d1, d2, d3, d4, d5, d6 = st.columns(6)
d1.metric("MTD 실적 (M)", fmt(il_mtd_act))
d2.metric("MTD 목표 (M)", fmt(il_mtd_tgt))
d3.metric("MTD 달성률",   fmt_pct(il_mtd_ach),
          delta=f"{il_mtd_ach-100:.1f}%p" if il_mtd_ach else None)
d4.metric("YTD 실적 (M)", fmt(il_ytd_act))
d5.metric("YTD 목표 (M)", fmt(il_ytd_tgt))
d6.metric("YTD 달성률",   fmt_pct(il_ytd_ach),
          delta=f"{il_ytd_ach-100:.1f}%p" if il_ytd_ach else None)

st.divider()


# ─── 브랜드 분석 탭 공통 렌더링 ────────────────────────────
def render_brand_tab(brand_name: str, ytd_data: pd.DataFrame):
    bdf = filter_by_brand(ytd_data, brand_name).copy()

    if bdf.empty:
        st.info(f"📭 {year}년 {sel_month}월까지 {brand_name} 실적이 없습니다.")
        return

    bdf["_month"] = bdf[COL_DATE].dt.month
    bdf["_sku"]   = bdf[COL_PRODUCT].apply(lambda x: normalize_sku(x, brand_name))
    month_range   = list(range(1, sel_month + 1))
    has_cust      = COL_CUSTOMER in bdf.columns

    # ── 월별 Summary 테이블 (metrics × months) ──────────────
    st.markdown("#### 월별 Summary")
    summary = {"수량": {}, "금액 (M)": {}, "판매처수": {}}
    for m in month_range:
        mdf = bdf[bdf["_month"] == m]
        col = MONTH_NAMES[m - 1]
        summary["수량"][col]     = f"{int(mdf[COL_INV_QTY].sum()):,}"
        summary["금액 (M)"][col] = fmt(mdf[COL_VALUE].sum())
        summary["판매처수"][col] = (
            f"{mdf[COL_CUSTOMER].nunique():,}" if has_cust else "-"
        )
    # YTD 합계 열 추가
    summary["수량"]["YTD 합계"]     = f"{int(bdf[COL_INV_QTY].sum()):,}"
    summary["금액 (M)"]["YTD 합계"] = fmt(bdf[COL_VALUE].sum())
    summary["판매처수"]["YTD 합계"] = (
        f"{bdf[COL_CUSTOMER].nunique():,}" if has_cust else "-"
    )

    st.dataframe(pd.DataFrame(summary).T, use_container_width=True)
    st.markdown("---")

    # ── SKU별 수량 Pivot ─────────────────────────────────────
    st.markdown("#### SKU별 수량 (Invoice Qty)")
    qty_piv = bdf.pivot_table(
        index="_sku", columns="_month",
        values=COL_INV_QTY, aggfunc="sum", fill_value=0,
    )
    qty_piv.columns = [MONTH_NAMES[c - 1] for c in qty_piv.columns]
    qty_piv["YTD 합계"] = qty_piv.sum(axis=1)
    qty_piv = qty_piv.sort_values("YTD 합계", ascending=False)
    qty_piv.index.name = "SKU (정규화)"

    qty_disp = qty_piv.copy()
    for col in qty_disp.columns:
        qty_disp[col] = qty_disp[col].apply(lambda x: f"{int(x):,}")
    st.dataframe(qty_disp, use_container_width=True)
    st.markdown("---")

    # ── SKU별 금액 Pivot ─────────────────────────────────────
    st.markdown("#### SKU별 금액 (M)")
    val_piv = bdf.pivot_table(
        index="_sku", columns="_month",
        values=COL_VALUE, aggfunc="sum", fill_value=0,
    )
    val_piv.columns = [MONTH_NAMES[c - 1] for c in val_piv.columns]
    val_piv["YTD 합계"] = val_piv.sum(axis=1)
    val_piv = val_piv.sort_values("YTD 합계", ascending=False)
    val_piv.index.name = "SKU (정규화)"

    val_disp = val_piv.copy()
    for col in val_disp.columns:
        val_disp[col] = val_disp[col].apply(fmt)
    st.dataframe(val_disp, use_container_width=True)


# ─── Tabs ───────────────────────────────────────────────────
tab_ov, tab_lms, tab_btd, tab_ntl, tab_cyc, tab_cis = st.tabs([
    "🏷️ 브랜드별 달성률",
    "💊 Lamisil",
    "🩺 Betadine",
    "🚭 Nicotinell",
    "🔵 Cyclogest",
    "💙 Cialis",
])

# ══════════════════════════════════════════════
with tab_ov:
    st.subheader(f"브랜드별 YTD / MTD 달성률 — {year}년 {sel_month}월 기준")

    if targets_df.empty:
        st.warning("`data/brand_target.xlsx` 파일이 없습니다.")
    else:
        brands = targets_df["Brand"].tolist()
        rows = []
        for brand in brands:
            kw     = BRAND_KEYWORDS.get(brand, [])
            has_kw = len(kw) > 0

            mtd_tgt   = get_mtd_target(targets_df, brand, sel_month_idx)
            ytd_tgt   = get_ytd_target(targets_df, brand, sel_month_idx)
            mtd_act   = filter_by_brand(mtd_df, brand)[COL_VALUE].sum() if has_kw else 0.0
            ytd_act   = filter_by_brand(ytd_df, brand)[COL_VALUE].sum() if has_kw else 0.0
            mtd_ach_b = (mtd_act / mtd_tgt * 100) if mtd_tgt > 0 else None
            ytd_ach_b = (ytd_act / ytd_tgt * 100) if ytd_tgt > 0 else None

            rows.append({
                "브랜드":        brand,
                "MTD 목표 (M)":  mtd_tgt,
                "MTD 실적 (M)":  mtd_act,
                "MTD 달성률(%)": round(mtd_ach_b, 1) if mtd_ach_b is not None else None,
                "YTD 목표 (M)":  ytd_tgt,
                "YTD 실적 (M)":  ytd_act,
                "YTD 달성률(%)": round(ytd_ach_b, 1) if ytd_ach_b is not None else None,
                "_no_kw":        not has_kw,
            })

        bdf_ov = pd.DataFrame(rows)

        valid = (
            bdf_ov[bdf_ov["YTD 달성률(%)"].notna() & ~bdf_ov["_no_kw"]]
            .sort_values("YTD 달성률(%)", ascending=True)
        )
        if not valid.empty:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                y=valid["브랜드"], x=valid["YTD 달성률(%)"],
                orientation="h", name="YTD 달성률",
                marker_color=[
                    "#00CC96" if v >= 100 else "#FFA15A" if v >= 80 else "#EF553B"
                    for v in valid["YTD 달성률(%)"]
                ],
                text=[f"{v:.1f}%" for v in valid["YTD 달성률(%)"]],
                textposition="outside",
            ))
            fig.add_trace(go.Scatter(
                y=valid["브랜드"], x=valid["MTD 달성률(%)"],
                mode="markers", name="MTD 달성률",
                marker=dict(symbol="diamond", size=10, color="#636EFA"),
            ))
            fig.add_vline(x=100, line_dash="dash", line_color="gray",
                          annotation_text="100%", annotation_position="top right")
            fig.update_layout(
                height=max(350, len(valid) * 50),
                margin=dict(l=0, r=90, t=30, b=0),
                xaxis_title="달성률 (%)",
                legend=dict(orientation="h", y=1.08),
                barmode="overlay",
            )
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("상세 테이블")
        disp = bdf_ov[[
            "브랜드","MTD 목표 (M)","MTD 실적 (M)","MTD 달성률(%)",
            "YTD 목표 (M)","YTD 실적 (M)","YTD 달성률(%)"
        ]].copy()
        for col in ["MTD 목표 (M)","MTD 실적 (M)","YTD 목표 (M)","YTD 실적 (M)"]:
            disp[col] = disp[col].apply(fmt)
        for col in ["MTD 달성률(%)","YTD 달성률(%)"]:
            disp[col] = disp[col].apply(fmt_pct)
        st.dataframe(disp, use_container_width=True, hide_index=True)

        no_kw = bdf_ov[bdf_ov["_no_kw"]]["브랜드"].tolist()
        if no_kw:
            st.warning(f"⚠️ **{', '.join(no_kw)}** 브랜드는 제품명 키워드가 없어 실적이 0입니다.")

# ══════════════════════════════════════════════
with tab_lms:
    st.subheader(f"Lamisil — {year}년 {sel_month}월 기준")
    render_brand_tab("Lamisil", ytd_df)

with tab_btd:
    st.subheader(f"Betadine — {year}년 {sel_month}월 기준")
    render_brand_tab("Betadine", ytd_df)

with tab_ntl:
    st.subheader(f"Nicotinell — {year}년 {sel_month}월 기준")
    render_brand_tab("Nicotinell", ytd_df)

with tab_cyc:
    st.subheader(f"Cyclogest — {year}년 {sel_month}월 기준")
    render_brand_tab("Cyclogest", ytd_df)

with tab_cis:
    st.subheader(f"Cialis — {year}년 {sel_month}월 기준")
    render_brand_tab("Cialis", ytd_df)

st.caption("새로고침(F5)하면 최신 데이터로 업데이트됩니다.")
