from pathlib import Path
from datetime import datetime
import re

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="PigCost Insight, 양돈 사료 원가를 한눈에",
    layout="wide"
)

# -------------------------------------------------
# 스타일
# -------------------------------------------------
st.markdown("""
<style>
.block-container {
    padding-top: 1.1rem;
    padding-bottom: 2.2rem;
    padding-left: 2rem;
    padding-right: 2rem;
    max-width: 1500px;
}

.main-title {
    font-size: 2.35rem;
    font-weight: 900;
    color: #0f172a;
    margin-bottom: 0.25rem;
    letter-spacing: -0.03em;
}

.sub-title {
    color: #64748b;
    font-size: 1rem;
    margin-bottom: 0;
}

.hero-box {
    background:
        radial-gradient(circle at top right, rgba(59,130,246,0.12), transparent 26%),
        radial-gradient(circle at bottom left, rgba(14,165,233,0.10), transparent 28%),
        linear-gradient(135deg, #f8fbff 0%, #f5f9ff 45%, #f8faff 100%);
    border: 1px solid #e6edf5;
    padding: 1.4rem 1.55rem;
    border-radius: 24px;
    margin-bottom: 1.1rem;
    box-shadow: 0 16px 40px rgba(15, 23, 42, 0.07);
}

.section-title {
    font-size: 1.15rem;
    font-weight: 900;
    color: #0f172a;
    margin: 0.35rem 0 0.9rem 0;
}

.kpi-card {
    border-radius: 22px;
    padding: 1rem 1.05rem 0.95rem 1.05rem;
    color: #0f172a;
    box-shadow: 0 14px 28px rgba(15, 23, 42, 0.08);
    border: 1px solid rgba(255,255,255,0.65);
    min-height: 148px;
    position: relative;
    overflow: hidden;
}

.kpi-card::after {
    content: "";
    position: absolute;
    right: -20px;
    top: -20px;
    width: 90px;
    height: 90px;
    border-radius: 999px;
    background: rgba(255,255,255,0.28);
}

.kpi-corn {
    background: linear-gradient(135deg, #fff7ed 0%, #ffedd5 100%);
}

.kpi-soy {
    background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
}

.kpi-fx {
    background: linear-gradient(135deg, #f5f3ff 0%, #ede9fe 100%);
}

.kpi-freight {
    background: linear-gradient(135deg, #fdf2f8 0%, #fce7f3 100%);
}

.kpi-label {
    font-size: 0.92rem;
    font-weight: 800;
    color: #334155;
    margin-bottom: 0.6rem;
}

.kpi-value {
    font-size: 2rem;
    font-weight: 900;
    color: #0f172a;
    line-height: 1.1;
    margin-bottom: 0.35rem;
    letter-spacing: -0.02em;
}

.kpi-delta-up {
    color: #16a34a;
    font-size: 0.92rem;
    font-weight: 800;
    margin-bottom: 0.25rem;
}

.kpi-delta-down {
    color: #dc2626;
    font-size: 0.92rem;
    font-weight: 800;
    margin-bottom: 0.25rem;
}

.kpi-delta-flat {
    color: #475569;
    font-size: 0.92rem;
    font-weight: 800;
    margin-bottom: 0.25rem;
}

.kpi-sub {
    color: #64748b;
    font-size: 0.82rem;
    font-weight: 600;
}

.small-card {
    background: #ffffff;
    border: 1px solid #ebf0f6;
    border-radius: 22px;
    padding: 0.9rem 1rem 0.7rem 1rem;
    box-shadow: 0 12px 26px rgba(15, 23, 42, 0.05);
    margin-bottom: 0.8rem;
}

.info-card {
    background: linear-gradient(135deg, #ffffff 0%, #f8fbff 100%);
    border: 1px solid #ebf0f6;
    border-radius: 20px;
    padding: 0.9rem 1rem;
    box-shadow: 0 12px 24px rgba(15, 23, 42, 0.04);
    margin-bottom: 0.8rem;
}

div.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
}

div.stTabs [data-baseweb="tab"] {
    background: #f8fafc;
    border: 1px solid #e5e7eb;
    border-radius: 12px;
    padding-left: 16px;
    padding-right: 16px;
    height: 42px;
    font-weight: 700;
}

div.stTabs [aria-selected="true"] {
    background: #e0f2fe !important;
    border-color: #7dd3fc !important;
}

div[data-baseweb="select"] > div {
    border-radius: 14px !important;
    border-color: #dbe4ee !important;
}

div[data-testid="stMetric"] {
    background: #ffffff;
    border: 1px solid #eef2f7;
    padding: 14px 16px;
    border-radius: 18px;
    box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
}

div[data-testid="stMetricLabel"] {
    font-weight: 800;
}

hr {
    margin-top: 1rem !important;
    margin-bottom: 1rem !important;
}

.hero-inner {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 1.2rem;
}

.hero-org {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    white-space: nowrap;
    color: #334155;
    font-size: 0.95rem;
    font-weight: 800;
}

.gov-logo {
    height: 46px;
    object-fit: contain;
}
            
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# 파일 경로
# -------------------------------------------------
DATA_DIR = Path("data")

FEED_PRICE_XLSX = DATA_DIR / "feed_price_monthly.xlsx"
FEED_PRICE_CSV = DATA_DIR / "feed_price_monthly.csv"

FEED_PROD_XLSX = DATA_DIR / "feed_production_detail.xlsx"
FEED_PROD_CSV = DATA_DIR / "feed_production_detail.csv"

CORN_XLSX = DATA_DIR / "corn_90d.xlsx"
CORN_CSV = DATA_DIR / "corn_90d.csv"

SOY_XLSX = DATA_DIR / "soymeal_90d.xlsx"
SOY_CSV = DATA_DIR / "soymeal_90d.csv"

FX_XLSX = DATA_DIR / "fx_90d.xlsx"
FX_CSV = DATA_DIR / "fx_90d.csv"

FREIGHT_XLSX = DATA_DIR / "freight_90d.xlsx"
FREIGHT_CSV = DATA_DIR / "freight_90d.csv"

# -------------------------------------------------
# 이름 통합 / 순서
# -------------------------------------------------
CATEGORY_ALIAS = {
    "성장기 전체": "양돈용사료 전체",
    "양돈사료 전체": "양돈용사료 전체",
    "양돈전체": "양돈용사료 전체",
    "전체": "양돈용사료 전체",
    "양돈용사료 전체": "양돈용사료 전체",
    "포유자돈": "포유돈",
    "포유돈": "포유돈",
    "이유돈": "이유돈",
    "육성돈": "육성돈",
    "비육돈": "비육돈",
    "번식용모돈": "번식용모돈",
    "번식용 모돈": "번식용모돈",
    "임신돈": "임신돈",
}

CATEGORY_ORDER = [
    "양돈용사료 전체",
    "포유돈",
    "이유돈",
    "육성돈",
    "비육돈",
    "번식용모돈",
    "임신돈",
]

CATEGORY_COLORS = {
    "양돈용사료 전체": "#2563eb",
    "포유돈": "#f59e0b",
    "이유돈": "#10b981",
    "육성돈": "#8b5cf6",
    "비육돈": "#ef4444",
    "번식용모돈": "#f97316",
    "임신돈": "#14b8a6",
}

# -------------------------------------------------
# 컬럼명 표준화
# -------------------------------------------------
def normalize_token(text: str) -> str:
    text = str(text).strip().lower()
    text = re.sub(r"[\\s_\\-\\/\\(\\)\\[\\]{}]", "", text)
    return text

COLUMN_CANDIDATES = {
    "month": ["month", "월", "기준월", "년월"],
    "price_krw_per_kg": ["price_krw_per_kg", "price", "가격", "가격원kg", "가격(원/kg)", "원kg", "배합사료가격"],
    "category": ["category", "사료구분", "구분", "성장단계", "품목"],
    "production_ton": ["production_ton", "production", "생산량", "생산량톤", "생산량(톤)", "톤"],
    "date": ["date", "날짜", "일자"],
    "price": ["price", "가격", "종가", "close", "closeprice", "값", "지수"]
}

def standardize_columns(df: pd.DataFrame, target_keys: list[str]) -> pd.DataFrame:
    cols = list(df.columns)
    normalized_map = {c: normalize_token(c) for c in cols}
    rename_map = {}

    for target in target_keys:
        candidates = [normalize_token(x) for x in COLUMN_CANDIDATES.get(target, [target])]
        for original, normalized in normalized_map.items():
            if normalized in candidates:
                rename_map[original] = target
                break

    return df.rename(columns=rename_map)

# -------------------------------------------------
# 공통 함수
# -------------------------------------------------
def file_signature(*paths):
    sig = []
    for p in paths:
        if p.exists():
            sig.append((p.name, int(p.stat().st_mtime)))
        else:
            sig.append((p.name, 0))
    return tuple(sig)

def read_csv_with_fallback(csv_path: Path):
    last_error = None
    for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
        try:
            return pd.read_csv(csv_path, encoding=enc)
        except Exception as e:
            last_error = e
    raise last_error

def load_table(xlsx_path: Path, csv_path: Path):
    if xlsx_path.exists():
        return pd.read_excel(xlsx_path)
    if csv_path.exists():
        return read_csv_with_fallback(csv_path)
    return None

@st.cache_data
def load_feed_price(_sig=None):
    df = load_table(FEED_PRICE_XLSX, FEED_PRICE_CSV)
    if df is None:
        return None

    df.columns = [str(c).strip() for c in df.columns]
    df = standardize_columns(df, ["month", "price_krw_per_kg"])

    required_cols = ["month", "price_krw_per_kg"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError("feed_price_monthly 파일 컬럼은 month/월, price_krw_per_kg/가격 형태여야 합니다.")

    df["month"] = pd.to_datetime(df["month"], format="%Y-%m", errors="coerce")
    if df["month"].isna().all():
        df["month"] = pd.to_datetime(df["month"], errors="coerce")

    df["price_krw_per_kg"] = pd.to_numeric(df["price_krw_per_kg"], errors="coerce")

    df = (
        df.dropna(subset=["month", "price_krw_per_kg"])
          .sort_values("month")
          .reset_index(drop=True)
    )

    if not df.empty:
        df["month_label"] = df["month"].dt.strftime("%Y-%m")

    return df

@st.cache_data
def load_feed_production(_sig=None):
    df = load_table(FEED_PROD_XLSX, FEED_PROD_CSV)
    if df is None:
        return None

    df.columns = [str(c).strip() for c in df.columns]
    df = standardize_columns(df, ["month", "category", "production_ton"])

    required_cols = ["month", "category", "production_ton"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError("feed_production_detail 파일 컬럼은 month/월, category/사료구분, production_ton/생산량 형태여야 합니다.")

    df["month"] = pd.to_datetime(df["month"], format="%Y-%m", errors="coerce")
    if df["month"].isna().all():
        df["month"] = pd.to_datetime(df["month"], errors="coerce")

    df["category"] = df["category"].astype(str).str.strip()
    df["category"] = df["category"].replace(CATEGORY_ALIAS)
    df["production_ton"] = pd.to_numeric(df["production_ton"], errors="coerce")

    df = df.dropna(subset=["month", "category", "production_ton"])

    df = (
        df.groupby(["month", "category"], as_index=False)["production_ton"]
          .sum()
    )

    all_months = sorted(df["month"].dropna().unique())
    add_rows = []

    for m in all_months:
        month_df = df[df["month"] == m].copy()
        has_total = (month_df["category"] == "양돈용사료 전체").any()

        if not has_total:
            subtotal = month_df[
                month_df["category"].isin(CATEGORY_ORDER[1:])
            ]["production_ton"].sum()

            if subtotal > 0:
                add_rows.append({
                    "month": m,
                    "category": "양돈용사료 전체",
                    "production_ton": subtotal
                })

    if add_rows:
        df = pd.concat([df, pd.DataFrame(add_rows)], ignore_index=True)

    df["category_order"] = df["category"].apply(
        lambda x: CATEGORY_ORDER.index(x) if x in CATEGORY_ORDER else 999
    )

    df = (
        df.sort_values(["month", "category_order", "category"])
          .drop(columns=["category_order"])
          .reset_index(drop=True)
    )

    if not df.empty:
        df["month_label"] = df["month"].dt.strftime("%Y-%m")

    return df

@st.cache_data
def load_series(xlsx_path: Path, csv_path: Path, _sig=None):
    df = load_table(xlsx_path, csv_path)
    if df is None:
        return None

    df.columns = [str(c).strip() for c in df.columns]
    df = standardize_columns(df, ["date", "price"])

    required_cols = ["date", "price"]
    for col in required_cols:
        if col not in df.columns:
            return None

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["price"] = pd.to_numeric(df["price"], errors="coerce")

    df = (
        df.dropna(subset=["date", "price"])
          .sort_values("date")
          .reset_index(drop=True)
    )

    if not df.empty:
        df["date_label"] = df["date"].dt.strftime("%Y-%m-%d")

    return df

def metric_delta_parts(current, previous, digits=2):
    if previous is None or pd.isna(previous):
        return "변동 없음", "kpi-delta-flat"

    delta = current - previous
    pct = 0 if previous == 0 else (delta / previous) * 100

    if delta > 0:
        return f"▲ {delta:+,.{digits}f} ({pct:+.2f}%)", "kpi-delta-up"
    elif delta < 0:
        return f"▼ {delta:+,.{digits}f} ({pct:+.2f}%)", "kpi-delta-down"
    return f"• {delta:+,.{digits}f} ({pct:+.2f}%)", "kpi-delta-flat"

def render_kpi_card(title, df, theme_class, unit="", digits=2, latest_prefix="최근 반영일"):
    if df is not None and len(df) >= 1:
        current = df.iloc[-1]["price"]
        previous = df.iloc[-2]["price"] if len(df) >= 2 else None
        latest_date = pd.to_datetime(df.iloc[-1]["date"]).strftime("%Y-%m-%d")
        delta_text, delta_class = metric_delta_parts(current, previous, digits=digits)
        value_text = f"{current:,.{digits}f}{unit}"
        sub_text = f"{latest_prefix} {latest_date}"
    else:
        delta_text, delta_class = "데이터 연결 대기", "kpi-delta-flat"
        value_text = "준비중"
        sub_text = "파일 생성 후 자동 반영"

    st.markdown(
        f"""
        <div class="kpi-card {theme_class}">
            <div class="kpi-label">{title}</div>
            <div class="kpi-value">{value_text}</div>
            <div class="{delta_class}">{delta_text}</div>
            <div class="kpi-sub">{sub_text}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

def apply_light_plot(fig):
    fig.update_layout(
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#ffffff",
        font=dict(color="#1f2937", family="Arial"),
        title_font=dict(size=18, color="#111827"),
        legend_font=dict(color="#475569"),
        hoverlabel=dict(bgcolor="white", font_size=13, font_color="#111827"),
        hovermode="x unified",
        margin=dict(l=12, r=12, t=58, b=10),
        xaxis=dict(
            showgrid=False,
            zeroline=False,
            tickfont=dict(color="#64748b")
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor="rgba(148,163,184,0.18)",
            zeroline=False,
            tickfont=dict(color="#64748b")
        ),
    )
    return fig

def add_latest_badge(fig, x, y, text, color):
    fig.add_annotation(
        x=x,
        y=y,
        text=text,
        showarrow=True,
        arrowhead=2,
        arrowsize=1,
        arrowwidth=1.2,
        arrowcolor=color,
        ax=25,
        ay=-35,
        bgcolor="rgba(255,255,255,0.92)",
        bordercolor=color,
        borderwidth=1,
        font=dict(color="#111827", size=12)
    )
    return fig

def make_pretty_line_chart(df, x_col, y_col, title, line_color, y_title, fill_alpha=0.10):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df[x_col],
        y=df[y_col],
        mode="lines+markers",
        name=title,
        line=dict(color=line_color, width=3.5, shape="spline", smoothing=1.0),
        marker=dict(size=7, color=line_color, line=dict(color="white", width=1.5)),
        fill="tozeroy",
        fillcolor=line_color.replace(")", f", {fill_alpha})").replace("rgb", "rgba") if "rgb" in line_color else None
    ))

    # hex color용 fill 처리
    if not ("rgb" in line_color):
        fig.data[0].fillcolor = {
            "#2563eb": "rgba(37, 99, 235, 0.10)",
            "#10b981": "rgba(16, 185, 129, 0.10)",
            "#8b5cf6": "rgba(139, 92, 246, 0.10)",
            "#ec4899": "rgba(236, 72, 153, 0.10)",
            "#f59e0b": "rgba(245, 158, 11, 0.10)",
            "#0ea5e9": "rgba(14, 165, 233, 0.10)",
        }.get(line_color, "rgba(59,130,246,0.10)")

    fig.update_layout(
        title=title,
        height=360,
        xaxis_title="기준일",
        yaxis_title=y_title,
        showlegend=False
    )

    apply_light_plot(fig)

    last_x = df[x_col].iloc[-1]
    last_y = df[y_col].iloc[-1]
    badge = f"{last_y:,.2f}" if isinstance(last_y, float) else str(last_y)
    add_latest_badge(fig, last_x, last_y, badge, line_color)

    return fig

def make_pretty_bar_chart(df, x_col, y_col, title, color, y_title):
    fig = px.bar(
        df,
        x=x_col,
        y=y_col,
        title=title,
        text_auto=".0f",
        color_discrete_sequence=[color]
    )
    fig.update_traces(
        texttemplate="%{y:,.0f}",
        marker_line_width=0,
        opacity=0.95,
        hovertemplate="%{x}<br>%{y:,.0f}<extra></extra>"
    )
    fig.update_layout(
        height=360,
        xaxis_title="기준월",
        yaxis_title=y_title,
        showlegend=False
    )
    apply_light_plot(fig)
    return fig

# -------------------------------------------------
# 데이터 로드
# -------------------------------------------------
feed_price_df = load_feed_price(file_signature(FEED_PRICE_XLSX, FEED_PRICE_CSV))
feed_prod_df = load_feed_production(file_signature(FEED_PROD_XLSX, FEED_PROD_CSV))

corn_df = load_series(CORN_XLSX, CORN_CSV, file_signature(CORN_XLSX, CORN_CSV))
soy_df = load_series(SOY_XLSX, SOY_CSV, file_signature(SOY_XLSX, SOY_CSV))
fx_df = load_series(FX_XLSX, FX_CSV, file_signature(FX_XLSX, FX_CSV))
freight_df = load_series(FREIGHT_XLSX, FREIGHT_CSV, file_signature(FREIGHT_XLSX, FREIGHT_CSV))

# -------------------------------------------------
# 헤더
# -------------------------------------------------
st.markdown(f"""
<div class="hero-box">
    <div class="hero-inner">
        <div>
            <div class="main-title">PigCost Insight, 양돈 사료 원가를 한눈에</div>
            <div class="sub-title">
                기준일: {datetime.today().strftime("%Y-%m-%d")} · 배합사료 가격/생산량 기준: 농림축산식품부 최근통계자료
            </div>
        </div>
        <div class="hero-org">
            <div>국립축산과학원 양돈과</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------
# 상단 KPI
# -------------------------------------------------
st.markdown('<div class="section-title">상단 지표</div>', unsafe_allow_html=True)

k1, k2, k3, k4 = st.columns(4)
with k1:
    render_kpi_card("옥수수 현재가", corn_df, "kpi-card kpi-corn")
with k2:
    render_kpi_card("대두박 현재가", soy_df, "kpi-card kpi-soy")
with k3:
    render_kpi_card("USD/KRW", fx_df, "kpi-card kpi-fx")
with k4:
    render_kpi_card("해상운임", freight_df, "kpi-card kpi-freight")

st.markdown("---")

# -------------------------------------------------
# 필수 데이터 체크
# -------------------------------------------------
if feed_price_df is None:
    st.error("data/feed_price_monthly.xlsx 또는 data/feed_price_monthly.csv 파일이 없습니다.")
    st.stop()

if feed_prod_df is None:
    st.error("data/feed_production_detail.xlsx 또는 data/feed_production_detail.csv 파일이 없습니다.")
    st.stop()

# -------------------------------------------------
# 중간 영역
# -------------------------------------------------
st.markdown('<div class="section-title">양돈 사료 단계별 생산량 및 가격</div>', unsafe_allow_html=True)

latest_price_month = feed_price_df["month"].max()
latest_price = feed_price_df.loc[
    feed_price_df["month"] == latest_price_month, "price_krw_per_kg"
].iloc[0]

available_categories = [c for c in CATEGORY_ORDER if c in feed_prod_df["category"].unique().tolist()]
if not available_categories:
    st.warning("생산량 카테고리 데이터가 없습니다.")
    st.stop()

default_category = "양돈용사료 전체" if "양돈용사료 전체" in available_categories else available_categories[0]

selected_category = st.selectbox(
    "사료 생산단계 선택",
    available_categories,
    index=available_categories.index(default_category)
)

selected_prod_df = feed_prod_df[feed_prod_df["category"] == selected_category].copy()

if selected_prod_df.empty:
    st.warning(f"{selected_category} 데이터가 없습니다.")
    st.stop()

latest_prod_month = selected_prod_df["month"].max()
latest_prod = selected_prod_df.loc[
    selected_prod_df["month"] == latest_prod_month, "production_ton"
].iloc[0]

m1, m2, m3 = st.columns(3)
with m1:
    st.metric("최신 가격 기준월", latest_price_month.strftime("%Y-%m"))
with m2:
    st.metric("배합사료 가격", f"{latest_price:,.0f} 원/kg")
with m3:
    st.metric(f"{selected_category} 생산량", f"{latest_prod:,.0f} 톤")

left_card, right_card = st.columns(2)

with left_card:
    st.markdown('<div class="small-card">', unsafe_allow_html=True)
    fig_price = go.Figure()
    fig_price.add_trace(go.Scatter(
        x=feed_price_df["month_label"],
        y=feed_price_df["price_krw_per_kg"],
        mode="lines+markers",
        line=dict(color="#2563eb", width=4, shape="spline", smoothing=1.0),
        marker=dict(size=8, color="#2563eb", line=dict(color="white", width=2)),
        fill="tozeroy",
        fillcolor="rgba(37, 99, 235, 0.10)",
        hovertemplate="기준월 %{x}<br>가격 %{y:,.0f} 원/kg<extra></extra>"
    ))
    fig_price.update_layout(
        title="양돈 배합사료 월별 가격 추이",
        height=400,
        xaxis_title="기준월",
        yaxis_title="가격(원/kg)",
        showlegend=False
    )
    apply_light_plot(fig_price)
    add_latest_badge(
        fig_price,
        feed_price_df["month_label"].iloc[-1],
        feed_price_df["price_krw_per_kg"].iloc[-1],
        f"{feed_price_df['price_krw_per_kg'].iloc[-1]:,.0f} 원/kg",
        "#2563eb"
    )
    st.plotly_chart(fig_price, width="stretch")
    st.markdown('</div>', unsafe_allow_html=True)

with right_card:
    st.markdown('<div class="small-card">', unsafe_allow_html=True)
    prod_color = CATEGORY_COLORS.get(selected_category, "#10b981")
    fig_prod = px.bar(
        selected_prod_df,
        x="month_label",
        y="production_ton",
        title=f"{selected_category} 월별 생산량 추이",
        text_auto=".0f",
        color_discrete_sequence=[prod_color]
    )
    fig_prod.update_traces(
        marker_line_width=0,
        hovertemplate="기준월 %{x}<br>생산량 %{y:,.0f} 톤<extra></extra>"
    )
    fig_prod.update_layout(
        height=400,
        xaxis_title="기준월",
        yaxis_title="생산량(톤)",
        showlegend=False
    )
    apply_light_plot(fig_prod)
    st.plotly_chart(fig_prod, width="stretch")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# -------------------------------------------------
# 하단 탭
# -------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(["통합", "환율", "양돈사료", "원본데이터"])

with tab1:
    st.subheader("통합 보기")

    row1_col1, row1_col2 = st.columns(2)
    row2_col1, row2_col2 = st.columns(2)

    with row1_col1:
        if corn_df is not None and not corn_df.empty:
            fig = make_pretty_line_chart(
                corn_df.tail(90),
                "date_label",
                "price",
                "옥수수 90일 추이",
                "#f59e0b",
                "가격"
            )
            st.plotly_chart(fig, width="stretch")
        else:
            st.info("옥수수 데이터 파일이 아직 없습니다.")

    with row1_col2:
        if soy_df is not None and not soy_df.empty:
            fig = make_pretty_line_chart(
                soy_df.tail(90),
                "date_label",
                "price",
                "대두박 90일 추이",
                "#10b981",
                "가격"
            )
            st.plotly_chart(fig, width="stretch")
        else:
            st.info("대두박 데이터 파일이 아직 없습니다.")

    with row2_col1:
        if fx_df is not None and not fx_df.empty:
            fig = make_pretty_line_chart(
                fx_df.tail(90),
                "date_label",
                "price",
                "환율 90일 추이",
                "#8b5cf6",
                "원/달러"
            )
            st.plotly_chart(fig, width="stretch")
        else:
            st.info("환율 데이터 파일이 아직 없습니다.")

    with row2_col2:
        if freight_df is not None and not freight_df.empty:
            fig = make_pretty_line_chart(
                freight_df.tail(90),
                "date_label",
                "price",
                "해상운임 90일 추이",
                "#ec4899",
                "값"
            )
            st.plotly_chart(fig, width="stretch")
        else:
            st.info("해상운임 데이터 파일이 아직 없습니다.")

    st.markdown("### 양돈 배합사료 월별 추이")
    bottom1, bottom2 = st.columns(2)

    with bottom1:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=feed_price_df["month_label"],
            y=feed_price_df["price_krw_per_kg"],
            mode="lines+markers",
            line=dict(color="#0ea5e9", width=4, shape="spline", smoothing=1.0),
            marker=dict(size=8, color="#0ea5e9", line=dict(color="white", width=2)),
            fill="tozeroy",
            fillcolor="rgba(14, 165, 233, 0.10)",
            hovertemplate="기준월 %{x}<br>가격 %{y:,.0f} 원/kg<extra></extra>"
        ))
        fig.update_layout(
            title="양돈 배합사료 월별 가격",
            height=340,
            xaxis_title="기준월",
            yaxis_title="가격(원/kg)",
            showlegend=False
        )
        apply_light_plot(fig)
        st.plotly_chart(fig, width="stretch")

    with bottom2:
        latest_snapshot = (
            feed_prod_df[feed_prod_df["month"] == feed_prod_df["month"].max()]
            .copy()
        )

        latest_snapshot["category"] = pd.Categorical(
            latest_snapshot["category"],
            categories=CATEGORY_ORDER,
            ordered=True
        )

        latest_snapshot = (
            latest_snapshot.dropna(subset=["category"])
            .sort_values("category")
            .reset_index(drop=True)
        )

        if latest_snapshot.empty:
            st.info("최신월 생산량 데이터가 없습니다.")
        else:
            fig = px.bar(
                latest_snapshot,
                x="category",
                y="production_ton",
                color="category",
                title=f'{feed_prod_df["month"].max().strftime("%Y-%m")} 성장단계별 생산량',
                text_auto=".0f",
                color_discrete_map=CATEGORY_COLORS
            )
            fig.update_traces(
                texttemplate="%{y:,.0f}",
                hovertemplate="구분 %{x}<br>생산량 %{y:,.0f} 톤<extra></extra>"
            )
            fig.update_layout(
                height=340,
                xaxis_title="사료 구분",
                yaxis_title="생산량(톤)",
                showlegend=False
            )
            apply_light_plot(fig)
            st.plotly_chart(fig, width="stretch")

with tab2:
    st.subheader("환율")
    if fx_df is not None and not fx_df.empty:
        fig = make_pretty_line_chart(
            fx_df.tail(90),
            "date_label",
            "price",
            "USD/KRW 90일 추이",
            "#8b5cf6",
            "원/달러"
        )
        fig.update_layout(height=440)
        st.plotly_chart(fig, width="stretch")
    else:
        st.info("data/fx_90d 파일을 넣으면 환율 탭이 표시됩니다.")

with tab3:
    st.subheader("양돈사료")

    top_a, top_b = st.columns(2)

    with top_a:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=feed_price_df["month_label"],
            y=feed_price_df["price_krw_per_kg"],
            mode="lines+markers",
            line=dict(color="#2563eb", width=4, shape="spline", smoothing=1.0),
            marker=dict(size=8, color="#2563eb", line=dict(color="white", width=2)),
            fill="tozeroy",
            fillcolor="rgba(37, 99, 235, 0.10)",
            hovertemplate="기준월 %{x}<br>가격 %{y:,.0f} 원/kg<extra></extra>"
        ))
        fig.update_layout(
            title="배합사료 가격 월별 추이",
            height=360,
            xaxis_title="기준월",
            yaxis_title="가격(원/kg)",
            showlegend=False
        )
        apply_light_plot(fig)
        st.plotly_chart(fig, width="stretch")

    with top_b:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=selected_prod_df["month_label"],
            y=selected_prod_df["production_ton"],
            mode="lines+markers",
            line=dict(
                color=CATEGORY_COLORS.get(selected_category, "#10b981"),
                width=4,
                shape="spline",
                smoothing=1.0
            ),
            marker=dict(
                size=8,
                color=CATEGORY_COLORS.get(selected_category, "#10b981"),
                line=dict(color="white", width=2)
            ),
            fill="tozeroy",
            fillcolor="rgba(16, 185, 129, 0.10)",
            hovertemplate="기준월 %{x}<br>생산량 %{y:,.0f} 톤<extra></extra>"
        ))
        fig.update_layout(
            title=f"{selected_category} 생산량 월별 추이",
            height=360,
            xaxis_title="기준월",
            yaxis_title="생산량(톤)",
            showlegend=False
        )
        apply_light_plot(fig)
        st.plotly_chart(fig, width="stretch")

    latest_snapshot = (
        feed_prod_df[feed_prod_df["month"] == feed_prod_df["month"].max()]
        .copy()
    )
    latest_snapshot["category"] = pd.Categorical(
        latest_snapshot["category"],
        categories=CATEGORY_ORDER,
        ordered=True
    )
    latest_snapshot = (
        latest_snapshot.dropna(subset=["category"])
        .sort_values("category")
        .reset_index(drop=True)
    )

    if latest_snapshot.empty:
        st.info("최신월 성장단계별 생산량 데이터가 없습니다.")
    else:
        fig = px.bar(
            latest_snapshot,
            x="category",
            y="production_ton",
            color="category",
            title=f'{feed_prod_df["month"].max().strftime("%Y-%m")} 성장단계별 생산량 비교',
            text_auto=".0f",
            color_discrete_map=CATEGORY_COLORS
        )
        fig.update_traces(
            texttemplate="%{y:,.0f}",
            hovertemplate="구분 %{x}<br>생산량 %{y:,.0f} 톤<extra></extra>"
        )
        fig.update_layout(
            height=400,
            xaxis_title="사료 구분",
            yaxis_title="생산량(톤)",
            showlegend=False
        )
        apply_light_plot(fig)
        st.plotly_chart(fig, width="stretch")

with tab4:
    st.subheader("원본데이터")

    st.markdown("### 배합사료 가격")
    show_price = feed_price_df.copy()
    show_price["month"] = show_price["month"].dt.strftime("%Y-%m")
    st.dataframe(show_price[["month", "price_krw_per_kg"]], width="stretch")

    st.markdown("### 배합사료 생산량")
    show_prod = feed_prod_df.copy()
    show_prod["month"] = show_prod["month"].dt.strftime("%Y-%m")
    st.dataframe(show_prod[["month", "category", "production_ton"]], width="stretch")

    if corn_df is not None and not corn_df.empty:
        st.markdown("### 옥수수")
        st.dataframe(corn_df[["date_label", "price"]].tail(20), width="stretch")

    if soy_df is not None and not soy_df.empty:
        st.markdown("### 대두박")
        st.dataframe(soy_df[["date_label", "price"]].tail(20), width="stretch")

    if fx_df is not None and not fx_df.empty:
        st.markdown("### 환율")
        st.dataframe(fx_df[["date_label", "price"]].tail(20), width="stretch")

    if freight_df is not None and not freight_df.empty:
        st.markdown("### 해상운임")
        st.dataframe(freight_df[["date_label", "price"]].tail(20), width="stretch")

# force redeploy
st.markdown("---")
st.caption("배합사료 가격·생산량은 농림축산식품부 최근통계자료를 기준으로 정리하는 구조입니다. [Source](https://www.mafra.go.kr/home/5102/subview.do)")
