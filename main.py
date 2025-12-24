import streamlit as st
import pandas as pd

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from pathlib import Path
import unicodedata
import io

# =========================
# 페이지 설정
# =========================
st.set_page_config(
    page_title="🌱 극지식물 최적 EC 농도 연구",
    layout="wide"
)

# =========================
# 한글 폰트 (Streamlit + Plotly)
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR&display=swap');
html, body, [class*="css"] {
    font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}
</style>
""", unsafe_allow_html=True)

PLOTLY_FONT = dict(
    family="Malgun Gothic, Apple SD Gothic Neo, Noto Sans KR, sans-serif"
)

# =========================
# 한글 파일명 안전 처리
# =========================
def normalize(text: str) -> str:
    return unicodedata.normalize("NFC", text)

def find_file(directory: Path, target_name: str):
    target = normalize(target_name)
    for file in directory.iterdir():
        if normalize(file.name) == target:
            return file
    return None

# =========================
# 데이터 로딩 (캐시)
# =========================
@st.cache_data
def load_env_data(data_dir: Path):
    result = {}
    for file in data_dir.iterdir():
        if file.suffix.lower() == ".csv":
            school = file.stem.replace("_환경데이터", "")
            result[school] = pd.read_csv(file)
    return result

@st.cache_data
def load_growth_data(xlsx_path: Path):
    xls = pd.ExcelFile(xlsx_path)
    return {sheet: pd.read_excel(xlsx_path, sheet_name=sheet) for sheet in xls.sheet_names}

# =========================
# 경로 설정
# =========================
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"

with st.spinner("📂 데이터 로딩 중..."):
    env_data = load_env_data(DATA_DIR)
    growth_file = find_file(DATA_DIR, "4개교_생육결과데이터.xlsx")
    growth_data = load_growth_data(growth_file) if growth_file else {}

if not env_data or not growth_data:
    st.error("❌ 데이터 파일을 찾을 수 없습니다. data 폴더를 확인하세요.")
    st.stop()

# =========================
# EC 정보
# =========================
EC_INFO = {
    "송도고": 1.0,
    "하늘고": 2.0,
    "아라고": 4.0,
    "동산고": 8.0,
}

# =========================
# 사이드바
# =========================
st.sidebar.header("🔍 학교 선택")
selected_school = st.sidebar.selectbox(
    "학교",
    ["전체"] + list(EC_INFO.keys())
)

# =========================
# 제목
# =========================
st.title("🌱 극지식물 최적 EC 농도 연구")

tab1, tab2, tab3 = st.tabs(["📖 실험 개요", "🌡️ 환경 데이터", "📊 생육 결과"])

# =========================
# TAB 1
# =========================
with tab1:
    st.subheader("연구 목적")
    st.write(
        "서로 다른 EC 조건에서 극지식물의 생육 반응을 비교하여 "
        "**최적 EC 농도(2.0)**를 도출한다."
    )

    overview = []
    for school, ec in EC_INFO.items():
        overview.append([school, ec, len(growth_data[school])])

    df_overview = pd.DataFrame(
        overview,
        columns=["학교", "EC 목표", "개체수"]
    )
    st.dataframe(df_overview, use_container_width=True)

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("총 개체수", sum(df_overview["개체수"]))
    col2.metric("평균 온도", f"{pd.concat(env_data.values())['temperature'].mean():.1f} ℃")
    col3.metric("평균 습도", f"{pd.concat(env_data.values())['humidity'].mean():.1f} %")
    col4.metric("최적 EC", "2.0 (하늘고) ⭐")

# =========================
# TAB 2
# =========================
with tab2:
    st.subheader("환경 데이터 비교")

    rows = []
    for school, df in env_data.items():
        rows.append([
            school,
            df["temperature"].mean(),
            df["humidity"].mean(),
            df["ph"].mean(),
            df["ec"].mean(),
            EC_INFO[school]
        ])

    avg_df = pd.DataFrame(
        rows,
        columns=["학교", "온도", "습도", "pH", "실측 EC", "목표 EC"]
    )

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=["평균 온도", "평균 습도", "평균 pH", "EC 비교"]
    )

    fig.add_bar(x=avg_df["학교"], y=avg_df["온도"], row=1, col=1)
    fig.add_bar(x=avg_df["학교"], y=avg_df["습도"], row=1, col=2)
    fig.add_bar(x=avg_df["학교"], y=avg_df["pH"], row=2, col=1)
    fig.add_bar(x=avg_df["학교"], y=avg_df["실측 EC"], name="실측", row=2, col=2)
    fig.add_bar(x=avg_df["학교"], y=avg_df["목표 EC"], name="목표", row=2, col=2)

    fig.update_layout(height=700, font=PLOTLY_FONT)
    st.plotly_chart(fig, use_container_width=True)

# =========================
# TAB 3
# =========================
with tab3:
    st.subheader("EC별 생육 결과")

    summary = []
    for school, df in growth_data.items():
        summary.append([
            school,
            EC_INFO[school],
            df["생중량(g)"].mean(),
            df["잎 수(장)"].mean(),
            df["지상부 길이(mm)"].mean(),
            len(df)
        ])

    gs = pd.DataFrame(
        summary,
        columns=["학교", "EC", "평균 생중량", "평균 잎 수", "평균 지상부 길이", "개체수"]
    )

    best = gs.loc[gs["평균 생중량"].idxmax()]
    st.metric("🥇 최고 생중량", f"{best['평균 생중량']:.2f} g", f"EC {best['EC']}")

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=["생중량", "잎 수", "지상부 길이", "개체수"]
    )

    fig.add_bar(x=gs["EC"], y=gs["평균 생중량"], row=1, col=1)
    fig.add_bar(x=gs["EC"], y=gs["평균 잎 수"], row=1, col=2)
    fig.add_bar(x=gs["EC"], y=gs["평균 지상부 길이"], row=2, col=1)
    fig.add_bar(x=gs["EC"], y=gs["개체수"], row=2, col=2)

    fig.update_layout(height=700, font=PLOTLY_FONT)
    st.plotly_chart(fig, use_container_width=True)



