import streamlit as st
import pandas as pd

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from pathlib import Path
import unicodedata
import io

# ==================================================
# 페이지 설정
# ==================================================
st.set_page_config(
    page_title="🌱 극지식물 최적 EC 농도 연구",
    layout="wide"
)

# ==================================================
# 한글 폰트 (Streamlit)
# ==================================================
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

# ==================================================
# 한글 파일명 안전 처리 (NFC/NFD)
# ==================================================
def normalize(text: str) -> str:
    return unicodedata.normalize("NFC", text)

def find_file(directory: Path, target_name: str):
    target = normalize(target_name)
    for file in directory.iterdir():
        if normalize(file.name) == target:
            return file
    return None

# ==================================================
# 데이터 로딩 (캐시)
# ==================================================
@st.cache_data
def load_environment_data(data_dir: Path):
    env = {}
    for file in data_dir.iterdir():
        if file.suffix.lower() == ".csv":
            school = file.stem.replace("_환경데이터", "")
            env[school] = pd.read_csv(file)
    return env

@st.cache_data
def load_growth_data(xlsx_path: Path):
    xls = pd.ExcelFile(xlsx_path)
    return {
        sheet: pd.read_excel(xlsx_path, sheet_name=sheet)
        for sheet in xls.sheet_names
    }

# ==================================================
# 경로 설정
# ==================================================
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"

with st.spinner("📂 데이터 로딩 중..."):
    env_data = load_environment_data(DATA_DIR)
    growth_file = find_file(DATA_DIR, "4개교_생육결과데이터.xlsx")
    growth_data = load_growth_data(growth_file) if growth_file else {}

if not env_data or not growth_data:
    st.error("❌ 데이터 파일을 찾을 수 없습니다. data 폴더를 확인하세요.")
    st.stop()

# ==================================================
# EC 조건 정보
# ==================================================
EC_INFO = {
    "송도고": 1.0,
    "하늘고": 2.0,   # 최적
    "아라고": 4.0,
    "동산고": 8.0,
}

# ==================================================
# 사이드바
# ==================================================
st.sidebar.header("🔍 학교 선택")
selected_school = st.sidebar.selectbox(
    "학교",
    ["전체"] + list(EC_INFO.keys())
)

# ==================================================
# 제목
# ==================================================
st.title("🌱 극지식물 최적 EC 농도 연구 대시보드")

tab1, tab2, tab3 = st.tabs([
    "📈 pH–EC 상관관계",
    "🌿 EC별 생중량 비교",
    "📝 실험 결과 해석"
])

# ==================================================
# TAB 1: pH와 EC 상관관계 (statsmodels 미사용)
# ==================================================
with tab1:
    st.subheader("pH와 EC의 관계")

    env_list = []
    for school, df in env_data.items():
        temp = df.copy()
        temp["학교"] = school
        env_list.append(temp)

    env_df = pd.concat(env_list)

    fig_scatter = px.scatter(
        env_df,
        x="ph",
        y="ec",
        color="학교",
        labels={"ph": "pH", "ec": "EC"}
    )
    fig_scatter.update_layout(font=PLOTLY_FONT)

    st.plotly_chart(fig_scatter, use_container_width=True)

    st.markdown("""
    **해석**  
    pH와 EC 값은 일정 범위 내에서 분포하지만, 뚜렷한 선형 상관관계보다는  
    각 학교별 EC 조건 차이가 더 크게 작용함을 확인할 수 있다.
    """)

# ==================================================
# TAB 2: EC별 생중량 비교
# ==================================================
with tab2:
    st.subheader("학교별 EC 조건에 따른 평균 생중량")

    summary = []
    for school, df in growth_data.items():
        summary.append([
            school,
            EC_INFO.get(school),
            df["생중량(g)"].mean()
        ])

    summary_df = pd.DataFrame(
        summary,
        columns=["학교", "EC", "평균 생중량"]
    )

    fig_bar = px.bar(
        summary_df,
        x="EC",
        y="평균 생중량",
        color="학교",
        text_auto=".2f"
    )

    # 최적 EC 강조 (하늘고, EC 2.0)
    fig_bar.add_vline(
        x=2.0,
        line_dash="dash",
        line_color="red",
        annotation_text="최적 EC (2.0)"
    )

    fig_bar.update_layout(font=PLOTLY_FONT)
    st.plotly_chart(fig_bar, use_container_width=True)

# ==================================================
# TAB 3: 실험 결과 글로 제시
# ==================================================
with tab3:
    st.subheader("실험 결과 및 결론")

    st.markdown("""
    ### 🔬 실험 결과 요약

    본 연구에서는 서로 다른 EC 농도 조건(1.0, 2.0, 4.0, 8.0)에서  
    극지식물의 생육 반응을 비교 분석하였다.

    - **EC 2.0(하늘고)** 조건에서 평균 생중량이 가장 높게 나타났다.
    - EC가 낮거나(1.0) 과도하게 높을 경우(4.0, 8.0),
      양분 흡수 효율 저하 또는 염류 스트레스로 인해 생육이 감소하였다.
    - pH와 EC의 분포를 분석한 결과,
      생중량에 가장 큰 영향을 미친 요인은 **EC 농도**임을 확인하였다.

    ### ✅ 결론

    극지식물의 안정적인 생육을 위해서는  
    **EC 2.0 수준의 양액 농도가 가장 적절한 조건**으로 판단된다.
    """)

    # 요약 데이터 다운로드
    buffer = io.BytesIO()
    summary_df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        label="📥 EC별 생육 결과 요약 XLSX 다운로드",
        data=buffer,
        file_name="EC별_생육결과_요약.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



