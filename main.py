import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from pathlib import Path
import unicodedata
import io

# =========================
# 기본 설정
# =========================
st.set_page_config(
    page_title="🌱 극지식물 최적 EC 농도 연구",
    layout="wide"
)

# 한글 폰트 (Streamlit + Plotly)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR&display=swap');
html, body, [class*="css"] {
    font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}
</style>
""", unsafe_allow_html=True)

PLOTLY_FONT = dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif")

# =========================
# 유틸: 한글 파일명 안전 로딩
# =========================
def normalize(text):
    return unicodedata.normalize("NFC", text)

def find_file_by_name(directory: Path, target_name: str):
    target_norm = normalize(target_name)
    for p in directory.iterdir():
        if normalize(p.name) == target_norm:
            return p
    return None

# =========================
# 데이터 로딩
# =========================
@st.cache_data
def load_environment_data(data_dir: Path):
    env_data = {}
    for file in data_dir.iterdir():
        if file.suffix.lower() == ".csv":
            try:
                df = pd.read_csv(file)
                school = file.stem.replace("_환경데이터", "")
                env_data[school] = df
            except Exception as e:
                st.error(f"{file.name} 로딩 실패: {e}")
    return env_data

@st.cache_data
def load_growth_data(xlsx_path: Path):
    try:
        xls = pd.ExcelFile(xlsx_path)
        growth = {}
        for sheet in xls.sheet_names:
            growth[sheet] = pd.read_excel(xlsx_path, sheet_name=sheet)
        return growth
    except Exception as e:
        st.error(f"생육 데이터 로딩 실패: {e}")
        return {}

# =========================
# 파일 경로
# =========================
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"

with st.spinner("데이터 로딩 중..."):
    env_data = load_environment_data(DATA_DIR)
    growth_file = find_file_by_name(DATA_DIR, "4개교_생육결과데이터.xlsx")
    growth_data = load_growth_data(growth_file) if growth_file else {}

if not env_data or not growth_data:
    st.error("필요한 데이터 파일을 찾을 수 없습니다.")
    st.stop()

# =========================
# 메타 정보
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
schools = ["전체"] + list(EC_INFO.keys())
selected_school = st.sidebar.selectbox("학교 선택", schools)

# =========================
# 제목
# =========================
st.title("🌱 극지식물 최적 EC 농도 연구")

tab1, tab2, tab3 = st.tabs(["📖 실험 개요", "🌡️ 환경 데이터", "📊 생육 결과"])

# =========================
# TAB 1: 실험 개요
# =========================
with tab1:
    st.subheader("연구 배경 및 목적")
    st.write(
        "본 연구는 서로 다른 EC 조건에서 극지식물의 생육 반응을 비교하여 "
        "**최적 EC 농도**를 도출하는 것을 목표로 한다."
    )

    summary_rows = []
    for school, ec in EC_INFO.items():
        count = len(growth_data.get(school, []))
        summary_rows.append([school, ec, count])

    summary_df = pd.DataFrame(
        summary_rows,
        columns=["학교명", "EC 목표", "개체수"]
    )
    st.dataframe(summary_df, use_container_width=True)

    total_plants = sum(len(df) for df in growth_data.values())
    avg_temp = pd.concat(env_data.values())["temperature"].mean()
    avg_hum = pd.concat(env_data.values())["humidity"].mean()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("총 개체수", total_plants)
    col2.metric("평균 온도 (℃)", f"{avg_temp:.1f}")
    col3.metric("평균 습도 (%)", f"{avg_hum:.1f}")
    col4.metric("최적 EC", "2.0 (하늘고)")

# =========================
# TAB 2: 환경 데이터
# =========================
with tab2:
    st.subheader("학교별 환경 데이터 비교")

    env_avg = []
    for school, df in env_data.items():
        env_avg.append([
            school,
            df["temperature"].mean(),
            df["humidity"].mean(),
            df["ph"].mean(),
            df["ec"].mean(),
            EC_INFO.get(school, None)
        ])

    avg_df = pd.DataFrame(
        env_avg,
        columns=["학교", "온도", "습도", "pH", "실측 EC", "목표 EC"]
    )

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=[
            "평균 온도", "평균 습도", "평균 pH", "목표 EC vs 실측 EC"
        ]
    )

    fig.add_bar(x=avg_df["학교"], y=avg_df["온도"], row=1, col=1)
    fig.add_bar(x=avg_df["학교"], y=avg_df["습도"], row=1, col=2)
    fig.add_bar(x=avg_df["학교"], y=avg_df["pH"], row=2, col=1)
    fig.add_bar(x=avg_df["학교"], y=avg_df["실측 EC"], name="실측", row=2, col=2)
    fig.add_bar(x=avg_df["학교"], y=avg_df["목표 EC"], name="목표", row=2, col=2)

    fig.update_layout(height=700, font=PLOTLY_FONT)
    st.plotly_chart(fig, use_container_width=True)

    if selected_school != "전체":
        df = env_data[selected_school]

        fig2 = px.line(df, x="time", y=["temperature", "humidity", "ec"])
        fig2.add_hline(
            y=EC_INFO[selected_school],
            line_dash="dash",
            annotation_text="목표 EC"
        )
        fig2.update_layout(font=PLOTLY_FONT)
        st.plotly_chart(fig2, use_container_width=True)

    with st.expander("📂 환경 데이터 원본"):
        combined_env = pd.concat(env_data, names=["학교"])
        st.dataframe(combined_env, use_container_width=True)

        buffer = io.BytesIO()
        combined_env.to_csv(buffer, index=False, encoding="utf-8-sig")
        buffer.seek(0)
        st.download_button(
            "CSV 다운로드",
            data=buffer,
            file_name="환경데이터_전체.csv",
            mime="text/csv"
        )

# =========================
# TAB 3: 생육 결과
# =========================
with tab3:
    st.subheader("EC별 생육 결과 분석")

    growth_summary = []
    for school, df in growth_data.items():
        growth_summary.append([
            school,
            EC_INFO.get(school),
            df["생중량(g)"].mean(),
            df["잎 수(장)"].mean(),
            df["지상부 길이(mm)"].mean(),
            len(df)
        ])

    gs_df = pd.DataFrame(
        growth_summary,
        columns=["학교", "EC", "평균 생중량", "평균 잎 수", "평균 지상부 길이", "개체수"]
    )

    best = gs_df.loc[gs_df["평균 생중량"].idxmax()]

    st.metric(
        "🥇 최고 평균 생중량",
        f"{best['평균 생중량']:.2f} g",
        f"EC {best['EC']}"
    )

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=[
            "평균 생중량", "평균 잎 수",
            "평균 지상부 길이", "개체수"
        ]
    )

    fig.add_bar(x=gs_df["EC"], y=gs_df["평균 생중량"], row=1, col=1)
    fig.add_bar(x=gs_df["EC"], y=gs_df["평균 잎 수"], row=1, col=2)
    fig.add_bar(x=gs_df["EC"], y=gs_df["평균 지상부 길이"], row=2, col=1)
    fig.add_bar(x=gs_df["EC"], y=gs_df["개체수"], row=2, col=2)

    fig.update_layout(height=700, font=PLOTLY_FONT)
    st.plotly_chart(fig, use_container_width=True)

    all_growth = pd.concat(growth_data, names=["학교"])
    fig_box = px.box(
        all_growth.reset_index(),
        x="학교",
        y="생중량(g)",
        color="학교"
    )
    fig_box.update_layout(font=PLOTLY_FONT)
    st.plotly_chart(fig_box, use_container_width=True)

    fig_corr1 = px.scatter(
        all_growth,
        x="잎 수(장)",
        y="생중량(g)",
        trendline="ols"
    )
    fig_corr2 = px.scatter(
        all_growth,
        x="지상부 길이(mm)",
        y="생중량(g)",
        trendline="ols"
    )
    fig_corr1.update_layout(font=PLOTLY_FONT)
    fig_corr2.update_layout(font=PLOTLY_FONT)

    col1, col2 = st.columns(2)
    col1.plotly_chart(fig_corr1, use_container_width=True)
    col2.plotly_chart(fig_corr2, use_container_width=True)

    with st.expander("📂 생육 데이터 원본"):
        st.dataframe(all_growth, use_container_width=True)

        buffer = io.BytesIO()
        all_growth.to_excel(buffer, engine="openpyxl")
        buffer.seek(0)
        st.download_button(
            "XLSX 다운로드",
            data=buffer,
            file_name="생육결과_전체.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from pathlib import Path
import unicodedata
import io

# =========================
# 기본 설정
# =========================
st.set_page_config(
    page_title="🌱 극지식물 최적 EC 농도 연구",
    layout="wide"
)

# 한글 폰트 (Streamlit + Plotly)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR&display=swap');
html, body, [class*="css"] {
    font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}
</style>
""", unsafe_allow_html=True)

PLOTLY_FONT = dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif")

# =========================
# 유틸: 한글 파일명 안전 로딩
# =========================
def normalize(text):
    return unicodedata.normalize("NFC", text)

def find_file_by_name(directory: Path, target_name: str):
    target_norm = normalize(target_name)
    for p in directory.iterdir():
        if normalize(p.name) == target_norm:
            return p
    return None

# =========================
# 데이터 로딩
# =========================
@st.cache_data
def load_environment_data(data_dir: Path):
    env_data = {}
    for file in data_dir.iterdir():
        if file.suffix.lower() == ".csv":
            try:
                df = pd.read_csv(file)
                school = file.stem.replace("_환경데이터", "")
                env_data[school] = df
            except Exception as e:
                st.error(f"{file.name} 로딩 실패: {e}")
    return env_data

@st.cache_data
def load_growth_data(xlsx_path: Path):
    try:
        xls = pd.ExcelFile(xlsx_path)
        growth = {}
        for sheet in xls.sheet_names:
            growth[sheet] = pd.read_excel(xlsx_path, sheet_name=sheet)
        return growth
    except Exception as e:
        st.error(f"생육 데이터 로딩 실패: {e}")
        return {}

# =========================
# 파일 경로
# =========================
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"

with st.spinner("데이터 로딩 중..."):
    env_data = load_environment_data(DATA_DIR)
    growth_file = find_file_by_name(DATA_DIR, "4개교_생육결과데이터.xlsx")
    growth_data = load_growth_data(growth_file) if growth_file else {}

if not env_data or not growth_data:
    st.error("필요한 데이터 파일을 찾을 수 없습니다.")
    st.stop()

# =========================
# 메타 정보
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
schools = ["전체"] + list(EC_INFO.keys())
selected_school = st.sidebar.selectbox("학교 선택", schools)

# =========================
# 제목
# =========================
st.title("🌱 극지식물 최적 EC 농도 연구")

tab1, tab2, tab3 = st.tabs(["📖 실험 개요", "🌡️ 환경 데이터", "📊 생육 결과"])

# =========================
# TAB 1: 실험 개요
# =========================
with tab1:
    st.subheader("연구 배경 및 목적")
    st.write(
        "본 연구는 서로 다른 EC 조건에서 극지식물의 생육 반응을 비교하여 "
        "**최적 EC 농도**를 도출하는 것을 목표로 한다."
    )

    summary_rows = []
    for school, ec in EC_INFO.items():
        count = len(growth_data.get(school, []))
        summary_rows.append([school, ec, count])

    summary_df = pd.DataFrame(
        summary_rows,
        columns=["학교명", "EC 목표", "개체수"]
    )
    st.dataframe(summary_df, use_container_width=True)

    total_plants = sum(len(df) for df in growth_data.values())
    avg_temp = pd.concat(env_data.values())["temperature"].mean()
    avg_hum = pd.concat(env_data.values())["humidity"].mean()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("총 개체수", total_plants)
    col2.metric("평균 온도 (℃)", f"{avg_temp:.1f}")
    col3.metric("평균 습도 (%)", f"{avg_hum:.1f}")
    col4.metric("최적 EC", "2.0 (하늘고)")

# =========================
# TAB 2: 환경 데이터
# =========================
with tab2:
    st.subheader("학교별 환경 데이터 비교")

    env_avg = []
    for school, df in env_data.items():
        env_avg.append([
            school,
            df["temperature"].mean(),
            df["humidity"].mean(),
            df["ph"].mean(),
            df["ec"].mean(),
            EC_INFO.get(school, None)
        ])

    avg_df = pd.DataFrame(
        env_avg,
        columns=["학교", "온도", "습도", "pH", "실측 EC", "목표 EC"]
    )

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=[
            "평균 온도", "평균 습도", "평균 pH", "목표 EC vs 실측 EC"
        ]
    )

    fig.add_bar(x=avg_df["학교"], y=avg_df["온도"], row=1, col=1)
    fig.add_bar(x=avg_df["학교"], y=avg_df["습도"], row=1, col=2)
    fig.add_bar(x=avg_df["학교"], y=avg_df["pH"], row=2, col=1)
    fig.add_bar(x=avg_df["학교"], y=avg_df["실측 EC"], name="실측", row=2, col=2)
    fig.add_bar(x=avg_df["학교"], y=avg_df["목표 EC"], name="목표", row=2, col=2)

    fig.update_layout(height=700, font=PLOTLY_FONT)
    st.plotly_chart(fig, use_container_width=True)

    if selected_school != "전체":
        df = env_data[selected_school]

        fig2 = px.line(df, x="time", y=["temperature", "humidity", "ec"])
        fig2.add_hline(
            y=EC_INFO[selected_school],
            line_dash="dash",
            annotation_text="목표 EC"
        )
        fig2.update_layout(font=PLOTLY_FONT)
        st.plotly_chart(fig2, use_container_width=True)

    with st.expander("📂 환경 데이터 원본"):
        combined_env = pd.concat(env_data, names=["학교"])
        st.dataframe(combined_env, use_container_width=True)

        buffer = io.BytesIO()
        combined_env.to_csv(buffer, index=False, encoding="utf-8-sig")
        buffer.seek(0)
        st.download_button(
            "CSV 다운로드",
            data=buffer,
            file_name="환경데이터_전체.csv",
            mime="text/csv"
        )

# =========================
# TAB 3: 생육 결과
# =========================
with tab3:
    st.subheader("EC별 생육 결과 분석")

    growth_summary = []
    for school, df in growth_data.items():
        growth_summary.append([
            school,
            EC_INFO.get(school),
            df["생중량(g)"].mean(),
            df["잎 수(장)"].mean(),
            df["지상부 길이(mm)"].mean(),
            len(df)
        ])

    gs_df = pd.DataFrame(
        growth_summary,
        columns=["학교", "EC", "평균 생중량", "평균 잎 수", "평균 지상부 길이", "개체수"]
    )

    best = gs_df.loc[gs_df["평균 생중량"].idxmax()]

    st.metric(
        "🥇 최고 평균 생중량",
        f"{best['평균 생중량']:.2f} g",
        f"EC {best['EC']}"
    )

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=[
            "평균 생중량", "평균 잎 수",
            "평균 지상부 길이", "개체수"
        ]
    )

    fig.add_bar(x=gs_df["EC"], y=gs_df["평균 생중량"], row=1, col=1)
    fig.add_bar(x=gs_df["EC"], y=gs_df["평균 잎 수"], row=1, col=2)
    fig.add_bar(x=gs_df["EC"], y=gs_df["평균 지상부 길이"], row=2, col=1)
    fig.add_bar(x=gs_df["EC"], y=gs_df["개체수"], row=2, col=2)

    fig.update_layout(height=700, font=PLOTLY_FONT)
    st.plotly_chart(fig, use_container_width=True)

    all_growth = pd.concat(growth_data, names=["학교"])
    fig_box = px.box(
        all_growth.reset_index(),
        x="학교",
        y="생중량(g)",
        color="학교"
    )
    fig_box.update_layout(font=PLOTLY_FONT)
    st.plotly_chart(fig_box, use_container_width=True)

    fig_corr1 = px.scatter(
        all_growth,
        x="잎 수(장)",
        y="생중량(g)",
        trendline="ols"
    )
    fig_corr2 = px.scatter(
        all_growth,
        x="지상부 길이(mm)",
        y="생중량(g)",
        trendline="ols"
    )
    fig_corr1.update_layout(font=PLOTLY_FONT)
    fig_corr2.update_layout(font=PLOTLY_FONT)

    col1, col2 = st.columns(2)
    col1.plotly_chart(fig_corr1, use_container_width=True)
    col2.plotly_chart(fig_corr2, use_container_width=True)

    with st.expander("📂 생육 데이터 원본"):
        st.dataframe(all_growth, use_container_width=True)

        buffer = io.BytesIO()
        all_growth.to_excel(buffer, engine="openpyxl")
        buffer.seek(0)
        st.download_button(
            "XLSX 다운로드",
            data=buffer,
            file_name="생육결과_전체.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

