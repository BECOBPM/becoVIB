import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# 진동 분석 모듈 모듈 불러오기
from vibration_analyzer import calculate_metrics, perform_fft, evaluate_status

st.set_page_config(page_title="진동 측정 데이터 분석 대시보드", layout="wide")

st.title("⚙️ 진동 측정 데이터 분석 대시보드")
st.markdown("---")

# 사이드바: 데이터 업로드 및 샘플 실행
st.sidebar.header("📂 데이터 입력")
uploaded_file = st.sidebar.file_uploader("진동 데이터 CSV 파일 업로드", type=["csv"])
sampling_rate = st.sidebar.number_input("샘플링 주파수 (Hz)", value=1000, step=100)

use_sample = st.sidebar.checkbox("테스트용 샘플 데이터 사용", value=(uploaded_file is None))

# 데이터 로드
if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    signal_col = st.sidebar.selectbox("분석할 진동 데이터 컬럼 선택", df.columns)
    signal_data = df[signal_col].values
elif use_sample:
    # 샘플 진동 데이터 생성 (60Hz, 120Hz 성분 + 노이즈)
    t = np.linspace(0, 1, sampling_rate, endpoint=False)
    signal_data = 1.8 * np.sin(2 * np.pi * 60 * t) + 0.9 * np.sin(2 * np.pi * 120 * t) + np.random.normal(0, 0.4, sampling_rate)
    df = pd.DataFrame({"Time_s": t, "Vibration": signal_data})
    signal_col = "Vibration"
else:
    st.info("CSV 파일을 업로드하거나 샘플 데이터 사용을 체크해 주세요.")
    st.stop()

# 1. 지표 계산 및 상태 평가
metrics = calculate_metrics(signal_data)
status_text, status_type = evaluate_status(metrics["RMS"])
fft_df = perform_fft(signal_data, sampling_rate)

# 2. 상단 KPI 메트릭 카드
col1, col2, col3, col4 = st.columns(4)
col1.metric("RMS (진동 실효값)", f"{metrics['RMS']} mm/s")
col2.metric("Peak (최대 피크값)", f"{metrics['Peak']} mm/s")
col3.metric("Crest Factor", f"{metrics['Crest_Factor']}")

if status_type == "success":
    col4.success(status_text)
elif status_type == "warning":
    col4.warning(status_text)
else:
    col4.error(status_text)

st.markdown("---")

# 3. 차트 영역 (시계열 & FFT 주파수 스펙트럼)
tab1, tab2, tab3 = st.tabs(["📈 시계열 파형 (Time Domain)", "🌊 주파수 스펙트럼 (FFT)", "📋 데이터 테이블"])

with tab1:
    st.subheader("시간 영역 진동 파형")
    fig_time = px.line(df, y=signal_col, labels={"index": "Sample Point", signal_col: "Amplitude"}, title="Time Domain Signal")
    fig_time.update_traces(line_color="#1f77b4")
    st.plotly_chart(fig_time, use_container_width=True)

with tab2:
    st.subheader("주파수 영역 분석 (FFT Spectrum)")
    fig_fft = px.line(fft_df, x="Frequency_Hz", y="Amplitude", labels={"Frequency_Hz": "Frequency (Hz)", "Amplitude": "Amplitude"}, title="Frequency Domain (FFT Spectrum)")
    fig_fft.update_traces(line_color="#ff7f0e")
    st.plotly_chart(fig_fft, use_container_width=True)

with tab3:
    st.subheader("원본 데이터 확인")
    st.dataframe(df, use_container_width=True)