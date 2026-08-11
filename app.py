import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from vibration_analyzer import calculate_metrics, perform_fft, evaluate_status, generate_site_sample_data

st.set_page_config(page_title="사업소별 진동 측정 데이터 분석", layout="wide")

st.title("🏭 사업소별 진동 측정 데이터 분석 대시보드")
st.markdown("---")

# 사이드바 설정
st.sidebar.header("⚙️ 설정 및 사업소 선택")
data_mode = st.sidebar.radio("데이터 불러오기 방식", ["시뮬레이션 (사업소 멀티 샘플)", "CSV 파일 직접 업로드"])

sampling_rate = st.sidebar.number_input("샘플링 주파수 (Hz)", value=1000, step=100)

site_data_dict = {}

if data_mode == "시뮬레이션 (사업소 멀티 샘플)":
    # 사업소별 샘플 데이터 자동 생성
    site_data_dict = generate_site_sample_data(sampling_rate)
    selected_site = st.sidebar.selectbox("📍 조회할 사업소 선택", list(site_data_dict.keys()))
    df_selected = site_data_dict[selected_site]
    signal_col = "Vibration"

else:
    # CSV 파일 직접 업로드 (사업소 구분 컬럼이 있거나, 사업소명을 직접 입력)
    uploaded_file = st.sidebar.file_uploader("사업소 진동 데이터 CSV 업로드", type=["csv"])
    selected_site = st.sidebar.text_input("사업소/설비명 입력", value="사업소 A")
    
    if uploaded_file is not None:
        df_selected = pd.read_csv(uploaded_file)
        signal_col = st.sidebar.selectbox("진동 데이터 컬럼 선택", df_selected.columns)
    else:
        st.info("CSV 파일을 업로드해 주세요.")
        st.stop()

# --- 1. 전체 사업소 요약 현황판 (시뮬레이션 모드일 때) ---
if data_mode == "시뮬레이션 (사업소 멀티 샘플)":
    st.subheader("📊 사업소별 전체 상태 현황")
    summary_list = []
    
    for site_name, s_df in site_data_dict.items():
        m = calculate_metrics(s_df["Vibration"].values)
        st_text, _ = evaluate_status(m["RMS"])
        summary_list.append({
            "사업소명": site_name,
            "RMS (mm/s)": m["RMS"],
            "Peak (mm/s)": m["Peak"],
            "Crest Factor": m["Crest_Factor"],
            "상태": st_text
        })
    
    summary_df = pd.DataFrame(summary_list)
    
    # 사업소 요약 메트릭 카드 배치
    cols = st.columns(len(site_data_dict))
    for idx, row in summary_df.iterrows():
        cols[idx].metric(label=row["사업소명"], value=f"{row['RMS (mm/s)']} mm/s", delta=row["상태"])
    
    st.markdown("---")

# --- 2. 선택한 사업소 상세 분석 ---
st.subheader(f"🔍 [{selected_site}] 상세 진동 분석")

signal_data = df_selected[signal_col].values
metrics = calculate_metrics(signal_data)
status_text, status_type = evaluate_status(metrics["RMS"])
fft_df = perform_fft(signal_data, sampling_rate)

# 상세 지표 표출
m_col1, m_col2, m_col3, m_col4 = st.columns(4)
m_col1.metric("RMS (실효값)", f"{metrics['RMS']} mm/s")
m_col2.metric("Peak (최대값)", f"{metrics['Peak']} mm/s")
m_col3.metric("Crest Factor", f"{metrics['Crest_Factor']}")

if status_type == "success":
    m_col4.success(status_text)
elif status_type == "warning":
    m_col4.warning(status_text)
else:
    m_col4.error(status_text)

# 차트 탭
tab1, tab2, tab3 = st.tabs(["📈 시계열 파형", "🌊 FFT 주파수 스펙트럼", "📋 원본 데이터"])

with tab1:
    fig_time = px.line(df_selected, y=signal_col, title=f"{selected_site} - Time Domain Signal")
    fig_time.update_traces(line_color="#1f77b4")
    st.plotly_chart(fig_time, use_container_width=True)

with tab2:
    fig_fft = px.line(fft_df, x="Frequency_Hz", y="Amplitude", title=f"{selected_site} - FFT Spectrum")
    fig_fft.update_traces(line_color="#ff7f0e")
    st.plotly_chart(fig_fft, use_container_width=True)

with tab3:
    st.dataframe(df_selected, use_container_width=True)