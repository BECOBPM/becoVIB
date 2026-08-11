import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os

# 1. 페이지 설정
st.set_page_config(
    page_title="부산환경공단 회전기기 진동 관리 시스템",
    page_icon="⚙️",
    layout="wide"
)

# 2. 이미지 기준 전체 사업소/사업단 목록
SITE_LIST = [
    "수영사업단", "강변사업단", "남부사업소", "녹산사업소", "기장사업소",
    "동부사업소", "중앙사업소", "영도사업소", "정관사업소", "서부사업소",
    "관로사업소", "하수자원사업소", "위생사업소", "해운대사업단", "생곡사업단",
    "명지사업소", "에너지사업소", "대기환경사업소"
]

# GitHub 레포지토리에 업로드된 기본 엑셀 파일명
DEFAULT_EXCEL_FILE = "설비점검 및 정비 관리대장(에너지사업소).xlsx"

# 3. 사이드바 구성
with st.sidebar:
    st.title("📌 메뉴 (Navigation)")
    
    menu_type = st.radio(
        "구분 선택",
        [
            "🏢 사업소별 설비 관리", 
            "📈 진동 데이터 분석 (파형/FFT)",
            "📏 진동 측정 기준 (ISO 10816)"
        ]
    )
    
    st.markdown("---")
    
    # 18개 사업소 선택 드롭다운 (기본값: 에너지사업소)
    selected_site = st.selectbox("🏢 사업소 선택", SITE_LIST, index=16)
    
    st.markdown("---")
    
    st.subheader("📂 사업소 데이터 업로드")
    uploaded_file = st.file_uploader(f"[{selected_site}] 측정 데이터 파일(CSV/Excel)", type=["csv", "xlsx"])


# 4. 데이터 로드 및 전처리 함수
def load_vibration_data(uploaded_file, default_file):
    df = pd.DataFrame()
    
    # 1) 사용자가 직접 파일 업로드한 경우 우선 처리
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            st.sidebar.success("✅ 업로드 파일 로드 완료")
        except Exception as e:
            st.sidebar.error(f"파일 로드 실패: {e}")
            
    # 2) 업로드 파일이 없으면 레포지토리의 기본 엑셀 대장 자동 로드
    elif os.path.exists(default_file):
        try:
            df = pd.read_excel(default_file)
            st.sidebar.info(f"📁 기본 대장 파일('{default_file}') 자동 적용 중")
        except Exception as e:
            st.sidebar.warning(f"기본 파일 읽기 실패: {e}")
            
    # 3) 둘 다 없을 경우 대체용 샘플 데이터 생성
    if df.empty:
        data = {
            "측정일자": ["2026-04-02", "2026-04-02", "2026-04-01", "2026-03-28", "2026-03-25", "2026-03-20"],
            "사업소명": [selected_site] * 6,
            "설비명": ["집단에너지 FD Fan #1", "집단에너지 FD Fan #1", "보일러 급수펌프 #1", "1차 냉각수펌프 #2", "재열기 유압펌프 #1", "공기압축기 #3"],
            "전동기(kW)": [310, 310, 95, 11, 45, 37],
            "측정위치": ["1(전동기) X", "1(전동기) Y", "2(펌프) Z", "1(전동기) X", "2(펌프) Y", "1(전동기) Z"],
            "속도(mm/s)": [12.3, 2.8, 4.8, 1.1, 7.2, 1.8],
            "판정": ["D (즉시점검)", "B (양호)", "C (보수필요)", "A (양호)", "C (보수필요)", "A (양호)"],
            "조치사항": ["베어링 수선 정비", "정상운전", "구리스 보충 및 트렌드 관찰", "정상운전", "축 정렬 재점검", "정상운전"]
        }
        df = pd.DataFrame(data)
        
    return df


# 5. 메인 화면 구성
# ==========================================
# 메뉴 1: 🏢 사업소별 설비 관리
# ==========================================
if menu_type == "🏢 사업소별 설비 관리":
    st.info(f"🏢 현재 선택된 사업소: **{selected_site}**")
    st.title(f"📜 [{selected_site}] 설비별 진동 점검 이력 및 현황")
    st.caption("부산환경공단 소속 회전기기 진동 상태 측정 데이터 관리 시스템")
    
    # 엑셀 데이터 불러오기
    df = load_vibration_data(uploaded_file, DEFAULT_EXCEL_FILE)
    
    # 사업소 컬럼 필터링 (컬럼명 유연 감지)
    site_col = next((col for col in df.columns if "사업소" in col), None)
    if site_col and selected_site in df[site_col].values:
        filtered_df = df[df[site_col] == selected_site]
    else:
        filtered_df = df
        
    # 주요 수치 및 상태 컬럼 자동 감지
    speed_col = next((col for col in df.columns if "속도" in col or "진동" in col or "RMS" in col), None)
    status_col = next((col for col in df.columns if "판정" in col or "상태" in col or "결과" in col), None)
    equipment_col = next((col for col in df.columns if "설비" in col or "기기" in col), None)
    
    # 상단 요약 카운트 메트릭
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📊 전체 점검 건수", f"{len(filtered_df)} 건")
    
    if status_col and status_col in filtered_df.columns:
        good_cnt = len(filtered_df[filtered_df[status_col].astype(str).str.contains("A|B|양호|정상", na=False)])
        warning_cnt = len(filtered_df[filtered_df[status_col].astype(str).str.contains("C|주의|보수", na=False)])
        danger_cnt = len(filtered_df[filtered_df[status_col].astype(str).str.contains("D|위험|즉시|점검", na=False)])
    else:
        good_cnt, warning_cnt, danger_cnt = 0, 0, 0

    col2.metric("✅ 양호 (A / B)", f"{good_cnt} 건")
    col3.metric("⚠️ 보수 필요 (C)", f"{warning_cnt} 건")
    col4.metric("🚨 즉시 점검 (D)", f"{danger_cnt} 건")
    
    st.markdown("---")
    
    # 판정 상태별 필터링
    if status_col and status_col in filtered_df.columns:
        status_options = filtered_df[status_col].dropna().unique()
        status_filter = st.multiselect("🔍 판정 상태별 필터", options=status_options, default=status_options)
        if status_filter:
            display_df = filtered_df[filtered_df[status_col].isin(status_filter)]
        else:
            display_df = filtered_df
    else:
        display_df = filtered_df

    # 차트 시각화
    if not display_df.empty and equipment_col and speed_col:
        chart_col1, chart_col2 = st.columns([2, 1])
        
        with chart_col1:
            st.subheader("📊 설비별 진동 수치 비교")
            fig_bar = px.bar(
                display_df, 
                x=equipment_col, 
                y=speed_col, 
                color=status_col if status_col else None,
                text=speed_col,
                title=f"{selected_site} 설비별 진동 측정 데이터"
            )
            fig_bar.update_traces(textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True)
            
        with chart_col2:
            if status_col:
                st.subheader("🍩 판정 비율")
                fig_pie = px.pie(display_df, names=status_col, title="상태별 비율", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)

    # 상세 데이터 테이블
    st.subheader("📋 상세 점검 이력 대장")
    st.dataframe(display_df, use_container_width=True)


# ==========================================
# 메뉴 2: 📈 진동 데이터 분석 (파형/FFT)
# ==========================================
elif menu_type == "📈 진동 데이터 분석 (파형/FFT)":
    st.title(f"🌊 [{selected_site}] 정밀 진동 파형 & FFT 주파수 분석")
    st.caption("진동 신호의 시간 영역 파형(Time Domain) 및 주파수 영역(FFT) 분석을 수행합니다.")
    
    sampling_rate = st.number_input("샘플링 주파수 (Hz)", value=1000, step=100)
    t = np.linspace(0, 1, sampling_rate, endpoint=False)
    
    vibration_signal = 2.5 * np.sin(2 * np.pi * 60 * t) + 1.2 * np.sin(2 * np.pi * 120 * t) + np.random.normal(0, 0.3, sampling_rate)
    
    rms_val = np.sqrt(np.mean(np.square(vibration_signal)))
    peak_val = np.max(np.abs(vibration_signal))
    
    m1, m2, m3 = st.columns(3)
    m1.metric("계산된 RMS (진동 실효값)", f"{rms_val:.2f} mm/s")
    m2.metric("Peak (최대값)", f"{peak_val:.2f} mm/s")
    m3.metric("Crest Factor", f"{(peak_val/rms_val):.2f}")
    
    n = len(vibration_signal)
    fft_vals = np.fft.rfft(vibration_signal)
    fft_freq = np.fft.rfftfreq(n, d=1.0 / sampling_rate)
    amplitude = np.abs(fft_vals) * 2.0 / n
    
    tab1, tab2 = st.tabs(["📈 시간 영역 (Time Domain)", "🌊 주파수 영역 (FFT Spectrum)"])
    
    with tab1:
        fig_time = px.line(x=t, y=vibration_signal, labels={'x': '시간 (초)', 'y': '진동 가속도/속도'}, title="시간 영역 파형")
        st.plotly_chart(fig_time, use_container_width=True)
        
    with tab2:
        fig_fft = px.line(x=fft_freq, y=amplitude, labels={'x': '주파수 (Hz)', 'y': '진폭 (Amplitude)'}, title="FFT 주파수 스펙트럼")
        fig_fft.update_traces(line_color="#ff7f0e")
        st.plotly_chart(fig_fft, use_container_width=True)


# ==========================================
# 메뉴 3: 📏 진동 측정 기준 (ISO 10816)
# ==========================================
elif menu_type == "📏 진동 측정 기준 (ISO 10816)":
    st.title("📏 ISO 10816 회전기기 진동 평가 기준")
    st.markdown("ISO 10816-3 표준은 산업용 회전기기의 진동 속도 실효값(RMS, mm/s)을 기준으로 설비의 건전성을 4단계로 평가합니다.")
    st.markdown("---")
    
    z1, z2, z3, z4 = st.columns(4)
    z1.success("🟢 **영역 A (Zone A)**\n\n신규 설치 또는 정비 직후의 우수한 상태")
    z2.info("🔵 **영역 B (Zone B)**\n\n장기 운전이 허용되는 양호한 상태")
    z3.warning("🟡 **영역 C (Zone C)**\n\n장기 운전 불가, 조만간 보수/정비 필요")
    z4.error("🔴 **영역 D (Zone D)**\n\n설비 손상 위험, 즉시 운전 정지 및 점검")
    
    st.markdown("---")
    
    st.subheader("📊 ISO 10816-3 진동 속도 기준표 (RMS mm/s)")
    
    iso_data = {
        "진동 구역 (Zone)": ["Zone A (우수)", "Zone B (양호)", "Zone C (보수필요)", "Zone D (위험/정지)"],
        "그룹 1: 대형 기기 (>300kW) [강성 기초]": ["< 2.3 mm/s", "2.3 ~ 4.5 mm/s", "4.5 ~ 7.1 mm/s", "> 7.1 mm/s"],
        "그룹 1: 대형 기기 (>300kW) [연성 기초]": ["< 3.5 mm/s", "3.5 ~ 7.1 mm/s", "7.1 ~ 11.0 mm/s", "> 11.0 mm/s"],
        "그룹 2: 중형 기기 (15kW~300kW) [강성 기초]": ["< 1.4 mm/s", "1.4 ~ 2.8 mm/s", "2.8 ~ 4.5 mm/s", "> 4.5 mm/s"],
        "그룹 2: 중형 기기 (15kW~300kW) [연성 기초]": ["< 2.3 mm/s", "2.3 ~ 4.5 mm/s", "4.5 ~ 7.1 mm/s", "> 7.1 mm/s"],
    }
    
    iso_df = pd.DataFrame(iso_data)
    st.table(iso_df)