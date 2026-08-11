import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os

# 1. 페이지 기본 설정
st.set_page_config(
    page_title="부산환경공단 회전기기 진동 관리 시스템",
    page_icon="⚙️",
    layout="wide"
)

# 2. 이미지 기준 18개 사업소/사업단 목록
SITE_LIST = [
    "수영사업단", "강변사업단", "남부사업소", "녹산사업소", "기장사업소",
    "동부사업소", "중앙사업소", "영도사업소", "정관사업소", "서부사업소",
    "관로사업소", "하수자원사업소", "위생사업소", "해운대사업단", "생곡사업단",
    "명지사업소", "에너지사업소", "대기환경사업소"
]

DEFAULT_EXCEL_FILE = "설비점검 및 정비 관리대장(에너지사업소).xlsx"

# 3. Session State (사업소별 데이터 저장소) 초기화
if "site_data_store" not in st.session_state:
    st.session_state["site_data_store"] = {}

# 4. 엑셀 데이터 정제 함수 (상단 제목 및 빈 행 자동 제거)
def clean_excel_data(df):
    if df is None or df.empty:
        return pd.DataFrame()
    
    # '설비명' 단어가 포함된 실제 헤더 행 위치 찾기
    header_idx = None
    for idx in range(min(10, len(df))):
        row_vals = df.iloc[idx].astype(str).values
        if any("설비명" in val for val in row_vals):
            header_idx = idx
            break
            
    if header_idx is not None:
        # 헤더 행으로 컬럼 설정 및 상단 제목/빈 행 삭제
        new_cols = [str(val).strip() if pd.notna(val) else f"Unnamed_{i}" for i, val in enumerate(df.iloc[header_idx].values)]
        df = df.iloc[header_idx + 1:].copy()
        df.columns = new_cols

    # '설비명' 컬럼 자동 찾기 및 유효 데이터만 필터링
    equip_col = next((c for c in df.columns if "설비명" in c), None)
    if equip_col:
        df = df[df[equip_col].notna()]
        df = df[~df[equip_col].astype(str).str.strip().isin(["None", "nan", "", "설비명"])]
        
    return df.reset_index(drop=True)


# 5. 사이드바 구성
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
    
    # 사업소 선택 (기본값: 에너지사업소)
    selected_site = st.selectbox("🏢 사업소 선택", SITE_LIST, index=16)
    
    st.markdown("---")
    
    st.subheader("📂 사업소 데이터 업로드")
    uploaded_file = st.file_uploader(f"[{selected_site}] 전용 파일 업로드 (CSV/Excel)", type=["csv", "xlsx"])

    # 파일 업로드 시 해당 사업소 전용 저장소에 저장
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                raw_df = pd.read_csv(uploaded_file)
            else:
                raw_df = pd.read_excel(uploaded_file)
            
            cleaned_df = clean_excel_data(raw_df)
            st.session_state["site_data_store"][selected_site] = cleaned_df
            st.success(f"✅ [{selected_site}] 데이터 {len(cleaned_df)}건 저장 완료!")
        except Exception as e:
            st.error(f"파일 처리 오류: {e}")


# 6. 현재 선택된 사업소 데이터 가져오기
current_df = pd.DataFrame()

if selected_site in st.session_state["site_data_store"]:
    current_df = st.session_state["site_data_store"][selected_site]
elif selected_site == "에너지사업소" and os.path.exists(DEFAULT_EXCEL_FILE):
    # 에너지사업소 선택 시 기본 대장 파일 자동 적용
    try:
        raw_df = pd.read_excel(DEFAULT_EXCEL_FILE)
        current_df = clean_excel_data(raw_df)
        st.session_state["site_data_store"]["에너지사업소"] = current_df
    except Exception as e:
        current_df = pd.DataFrame()


# 7. 메인 화면 구성
# ==========================================
# 메뉴 1: 🏢 사업소별 설비 관리
# ==========================================
if menu_type == "🏢 사업소별 설비 관리":
    st.info(f"🏢 현재 선택된 사업소: **{selected_site}**")
    st.title(f"📜 [{selected_site}] 설비별 진동 점검 이력 및 현황")
    st.caption("부산환경공단 소속 회전기기 진동 상태 측정 데이터 관리 시스템")
    
    if current_df.empty:
        st.warning(f"💡 [{selected_site}]에 등록된 진동 점검 데이터가 없습니다.")
        st.info("👈 왼쪽 사이드바의 **'사업소 데이터 업로드'**에서 해당 사업소의 엑셀/CSV 파일을 등록해 주세요.")
    else:
        # 주요 컬럼 자동 인식
        status_col = next((c for c in current_df.columns if "판정" in c or "상태" in c), None)
        problem_col = next((c for c in current_df.columns if "문제점" in c), None)
        speed_col = next((c for c in current_df.columns if "속도" in c or "진동" in c or "RMS" in c), None)
        equip_col = next((c for c in current_df.columns if "설비" in c), None)

        # 건수 집계 로직
        total_cnt = len(current_df)
        good_cnt, warning_cnt, danger_cnt = 0, 0, 0

        if status_col:
            s_str = current_df[status_col].astype(str)
            good_cnt = len(current_df[s_str.str.contains("A|B|양호|정상", case=False, na=False)])
            warning_cnt = len(current_df[s_str.str.contains("C|주의|보수", case=False, na=False)])
            danger_cnt = len(current_df[s_str.str.contains("D|위험|즉시|점검", case=False, na=False)])
        elif problem_col:
            p_str = current_df[problem_col].astype(str).str.strip()
            # 문제점이 기록된 건수만 보수필요로 분류
            has_problem = p_str.notna() & ~p_str.isin(["None", "nan", "", "정상"])
            warning_cnt = len(current_df[has_problem])
            good_cnt = total_cnt - warning_cnt

        # 상단 요약 카운트 카드
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📊 점검 설비 건수", f"{total_cnt} 건")
        c2.metric("✅ 양호 (정상)", f"{good_cnt} 건")
        c3.metric("⚠️ 보수 필요 (주의)", f"{warning_cnt} 건")
        c4.metric("🚨 즉시 점검 (위험)", f"{danger_cnt} 건")
        
        st.markdown("---")
        
        # 설비 진동 수치 차트 (속도/진동 데이터가 있는 경우)
        if speed_col and equip_col:
            st.subheader("📊 설비별 진동 측정치 비교")
            fig_bar = px.bar(
                current_df, 
                x=equip_col, 
                y=speed_col, 
                color=status_col if status_col else None,
                text=speed_col,
                title=f"{selected_site} 설비 진동 현황"
            )
            fig_bar.update_traces(textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True)

        # 상세 데이터 테이블
        st.subheader("📋 상세 점검 이력 대장")
        st.dataframe(current_df, use_container_width=True)


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