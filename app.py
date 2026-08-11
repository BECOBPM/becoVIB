import streamlit as st
import pandas as pd
import numpy as np

# 페이지 기본 설정
st.set_page_config(
    page_title="부산환경공단 설비 관리 시스템",
    page_icon="🏢",
    layout="wide"
)

# -------------------------------------------------------------------
# 1. 사이드바 (Navigation & 사업소 선택)
# -------------------------------------------------------------------
st.sidebar.title("📌 메뉴 (Navigation)")

# 구분 선택 라디오 버튼
menu_option = st.sidebar.radio(
    "구분 선택",
    ["🏢 사업소별 설비 관리", "📏 진동 측정 기준 (ISO 10816)"]
)

st.sidebar.markdown("---")

# 부산환경공단 공식 홈페이지 기준 개별 사업소 리스트 (그룹화 없이 1:1 구성)
beco_sites = [
    "에너지사업소",
    "해운대사업단",
    "명지사업소",
    "생곡사업단",
    "대기환경사업소",
    "수영사업단",
    "강변사업단",
    "남부사업소",
    "녹산사업소",
    "기장사업소",
    "동부사업소",
    "중앙사업소",
    "영도사업소",
    "정관사업소",
    "서부사업소",
    "관로사업소",
    "하수자원사업소",
    "위생사업소",
    "전체 / 통합"
]

# 드롭다운 사업소 선택 (기본값: 에너지사업소)
selected_site = st.sidebar.selectbox(
    "🏢 사업소 선택",
    options=beco_sites,
    index=0  # 에너지사업소가 기본 선택됨
)

# -------------------------------------------------------------------
# 2. 샘플 데이터 생성 함수 (에너지사업소 및 타 사업소 데이터)
# -------------------------------------------------------------------
@st.cache_data
def load_vibration_data():
    # 에너지사업소 중심 및 타 사업소 예시 데이터
    data = [
        {"측정일자": "2026-04-02", "사업소명": "에너지사업소", "설비명": "집단에너지 FD Fan #1", "전동기(kW)": 310, "측정위치": "1(전동기) X", "속도(mm/s)": 12.3, "판정": "D (즉시점검)", "조치사항": "베어링 수선 정비"},
        {"측정일자": "2026-04-02", "사업소명": "에너지사업소", "설비명": "집단에너지 FD Fan #1", "전동기(kW)": 310, "측정위치": "1(전동기) Y", "속도(mm/s)": 2.8, "판정": "B (양호)", "조치사항": "정상운전"},
        {"측정일자": "2026-04-01", "사업소명": "에너지사업소", "설비명": "보일러 급수펌프 #1", "전동기(kW)": 95, "측정위치": "2(펌프) Z", "속도(mm/s)": 4.8, "판정": "C (보수필요)", "조치사항": "구리스 보충 및 트렌드 관찰"},
        {"측정일자": "2026-03-28", "사업소명": "에너지사업소", "설비명": "1차 냉각수펌프 #2", "전동기(kW)": 11, "측정위치": "1(전동기) X", "속도(mm/s)": 1.1, "판정": "A (양호)", "조치사항": "정상운전"},
        {"측정일자": "2026-03-15", "사업소명": "해운대사업단", "설비명": "소각로 유인풍기 #2", "전동기(kW)": 220, "측정위치": "2(풍기) Y", "속도(mm/s)": 5.2, "판정": "C (보수필요)", "조치사항": "진동 관찰"},
        {"측정일자": "2026-03-10", "사업소명": "명지사업소", "설비명": "슬러지 수송펌프 #1", "전동기(kW)": 45, "측정위치": "1(전동기) X", "속도(mm/s)": 1.4, "판정": "A (양호)", "조치사항": "정상운전"}
    ]
    return pd.DataFrame(data)

df_all = load_vibration_data()

# 선택된 사업소에 맞게 데이터 필터링
if selected_site == "전체 / 통합":
    df_filtered = df_all
else:
    df_filtered = df_all[df_all["사업소명"] == selected_site]

# -------------------------------------------------------------------
# 3. 메인 화면 출력
# -------------------------------------------------------------------
if menu_option == "🏢 사업소별 설비 관리":
    # 상단 배지 표시
    st.info(f"🏢 현재 선택된 사업소: **{selected_site}**")
    
    st.title(f"📜 [{selected_site}] 설비별 진동 점검 이력 및 현황")
    st.caption(f"부산환경공단 {selected_site} 소속 회전기기(보일러 급수펌프, 송풍기, 펌프 등) 진동 상태 측정 데이터")

    # 메트릭 요약 카드
    col1, col2, col3, col4 = st.columns(4)
    
    total_cnt = len(df_filtered)
    good_cnt = len(df_filtered[df_filtered['판정'].str.startswith(('A', 'B'))])
    warn_cnt = len(df_filtered[df_filtered['판정'].str.startswith('C')])
    crit_cnt = len(df_filtered[df_filtered['판정'].str.startswith('D')])

    col1.metric("점검 대상 설비 데이터", f"{total_cnt} 건", selected_site)
    col2.metric("양호 (A / B)", f"{good_cnt} 건", "정상 운전중")
    col3.metric("보수 필요 (C)", f"{warn_cnt} 건", "계획 정비", delta_color="off")
    col4.metric("즉시 점검 (D)", f"{crit_cnt} 건", "긴급 정비", delta_color="inverse")

    st.markdown("---")

    # 데이터 테이블 출력
    st.subheader("📋 상세 점검 이력")
    
    if df_filtered.empty:
        st.warning(f" 현재 {selected_site}에 등록된 측정 데이터가 없습니다.")
    else:
        # 조건별 스타일 적용 (D등급 빨간색 강조 등)
        def highlight_status(val):
            if 'D (' in str(val):
                return 'background-color: #FFCDD2; color: #B71C1C; font-weight: bold;'
            elif 'C (' in str(val):
                return 'background-color: #FFE0B2; color: #E65100; font-weight: bold;'
            elif 'A (' in str(val) or 'B (' in str(val):
                return 'color: #2E7D32;'
            return ''

        styled_df = df_filtered.style.map(highlight_status, subset=['판정'])
        st.dataframe(styled_df, use_container_width=True)

elif menu_option == "📏 진동 측정 기준 (ISO 10816)":
    st.title("📏 진동 측정 및 평가 기준 (ISO 10816)")
    st.write("ISO 10816-3 회전기기 진동 속도(mm/s RMS) 판정 기준표입니다.")
    
    iso_data = {
        "구분": ["A 등급 (Good)", "B 등급 (Satisfactory)", "C 등급 (Unsatisfactory)", "D 등급 (Unacceptable)"],
        "진동 속도 범위 (mm/s)": ["0.0 ~ 1.8 mm/s", "1.8 ~ 4.5 mm/s", "4.5 ~ 11.2 mm/s", "11.2 mm/s 초과"],
        "설비 상태": ["신규 설치 및 양호한 상태", "장기 운전 가능 상태", "조만간 정비 및 조치 필요", "즉시 정지 및 긴급 정비"]
    }
    st.table(pd.DataFrame(iso_data))