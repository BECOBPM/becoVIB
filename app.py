import streamlit as st
import pandas as pd

# 1. 페이지 설정
st.set_page_config(page_title="부산환경공단 회전기기 진동 관리", layout="wide")

# 2. 사이드바 구성
with st.sidebar:
    st.title("📌 메뉴 (Navigation)")
    
    # 구분 선택
    menu_type = st.radio(
        "구분 선택",
        ["🏢 사업소별 설비 관리", "📏 진동 측정 기준 (ISO 10816)"]
    )
    
    st.markdown("---")
    
    # 사업소 선택 드롭다운
    site_list = ["에너지사업소", "생곡사업소", "명지사업소", "해운대사업소", "남부사업소", "수영사업소"]
    selected_site = st.selectbox("🏢 사업소 선택", site_list)
    
    st.markdown("---")
    
    # 사업소 데이터 파일 업로드 (사이드바에만 위치)
    st.subheader("📂 사업소 데이터 업로드")
    uploaded_file = st.file_uploader(f"[{selected_site}] 측정 데이터 파일(CSV/Excel) 첨부", type=["csv", "xlsx"])

# 3. 메인 화면 구성
if menu_type == "🏢 사업소별 설비 관리":
    # 선택된 사업소 표시
    st.info(f"🏢 현재 선택된 사업소: **{selected_site}**")
    
    st.title(f"📜 [{selected_site}] 설비별 진동 점검 이력 및 현황")
    st.caption("부산환경공단 소속 회전기기 진동 상태 측정 데이터 관리 시스템")
    
    # 데이터 로드 (업로드된 파일이 있으면 사용, 없으면 기본/예시 데이터 사용)
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            st.success("데이터 파일이 성공적으로 로드되었습니다.")
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            df = pd.DataFrame()
    else:
        # 예시 mock 데이터
        data = {
            "측정일자": ["2026-04-02", "2026-04-02", "2026-04-01", "2026-03-28"],
            "사업소명": [selected_site, selected_site, selected_site, selected_site],
            "설비명": ["집단에너지 FD Fan #1", "집단에너지 FD Fan #1", "보일러 급수펌프 #1", "1차 냉각수펌프 #2"],
            "전동기(kW)": [310, 310, 95, 11],
            "측정위치": ["1(전동기) X", "1(전동기) Y", "2(펌프) Z", "1(전동기) X"],
            "속도(mm/s)": [12.3, 2.8, 4.8, 1.1],
            "판정": ["D (즉시점검)", "B (양호)", "C (보수필요)", "A (양호)"],
            "조치사항": ["베어링 수선 정비", "정상운전", "구리스 보충 및 트렌드 관찰", "정상운전"]
        }
        df = pd.DataFrame(data)
    
    # 사업소 데이터 필터링 (데이터프레임에 '사업소명' 컬럼이 있는 경우)
    if "사업소명" in df.columns:
        filtered_df = df[df["사업소명"] == selected_site]
    else:
        filtered_df = df

    # 메인 상단 요약 지표 (Metric Cards)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📊 전체 설비 데이터", f"{len(filtered_df)} 건")
    
    if "판정" in filtered_df.columns:
        good_cnt = len(filtered_df[filtered_df["판정"].str.contains("A|B", na=False)])
        warning_cnt = len(filtered_df[filtered_df["판정"].str.contains("C", na=False)])
        danger_cnt = len(filtered_df[filtered_df["판정"].str.contains("D", na=False)])
    else:
        good_cnt, warning_cnt, danger_cnt = 0, 0, 0

    col2.metric("✅ 양호 (A / B)", f"{good_cnt} 건")
    col3.metric("⚠️ 보수 필요 (C)", f"{warning_cnt} 건")
    col4.metric("🚨 즉시 점검 (D)", f"{danger_cnt} 건")
    
    st.markdown("---")
    
    # 상세 데이터 테이블
    st.subheader("📋 상세 점검 이력")
    st.dataframe(filtered_df, use_container_width=True)

elif menu_type == "📏 진동 측정 기준 (ISO 10816)":
    st.title("📏 ISO 10816 회전기기 진동 평가 기준")
    # ISO 10816 안내 표 및 기준 내용 출력