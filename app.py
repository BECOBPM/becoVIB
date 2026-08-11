import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os
import io

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

# 4. 엑셀 데이터 정제 함수 (상단 제목 및 빈 행 자동 제거, float 타입 예외 처리)
def clean_excel_data(df):
    if df is None or df.empty:
        return pd.DataFrame()
    
    # 1) '설비명' 단어가 포함된 실제 헤더 행 위치 찾기
    header_idx = None
    for idx in range(min(15, len(df))):
        row_str_list = [str(val) for val in df.iloc[idx].tolist()]
        if any("설비명" in val for val in row_str_list):
            header_idx = idx
            break
            
    # 2) 헤더 위치 지정 및 상단 비어있는 행 삭제
    if header_idx is not None:
        raw_headers = df.iloc[header_idx].tolist()
        new_cols = []
        for i, h in enumerate(raw_headers):
            h_str = str(h).strip() if pd.notna(h) and str(h).lower() != "nan" else f"Unused_{i}"
            new_cols.append(h_str)
        
        df = df.iloc[header_idx + 1:].copy()
        df.columns = new_cols

    # 3) '설비명' 컬럼 기준 유효 데이터만 정제
    equip_col = next((c for c in df.columns if "설비명" in str(c)), None)
    if equip_col:
        mask = df[equip_col].apply(lambda x: pd.notna(x) and str(x).strip().lower() not in ["none", "nan", "", "설비명"])
        df = df[mask]
        
    return df.reset_index(drop=True)


# 5. 샘플 엑셀 양식 생성 함수
def generate_template_excel():
    sample_df = pd.DataFrame({
        "설비명": ["FD Fan #1", "FD Fan #2", "보일러 급수펌프 #1"],
        "점검계획": ["반기", "반기", "반기"],
        "점검내역": ["오일, 그리스, 진동 상태 등", "오일, 그리스, 진동 상태 등", "오일, 그리스, 진동 상태 등"],
        "점검일자": ["2026-04-22", "2026-04-22", "2026-04-23"],
        "문제점": ["전동기 진동", "정상", "정상"],
        "원인분석": ["베어링 파손", "-", "-"],
        "조치사항": ["분해정비", "-", "-"],
        "예방정비내역": ["-", "오일, 그리스 주입", "오일, 그리스 주입"],
        "차기점검일": ["2026-10-01", "2026-10-01", "2026-10-01"],
        "진동속도(mm/s)": [4.8, 1.1, 0.9],
        "판정": ["C (보수필요)", "A (양호)", "A (양호)"]
    })
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sample_df.to_excel(writer, index=False, sheet_name='점검대장양식')
    return output.getvalue()


# 6. 사이드바 구성
with st.sidebar:
    st.title("📌 메뉴 (Navigation)")
    
    menu_type = st.radio(
        "구분 선택",
        [
            "🏢 사업소별 설비 현황 및 이력", 
            "📂 데이터 업로드 및 양식 다운로드",
            "📏 진동 측정 기준 (ISO 10816)"
        ]
    )
    
    st.markdown("---")
    
    # 사업소 선택 (기본값: 에너지사업소)
    selected_site = st.selectbox("🏢 사업소 선택", SITE_LIST, index=16)
    
    st.markdown("---")
    st.caption("부산환경공단 회전기기 진동 관리 시스템 v1.2")


# 7. 기본 데이터 로드 (에너지사업소 파일)
current_df = pd.DataFrame()

if selected_site in st.session_state["site_data_store"]:
    current_df = st.session_state["site_data_store"][selected_site]
elif selected_site == "에너지사업소" and os.path.exists(DEFAULT_EXCEL_FILE):
    try:
        raw_df = pd.read_excel(DEFAULT_EXCEL_FILE)
        current_df = clean_excel_data(raw_df)
        st.session_state["site_data_store"]["에너지사업소"] = current_df
    except Exception as e:
        current_df = pd.DataFrame()


# 8. 메인 화면 구성
# ==========================================
# 메뉴 1: 🏢 사업소별 설비 현황 및 이력
# ==========================================
if menu_type == "🏢 사업소별 설비 현황 및 이력":
    st.info(f"🏢 현재 선택된 사업소: **{selected_site}**")
    st.title(f"📜 [{selected_site}] 설비별 진동 점검 이력 및 현황")
    st.caption("부산환경공단 소속 회전기기 진동 상태 측정 데이터 관리 시스템")
    
    if current_df.empty:
        st.warning(f"💡 [{selected_site}]에 등록된 진동 점검 데이터가 없습니다.")
        st.info("👈 메뉴에서 **'📂 데이터 업로드 및 양식 다운로드'**로 이동하여 파일을 등록해 주세요.")
    else:
        # 주요 컬럼 자동 인식
        status_col = next((c for c in current_df.columns if "판정" in str(c) or "상태" in str(c)), None)
        problem_col = next((c for c in current_df.columns if "문제점" in str(c)), None)
        speed_col = next((c for c in current_df.columns if "속도" in str(c) or "진동" in str(c) or "RMS" in str(c)), None)
        equip_col = next((c for c in current_df.columns if "설비" in str(c)), None)

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
            has_problem = p_str.notna() & ~p_str.isin(["None", "nan", "", "정상", "-"])
            warning_cnt = len(current_df[has_problem])
            good_cnt = total_cnt - warning_cnt

        # 상단 요약 카운트 카드
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📊 점검 설비 건수", f"{total_cnt} 건")
        c2.metric("✅ 양호 (정상)", f"{good_cnt} 건")
        c3.metric("⚠️ 보수 필요 (주의)", f"{warning_cnt} 건")
        c4.metric("🚨 즉시 점검 (위험)", f"{danger_cnt} 건")
        
        st.markdown("---")
        
        # 설비 진동 수치 차트
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
# 메뉴 2: 📂 데이터 업로드 및 양식 다운로드
# ==========================================
elif menu_type == "📂 데이터 업로드 및 양식 다운로드":
    st.title("📂 데이터 업로드 및 엑셀 양식 안내")
    st.caption("각 사업소별 설비 점검 대장을 등록하거나 표준 업로드 양식을 다운로드합니다.")
    st.markdown("---")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1️⃣ 표준 업로드 양식 다운로드")
        st.markdown("""
        정확한 데이터 인식을 위해 아래 **표준 양식 서식**에 맞춰 작성된 파일 업로드를 권장합니다.
        
        * **필수 포함 항목**: `설비명`, `점검일자`
        * **권장 포함 항목**: `점검계획`, `점검내역`, `문제점`, `원인분석`, `조치사항`, `차기점검일`, `진동속도(mm/s)`, `판정`
        """)
        
        # 샘플 엑셀 파일 다운로드 버튼
        excel_bytes = generate_template_excel()
        st.download_button(
            label="📥 표준 엑셀 양식 다운로드 (.xlsx)",
            data=excel_bytes,
            file_name="회전기기_진동점검대장_표준양식.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        st.subheader(f"2️⃣ [{selected_site}] 데이터 파일 업로드")
        st.markdown(f"현재 선택된 사업소(**{selected_site}**)에 데이터를 올립니다.")
        
        uploaded_file = st.file_uploader(
            f"[{selected_site}] 전용 점검대장 업로드 (CSV/XLSX)", 
            type=["csv", "xlsx"]
        )

        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    raw_df = pd.read_csv(uploaded_file)
                else:
                    raw_df = pd.read_excel(uploaded_file)
                
                cleaned_df = clean_excel_data(raw_df)
                st.session_state["site_data_store"][selected_site] = cleaned_df
                st.success(f"🎉 [{selected_site}] 데이터 {len(cleaned_df)}건이 성공적으로 등록되었습니다!")
                
                # 등록된 미리보기
                st.markdown("##### 🔍 업로드 데이터 미리보기")
                st.dataframe(cleaned_df.head(5), use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ 파일 업로드 중 오류가 발생했습니다: {e}")

    st.markdown("---")
    st.subheader("💡 작성 시 유의사항")
    st.markdown("""
    1. 엑셀 제목 행(Header) 위치는 상단 1~10번째 줄 내에 **'설비명'** 단어가 들어가 있으면 자동으로 인식됩니다.
    2. 데이터가 없는 빈 셀은 `None` 또는 `-` 로 표시해도 무방합니다.
    3. 등록된 데이터는 상단 **'🏢 사업소별 설비 현황 및 이력'** 메뉴에서 즉시 확인 가능합니다.
    """)


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