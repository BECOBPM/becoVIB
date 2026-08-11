import streamlit as st
import pandas as pd

# 페이지 기본 설정
st.set_page_config(page_title="설비별 진동 점검 현황 대시보드", layout="wide")

# Custom CSS로 지표 카드 스타일링 강화
st.markdown("""
<style>
    div[data-testid="stMetric"] {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        padding: 15px 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 1. 사이드바 메뉴 및 파일 업로드 기능
# ---------------------------------------------------------
st.sidebar.title("📌 메뉴 (Navigation)")

st.sidebar.subheader("📂 데이터 파일 업로드")
uploaded_file = st.sidebar.file_uploader(
    "CSV 또는 엑셀 파일(.xlsx)을 업로드하세요", 
    type=["csv", "xlsx", "xls"]
)

# 데이터 로드 캐싱 함수
@st.cache_data
def load_data(file):
    if file is not None:
        if file.name.endswith(".csv"):
            return pd.read_csv(file)
        else:
            return pd.read_excel(file)
    else:
        # 업로드된 파일이 없을 때 기본 예시 데이터
        data = {
            "측정일자": ["2026-04-02", "2026-04-02", "2026-04-01", "2026-03-28"],
            "사업소명": ["에너지사업소", "에너지사업소", "에너지사업소", "에너지사업소"],
            "설비명": ["집단에너지 FD Fan #1", "집단에너지 FD Fan #1", "보일러 급수펌프 #1", "1차 냉각수펌프 #2"],
            "전동기(kW)": [310, 310, 95, 11],
            "측정위치": ["1(전동기) X", "1(전동기) Y", "2(펌프) Z", "1(전동기) X"],
            "속도(mm/s)": [12.3, 2.8, 4.8, 1.1],
            "판정": ["D (즉시점검)", "B (양호)", "C (보수필요)", "A (양호)"],
            "조치사항": ["베어링 수선 정비", "정상운전", "구리스 보충 및 트렌드 관찰", "정상운전"]
        }
        return pd.DataFrame(data)

df = load_data(uploaded_file)

st.sidebar.markdown("---")

# 사업소 선택 드롭다운
if "사업소명" in df.columns:
    site_list = df["사업소명"].unique()
    selected_site = st.sidebar.selectbox("🏢 사업소 선택", site_list)
    filtered_df = df[df["사업소명"] == selected_site].copy()
else:
    selected_site = "전체 사업소"
    filtered_df = df.copy()

# ---------------------------------------------------------
# 2. 메인 화면 - 상단 헤더 및 현황 요약 카드
# ---------------------------------------------------------
st.info(f"🏢 현재 선택된 사업소: **{selected_site}**")

st.title(f"📜 [{selected_site}] 설비별 진동 점검 이력 및 현황")
st.caption("회전기기(보일러 급수펌프, 송풍기, 펌프 등) 진동 상태 측정 데이터 및 점검 이력 관리")

# 데이터 현황 집계
total_count = len(filtered_df)
good_count = len(filtered_df[filtered_df["판정"].astype(str).str.contains("A|B|양호", na=False)])
repair_count = len(filtered_df[filtered_df["판정"].astype(str).str.contains("C|보수", na=False)])
urgent_count = len(filtered_df[filtered_df["판정"].astype(str).str.contains("D|즉시|긴급", na=False)])

# 4열 지표 카드 레이아웃
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="📊 전체 점검 건수", 
        value=f"{total_count} 건", 
        delta=f"{selected_site} 총 데이터"
    )

with col2:
    st.metric(
        label="✅ 양호 (A / B)", 
        value=f"{good_count} 건", 
        delta="정상 운전 중",
        delta_color="normal"
    )

with col3:
    st.metric(
        label="⚠️ 보수 필요 (C)", 
        value=f"{repair_count} 건", 
        delta="계획 정비 대상",
        delta_color="off"
    )

with col4:
    st.metric(
        label="🚨 즉시 점검 (D)", 
        value=f"{urgent_count} 건", 
        delta="긴급 점검 필요",
        delta_color="inverse"
    )

st.markdown("---")

# ---------------------------------------------------------
# 3. 상세 점검 이력 테이블
# ---------------------------------------------------------
st.subheader("📋 상세 점검 이력")

# 판정 결과별 강조 스타일 적용 함수
def highlight_status(val):
    val_str = str(val)
    if "D" in val_str or "즉시" in val_str or "긴급" in val_str:
        return 'background-color: #FFCDD2; color: #B71C1C; font-weight: bold;'
    elif "C" in val_str or "보수" in val_str:
        return 'background-color: #FFE0B2; color: #E65100; font-weight: bold;'
    elif "A" in val_str or "B" in val_str or "양호" in val_str:
        return 'background-color: #C8E6C9; color: #1B5E20; font-weight: bold;'
    return ''

if "판정" in filtered_df.columns:
    styled_df = filtered_df.style.applymap(highlight_status, subset=['판정'])
    st.dataframe(styled_df, use_container_width=True, height=350)
else:
    st.dataframe(filtered_df, use_container_width=True, height=350)