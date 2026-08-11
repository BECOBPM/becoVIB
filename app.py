import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 페이지 기본 설정
st.set_page_config(
    page_title="부산환경공단 - 설비 진동 & 이력 관리 시스템",
    page_icon="⚙️",
    layout="wide"
)

# 데이터 저장 경로
DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "vibration_data.csv")

# 데이터 로드 및 초기화 함수
@st.cache_data(ttl=1) # 데이터 업데이트 반영
def load_data():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    
    # DB 파일이 없는 경우 기존 엑셀 데이터를 정제하여 초기 CSV 생성
    if not os.path.exists(DATA_FILE):
        excel_path = "설비점검 및 정비 관리대장(에너지사업소).xlsx"
        if os.path.exists(excel_path):
            df_excel = pd.read_excel(excel_path, sheet_name='진동측정결과')
            df_excel['설비명'] = df_excel['설비명'].ffill()
            df_excel['측정일자'] = pd.to_datetime(df_excel['측정일자']).ffill()
            df_excel['사업소'] = "에너지사업소"  # 기본 사업소 지정
            df_excel.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
            return df_excel
        else:
            # 엑셀도 없으면 빈 프레임 생성
            cols = ['측정일자', '사업소', '설비명', '전동기(kW)', '측정위치 (1~4)', '축', '진동속도 (rms)', '부하상태 (%)', '판정 (A~D)', '이상 원인 추정', '정비 우선순위']
            df_empty = pd.DataFrame(columns=cols)
            df_empty.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
            return df_empty
    else:
        df = pd.read_csv(DATA_FILE)
        df['측정일자'] = pd.to_datetime(df['측정일자'])
        return df

df = load_data()

# ==========================================
# 🌲 [좌측 사이드바] 사업소 목차 및 설비 추가
# ==========================================
st.sidebar.title("🏢 부산환경공단")
st.sidebar.markdown("**사업소 목록**")

plants = [
    "강변사업소", "수영사업소", "해운대사업소", "남부사업소", 
    "명지사업소", "생곡사업소", "정관사업소", "중앙사업소", "서부사업소", "에너지사업소"
]
selected_plant = st.sidebar.selectbox("사업소를 선택하세요", plants)

st.sidebar.divider()

# ➕ 동적 설비 추가 기능
st.sidebar.subheader("➕ 신규 설비 등록")
with st.sidebar.form("add_equipment_form"):
    new_eq_name = st.text_input("설비명 (예: #1 급수펌프)")
    new_eq_kw = st.number_input("전동기 용량 (kW)", min_value=0.0, step=5.0)
    submit_eq = st.form_submit_button("설비 등록")
    
    if submit_eq:
        if new_eq_name:
            # 신규 설비 등록 기본 1건 입력
            new_row = pd.DataFrame([{
                '측정일자': pd.Timestamp.now().strftime('%Y-%m-%d'),
                '사업소': selected_plant,
                '설비명': new_eq_name,
                '전동기(kW)': new_eq_kw,
                '측정위치 (1~4)': '1 (전동기 부하측)',
                '축': 'x',
                '진동속도 (rms)': 0.0,
                '부하상태 (%)': 0.0,
                '판정 (A~D)': 'A',
                '이상 원인 추정': '신규 설비 등록',
                '정비 우선순위': '양호'
            }])
            df_updated = pd.concat([df, new_row], ignore_index=True)
            df_updated.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
            st.success(f"[{selected_plant}] {new_eq_name} 등록 완료!")
            st.rerun()

# ==========================================
# 💻 [메인 화면] 대시보드 & 이력 관리
# ==========================================
st.title(f"🛠️ {selected_plant} 설비 진동 & 이력 관리")

# 해당 사업소 데이터 필터링
plant_df = df[df['사업소'] == selected_plant]
available_equipments = plant_df['설비명'].dropna().unique().tolist()

# 탭 메뉴 구성
tab1, tab2, tab3 = st.tabs(["📊 진동 모니터링 & 추이", "📝 건별 데이터 직접 입력", "📤 엑셀 파일 대량 업로드"])

# ------------------------------------------
# TAB 1: 진동 시각화 차트
# ------------------------------------------
with tab1:
    if len(plant_df) == 0 or len(available_equipments) == 0:
        st.info("📌 등록된 설비나 진동 데이터가 없습니다. 좌측 메뉴에서 설비를 새로 등록해 주세요.")
    else:
        # 상단 현황 요약
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("등록 설비 수", f"{len(available_equipments)} 개")
        col2.metric("🟢 양호", f"{len(plant_df[plant_df['정비 우선순위']=='양호'])} 건")
        col3.metric("🟠 보수필요", f"{len(plant_df[plant_df['정비 우선순위']=='보수필요'])} 건")
        col4.metric("🔴 즉시점검", f"{len(plant_df[plant_df['정비 우선순위']=='즉시점검'])} 건")
        
        st.divider()
        
        # 설비 선택 드롭다운
        selected_eq = st.selectbox("조회할 설비를 선택하세요", available_equipments)
        eq_df = plant_df[plant_df['설비명'] == selected_eq]
        
        # RMS 진동속도 추이 그래프
        st.subheader(f"📈 [{selected_eq}] 진동속도(RMS) 추이")
        fig = px.line(
            eq_df, 
            x='측정일자', 
            y='진동속도 (rms)', 
            color='축',
            markers=True,
            hover_data=['측정위치 (1~4)', '판정 (A~D)', '정비 우선순위', '이상 원인 추정'],
            title=f"{selected_eq} 시계열 진동 변화"
        )
        # 위험 기준선 추가 (D등급 한계선 예시: 7.1 mm/s 이상)
        fig.add_hline(y=7.1, line_dash="dash", line_color="red", annotation_text="D등급 (위험/즉시점검)")
        fig.add_hline(y=4.5, line_dash="dash", line_color="orange", annotation_text="C등급 (경고/보수필요)")
        st.plotly_chart(fig, use_container_width=True)
        
        # 상세 데이터 테이블
        st.subheader("📋 측정 이력 대장")
        st.dataframe(eq_df.sort_values(by='측정일자', ascending=False), use_container_width=True)

# ------------------------------------------
# TAB 2: 건별 직접 입력
# ------------------------------------------
with tab2:
    st.subheader("📝 신규 진동 측정 결과 등록")
    if len(available_equipments) == 0:
        st.warning("먼저 설비를 좌측에서 등록해야 데이터를 입력할 수 있습니다.")
    else:
        with st.form("input_form"):
            c1, c2 = st.columns(2)
            m_date = c1.date_input("측정일자")
            eq_target = c2.selectbox("대상 설비", available_equipments)
            
            pos = c1.selectbox("측정위치", ["1 (전동기 부하측)", "2 (전동기 반부하측)", "3 (피동기 부하측)", "4 (피동기 반부하측)"])
            axis = c2.selectbox("축 방향", ["x", "y", "z"])
            
            rms_val = c1.number_input("진동속도 (rms)", min_value=0.0, step=0.1)
            load_val = c2.number_input("부하상태 (%)", min_value=0.0, max_value=100.0, value=80.0)
            
            grade = c1.selectbox("판정 등급", ["A", "B", "C", "D"])
            priority = c2.selectbox("정비 우선순위", ["양호", "보수필요", "즉시점검"])
            
            cause = st.text_input("이상 원인 추정 및 정비 소견")
            
            submit_data = st.form_submit_button("저장하기")
            
            if submit_data:
                # kW 용량 가져오기
                eq_kw = eq_df['전동기(kW)'].values[0] if len(eq_df) > 0 else 0.0
                
                new_record = pd.DataFrame([{
                    '측정일자': m_date.strftime('%Y-%m-%d'),
                    '사업소': selected_plant,
                    '설비명': eq_target,
                    '전동기(kW)': eq_kw,
                    '측정위치 (1~4)': pos,
                    '축': axis,
                    '진동속도 (rms)': rms_val,
                    '부하상태 (%)': load_val,
                    '판정 (A~D)': grade,
                    '이상 원인 추정': cause,
                    '정비 우선순위': priority
                }])
                
                df_updated = pd.concat([df, new_record], ignore_index=True)
                df_updated.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
                st.success("새로운 진동 데이터가 성공적으로 저장되었습니다!")
                st.rerun()

# ------------------------------------------
# TAB 3: 엑셀 파일 대량 업로드
# ------------------------------------------
with tab3:
    st.subheader("📤 기존 엑셀 대장 파일 업로드")
    st.write("작성된 진동 측정 엑셀 파일(.xlsx)을 업로드하면 DB에 일괄 추가됩니다.")
    
    uploaded_file = st.file_uploader("엑셀 파일 선택", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            up_df = pd.read_excel(uploaded_file, sheet_name='진동측정결과')
            up_df['설비명'] = up_df['설비명'].ffill()
            up_df['측정일자'] = pd.to_datetime(up_df['측정일자']).ffill()
            up_df['사업소'] = selected_plant
            
            st.write("📋 **업로드할 데이터 미리보기**")
            st.dataframe(up_df.head(10))
            
            if st.button("이 데이터를 시스템에 일괄 등록"):
                df_updated = pd.concat([df, up_df], ignore_index=True)
                df_updated.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
                st.success("엑셀 데이터 일괄 업로드가 완료되었습니다!")
                st.rerun()
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다. 양식을 확인해주세요: {e}")