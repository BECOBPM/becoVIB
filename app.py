import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io

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
@st.cache_data(ttl=1)
def load_data():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    
    if not os.path.exists(DATA_FILE):
        excel_path = "설비점검 및 정비 관리대장(에너지사업소).xlsx"
        if os.path.exists(excel_path):
            df_excel = pd.read_excel(excel_path, sheet_name='진동측정결과')
            df_excel['설비명'] = df_excel['설비명'].ffill()
            df_excel['측정일자'] = pd.to_datetime(df_excel['측정일자']).ffill()
            df_excel['사업소'] = "에너지사업소"
            df_excel.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
            return df_excel
        else:
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

# 4개 탭 구성 (첫 번째 탭: 사업소 요약)
tab0, tab1, tab2, tab3 = st.tabs([
    "📋 사업소 종합 요약", 
    "📊 설비별 진동 추이", 
    "📝 건별 데이터 입력", 
    "📤 엑셀 대량 업로드 & 샘플 다운로드"
])

# ------------------------------------------
# TAB 0: 사업소 종합 요약 대시보드
# ------------------------------------------
with tab0:
    if len(plant_df) == 0:
        st.info("📌 등록된 설비나 진동 데이터가 없습니다. 좌측 메뉴에서 설비를 등록해 주세요.")
    else:
        # 1. 핵심 KPI 카운트
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("총 등록 설비", f"{len(available_equipments)} 개")
        c2.metric("🟢 정상 (A/B)", f"{len(plant_df[plant_df['판정 (A~D)'].isin(['A', 'B'])])} 건")
        c3.metric("🟠 주의 (C등급/보수필요)", f"{len(plant_df[plant_df['판정 (A~D)'] == 'C'])} 건")
        c4.metric("🔴 경고 (D등급/즉시점검)", f"{len(plant_df[plant_df['판정 (A~D)'] == 'D'])} 건")
        
        st.divider()
        
        # 2. 🚨 긴급 점검 필요 설비 알림판
        urgent_df = plant_df[plant_df['정비 우선순위'] == '즉시점검']
        if len(urgent_df) > 0:
            st.error(f"🚨 **[긴급 점검 필요] 즉시 점검 대상이 {len(urgent_df)}건 발견되었습니다!**")
            st.dataframe(
                urgent_df[['측정일자', '설비명', '측정위치 (1~4)', '축', '진동속도 (rms)', '이상 원인 추정']], 
                use_container_width=True
            )
        else:
            st.success("✅ 현재 긴급 점검(D등급)이 필요한 설비가 없습니다.")
            
        st.divider()
        
        # 3. 차트 요약
        col_summary1, col_summary2 = st.columns(2)
        
        with col_summary1:
            st.subheader("📊 설비별 최근 최대 진동속도 (RMS)")
            # 설비별 최근/최대 진동값 요약
            max_rms_df = plant_df.groupby('설비명')['진동속도 (rms)'].max().reset_index()
            fig_bar = px.bar(
                max_rms_df,
                x='설비명',
                y='진동속도 (rms)',
                title="설비별 최고 진동속도 현황",
                color='진동속도 (rms)',
                color_continuous_scale=['#2ecc71', '#e67e22', '#e74c3c']
            )
            fig_bar.add_hline(y=7.1, line_dash="dash", line_color="red", annotation_text="D등급 임계치(7.1)")
            st.plotly_chart(fig_bar, use_container_width=True)
            
        with col_summary2:
            st.subheader("🎯 사업소 전체 판정 등급 비율")
            fig_pie = px.pie(
                plant_df,
                names='판정 (A~D)',
                color='판정 (A~D)',
                color_discrete_map={'A':'#2ecc71', 'B':'#3498db', 'C':'#e67e22', 'D':'#e74c3c'},
                hole=0.4
            )
            st.plotly_chart(fig_pie, use_container_width=True)

# ------------------------------------------
# TAB 1: 설비별 진동 추이
# ------------------------------------------
with tab1:
    if len(available_equipments) == 0:
        st.info("등록된 설비가 없습니다.")
    else:
        selected_eq = st.selectbox("조회할 설비를 선택하세요", available_equipments)
        eq_df = plant_df[plant_df['설비명'] == selected_eq]
        
        st.subheader(f"📈 [{selected_eq}] 진동속도(RMS) 시계열 추이")
        fig = px.line(
            eq_df, 
            x='측정일자', 
            y='진동속도 (rms)', 
            color='축',
            markers=True,
            hover_data=['측정위치 (1~4)', '판정 (A~D)', '정비 우선순위', '이상 원인 추정']
        )
        fig.add_hline(y=7.1, line_dash="dash", line_color="red", annotation_text="D등급 (위험/즉시점검)")
        fig.add_hline(y=4.5, line_dash="dash", line_color="orange", annotation_text="C등급 (경고/보수필요)")
        st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("📋 측정 이력 대장")
        st.dataframe(eq_df.sort_values(by='측정일자', ascending=False), use_container_width=True)

# ------------------------------------------
# TAB 2: 건별 직접 입력
# ------------------------------------------
with tab2:
    st.subheader("📝 신규 진동 측정 결과 등록")
    if len(available_equipments) == 0:
        st.warning("먼저 좌측에서 설비를 등록해 주세요.")
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
                eq_kw = eq_df['전동기(kW)'].values[0] if 'eq_df' in locals() and len(eq_df) > 0 else 0.0
                
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
                st.success("새로운 진동 데이터가 저장되었습니다!")
                st.rerun()

# ------------------------------------------
# TAB 3: 엑셀 파일 대량 업로드 및 샘플 다운로드
# ------------------------------------------
with tab3:
    st.subheader("📤 기존 엑셀 대장 파일 업로드 및 샘플 다운로드")
    col_down, col_up = st.columns([1, 2])
    
    with col_down:
        st.markdown("##### 📥 1. 양식 샘플 다운로드")
        st.write("샘플 양식을 받아 데이터를 작성해 주세요.")
        
        sample_data = pd.DataFrame([
            {
                '측정일자': '2026-04-22',
                '설비명': 'FD Fan #1',
                '전동기(kW)': 310.0,
                '측정위치 (1~4)': '1 (전동기 부하측)',
                '축': 'x',
                '진동속도 (rms)': 12.3,
                '부하상태 (%)': 30.0,
                '판정 (A~D)': 'D',
                '이상 원인 추정': '베어링 진동',
                '정비 우선순위': '즉시점검'
            }
        ])
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            sample_data.to_excel(writer, sheet_name='진동측정결과', index=False)
        buffer.seek(0)
        
        st.download_button(
            label="📄 진동측정대장 샘플 양식 (.xlsx)",
            data=buffer,
            file_name="진동측정대장_양식_샘플.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_up:
        st.markdown("##### 📤 2. 작성된 엑셀 파일 업로드")
        uploaded_file = st.file_uploader("작성된 엑셀 파일을 선택하세요", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            try:
                try:
                    up_df = pd.read_excel(uploaded_file, sheet_name='진동측정결과')
                except:
                    up_df = pd.read_excel(uploaded_file)
                
                up_df['설비명'] = up_df['설비명'].ffill()
                up_df['측정일자'] = pd.to_datetime(up_df['측정일자']).ffill()
                up_df['사업소'] = selected_plant
                
                st.write("📋 **업로드할 데이터 미리보기**")
                st.dataframe(up_df.head(10))
                
                if st.button("이 데이터를 시스템에 일괄 저장"):
                    df_updated = pd.concat([df, up_df], ignore_index=True)
                    df_updated.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
                    st.success("엑셀 데이터 일괄 업로드가 완료되었습니다!")
                    st.rerun()
            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")