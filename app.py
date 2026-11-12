import streamlit as st
import streamlit.components.v1 as components
import sqlite3
import json
import os
import re
from datetime import datetime

# 엑셀 서버 직접 파싱을 위한 pandas / openpyxl 활용
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

# ==============================================================================
# 1. 페이지 기본 설정 및 서버 내부 SQLite DB 초기화
# ==============================================================================
st.set_page_config(
    page_title="BECO-VIB 전사 통합 진동진단 서버",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_FILE = "beco_vib_server.db"

# 부산환경공단 18개 사업단·사업소 기본 목록
DEFAULT_OFFICES = [
    {"name": "수영사업단", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "강변사업단", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "남부사업소", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "녹산사업소", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "기장사업소", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "동부사업소", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "중앙사업소", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "영도사업소", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "정관사업소", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "서부사업소", "icon": "fa-water", "color": "text-cyan-400"},
    {"name": "관로사업소", "icon": "fa-route", "color": "text-indigo-400"},
    {"name": "하수자원사업소", "icon": "fa-recycle", "color": "text-emerald-400"},
    {"name": "위생사업소", "icon": "fa-industry", "color": "text-blue-400"},
    {"name": "해운대사업단", "icon": "fa-fire-flame-curved", "color": "text-amber-400"},
    {"name": "생곡사업단", "icon": "fa-solar-panel", "color": "text-emerald-400"},
    {"name": "명지사업소", "icon": "fa-fire-flame-curved", "color": "text-amber-400"},
    {"name": "에너지사업소", "icon": "fa-bolt", "color": "text-amber-400"},
    {"name": "대기환경사업소", "icon": "fa-wind", "color": "text-sky-400"},
]

DEFAULT_PERIODS = ["2026년 상반기", "2026년 하반기"]


def calc_iso_grade_py(power_kw, rms_val, foundation="rigid"):
    group = "G2"
    if power_kw > 300:
        group = "G1"
    elif power_kw < 15:
        group = "G4"

    rms = float(rms_val or 0.0)
    grade = "A"
    if group == "G1":
        if foundation == "rigid":
            if rms <= 2.3: grade = "A"
            elif rms <= 4.5: grade = "B"
            elif rms <= 7.1: grade = "C"
            else: grade = "D"
        else:
            if rms <= 3.5: grade = "A"
            elif rms <= 7.1: grade = "B"
            elif rms <= 11.0: grade = "C"
            else: grade = "D"
    else:
        if foundation == "rigid":
            if rms <= 1.4: grade = "A"
            elif rms <= 2.8: grade = "B"
            elif rms <= 4.5: grade = "C"
            else: grade = "D"
        else:
            if rms <= 2.3: grade = "A"
            elif rms <= 4.5: grade = "B"
            elif rms <= 7.1: grade = "C"
            else: grade = "D"
    return group, grade


def parse_excel_directly_in_python(uploaded_file, upload_office, default_period, existing_periods):
    """파이썬 서버에서 직접 엑셀(.xlsx) 다중 시트를 파싱하여 SQLite DB 저장용 리스트 생성"""
    xl = pd.ExcelFile(uploaded_file)
    parsed_by_sheet = {}
    new_periods = list(existing_periods)

    for sheet_name in xl.sheet_names:
        if any(k in sheet_name for k in ["가이드", "양식", "설명"]):
            continue

        sheet_period = default_period
        ym = re.search(r"(\d{2,4})\s*년?", sheet_name)
        y_str = default_period.split("년")[0] + "년"
        if ym:
            y_val = ym.group(1)
            y_str = ("20" + y_val if len(y_val) == 2 else y_val) + "년"

        if "상반기" in sheet_name:
            sheet_period = f"{y_str} 상반기"
        elif "하반기" in sheet_name:
            sheet_period = f"{y_str} 하반기"

        if sheet_period not in new_periods:
            new_periods.append(sheet_period)
        if sheet_period not in parsed_by_sheet:
            parsed_by_sheet[sheet_period] = {}

        df = xl.parse(sheet_name, header=None).fillna("")
        rows = df.values.tolist()

        current_equip = None
        last_power = 55.0
        is_header_found = False

        for row in rows:
            if not row or len(row) < 4:
                continue
            row_str = "".join(str(c) for c in row).replace(" ", "")
            if "설비명" in row_str and ("측정위치" in row_str or "전동기" in row_str):
                is_header_found = True
                continue
            if not is_header_found:
                continue

            equip_name = re.sub(r"\s+", " ", str(row[1] if len(row) > 1 else "").strip())
            try:
                p_val = float(row[2]) if len(row) > 2 and str(row[2]).strip() != "" else None
                if p_val and p_val > 0:
                    last_power = p_val
            except Exception:
                p_val = None

            point_str = str(row[3] if len(row) > 3 else "").strip().replace(" ", "")
            axis_str = str(row[4] if len(row) > 4 else "").strip().lower()
            cause_str = str(row[8] if len(row) > 8 else "").strip()
            action_str = str(row[9] if len(row) > 9 else "").strip()

            if (
                len(equip_name) > 28
                or any(g in equip_name for g in ["점검실시", "업데이트", "페이지", "결재", "비고"])
                or equip_name == "반기"
            ):
                current_equip = None
                continue

            has_point = (
                any(k in point_str for k in ["전동기", "펌프", "송풍기", "부하", "가동정지", "미가동", "정비", "수리", "고장", "휴지", "대기"])
                or bool(re.match(r"^[1-4]", point_str))
            )

            if equip_name and equip_name != "설비명":
                if equip_name not in parsed_by_sheet[sheet_period]:
                    parsed_by_sheet[sheet_period][equip_name] = {
                        "name": equip_name,
                        "period": sheet_period,
                        "power": last_power,
                        "points": {
                            "m_de": {"h": 0.0, "v": 0.0, "a": 0.0},
                            "m_nde": {"h": 0.0, "v": 0.0, "a": 0.0},
                            "p_de": {"h": 0.0, "v": 0.0, "a": 0.0},
                            "p_nde": {"h": 0.0, "v": 0.0, "a": 0.0},
                        },
                        "summary": "",
                        "action": "",
                    }
                current_equip = parsed_by_sheet[sheet_period][equip_name]
                if p_val and p_val > 0:
                    current_equip["power"] = p_val

            if not current_equip:
                continue

            if "가동정지" in point_str or "미가동" in point_str:
                if "가동정지" not in current_equip["summary"]:
                    current_equip["summary"] = "가동정지" + (f" ({cause_str})" if cause_str else "")

            try:
                rms_val = float(row[5]) if len(row) > 5 and str(row[5]).strip() != "" else None
            except Exception:
                rms_val = None

            if has_point and rms_val is not None and 0 <= rms_val < 50:
                pts = current_equip["points"]
                tp = None
                if "전동기" in point_str and "반부하" in point_str:
                    tp = pts["m_nde"]
                elif "전동기" in point_str or point_str.startswith("1") or point_str.startswith("2"):
                    tp = pts["m_nde"] if "반부하" in point_str else pts["m_de"]
                elif any(k in point_str for k in ["펌프", "송풍기"]) or point_str.startswith("3") or point_str.startswith("4"):
                    tp = pts["p_nde"] if "반부하" in point_str else pts["p_de"]

                if tp is not None:
                    if axis_str in ("x", "h"):
                        tp["h"] = max(tp["h"], rms_val)
                    elif axis_str in ("y", "v"):
                        tp["v"] = max(tp["v"], rms_val)
                    elif axis_str in ("z", "a"):
                        tp["a"] = max(tp["a"], rms_val)

            if cause_str and cause_str != "-" and cause_str not in current_equip["summary"]:
                current_equip["summary"] += (", " if current_equip["summary"] else "") + cause_str
            if action_str and action_str not in ("양호", "-") and action_str not in current_equip["action"]:
                current_equip["action"] += (", " if current_equip["action"] else "") + action_str

    result_items = []
    sheet_counts = {}
    for p_key, eq_map in parsed_by_sheet.items():
        if not eq_map:
            continue
        sheet_counts[p_key] = len(eq_map)
        for eq in eq_map.values():
            p = eq["points"]
            max_rms = max(
                p["m_de"]["h"], p["m_de"]["v"], p["m_de"]["a"],
                p["m_nde"]["h"], p["m_nde"]["v"], p["m_nde"]["a"],
                p["p_de"]["h"], p["p_de"]["v"], p["p_de"]["a"],
                p["p_nde"]["h"], p["p_nde"]["v"], p["p_nde"]["a"],
            )
            grp, grd = calc_iso_grade_py(eq["power"], max_rms, "rigid")
            result_items.append({
                "id": f"{upload_office}_{p_key}_{eq['name']}",
                "office": upload_office,
                "period": p_key,
                "name": eq["name"],
                "power": eq["power"],
                "group": grp,
                "foundation": "rigid",
                "overallGrade": "A" if max_rms == 0 else grd,
                "maxRms": max_rms,
                "points": p,
                "summary": (eq["summary"] or "가동정지") if max_rms == 0 else (eq["summary"] or "측정 완료"),
                "action": (eq["action"] or "정비 대기") if max_rms == 0 else (eq["action"] or "추세 관리"),
            })

    return result_items, new_periods, sheet_counts


def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS equipment_records (
            id TEXT PRIMARY KEY,
            office TEXT,
            period TEXT,
            name TEXT,
            data_json TEXT,
            updated_at TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS system_meta (
            meta_key TEXT PRIMARY KEY,
            meta_value TEXT
        )
    """)
    c.execute("SELECT meta_value FROM system_meta WHERE meta_key = 'offices'")
    if not c.fetchone():
        c.execute("INSERT INTO system_meta VALUES ('offices', ?)", (json.dumps(DEFAULT_OFFICES, ensure_ascii=False),))
    c.execute("SELECT meta_value FROM system_meta WHERE meta_key = 'periods'")
    if not c.fetchone():
        c.execute("INSERT INTO system_meta VALUES ('periods', ?)", (json.dumps(DEFAULT_PERIODS, ensure_ascii=False),))
    conn.commit()
    conn.close()


def load_server_data():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT data_json FROM equipment_records")
    rows = c.fetchall()
    equip_list = []
    for r in rows:
        try:
            equip_list.append(json.loads(r[0]))
        except Exception:
            pass

    c.execute("SELECT meta_value FROM system_meta WHERE meta_key = 'offices'")
    off_row = c.fetchone()
    offices = json.loads(off_row[0]) if off_row else DEFAULT_OFFICES

    c.execute("SELECT meta_value FROM system_meta WHERE meta_key = 'periods'")
    per_row = c.fetchone()
    periods = json.loads(per_row[0]) if per_row else DEFAULT_PERIODS

    conn.close()
    return equip_list, offices, periods


def upsert_items_to_server_db(equip_items, periods=None, replace_office_periods=None):
    """서버 SQLite DB에 직접 저장 (타 컴퓨터 접속 시 즉시 조회됨)"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if replace_office_periods:
        for off_name, per_name in replace_office_periods:
            c.execute("DELETE FROM equipment_records WHERE office = ? AND period = ?", (off_name, per_name))

    for item in equip_items:
        rec_id = f"{item.get('office')}_{item.get('period')}_{item.get('name')}"
        item["id"] = rec_id
        c.execute(
            "INSERT OR REPLACE INTO equipment_records (id, office, period, name, data_json, updated_at) VALUES (?, ?, ?, ?, ?, ?)",
            (
                rec_id,
                item.get("office", "에너지사업소"),
                item.get("period", "2026년 상반기"),
                item.get("name", "미지정"),
                json.dumps(item, ensure_ascii=False),
                now_str,
            ),
        )

    if periods:
        c.execute("INSERT OR REPLACE INTO system_meta VALUES ('periods', ?)", (json.dumps(periods, ensure_ascii=False),))

    conn.commit()
    conn.close()


def delete_office_period_from_server(target_office, target_period):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    if target_office == "ALL":
        c.execute("DELETE FROM equipment_records WHERE period = ?", (target_period,))
    else:
        c.execute("DELETE FROM equipment_records WHERE office = ? AND period = ?", (target_office, target_period))
    conn.commit()
    conn.close()


init_db()
server_equip_data, server_offices, server_periods = load_server_data()

# ==============================================================================
# 2. 좌측 사이드바: 서버 DB 직결 엑셀 업로드 (타 PC 실시간 공유용)
# ==============================================================================
with st.sidebar:
    st.markdown("### 🏢 18개 사업소 통합 서버 DB")
    st.success(f"🟢 서버 DB 정상 가동 중 (총 **{len(server_equip_data)}**건 저장됨)")
    st.caption("여기에 엑셀을 올리면 **모든 컴퓨터(타 사업소)**에서 즉시 똑같이 조회됩니다!")

    office_names = [o["name"] for o in server_offices]
    selected_my_office = st.selectbox(
        "👤 1. 업로드/조회할 사업소 선택",
        options=office_names,
        index=office_names.index("에너지사업소") if "에너지사업소" in office_names else 0
    )

    selected_default_period = st.selectbox(
        "📅 2. 기본 측정시기 선택",
        options=server_periods,
        index=0
    )

    st.markdown("---")
    st.markdown("#### 📥 서버 DB에 엑셀 직접 업로드 (권장)")
    direct_excel_file = st.file_uploader(
        f"[{selected_my_office}] 현장 엑셀(.xlsx) 선택",
        type=["xlsx", "xls"],
        key="direct_server_excel"
    )

    if direct_excel_file is not None:
        if st.button("🚀 서버 DB에 즉시 저장 (전 컴퓨터 공유)", use_container_width=True, type="primary"):
            if not HAS_PANDAS:
                st.error("서버에 pandas/openpyxl 패키지가 없습니다. 우측 대시보드에서 업로드 후 아래 동기화 박스를 이용해 주세요.")
            else:
                try:
                    items, updated_periods, counts_map = parse_excel_directly_in_python(
                        direct_excel_file, selected_my_office, selected_default_period, server_periods
                    )
                    if not items:
                        st.warning("엑셀에서 인식된 설비가 없습니다. 파일 형식을 확인해 주세요.")
                    else:
                        replace_pairs = [(selected_my_office, p_k) for p_k in counts_map.keys()]
                        upsert_items_to_server_db(items, updated_periods, replace_office_periods=replace_pairs)
                        summary_txt = ", ".join([f"{k}: {v}대" for k, v in counts_map.items()])
                        st.success(f"✅ [{selected_my_office}] 서버 DB 저장 완료! ({summary_txt})")
                        st.rerun()
                except Exception as e:
                    st.error(f"엑셀 서버 저장 오류: {e}")

    st.markdown("---")
    with st.expander("🔄 우측 웹 화면 수정분 서버 DB 반영 / 초기화"):
        st.caption("우측 대시보드 안에서 직접 등록·수정한 경우, 우측 상단 **[서버반영 코드복사]** 버튼을 누른 뒤 아래 칸에 붙여넣고 저장하세요.")
        pasted_json = st.text_area("서버 반영 데이터 붙여넣기", height=80, placeholder="여기에 붙여넣기 (Ctrl+V)")
        if st.button("💾 붙여넣은 데이터 서버 DB 반영", use_container_width=True):
            try:
                payload = json.loads(pasted_json)
                items = payload.get("equipmentData", [])
                pers = payload.get("periodList", server_periods)
                conn = sqlite3.connect(DB_FILE)
                conn.execute("DELETE FROM equipment_records")
                conn.commit()
                conn.close()
                upsert_items_to_server_db(items, pers)
                st.success(f"서버 DB 동기화 완료! (총 {len(items)}건)")
                st.rerun()
            except Exception as e:
                st.error(f"데이터 형식 오류: {e}")

        if st.button(f"🗑️ [{selected_my_office} - {selected_default_period}] 서버 DB 비우기", use_container_width=True):
            delete_office_period_from_server(selected_my_office, selected_default_period)
            st.warning("선택한 사업소·시기의 서버 데이터가 삭제되었습니다.")
            st.rerun()

    # 전체 서버 DB 백업 다운로드
    full_backup_json = json.dumps({
        "targetOffice": "ALL",
        "equipmentData": server_equip_data,
        "periodList": server_periods,
        "savedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }, ensure_ascii=False, indent=2)

    st.download_button(
        label="💾 서버 DB 전체 백업 파일 받기 (.json)",
        data=full_backup_json,
        file_name=f"BECO_VIB_서버DB백업_{datetime.now().strftime('%Y%m%d')}.json",
        mime="application/json",
        use_container_width=True
    )

# ==============================================================================
# 3. 메인 웹 대시보드 (서버 DB 초기 주입 + 18개 사업소 + 전사 비교 + AI 진단)
# ==============================================================================
initial_office_val = selected_my_office

dashboard_html = r"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>부산환경공단 18개 사업소 회전설비 진동진단 통합 서버 시스템</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&family=Noto+Sans+KR:wght@400;500;700;800&display=swap" rel="stylesheet">
  <style>
    body { background-color: #0b1120; color: #f1f5f9; font-family: 'Noto Sans KR', 'Inter', sans-serif; }
    .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
    .custom-scrollbar::-webkit-scrollbar-track { background: #1e293b; }
    .custom-scrollbar::-webkit-scrollbar-thumb { background: #475569; border-radius: 4px; }
    .glass-card { background: rgba(30, 41, 59, 0.75); backdrop-filter: blur(12px); border: 1px solid rgba(255, 255, 255, 0.08); }
    .grade-badge-A { background: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.4); }
    .grade-badge-B { background: rgba(14, 165, 233, 0.15); color: #38bdf8; border: 1px solid rgba(14, 165, 233, 0.4); }
    .grade-badge-C { background: rgba(245, 158, 11, 0.15); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.4); }
    .grade-badge-D { background: rgba(239, 68, 68, 0.18); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.45); }
  </style>
</head>
<body class="min-h-screen custom-scrollbar text-slate-200">

  <!-- Top Sync Notice Banner (브라우저에만 있고 서버에 아직 안 올라간 데이터가 있을 때 자동 안내) -->
  <div id="unsyncedAlertBanner" class="hidden bg-amber-500/20 border-b border-amber-500/50 px-4 py-2 text-xs text-amber-200">
    <div class="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-2">
      <div>
        <i class="fa-solid fa-triangle-exclamation text-amber-400 mr-1.5"></i>
        <strong>알림:</strong> 현재 내 브라우저에만 임시 저장된 변경사항이 있습니다. <strong>다른 컴퓨터에서도 보이게 하려면</strong> 좌측 사이드바에서 엑셀을 직접 업로드하거나 우측 버튼을 눌러 서버 DB에 반영하세요!
      </div>
      <button onclick="copyServerSyncCode()" class="bg-amber-500 hover:bg-amber-400 text-slate-950 font-extrabold px-3 py-1 rounded transition shrink-0">
        <i class="fa-solid fa-copy mr-1"></i>서버반영 코드 복사하기
      </button>
    </div>
  </div>

  <!-- Top Header -->
  <header class="glass-card sticky top-0 z-40 border-b border-slate-700/60 bg-slate-900/95 px-4 py-3 sm:px-6">
    <div class="max-w-7xl mx-auto flex flex-col lg:flex-row items-center justify-between gap-3">
      <div class="flex items-center gap-3">
        <div class="h-10 w-10 rounded-xl bg-gradient-to-br from-cyan-500 to-blue-600 flex items-center justify-center text-white shadow-lg">
          <i class="fa-solid fa-server text-lg"></i>
        </div>
        <div>
          <div class="flex items-center gap-2 flex-wrap">
            <h1 class="text-lg sm:text-xl font-bold tracking-tight text-white">BECO-VIB Enterprise Server</h1>
            <span class="text-xs font-semibold px-2 py-0.5 rounded-full bg-cyan-950 text-cyan-400 border border-cyan-800">18개 사업소 통합</span>
            <span id="serverSyncBadge" class="text-[11px] font-semibold px-2.5 py-0.5 rounded-full bg-emerald-950 text-emerald-300 border border-emerald-700">
              <i class="fa-solid fa-database mr-1"></i>서버 DB 저장건수: <strong id="serverDbCountLabel">0</strong>건
            </span>
          </div>
          <p class="text-xs text-slate-400">부산환경공단 전 사업소 회전설비 12축 정밀진단 · AI 결함추론 · 반기별 수선성과 통합 플랫폼</p>
        </div>
      </div>

      <!-- Right Action Buttons -->
      <div class="flex items-center gap-2 flex-wrap justify-end">
        <button onclick="openOfficeOverviewModal()" class="flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-500 text-white text-xs sm:text-sm font-bold py-2 px-3 rounded-lg transition shadow">
          <i class="fa-solid fa-building-shield"></i>
          <span>18개 사업소 총괄현황</span>
        </button>
        <button onclick="openRepairSummaryModal()" class="flex items-center gap-1.5 bg-amber-600 hover:bg-amber-500 text-white text-xs sm:text-sm font-bold py-2 px-3 rounded-lg transition shadow">
          <i class="fa-solid fa-wand-magic-sparkles"></i>
          <span>수선/AI 요약시트</span>
        </button>
        <button onclick="openUploadModal()" class="flex items-center gap-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs sm:text-sm font-bold py-2 px-3 rounded-lg transition shadow">
          <i class="fa-solid fa-file-excel"></i>
          <span>엑셀 업로드</span>
        </button>
        <button onclick="openAddModal()" class="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-medium py-2 px-3 rounded-lg transition shadow">
          <i class="fa-solid fa-plus"></i>
          <span>설비 직접등록</span>
        </button>
        <button id="btnCopySync" onclick="copyServerSyncCode()" class="flex items-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-cyan-300 text-xs sm:text-sm font-semibold py-2 px-3 rounded-lg border border-cyan-700/60 transition shadow" title="웹에서 수정한 데이터를 서버 DB에 반영하기 위해 클립보드에 복사합니다">
          <i class="fa-solid fa-cloud-arrow-up"></i>
          <span>서버반영 코드복사</span>
        </button>
      </div>
    </div>

    <!-- 18개 사업소 선택 바 & 측정시기 컨트롤 -->
    <div class="max-w-7xl mx-auto mt-3 pt-3 border-t border-slate-700/60 flex flex-col gap-2">
      <div class="flex items-center justify-between flex-wrap gap-2">
        <div class="flex items-center gap-1.5 text-xs text-slate-300 font-semibold">
          <i class="fa-solid fa-map-location-dot text-cyan-400"></i>
          <span>사업소 선택 (18개소):</span>
        </div>

        <div class="flex items-center gap-2 text-xs flex-wrap">
          <div class="flex items-center gap-1.5 bg-slate-800 px-2.5 py-1 rounded-lg border border-slate-700">
            <span class="text-slate-400"><i class="fa-solid fa-clock-rotate-left text-cyan-400 mr-1"></i>측정 시기:</span>
            <select id="periodFilterSelect" onchange="switchPeriod(this.value)" class="bg-transparent text-cyan-300 font-bold focus:outline-none cursor-pointer"></select>
          </div>
          <button onclick="openManagePeriodModal()" class="bg-slate-800 hover:bg-slate-700 text-slate-300 px-2.5 py-1 rounded border border-slate-700 transition">
            <i class="fa-regular fa-calendar-plus mr-1 text-cyan-400"></i>시기 관리
          </button>
          <button onclick="exportToCSV()" class="bg-slate-800 hover:bg-slate-700 text-emerald-400 font-semibold px-2.5 py-1 rounded border border-slate-700 transition">
            <i class="fa-solid fa-file-excel mr-1"></i>엑셀 내보내기
          </button>
          <button id="btnResetPeriod" onclick="resetCurrentSelectionData()" class="bg-slate-800 hover:bg-rose-950 text-slate-400 hover:text-rose-400 px-2.5 py-1 rounded border border-slate-700 transition text-[11px] font-semibold">
            <i class="fa-solid fa-eraser mr-1"></i>현재 사업소·시기 비우기(0)
          </button>
        </div>
      </div>

      <!-- 18개 사업소 탭 스크롤 영역 -->
      <div class="flex items-center gap-1.5 overflow-x-auto custom-scrollbar pb-1.5 pt-0.5" id="officeTabs"></div>
    </div>
  </header>

  <main class="max-w-7xl mx-auto px-4 sm:px-6 py-6 space-y-6">

    <!-- KPI Cards -->
    <section class="grid grid-cols-2 lg:grid-cols-5 gap-3.5 sm:gap-4">
      <div class="glass-card rounded-xl p-4 border border-slate-700/50">
        <div class="flex items-center justify-between text-slate-400 text-xs mb-1">
          <span id="kpiOfficeLabel">진단 대상 설비</span>
          <i class="fa-solid fa-cubes text-slate-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-total" class="text-2xl sm:text-3xl font-extrabold text-white">0</span>
          <span class="text-xs text-slate-400">대</span>
        </div>
        <p class="text-[11px] text-slate-400 mt-2 truncate">전동기·펌프/송풍기 12축 계측</p>
      </div>

      <div class="glass-card rounded-xl p-4 border border-emerald-500/20 bg-emerald-950/10">
        <div class="flex items-center justify-between text-emerald-400 text-xs mb-1">
          <span>양호 상태 (A·B)</span>
          <i class="fa-solid fa-circle-check text-emerald-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-ab" class="text-2xl sm:text-3xl font-extrabold text-emerald-400">0</span>
          <span class="text-xs text-emerald-300/70">대</span>
        </div>
        <p class="text-[11px] text-emerald-400/80 mt-2 font-medium"><span id="metric-ab-pct">0%</span> 정상 가동 유지</p>
      </div>

      <div class="glass-card rounded-xl p-4 border border-amber-500/20 bg-amber-950/10">
        <div class="flex items-center justify-between text-amber-400 text-xs mb-1">
          <span>보수필요 (C등급)</span>
          <i class="fa-solid fa-triangle-exclamation text-amber-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-c" class="text-2xl sm:text-3xl font-extrabold text-amber-400">0</span>
          <span class="text-xs text-amber-300/70">대</span>
        </div>
        <p class="text-[11px] text-amber-400/80 mt-2">추세관리 및 계획보수</p>
      </div>

      <div class="glass-card rounded-xl p-4 border border-red-500/30 bg-red-950/20">
        <div class="flex items-center justify-between text-red-400 text-xs mb-1">
          <span>즉시점검 (D등급)</span>
          <i class="fa-solid fa-circle-exclamation text-red-400 animate-pulse"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-d" class="text-2xl sm:text-3xl font-extrabold text-red-400">0</span>
          <span class="text-xs text-red-300/70">대</span>
        </div>
        <p class="text-[11px] text-red-300 mt-2 font-medium">긴급 점검 및 오버홀</p>
      </div>

      <div class="glass-card rounded-xl p-4 border border-slate-700/50 col-span-2 lg:col-span-1">
        <div class="flex items-center justify-between text-slate-400 text-xs mb-1">
          <span>정비/미가동 설비</span>
          <i class="fa-solid fa-wrench text-slate-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-stop" class="text-2xl sm:text-3xl font-extrabold text-slate-300">0</span>
          <span class="text-xs text-slate-400">대</span>
        </div>
        <p id="metric-stop-desc" class="text-[11px] text-slate-400 mt-2 truncate">가동정지 및 수리 대기</p>
      </div>
    </section>

    <!-- Chart & ISO Guide -->
    <section class="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div class="lg:col-span-2 glass-card rounded-xl p-5 border border-slate-700/60">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-4 border-b border-slate-700/60 gap-3">
          <div>
            <h2 class="text-base font-semibold text-white flex items-center gap-2">
              <i class="fa-solid fa-chart-line text-cyan-400"></i>
              <span id="chartMainTitle">주요 설비 반기별 진동 추세 (Trend)</span>
            </h2>
            <p class="text-xs text-slate-400" id="chartSubTitle">C·D등급 이력 설비 또는 개별 설비의 반기별 진동 변화 추이</p>
          </div>

          <div class="flex flex-wrap items-center gap-2">
            <select id="trendTargetFilter" onchange="renderChart()" class="bg-slate-800 border border-slate-700 text-cyan-300 text-xs font-semibold rounded-lg px-2.5 py-1.5 focus:outline-none">
              <option value="CD_ONLY">🚨 C·D등급(주의/위험) 이력 설비만</option>
              <option value="TOP_RMS">🔥 최대 진동 상위 7대</option>
            </select>

            <div class="flex items-center gap-1 bg-slate-800 p-1 rounded-lg border border-slate-700 text-xs">
              <button id="btnModeTrend" onclick="setChartMode('trend')" class="px-2.5 py-1 rounded font-bold bg-cyan-600 text-white transition">반기별 추이선</button>
              <button id="btnModeBar" onclick="setChartMode('bar')" class="px-2.5 py-1 rounded font-medium text-slate-300 hover:text-white transition">현재 상태 막대</button>
            </div>
          </div>
        </div>
        <div class="pt-4 h-[290px] w-full">
          <canvas id="vibrationChart"></canvas>
        </div>
      </div>

      <!-- ISO 20816-3 Reference -->
      <div class="glass-card rounded-xl p-5 border border-slate-700/60 flex flex-col justify-between">
        <div>
          <div class="flex items-center justify-between pb-3 border-b border-slate-700/60">
            <h3 class="text-sm font-semibold text-white flex items-center gap-2">
              <i class="fa-solid fa-book-bookmark text-amber-400"></i>
              ISO 20816-3 진동 판정 기준표
            </h3>
            <span class="text-[10px] text-slate-400">속도 RMS(mm/s)</span>
          </div>
          <div class="mt-4 space-y-2.5 text-xs">
            <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
              <div class="flex justify-between font-bold text-cyan-300"><span>Group 1 (&gt; 300 kW)</span><span class="text-[11px] text-slate-400">대형 송풍기·순환펌프</span></div>
              <div class="grid grid-cols-3 gap-1 mt-1.5 text-center text-[11px]">
                <div class="p-1 rounded grade-badge-A">A/B ≤ 4.5</div>
                <div class="p-1 rounded grade-badge-C">C ≤ 7.1</div>
                <div class="p-1 rounded grade-badge-D">D &gt; 7.1</div>
              </div>
            </div>
            <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
              <div class="flex justify-between font-bold text-emerald-300"><span>Group 2 (15 ~ 300 kW)</span><span class="text-[11px] text-slate-400">급수·슬러지·히트펌프</span></div>
              <div class="grid grid-cols-3 gap-1 mt-1.5 text-center text-[11px]">
                <div class="p-1 rounded grade-badge-A">A/B ≤ 2.8</div>
                <div class="p-1 rounded grade-badge-C">C ≤ 4.5</div>
                <div class="p-1 rounded grade-badge-D">D &gt; 4.5</div>
              </div>
            </div>
            <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
              <div class="flex justify-between font-bold text-purple-300"><span>Group 4 (&lt; 15 kW)</span><span class="text-[11px] text-slate-400">소형 냉각수·약액펌프</span></div>
              <div class="grid grid-cols-3 gap-1 mt-1.5 text-center text-[11px]">
                <div class="p-1 rounded grade-badge-A">A/B ≤ 2.8</div>
                <div class="p-1 rounded grade-badge-C">C ≤ 4.5</div>
                <div class="p-1 rounded grade-badge-D">D &gt; 4.5</div>
              </div>
            </div>
          </div>
        </div>
        <div class="mt-4 pt-3 border-t border-slate-700/60 text-[11px] text-slate-400 flex justify-between">
          <span>* AI 12축 방향별(X/Y/Z) 결함 자동 추론</span>
          <span class="text-cyan-400 font-semibold">18개 사업소 표준</span>
        </div>
      </div>
    </section>

    <!-- Main Equipment Table -->
    <section class="glass-card rounded-xl border border-slate-700/60 overflow-hidden">
      <div class="p-4 sm:p-5 border-b border-slate-700/60 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 class="text-base font-bold text-white flex items-center gap-2">
            <i class="fa-solid fa-list-check text-cyan-400"></i>
            설비별 4개 지점(12축) 정밀 진단 및 AI 권고 대장
          </h2>
          <div class="flex items-center gap-2 mt-0.5">
            <p id="tableSubtitle" class="text-xs text-slate-400">사업소별 회전기계 상태진단 및 정비 대장</p>
            <span class="text-slate-600">|</span>
            <span id="filteredCount" class="text-xs font-mono text-cyan-400 font-semibold">총 0개 설비</span>
          </div>
        </div>

        <div class="flex flex-wrap items-center gap-2">
          <select id="officeFilterDropdown" onchange="switchOffice(this.value)" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2"></select>
          <select id="groupFilter" onchange="renderTable()" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2">
            <option value="ALL">전체 용량 구분</option>
            <option value="G1">Group 1 (&gt;300kW)</option>
            <option value="G2">Group 2 (15~300kW)</option>
            <option value="G4">Group 4 (&lt;15kW)</option>
          </select>
          <select id="statusFilter" onchange="renderTable()" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2">
            <option value="ALL">전체 평가 등급</option>
            <option value="D">D (즉시점검)</option>
            <option value="C">C (보수필요)</option>
            <option value="AB">A/B (양호)</option>
          </select>
          <div class="relative w-full sm:w-44">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
            <input type="text" id="searchInput" oninput="renderTable()" placeholder="설비명/원인 검색..." class="w-full bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg pl-8 pr-3 py-2">
          </div>
        </div>
      </div>

      <div class="overflow-x-auto custom-scrollbar">
        <table class="w-full text-left border-collapse text-xs">
          <thead>
            <tr class="bg-slate-800/80 text-slate-400 border-b border-slate-700 text-[11px] uppercase font-semibold">
              <th class="py-3 px-3 cursor-pointer hover:text-white" onclick="toggleSort('office')">사업소 <i class="fa-solid fa-sort ml-1 opacity-50"></i></th>
              <th class="py-3 px-3 cursor-pointer hover:text-white" onclick="toggleSort('period')">측정시기 <i class="fa-solid fa-sort ml-1 opacity-50"></i></th>
              <th class="py-3 px-4 cursor-pointer hover:text-white" onclick="toggleSort('name')">설비명 <i class="fa-solid fa-sort ml-1 opacity-50"></i></th>
              <th class="py-3 px-3">정격용량</th>
              <th class="py-3 px-3 text-center">전동기 Max</th>
              <th class="py-3 px-3 text-center">펌프/팬 Max</th>
              <th class="py-3 px-3 text-center cursor-pointer hover:text-white" onclick="toggleSort('maxRms')">최대 RMS <i class="fa-solid fa-sort ml-1 opacity-50"></i></th>
              <th class="py-3 px-3 text-center cursor-pointer hover:text-white" onclick="toggleSort('grade')">종합평가 <i class="fa-solid fa-sort ml-1 opacity-50"></i></th>
              <th class="py-3 px-4">🤖 AI 결함 추정 및 개선 권고</th>
              <th class="py-3 px-4">현장 진단 요약 및 정비대책</th>
              <th class="py-3 px-3 text-center">관리</th>
            </tr>
          </thead>
          <tbody id="tableBody" class="divide-y divide-slate-800 text-slate-300"></tbody>
        </table>
      </div>
    </section>
  </main>

  <!-- MODAL A: 18개 사업소 총괄 현황판 -->
  <div id="officeOverviewModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-5xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeOfficeOverviewModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
      <div class="flex items-center gap-3 mb-5 border-b border-slate-700 pb-4">
        <div class="h-10 w-10 rounded-xl bg-indigo-500/20 text-indigo-400 border border-indigo-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-building-shield"></i>
        </div>
        <div>
          <h3 class="text-base sm:text-lg font-bold text-white">부산환경공단 18개 사업단·사업소 진동진단 총괄 현황판</h3>
          <p class="text-xs text-slate-400">현재 측정시기 기준 각 사업소별 데이터 등록 대수 및 위험(C·D등급) 설비 분포를 한눈에 비교합니다.</p>
        </div>
      </div>
      <div id="officeOverviewGrid" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 text-xs"></div>
    </div>
  </div>

  <!-- MODAL B: 수선/AI 요약시트 -->
  <div id="repairSummaryModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-5xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeRepairSummaryModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
      <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-5 border-b border-slate-700 pb-4">
        <div class="flex items-center gap-3">
          <div class="h-10 w-10 rounded-xl bg-amber-500/20 text-amber-400 border border-amber-500/40 flex items-center justify-center text-lg">
            <i class="fa-solid fa-screwdriver-wrench"></i>
          </div>
          <div>
            <h3 class="text-base sm:text-lg font-bold text-white">반기별 설비 수선/개선 성과 및 AI 정밀진단 요약시트</h3>
            <p class="text-xs text-slate-400">상반기 ↔ 하반기 진동 변화를 비교하여 수선 완료(개선) 설비와 AI 결함 추정 결과를 요약합니다.</p>
          </div>
        </div>
        <button onclick="exportRepairSummaryExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold px-3.5 py-2 rounded-lg transition flex items-center gap-1.5 shadow">
          <i class="fa-solid fa-file-excel"></i><span>요약시트 엑셀 다운로드</span>
        </button>
      </div>

      <div class="flex flex-wrap items-center justify-between gap-3 bg-slate-800/80 p-3 rounded-xl border border-slate-700 mb-4 text-xs">
        <div class="flex items-center gap-2 flex-wrap">
          <span class="text-slate-300 font-semibold">비교 기준 시기:</span>
          <select id="sumBasePeriod" onchange="renderRepairSummarySheet()" class="bg-slate-900 border border-slate-600 text-slate-200 font-bold rounded px-2.5 py-1.5"></select>
          <i class="fa-solid fa-arrow-right text-slate-400 mx-1"></i>
          <select id="sumTargetPeriod" onchange="renderRepairSummarySheet()" class="bg-slate-900 border border-cyan-500 text-cyan-300 font-bold rounded px-2.5 py-1.5"></select>
        </div>
        <div class="flex items-center gap-1.5">
          <button onclick="setRepairFilter('ALL')" id="btnRepAll" class="px-2.5 py-1 rounded font-bold bg-cyan-600 text-white">전체 비교</button>
          <button onclick="setRepairFilter('IMPROVED')" id="btnRepImproved" class="px-2.5 py-1 rounded font-medium bg-slate-900 text-emerald-400 border border-emerald-700/50">🛠️ 수선/개선 완료만</button>
          <button onclick="setRepairFilter('NEED_REPAIR')" id="btnRepNeed" class="px-2.5 py-1 rounded font-medium bg-slate-900 text-rose-400 border border-rose-700/50">🚨 개선 필요(C·D)만</button>
        </div>
      </div>

      <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4 text-xs">
        <div class="p-3 rounded-xl bg-emerald-950/20 border border-emerald-500/30">
          <span class="text-emerald-400 font-semibold block">🛠️ 수선/개선 완료</span>
          <span id="sumKpiImproved" class="text-2xl font-black text-emerald-400 mt-1 block">0 대</span>
        </div>
        <div class="p-3 rounded-xl bg-cyan-950/20 border border-cyan-500/30">
          <span class="text-cyan-400 font-semibold block">📉 평균 진동 저감률</span>
          <span id="sumKpiReduction" class="text-2xl font-black text-cyan-300 mt-1 block">0 %</span>
        </div>
        <div class="p-3 rounded-xl bg-rose-950/20 border border-rose-500/30">
          <span class="text-rose-400 font-semibold block">🚨 현재 개선 필요 (C·D)</span>
          <span id="sumKpiNeed" class="text-2xl font-black text-rose-400 mt-1 block">0 대</span>
        </div>
        <div class="p-3 rounded-xl bg-slate-800/80 border border-slate-700">
          <span class="text-slate-300 font-semibold block">⏸️ 가동정지 설비</span>
          <span id="sumKpiStopped" class="text-2xl font-black text-slate-200 mt-1 block">0 대</span>
        </div>
      </div>

      <div class="overflow-x-auto custom-scrollbar border border-slate-700 rounded-xl">
        <table class="w-full text-left border-collapse text-xs">
          <thead>
            <tr class="bg-slate-800 text-slate-300 border-b border-slate-700 text-[11px]">
              <th class="py-2.5 px-3">사업소</th>
              <th class="py-2.5 px-3">설비명 (용량)</th>
              <th class="py-2.5 px-3 text-center">이전 반기</th>
              <th class="py-2.5 px-3 text-center">현재 반기</th>
              <th class="py-2.5 px-3 text-center">수선/증감 결과</th>
              <th class="py-2.5 px-3">🤖 AI 진단 추정 부위 및 개선 권고</th>
            </tr>
          </thead>
          <tbody id="repairSummaryTableBody" class="divide-y divide-slate-800 text-slate-300"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- MODAL C: 사업소별 엑셀 스마트 업로드 -->
  <div id="uploadModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-lg rounded-2xl shadow-2xl p-6 relative">
      <button onclick="closeUploadModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
      <div class="flex items-center gap-3 mb-4 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-blue-500/20 text-blue-400 border border-blue-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-file-arrow-up"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">사업소별 현장 엑셀 업로드</h3>
          <p class="text-xs text-slate-400">💡 <strong>타 컴퓨터와 즉시 공유</strong>하려면 <strong>좌측 사이드바</strong>에서 엑셀을 업로드하세요!</p>
        </div>
      </div>

      <div class="space-y-3.5 text-xs">
        <div class="flex items-center justify-between bg-slate-800/90 p-2.5 rounded-lg border border-slate-700">
          <label class="text-slate-300 font-bold"><i class="fa-solid fa-building text-amber-400 mr-1.5"></i>업로드 대상 사업소:</label>
          <select id="uploadOfficeSelect" class="bg-slate-900 border border-amber-500/60 rounded px-2.5 py-1 text-amber-300 font-bold focus:outline-none"></select>
        </div>

        <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
          <label class="text-slate-300 font-medium">기본 반영 측정시기 (시트명 미표기 시):</label>
          <select id="uploadPeriodSelect" class="bg-slate-900 border border-slate-600 rounded px-2.5 py-1 text-cyan-300 font-bold focus:outline-none"></select>
        </div>

        <div id="dropZone" onclick="document.getElementById('fileInput').click()" class="border-2 border-dashed border-slate-600 hover:border-cyan-500 rounded-2xl p-8 text-center cursor-pointer transition bg-slate-800/40 hover:bg-slate-800/70">
          <i class="fa-solid fa-cloud-arrow-up text-3xl text-cyan-400 mb-2"></i>
          <p class="text-slate-200 font-bold mb-1">현장 엑셀 파일(.xlsx)을 여기로 드래그하거나 클릭하세요</p>
          <p class="text-slate-400 text-[11px]">세로형 다중 행(병합셀) 및 가동정지 설비 100% 자동 인식</p>
          <input type="file" id="fileInput" accept=".xlsx, .xls, .csv" onchange="handleFileSelect(event)" class="hidden">
        </div>

        <div id="uploadStatusBox" class="hidden p-3 rounded-lg border text-xs"></div>
      </div>
    </div>
  </div>

  <!-- MODAL D: 설비 직접 등록 -->
  <div id="addModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-2xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeAddModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
      <h3 class="text-base font-bold text-white mb-4 border-b border-slate-700 pb-3">신규 설비 직접 등록 (18개 사업소 공통)</h3>
      <form id="addEquipForm" onsubmit="handleEquipmentSubmit(event)" class="space-y-4 text-xs">
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div>
            <label class="block text-slate-300 font-medium mb-1">소속 사업소</label>
            <select id="inpOffice" required class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200"></select>
          </div>
          <div>
            <label class="block text-slate-300 font-medium mb-1">측정 시기</label>
            <select id="inpPeriod" required class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200"></select>
          </div>
          <div>
            <label class="block text-slate-300 font-medium mb-1">설비명</label>
            <input type="text" id="inpName" required placeholder="예: #1 유입펌프" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200">
          </div>
        </div>
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="block text-slate-300 font-medium mb-1">정격 용량 (kW)</label>
            <input type="number" id="inpPower" required step="0.1" value="75" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 font-mono">
          </div>
          <div>
            <label class="block text-slate-300 font-medium mb-1">기초 지지 조건</label>
            <select id="inpFoundation" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200">
              <option value="rigid">강성 기초 (Rigid)</option>
              <option value="flexible">탄성 기초 (Flexible)</option>
            </select>
          </div>
        </div>
        <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-slate-800/80 p-3 rounded-xl border border-slate-700">
          <div><span class="text-amber-300 font-bold block mb-1">1. M-DE (H/V/A)</span>
            <div class="grid grid-cols-3 gap-1">
              <input type="number" id="inp_m_de_h" step="0.01" value="2.1" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
              <input type="number" id="inp_m_de_v" step="0.01" value="1.5" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
              <input type="number" id="inp_m_de_a" step="0.01" value="1.1" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
            </div>
          </div>
          <div><span class="text-amber-300 font-bold block mb-1">2. M-NDE (H/V/A)</span>
            <div class="grid grid-cols-3 gap-1">
              <input type="number" id="inp_m_nde_h" step="0.01" value="1.8" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
              <input type="number" id="inp_m_nde_v" step="0.01" value="1.2" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
              <input type="number" id="inp_m_nde_a" step="0.01" value="0.9" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
            </div>
          </div>
          <div><span class="text-cyan-300 font-bold block mb-1">3. P-DE (H/V/A)</span>
            <div class="grid grid-cols-3 gap-1">
              <input type="number" id="inp_p_de_h" step="0.01" value="3.2" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
              <input type="number" id="inp_p_de_v" step="0.01" value="2.4" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
              <input type="number" id="inp_p_de_a" step="0.01" value="1.8" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
            </div>
          </div>
          <div><span class="text-cyan-300 font-bold block mb-1">4. P-NDE (H/V/A)</span>
            <div class="grid grid-cols-3 gap-1">
              <input type="number" id="inp_p_nde_h" step="0.01" value="2.0" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
              <input type="number" id="inp_p_nde_v" step="0.01" value="1.6" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
              <input type="number" id="inp_p_nde_a" step="0.01" value="1.2" class="bg-slate-900 border border-slate-700 rounded p-1 text-center font-mono">
            </div>
          </div>
        </div>
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <input type="text" id="inpSummary" placeholder="현장 진단 요약" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200">
          <input type="text" id="inpAction" placeholder="후속 정비 대책" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200">
        </div>
        <div class="flex justify-end gap-2 pt-3 border-t border-slate-700">
          <button type="button" onclick="closeAddModal()" class="px-4 py-2 bg-slate-800 text-slate-300 rounded-lg">취소</button>
          <button type="submit" class="px-5 py-2 bg-emerald-600 text-white font-bold rounded-lg">등록 완료</button>
        </div>
      </form>
    </div>
  </div>

  <!-- MODAL E: 측정시기 관리 -->
  <div id="managePeriodModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-md rounded-2xl shadow-2xl p-6 relative">
      <button onclick="closeManagePeriodModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
      <h3 class="text-base font-bold text-white mb-4">측정시기(점검주기) 추가 및 삭제</h3>
      <form id="addPeriodForm" onsubmit="handlePeriodSubmit(event)" class="flex gap-2 mb-4 text-xs">
        <input type="text" id="inpNewPeriodName" required placeholder="예: 2027년 상반기" class="flex-1 bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 font-bold">
        <button type="submit" class="px-4 py-2 bg-cyan-600 text-white font-bold rounded-lg">추가</button>
      </form>
      <div id="periodManageList" class="space-y-1.5 max-h-52 overflow-y-auto custom-scrollbar text-xs"></div>
    </div>
  </div>

  <!-- MODAL F: 상세 보기 -->
  <div id="detailModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-3xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeDetailModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
      <div id="detailModalContent"></div>
    </div>
  </div>

  <script>
    // 파이썬 서버 DB(SQLite)에서 주입된 초기 데이터
    const SERVER_INJECTED_DATA = __SERVER_EQUIP_JSON__;
    const SERVER_INJECTED_OFFICES = __SERVER_OFFICES_JSON__;
    const SERVER_INJECTED_PERIODS = __SERVER_PERIODS_JSON__;
    const INITIAL_SELECTED_OFFICE = "__INITIAL_OFFICE__";

    let officeList = SERVER_INJECTED_OFFICES;
    let periodList = SERVER_INJECTED_PERIODS;
    let currentPeriod = periodList[0] || '2026년 상반기';
    let currentOffice = INITIAL_SELECTED_OFFICE || '에너지사업소';
    let equipmentData = [];
    let currentChartMode = 'trend';
    let vibrationChart = null;
    let sortCol = 'name';
    let sortAsc = true;
    let currentRepairFilter = 'ALL';

    function loadDataWithServerMerge() {
      document.getElementById('serverDbCountLabel').innerText = SERVER_INJECTED_DATA.length;
      try {
        const localRaw = localStorage.getItem('becovib_enterprise_equip');
        const localPeriods = localStorage.getItem('becovib_enterprise_periods');
        if (localPeriods) {
          const parsedP = JSON.parse(localPeriods);
          parsedP.forEach(p => { if (!periodList.includes(p)) periodList.push(p); });
        }

        const mergedMap = {};
        // 1) 로컬 작업분을 먼저 넣고
        let localCount = 0;
        if (localRaw) {
          const parsedL = JSON.parse(localRaw);
          localCount = parsedL.length;
          parsedL.forEach(item => {
            const key = `${item.office}||${item.period}||${item.name}`;
            mergedMap[key] = item;
          });
        }
        // 2) 파이썬 서버 DB 데이터를 최신 기준으로 덮어쓰기
        SERVER_INJECTED_DATA.forEach(item => {
          const key = `${item.office}||${item.period}||${item.name}`;
          mergedMap[key] = item;
        });

        equipmentData = Object.values(mergedMap);
        if (equipmentData.length > SERVER_INJECTED_DATA.length) {
          document.getElementById('unsyncedAlertBanner').classList.remove('hidden');
        }
      } catch (e) {
        equipmentData = SERVER_INJECTED_DATA;
      }
    }

    function saveToLocalStorage() {
      try {
        localStorage.setItem('becovib_enterprise_equip', JSON.stringify(equipmentData));
        localStorage.setItem('becovib_enterprise_periods', JSON.stringify(periodList));
        if (equipmentData.length !== SERVER_INJECTED_DATA.length) {
          document.getElementById('unsyncedAlertBanner').classList.remove('hidden');
        }
      } catch (e) {}
    }

    function copyServerSyncCode() {
      const payload = JSON.stringify({
        targetOffice: currentOffice,
        equipmentData: equipmentData,
        periodList: periodList
      });
      navigator.clipboard.writeText(payload).then(() => {
        const btn = document.getElementById('btnCopySync');
        btn.innerHTML = '<i class="fa-solid fa-check mr-1 text-emerald-400"></i>복사완료! 좌측 사이드바에 붙여넣으세요';
        setTimeout(() => {
          btn.innerHTML = '<i class="fa-solid fa-cloud-arrow-up"></i><span>서버반영 코드복사</span>';
        }, 4000);
      });
    }

    function calculateIsoGrade(powerKw, rmsValue, foundation = 'rigid') {
      let group = 'G2';
      if (powerKw > 300) group = 'G1';
      else if (powerKw < 15) group = 'G4';

      let grade = 'A';
      const rms = parseFloat(rmsValue) || 0;
      if (group === 'G1') {
        if (foundation === 'rigid') {
          if (rms <= 2.3) grade = 'A';
          else if (rms <= 4.5) grade = 'B';
          else if (rms <= 7.1) grade = 'C';
          else grade = 'D';
        } else {
          if (rms <= 3.5) grade = 'A';
          else if (rms <= 7.1) grade = 'B';
          else if (rms <= 11.0) grade = 'C';
          else grade = 'D';
        }
      } else {
        if (foundation === 'rigid') {
          if (rms <= 1.4) grade = 'A';
          else if (rms <= 2.8) grade = 'B';
          else if (rms <= 4.5) grade = 'C';
          else grade = 'D';
        } else {
          if (rms <= 2.3) grade = 'A';
          else if (rms <= 4.5) grade = 'B';
          else if (rms <= 7.1) grade = 'C';
          else grade = 'D';
        }
      }
      return { group, grade, foundation };
    }

    function diagnoseByAI(item) {
      const maxRms = parseFloat(item.maxRms) || 0;
      if (maxRms === 0) {
        return {
          status: 'STOP',
          part: '가동정지 / 미가동',
          cause: item.summary && item.summary !== '측정 완료' ? item.summary : '운전 휴지 또는 정비 대기',
          recommendation: '재가동 전 절연저항 및 수동 회전 점검'
        };
      }
      const p = item.points || {
        m_de: { h: maxRms, v: 0, a: 0 }, m_nde: { h: 0, v: 0, a: 0 },
        p_de: { h: maxRms, v: 0, a: 0 }, p_nde: { h: 0, v: 0, a: 0 }
      };
      const candidates = [
        { key: 'm_de', label: '전동기 부하측(M-DE)', h: p.m_de.h||0, v: p.m_de.v||0, a: p.m_de.a||0, max: Math.max(p.m_de.h||0, p.m_de.v||0, p.m_de.a||0) },
        { key: 'm_nde', label: '전동기 반부하측(M-NDE)', h: p.m_nde.h||0, v: p.m_nde.v||0, a: p.m_nde.a||0, max: Math.max(p.m_nde.h||0, p.m_nde.v||0, p.m_nde.a||0) },
        { key: 'p_de', label: '펌프/팬 부하측(P-DE)', h: p.p_de.h||0, v: p.p_de.v||0, a: p.p_de.a||0, max: Math.max(p.p_de.h||0, p.p_de.v||0, p.p_de.a||0) },
        { key: 'p_nde', label: '펌프/팬 반부하측(P-NDE)', h: p.p_nde.h||0, v: p.p_nde.v||0, a: p.p_nde.a||0, max: Math.max(p.p_nde.h||0, p.p_nde.v||0, p.p_nde.a||0) }
      ];
      candidates.sort((a, b) => b.max - a.max);
      const worst = candidates[0];

      if (item.overallGrade === 'A' || item.overallGrade === 'B') {
        return {
          status: 'GOOD',
          part: maxRms >= 2.2 ? `${worst.label} (상한 근접)` : '전 부위 양호',
          cause: '정상 운전 허용 범위 내 가동 중',
          recommendation: '정기 그리스 급유 및 주기적 모니터링 유지'
        };
      }

      if (worst.a > 0 && worst.a >= worst.h * 0.85) {
        return {
          status: 'DANGER',
          part: `${worst.label} [축방향(Z)]`,
          cause: '커플링 축정렬 불량(Misalignment) 또는 축 휘어짐 추정',
          recommendation: '레이저 축정렬 교정 및 커플링 점검 요망'
        };
      } else if (worst.v > worst.h * 1.12) {
        return {
          status: 'WARNING',
          part: `${worst.label} [수직(Y)]`,
          cause: '기초 볼트 풀림 또는 베이스 구조적 이완(Looseness) 추정',
          recommendation: '앵커볼트 체결 확인 및 소프트풋 보강 필요'
        };
      } else {
        return {
          status: 'DANGER',
          part: `${worst.label} [수평(X)]`,
          cause: '회전체 불평형(Unbalance) 또는 베어링 마모 진동 추정',
          recommendation: `${worst.label} 베어링 교체 및 임펠러 밸런싱 점검 필요`
        };
      }
    }

    function renderPeriodSelectors() {
      const html = periodList.map(p => `<option value="${p}" ${p === currentPeriod ? 'selected' : ''}>${p}</option>`).join('');
      ['periodFilterSelect', 'inpPeriod', 'uploadPeriodSelect'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = html;
      });
    }

    function renderOfficeTabs() {
      const container = document.getElementById('officeTabs');
      const dropdown = document.getElementById('officeFilterDropdown');
      const inpOff = document.getElementById('inpOffice');
      const upOff = document.getElementById('uploadOfficeSelect');

      const totalCountAll = equipmentData.filter(e => e.period === currentPeriod).length;
      let tabsHtml = `
        <button onclick="switchOffice('ALL')" class="px-3 py-1.5 rounded-lg text-xs font-bold transition shrink-0 flex items-center gap-1.5 ${currentOffice === 'ALL' ? 'bg-cyan-600 text-white shadow' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'}">
          <span>전체 사업소</span>
          <span class="px-1.5 py-0.2 rounded-full text-[10px] bg-slate-950/60 text-cyan-300">${totalCountAll}</span>
        </button>
      `;

      officeList.forEach(off => {
        const isSel = currentOffice === off.name;
        const cnt = equipmentData.filter(e => e.office === off.name && e.period === currentPeriod).length;
        tabsHtml += `
          <button onclick="switchOffice('${off.name}')" class="px-2.5 py-1.5 rounded-lg text-xs font-semibold transition shrink-0 flex items-center gap-1.5 ${isSel ? 'bg-cyan-600 text-white shadow' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'}">
            <i class="fa-solid ${off.icon} ${isSel ? 'text-white' : off.color}"></i>
            <span>${off.name}</span>
            ${cnt > 0 ? `<span class="px-1.5 py-0.2 rounded-full text-[10px] bg-slate-950/70 text-emerald-300 font-mono">${cnt}</span>` : ''}
          </button>
        `;
      });
      container.innerHTML = tabsHtml;

      dropdown.innerHTML = `<option value="ALL">전체 18개 사업소</option>` + officeList.map(o => `<option value="${o.name}" ${o.name === currentOffice ? 'selected' : ''}>${o.name}</option>`).join('');

      const targetOff = currentOffice === 'ALL' ? '에너지사업소' : currentOffice;
      const offOptions = officeList.map(o => `<option value="${o.name}" ${o.name === targetOff ? 'selected' : ''}>${o.name}</option>`).join('');
      if (inpOff) inpOff.innerHTML = offOptions;
      if (upOff) upOff.innerHTML = offOptions;
    }

    function switchOffice(officeName) {
      currentOffice = officeName;
      renderOfficeTabs();
      updateDashboard();
    }

    function switchPeriod(val) {
      currentPeriod = val;
      renderOfficeTabs();
      updateDashboard();
    }

    let resetConfirmTimer = null;
    function resetCurrentSelectionData() {
      const btn = document.getElementById('btnResetPeriod');
      const targetDesc = currentOffice === 'ALL' ? `전체 사업소(${currentPeriod})` : `${currentOffice}(${currentPeriod})`;
      if (btn.dataset.confirming !== 'true') {
        btn.dataset.confirming = 'true';
        btn.className = 'bg-rose-700 text-white px-2.5 py-1 rounded border border-rose-500 text-[11px] font-bold animate-pulse';
        btn.innerHTML = `<i class="fa-solid fa-triangle-exclamation mr-1"></i>한 번 더 누르면 [${targetDesc}] 0대 초기화`;
        if (resetConfirmTimer) clearTimeout(resetConfirmTimer);
        resetConfirmTimer = setTimeout(() => {
          btn.dataset.confirming = 'false';
          btn.className = 'bg-slate-800 hover:bg-rose-950 text-slate-400 hover:text-rose-400 px-2.5 py-1 rounded border border-slate-700 text-[11px] font-semibold';
          btn.innerHTML = '<i class="fa-solid fa-eraser mr-1"></i>현재 사업소·시기 비우기(0)';
        }, 3500);
        return;
      }
      btn.dataset.confirming = 'false';
      btn.className = 'bg-slate-800 hover:bg-rose-950 text-slate-400 hover:text-rose-400 px-2.5 py-1 rounded border border-slate-700 text-[11px] font-semibold';
      btn.innerHTML = '<i class="fa-solid fa-eraser mr-1"></i>현재 사업소·시기 비우기(0)';

      equipmentData = equipmentData.filter(e => {
        const matchOff = (currentOffice === 'ALL' || e.office === currentOffice);
        const matchPer = (!e.period || e.period === currentPeriod);
        return !(matchOff && matchPer);
      });
      saveToLocalStorage();
      renderOfficeTabs();
      updateDashboard();
    }

    function updateDashboard() {
      const activeData = equipmentData.filter(e => {
        const offMatch = (currentOffice === 'ALL' || e.office === currentOffice);
        const perMatch = (!e.period || e.period === currentPeriod);
        return offMatch && perMatch;
      });

      document.getElementById('kpiOfficeLabel').innerText = currentOffice === 'ALL' ? '전사 진단 대상 설비' : `${currentOffice} 설비`;
      const totalUnits = activeData.length;
      const stopUnits = activeData.filter(e => (parseFloat(e.maxRms) || 0) === 0).length;
      const operated = activeData.filter(e => (parseFloat(e.maxRms) || 0) > 0);

      const gA = operated.filter(e => e.overallGrade === 'A').length;
      const gB = operated.filter(e => e.overallGrade === 'B').length;
      const gC = operated.filter(e => e.overallGrade === 'C').length;
      const gD = operated.filter(e => e.overallGrade === 'D').length;

      document.getElementById('metric-total').innerText = totalUnits;
      document.getElementById('metric-ab').innerText = gA + gB;
      document.getElementById('metric-c').innerText = gC;
      document.getElementById('metric-d').innerText = gD;
      document.getElementById('metric-stop').innerText = stopUnits;

      const stopNames = activeData.filter(e => (parseFloat(e.maxRms) || 0) === 0).map(e => e.name).slice(0, 4).join(', ');
      document.getElementById('metric-stop-desc').innerText = stopUnits > 0 ? stopNames : '미가동 설비 없음';
      document.getElementById('metric-ab-pct').innerText = (operated.length > 0 ? Math.round(((gA + gB) / operated.length) * 100) : 0) + '%';

      updateTrendFilterDropdown();
      renderTable();
      renderChart();
    }

    function updateTrendFilterDropdown() {
      const sel = document.getElementById('trendTargetFilter');
      if (!sel) return;
      const prev = sel.value || 'CD_ONLY';
      const names = Array.from(new Set(
        equipmentData.filter(e => currentOffice === 'ALL' || e.office === currentOffice).map(e => e.name)
      )).sort((a, b) => a.localeCompare(b, 'ko'));

      let html = `
        <option value="CD_ONLY">🚨 C·D등급(주의/위험) 이력 설비만</option>
        <option value="D_ONLY">⛔ D등급(즉시점검) 이력 설비만</option>
        <option value="TOP_RMS">🔥 최대 진동 상위 7대</option>
        <optgroup label="── 개별 설비 선택 ──">
      `;
      names.forEach(n => { html += `<option value="EQ:${n}">🔍 ${n}</option>`; });
      html += `</optgroup>`;
      sel.innerHTML = html;
      if ([...sel.options].some(o => o.value === prev)) sel.value = prev;
    }

    function setChartMode(mode) {
      currentChartMode = mode;
      document.getElementById('btnModeTrend').className = mode === 'trend' ? 'px-2.5 py-1 rounded font-bold bg-cyan-600 text-white' : 'px-2.5 py-1 rounded font-medium text-slate-300';
      document.getElementById('btnModeBar').className = mode === 'bar' ? 'px-2.5 py-1 rounded font-bold bg-cyan-600 text-white' : 'px-2.5 py-1 rounded font-medium text-slate-300';
      renderChart();
    }

    function renderChart() {
      const ctx = document.getElementById('vibrationChart').getContext('2d');
      if (vibrationChart) vibrationChart.destroy();

      const activeData = equipmentData.filter(e => (currentOffice === 'ALL' || e.office === currentOffice) && (!e.period || e.period === currentPeriod));

      if (currentChartMode === 'trend') {
        const filterMode = document.getElementById('trendTargetFilter')?.value || 'CD_ONLY';
        const equipMap = {};
        equipmentData.forEach(e => {
          if (currentOffice !== 'ALL' && e.office !== currentOffice) return;
          const key = `${e.office}||${e.name}`;
          if (!equipMap[key]) equipMap[key] = { name: e.name, office: e.office, maxEver: 0, hadCD: false, hadD: false, history: {} };
          const v = parseFloat(e.maxRms) || 0;
          if (v > equipMap[key].maxEver) equipMap[key].maxEver = v;
          if (e.overallGrade === 'C' || e.overallGrade === 'D') equipMap[key].hadCD = true;
          if (e.overallGrade === 'D') equipMap[key].hadD = true;
          if (e.period && v > 0) equipMap[key].history[e.period] = v;
        });

        let list = Object.values(equipMap);
        if (filterMode.startsWith('EQ:')) {
          const targetName = filterMode.replace('EQ:', '');
          list = list.filter(c => c.name === targetName);
        } else if (filterMode === 'D_ONLY') {
          list = list.filter(c => c.hadD).sort((a, b) => b.maxEver - a.maxEver).slice(0, 10);
        } else if (filterMode === 'CD_ONLY') {
          list = list.filter(c => c.hadCD).sort((a, b) => b.maxEver - a.maxEver).slice(0, 10);
        } else {
          list = list.sort((a, b) => b.maxEver - a.maxEver).slice(0, 7);
        }

        const palette = ['#ef4444', '#f59e0b', '#06b6d4', '#10b981', '#8b5cf6', '#ec4899', '#3b82f6'];
        vibrationChart = new Chart(ctx, {
          type: 'line',
          data: {
            labels: periodList,
            datasets: list.map((eq, i) => ({
              label: currentOffice === 'ALL' ? `${eq.name}(${eq.office.replace('사업소','').replace('사업단','')})` : eq.name,
              data: periodList.map(p => eq.history[p] !== undefined ? eq.history[p] : null),
              borderColor: palette[i % palette.length],
              backgroundColor: palette[i % palette.length],
              borderWidth: 2.5,
              pointRadius: 4.5,
              tension: 0.25,
              spanGaps: true
            }))
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#cbd5e1', font: { size: 11 } } } },
            scales: {
              x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } },
              y: { beginAtZero: true, suggestedMax: 9, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } }
            }
          }
        });
      } else {
        vibrationChart = new Chart(ctx, {
          type: 'bar',
          data: {
            labels: activeData.map(e => e.name),
            datasets: [{
              label: '최대 RMS (mm/s)',
              data: activeData.map(e => e.maxRms),
              backgroundColor: activeData.map(e => e.overallGrade === 'D' ? '#ef4444' : (e.overallGrade === 'C' ? '#f59e0b' : '#06b6d4')),
              borderRadius: 5
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
              x: { ticks: { color: '#94a3b8', font: { size: 10 } } },
              y: { beginAtZero: true, suggestedMax: 9, ticks: { color: '#94a3b8' } }
            }
          }
        });
      }
    }

    function toggleSort(col) {
      if (sortCol === col) sortAsc = !sortAsc;
      else { sortCol = col; sortAsc = true; }
      renderTable();
    }

    function renderTable() {
      const gFilter = document.getElementById('groupFilter').value;
      const sFilter = document.getElementById('statusFilter').value;
      const query = document.getElementById('searchInput').value.trim().toLowerCase();
      const tbody = document.getElementById('tableBody');

      let filtered = equipmentData.filter(item => {
        if (currentOffice !== 'ALL' && item.office !== currentOffice) return false;
        if (item.period && item.period !== currentPeriod) return false;
        if (gFilter !== 'ALL' && item.group !== gFilter) return false;
        if (sFilter === 'D' && item.overallGrade !== 'D') return false;
        if (sFilter === 'C' && item.overallGrade !== 'C') return false;
        if (sFilter === 'AB' && (item.overallGrade !== 'A' && item.overallGrade !== 'B')) return false;
        if (query && !(`${item.name} ${item.office} ${item.summary} ${item.action}`).toLowerCase().includes(query)) return false;
        return true;
      });

      filtered.sort((a, b) => {
        let vA = a[sortCol] || '', vB = b[sortCol] || '';
        if (sortCol === 'maxRms') { vA = parseFloat(a.maxRms) || 0; vB = parseFloat(b.maxRms) || 0; }
        else if (sortCol === 'grade') { const r = {A:1,B:2,C:3,D:4}; vA = r[a.overallGrade]||0; vB = r[b.overallGrade]||0; }
        else { vA = String(vA).toLowerCase(); vB = String(vB).toLowerCase(); }
        return vA < vB ? (sortAsc ? -1 : 1) : (vA > vB ? (sortAsc ? 1 : -1) : 0);
      });

      document.getElementById('filteredCount').innerText = `총 ${filtered.length}개 설비`;

      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="11" class="text-center py-10 text-slate-400">해당 사업소 및 측정시기에 등록된 데이터가 없습니다.<br><span class="text-cyan-400 font-bold">좌측 사이드바의 [서버 DB에 엑셀 직접 업로드]</span>를 이용하시면 모든 PC에서 즉시 공유됩니다!</td></tr>`;
        return;
      }

      tbody.innerHTML = filtered.map(item => {
        const p = item.points || { m_de:{h:0,v:0,a:0}, m_nde:{h:0,v:0,a:0}, p_de:{h:0,v:0,a:0}, p_nde:{h:0,v:0,a:0} };
        const mMax = Math.max(p.m_de.h, p.m_de.v, p.m_de.a, p.m_nde.h, p.m_nde.v, p.m_nde.a);
        const pMax = Math.max(p.p_de.h, p.p_de.v, p.p_de.a, p.p_nde.h, p.p_nde.v, p.p_nde.a);
        const ai = diagnoseByAI(item);
        const badgeClass = `grade-badge-${item.overallGrade}`;

        return `
          <tr class="hover:bg-slate-800/50 transition">
            <td class="py-3 px-3"><span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-800 text-cyan-300 border border-slate-700">${item.office}</span></td>
            <td class="py-3 px-3 font-mono text-[11px] text-slate-400">${item.period}</td>
            <td class="py-3 px-4 font-bold text-white">${item.name}</td>
            <td class="py-3 px-3 font-mono text-cyan-300">${item.power} kW</td>
            <td class="py-3 px-3 text-center font-mono text-amber-300">${item.maxRms === 0 ? '-' : mMax.toFixed(2)}</td>
            <td class="py-3 px-3 text-center font-mono text-cyan-300">${item.maxRms === 0 ? '-' : pMax.toFixed(2)}</td>
            <td class="py-3 px-3 text-center font-mono font-black ${item.maxRms === 0 ? 'text-slate-500' : (item.overallGrade === 'D' ? 'text-red-400' : (item.overallGrade === 'C' ? 'text-amber-400' : 'text-emerald-400'))}">
              ${item.maxRms === 0 ? '가동정지' : parseFloat(item.maxRms).toFixed(2)}
            </td>
            <td class="py-3 px-3 text-center">
              <span class="px-2.5 py-1 rounded text-[11px] font-bold ${item.maxRms === 0 ? 'bg-slate-800 text-slate-400' : badgeClass}">
                ${item.maxRms === 0 ? '미가동' : item.overallGrade + '등급'}
              </span>
            </td>
            <td class="py-3 px-4 max-w-xs">
              <div class="text-[11px] font-bold text-cyan-300 truncate" title="${ai.part} - ${ai.cause}"><i class="fa-solid fa-robot mr-1"></i>${ai.part}</div>
              <div class="text-[10px] text-slate-400 truncate" title="${ai.recommendation}">${ai.recommendation}</div>
            </td>
            <td class="py-3 px-4 max-w-xs truncate" title="${item.summary} / ${item.action}">
              <div>${item.summary}</div>
              <div class="text-[10px] text-slate-400">${item.action}</div>
            </td>
            <td class="py-3 px-3 text-center">
              <button onclick="viewEquipmentDetail('${item.id}')" class="px-2 py-1 rounded bg-slate-800 hover:bg-slate-700 text-cyan-400 text-[11px] border border-slate-700 mr-1">상세</button>
              <button onclick="deleteEquipment('${item.id}')" class="px-2 py-1 rounded bg-slate-800 hover:bg-rose-950 text-slate-400 hover:text-rose-400 text-[11px] border border-slate-700"><i class="fa-solid fa-trash-can"></i></button>
            </td>
          </tr>
        `;
      }).join('');
    }

    function openOfficeOverviewModal() {
      const grid = document.getElementById('officeOverviewGrid');
      grid.innerHTML = officeList.map(off => {
        const list = equipmentData.filter(e => e.office === off.name && e.period === currentPeriod);
        const dCnt = list.filter(e => e.overallGrade === 'D' && e.maxRms > 0).length;
        const cCnt = list.filter(e => e.overallGrade === 'C' && e.maxRms > 0).length;
        const abCnt = list.filter(e => (e.overallGrade === 'A' || e.overallGrade === 'B') && e.maxRms > 0).length;
        return `
          <div onclick="switchOffice('${off.name}'); closeOfficeOverviewModal();" class="p-3.5 rounded-xl bg-slate-800/80 border ${list.length > 0 ? 'border-cyan-500/40' : 'border-slate-700'} hover:border-cyan-400 cursor-pointer transition">
            <div class="flex items-center justify-between mb-2">
              <div class="flex items-center gap-2 font-bold text-white">
                <i class="fa-solid ${off.icon} ${off.color}"></i>
                <span>${off.name}</span>
              </div>
              <span class="px-2 py-0.5 rounded-full text-[11px] font-bold ${list.length > 0 ? 'bg-cyan-950 text-cyan-300 border border-cyan-700' : 'bg-slate-900 text-slate-500'}">총 ${list.length}대</span>
            </div>
            <div class="grid grid-cols-3 gap-1 text-center text-[11px] mt-2">
              <div class="p-1 rounded bg-emerald-950/30 text-emerald-400">양호: ${abCnt}</div>
              <div class="p-1 rounded bg-amber-950/30 text-amber-400">C주의: ${cCnt}</div>
              <div class="p-1 rounded bg-rose-950/30 text-rose-400 font-bold">D위험: ${dCnt}</div>
            </div>
          </div>
        `;
      }).join('');
      document.getElementById('officeOverviewModal').classList.remove('hidden');
    }
    function closeOfficeOverviewModal() { document.getElementById('officeOverviewModal').classList.add('hidden'); }

    function openUploadModal() {
      renderOfficeTabs();
      document.getElementById('uploadModal').classList.remove('hidden');
      document.getElementById('uploadStatusBox').classList.add('hidden');
    }
    function closeUploadModal() { document.getElementById('uploadModal').classList.add('hidden'); }

    function handleFileSelect(e) {
      const f = e.target.files[0];
      if (f) processDataFile(f);
    }

    function processDataFile(file) {
      const statusBox = document.getElementById('uploadStatusBox');
      const targetPeriod = document.getElementById('uploadPeriodSelect').value;
      const uploadOffice = document.getElementById('uploadOfficeSelect').value;
      statusBox.classList.remove('hidden');
      statusBox.innerHTML = `<span class="text-cyan-400 font-semibold"><i class="fa-solid fa-spinner fa-spin mr-1"></i>[${uploadOffice}] 엑셀 정밀 파싱 중...</span>`;

      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          let parsedBySheet = {};

          for (let sheetName of workbook.SheetNames) {
            if (sheetName.includes('가이드') || sheetName.includes('양식')) continue;
            let sheetPeriod = targetPeriod;
            const ym = sheetName.match(/(\d{2,4})\s*년?/);
            let yStr = targetPeriod.split('년')[0] + '년';
            if (ym) yStr = (ym[1].length === 2 ? '20' + ym[1] : ym[1]) + '년';
            if (sheetName.includes('상반기')) sheetPeriod = `${yStr} 상반기`;
            else if (sheetName.includes('하반기')) sheetPeriod = `${yStr} 하반기`;

            if (!periodList.includes(sheetPeriod)) periodList.push(sheetPeriod);
            if (!parsedBySheet[sheetPeriod]) parsedBySheet[sheetPeriod] = {};

            const ws = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            let currentEquip = null;
            let lastPower = 55;
            let isHeaderFound = false;

            for (let i = 0; i < jsonData.length; i++) {
              const row = jsonData[i];
              if (!row || row.length < 4) continue;
              const rowStr = row.join('').replace(/\s/g, '');
              if (rowStr.includes('설비명') && (rowStr.includes('측정위치') || rowStr.includes('전동기'))) {
                isHeaderFound = true;
                continue;
              }
              if (!isHeaderFound) continue;

              let equipName = String(row[1] || '').trim().replace(/\s+/g, ' ');
              let pVal = parseFloat(row[2]);
              if (!isNaN(pVal) && pVal > 0) lastPower = pVal;

              let pointStr = String(row[3] || '').trim().replace(/\s/g, '');
              let axisStr = String(row[4] || '').toLowerCase().trim();
              let causeStr = String(row[8] || '').trim();
              let actionStr = String(row[9] || '').trim();

              if (equipName.length > 28 || equipName.includes('점검실시') || equipName.includes('업데이트') || equipName === '반기' || equipName.includes('페이지')) {
                currentEquip = null;
                continue;
              }

              const hasPointFormat = pointStr.includes('전동기') || pointStr.includes('펌프') || pointStr.includes('송풍기') || pointStr.includes('부하') || /^[1-4]/.test(pointStr) || pointStr.includes('가동정지') || pointStr.includes('미가동') || pointStr.includes('정비');

              if (equipName !== '' && equipName !== '설비명') {
                if (!parsedBySheet[sheetPeriod][equipName]) {
                  parsedBySheet[sheetPeriod][equipName] = {
                    name: equipName,
                    period: sheetPeriod,
                    power: lastPower,
                    points: { m_de:{h:0,v:0,a:0}, m_nde:{h:0,v:0,a:0}, p_de:{h:0,v:0,a:0}, p_nde:{h:0,v:0,a:0} },
                    summary: '',
                    action: ''
                  };
                }
                currentEquip = parsedBySheet[sheetPeriod][equipName];
                if (!isNaN(pVal) && pVal > 0) currentEquip.power = pVal;
              }
              if (!currentEquip) continue;

              if (pointStr.includes('가동정지') || pointStr.includes('미가동')) {
                if (!currentEquip.summary.includes('가동정지')) currentEquip.summary = '가동정지' + (causeStr ? ` (${causeStr})` : '');
              }

              let rmsVal = parseFloat(row[5]);
              if (hasPointFormat && !isNaN(rmsVal) && rmsVal >= 0 && rmsVal < 50) {
                let tp = null;
                if (pointStr.includes('전동기') && pointStr.includes('반부하')) tp = currentEquip.points.m_nde;
                else if (pointStr.includes('전동기') || pointStr.startsWith('1') || pointStr.startsWith('2')) tp = pointStr.includes('반부하') ? currentEquip.points.m_nde : currentEquip.points.m_de;
                else if ((pointStr.includes('펌프') || pointStr.includes('송풍기') || pointStr.startsWith('3') || pointStr.startsWith('4')) && pointStr.includes('반부하')) tp = currentEquip.points.p_nde;
                else if (pointStr.includes('펌프') || pointStr.includes('송풍기') || pointStr.startsWith('3') || pointStr.startsWith('4')) tp = currentEquip.points.p_de;

                if (tp) {
                  if (axisStr === 'x' || axisStr === 'h') tp.h = Math.max(tp.h, rmsVal);
                  if (axisStr === 'y' || axisStr === 'v') tp.v = Math.max(tp.v, rmsVal);
                  if (axisStr === 'z' || axisStr === 'a') tp.a = Math.max(tp.a, rmsVal);
                }
              }
              if (causeStr && causeStr !== '-' && !currentEquip.summary.includes(causeStr)) currentEquip.summary += (currentEquip.summary ? ', ' : '') + causeStr;
              if (actionStr && actionStr !== '양호' && actionStr !== '-' && !currentEquip.action.includes(actionStr)) currentEquip.action += (currentEquip.action ? ', ' : '') + actionStr;
            }
          }

          const foundPeriods = Object.keys(parsedBySheet).filter(k => Object.keys(parsedBySheet[k]).length > 0);
          let summaryMsg = [];

          foundPeriods.forEach(pKey => {
            const list = Object.values(parsedBySheet[pKey]);
            equipmentData = equipmentData.filter(e => !(e.office === uploadOffice && e.period === pKey));
            list.forEach((eq, idx) => {
              const p = eq.points;
              const maxRms = Math.max(p.m_de.h, p.m_de.v, p.m_de.a, p.m_nde.h, p.m_nde.v, p.m_nde.a, p.p_de.h, p.p_de.v, p.p_de.a, p.p_nde.h, p.p_nde.v, p.p_nde.a);
              const evalRes = calculateIsoGrade(eq.power, maxRms, 'rigid');
              equipmentData.push({
                id: `${uploadOffice}_${pKey}_${eq.name}`,
                office: uploadOffice,
                period: eq.period,
                name: eq.name,
                power: eq.power,
                group: evalRes.group,
                foundation: 'rigid',
                overallGrade: maxRms === 0 ? 'A' : evalRes.grade,
                maxRms: maxRms,
                points: eq.points,
                summary: maxRms === 0 ? (eq.summary || '가동정지') : (eq.summary || '측정 완료'),
                action: maxRms === 0 ? (eq.action || '정비 대기') : (eq.action || '추세 관리')
              });
            });
            summaryMsg.push(`${pKey}: ${list.length}대`);
          });

          currentOffice = uploadOffice;
          if (foundPeriods.length > 0) currentPeriod = foundPeriods[0];
          saveToLocalStorage();
          copyServerSyncCode();
          renderPeriodSelectors();
          renderOfficeTabs();
          updateDashboard();

          statusBox.className = 'p-3 rounded-lg border text-xs bg-emerald-950/30 border-emerald-500/40 text-emerald-300 font-bold';
          statusBox.innerHTML = `<i class="fa-solid fa-circle-check mr-1"></i>[${uploadOffice}] 파싱 완료 (${summaryMsg.join(' / ')})<br><span class="text-amber-300">※ 서버반영 코드가 자동 복사되었습니다! 좌측 사이드바에서 바로 엑셀을 올리시면 이 과정 없이 즉시 공유됩니다.</span>`;
          setTimeout(() => closeUploadModal(), 2500);
        } catch (err) {
          statusBox.className = 'p-3 rounded-lg border text-xs bg-rose-950/30 border-rose-500/40 text-rose-300';
          statusBox.innerHTML = '엑셀 파싱 실패: 파일 형식을 확인해 주세요.';
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function openRepairSummaryModal() {
      const bSel = document.getElementById('sumBasePeriod');
      const tSel = document.getElementById('sumTargetPeriod');
      bSel.innerHTML = periodList.map((p, i) => `<option value="${p}" ${i === 0 ? 'selected' : ''}>${p}</option>`).join('');
      tSel.innerHTML = periodList.map((p, i) => `<option value="${p}" ${i === periodList.length - 1 ? 'selected' : ''}>${p}</option>`).join('');
      document.getElementById('repairSummaryModal').classList.remove('hidden');
      renderRepairSummarySheet();
    }
    function closeRepairSummaryModal() { document.getElementById('repairSummaryModal').classList.add('hidden'); }

    function setRepairFilter(m) {
      currentRepairFilter = m;
      renderRepairSummarySheet();
    }

    function buildRepairComparisonRows() {
      const baseP = document.getElementById('sumBasePeriod')?.value || periodList[0];
      const targetP = document.getElementById('sumTargetPeriod')?.value || periodList[periodList.length - 1];
      const map = {};
      equipmentData.forEach(item => {
        if (currentOffice !== 'ALL' && item.office !== currentOffice) return;
        const k = `${item.office}||${item.name}`;
        if (!map[k]) map[k] = { office: item.office, name: item.name, power: item.power, base: null, target: null };
        if (item.period === baseP) map[k].base = item;
        if (item.period === targetP) map[k].target = item;
      });

      const rank = { A: 1, B: 2, C: 3, D: 4 };
      return Object.values(map).sort((a, b) => a.name.localeCompare(b.name, 'ko')).map(entry => {
        const b = entry.base, t = entry.target;
        const bRms = b ? (parseFloat(b.maxRms) || 0) : null;
        const tRms = t ? (parseFloat(t.maxRms) || 0) : null;
        const bGrade = b ? (bRms === 0 ? 'STOP' : b.overallGrade) : '-';
        const tGrade = t ? (tRms === 0 ? 'STOP' : t.overallGrade) : '-';
        let changeType = 'SAME', diffPct = 0;

        if (bRms > 0 && tRms > 0) {
          diffPct = Math.round(((bRms - tRms) / bRms) * 100);
          if ((rank[bGrade] >= 3 && rank[tGrade] <= 2) || (bRms - tRms) >= 0.5) changeType = 'IMPROVED';
          else if (rank[tGrade] >= 3) changeType = 'NEED_REPAIR';
        } else if (t && (t.overallGrade === 'C' || t.overallGrade === 'D') && tRms > 0) {
          changeType = 'NEED_REPAIR';
        }
        const aiDiag = t ? diagnoseByAI(t) : (b ? diagnoseByAI(b) : { part: '-', cause: '-', recommendation: '-' });
        return { ...entry, bRms, tRms, bGrade, tGrade, changeType, diffPct, aiDiag };
      });
    }

    function renderRepairSummarySheet() {
      const rows = buildRepairComparisonRows();
      const imp = rows.filter(r => r.changeType === 'IMPROVED');
      const need = rows.filter(r => r.tGrade === 'C' || r.tGrade === 'D');
      const stop = rows.filter(r => r.tGrade === 'STOP');
      const avgRed = imp.filter(r => r.diffPct > 0).length > 0
        ? Math.round(imp.filter(r => r.diffPct > 0).reduce((a, c) => a + c.diffPct, 0) / imp.filter(r => r.diffPct > 0).length) : 0;

      document.getElementById('sumKpiImproved').innerText = `${imp.length} 대`;
      document.getElementById('sumKpiReduction').innerText = `${avgRed} % 감소`;
      document.getElementById('sumKpiNeed').innerText = `${need.length} 대`;
      document.getElementById('sumKpiStopped').innerText = `${stop.length} 대`;

      let list = rows;
      if (currentRepairFilter === 'IMPROVED') list = imp;
      else if (currentRepairFilter === 'NEED_REPAIR') list = need;

      const tbody = document.getElementById('repairSummaryTableBody');
      tbody.innerHTML = list.map(r => `
        <tr class="hover:bg-slate-800/60">
          <td class="py-2.5 px-3">${r.office}</td>
          <td class="py-2.5 px-3 font-bold text-white">${r.name} <span class="text-[10px] text-cyan-400">${r.power}kW</span></td>
          <td class="py-2.5 px-3 text-center font-mono">${r.bRms === null ? '-' : (r.bRms === 0 ? '가동정지' : r.bRms.toFixed(2) + ' (' + r.bGrade + ')')}</td>
          <td class="py-2.5 px-3 text-center font-mono">${r.tRms === null ? '-' : (r.tRms === 0 ? '가동정지' : r.tRms.toFixed(2) + ' (' + r.tGrade + ')')}</td>
          <td class="py-2.5 px-3 text-center">
            ${r.changeType === 'IMPROVED' ? `<span class="px-2 py-0.5 rounded bg-emerald-950 text-emerald-300 font-bold">🛠️ 수선완료 (-${r.diffPct}%)</span>` : (r.tGrade === 'C' || r.tGrade === 'D' ? `<span class="px-2 py-0.5 rounded bg-rose-950 text-rose-300 font-bold">🚨 정비요망</span>` : '<span class="text-slate-400">양호유지</span>')}
          </td>
          <td class="py-2.5 px-3">
            <div class="text-[11px] font-bold text-cyan-300">${r.aiDiag.part} : ${r.aiDiag.cause}</div>
            <div class="text-[10px] text-amber-300">${r.aiDiag.recommendation}</div>
          </td>
        </tr>
      `).join('');
    }

    function exportRepairSummaryExcel() {
      const rows = buildRepairComparisonRows();
      const aoa = [["사업소", "설비명", "용량(kW)", "이전반기RMS", "이전등급", "현재반기RMS", "현재등급", "수선결과", "저감률(%)", "AI추정원인", "AI개선권고"]];
      rows.forEach(r => {
        aoa.push([r.office, r.name, r.power, r.bRms ?? '-', r.bGrade, r.tRms ?? '-', r.tGrade, r.changeType, r.diffPct + '%', r.aiDiag.cause, r.aiDiag.recommendation]);
      });
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "수선및AI요약");
      XLSX.writeFile(wb, `BECO_수선및AI요약시트_${currentOffice}.xlsx`);
    }

    function openAddModal() { renderOfficeTabs(); document.getElementById('addModal').classList.remove('hidden'); }
    function closeAddModal() { document.getElementById('addModal').classList.add('hidden'); }

    function handleEquipmentSubmit(e) {
      e.preventDefault();
      const office = document.getElementById('inpOffice').value;
      const period = document.getElementById('inpPeriod').value;
      const name = document.getElementById('inpName').value.trim();
      const power = parseFloat(document.getElementById('inpPower').value) || 55;
      const pts = {
        m_de: { h: parseFloat(document.getElementById('inp_m_de_h').value)||0, v: parseFloat(document.getElementById('inp_m_de_v').value)||0, a: parseFloat(document.getElementById('inp_m_de_a').value)||0 },
        m_nde: { h: parseFloat(document.getElementById('inp_m_nde_h').value)||0, v: parseFloat(document.getElementById('inp_m_nde_v').value)||0, a: parseFloat(document.getElementById('inp_m_nde_a').value)||0 },
        p_de: { h: parseFloat(document.getElementById('inp_p_de_h').value)||0, v: parseFloat(document.getElementById('inp_p_de_v').value)||0, a: parseFloat(document.getElementById('inp_p_de_a').value)||0 },
        p_nde: { h: parseFloat(document.getElementById('inp_p_nde_h').value)||0, v: parseFloat(document.getElementById('inp_p_nde_v').value)||0, a: parseFloat(document.getElementById('inp_p_nde_a').value)||0 }
      };
      const maxRms = Math.max(pts.m_de.h, pts.m_de.v, pts.m_de.a, pts.m_nde.h, pts.m_nde.v, pts.m_nde.a, pts.p_de.h, pts.p_de.v, pts.p_de.a, pts.p_nde.h, pts.p_nde.v, pts.p_nde.a);
      const res = calculateIsoGrade(power, maxRms, 'rigid');
      equipmentData.unshift({
        id: `${office}_${period}_${name}`,
        office, period, name, power, group: res.group, overallGrade: res.grade, maxRms, points: pts,
        summary: document.getElementById('inpSummary').value || '직접 등록',
        action: document.getElementById('inpAction').value || '추세 관리'
      });
      saveToLocalStorage();
      closeAddModal();
      renderOfficeTabs();
      updateDashboard();
    }

    function deleteEquipment(id) {
      equipmentData = equipmentData.filter(e => e.id !== id);
      saveToLocalStorage();
      renderOfficeTabs();
      updateDashboard();
    }

    function openManagePeriodModal() {
      const listEl = document.getElementById('periodManageList');
      listEl.innerHTML = periodList.map(p => `
        <div class="flex justify-between items-center p-2 bg-slate-800 rounded border border-slate-700">
          <span class="font-bold">${p}</span>
          <button onclick="deletePeriod('${p}')" class="px-2 py-0.5 bg-rose-950 text-rose-300 rounded text-[11px]">삭제</button>
        </div>
      `).join('');
      document.getElementById('managePeriodModal').classList.remove('hidden');
    }
    function closeManagePeriodModal() { document.getElementById('managePeriodModal').classList.add('hidden'); }

    function handlePeriodSubmit(e) {
      e.preventDefault();
      const val = document.getElementById('inpNewPeriodName').value.trim();
      if (val && !periodList.includes(val)) periodList.push(val);
      currentPeriod = val;
      saveToLocalStorage();
      renderPeriodSelectors();
      openManagePeriodModal();
      updateDashboard();
    }

    function deletePeriod(p) {
      periodList = periodList.filter(x => x !== p);
      equipmentData = equipmentData.filter(e => e.period !== p);
      if (currentPeriod === p) currentPeriod = periodList[0] || '2026년 상반기';
      saveToLocalStorage();
      renderPeriodSelectors();
      openManagePeriodModal();
      updateDashboard();
    }

    function viewEquipmentDetail(id) {
      const item = equipmentData.find(e => e.id === id);
      if (!item) return;
      const ai = diagnoseByAI(item);
      const pts = item.points || { m_de:{h:0,v:0,a:0}, m_nde:{h:0,v:0,a:0}, p_de:{h:0,v:0,a:0}, p_nde:{h:0,v:0,a:0} };
      document.getElementById('detailModalContent').innerHTML = `
        <h3 class="text-lg font-bold text-white mb-2">[${item.office}] ${item.name} (${item.period})</h3>
        <div class="p-3.5 rounded-xl bg-cyan-950/40 border border-cyan-500/40 text-xs space-y-1 mb-4">
          <div class="font-bold text-cyan-300"><i class="fa-solid fa-robot mr-1"></i>AI 12축 정밀 진단 결과</div>
          <div>• 집중 부위 및 추정 원인: <span class="text-amber-300 font-semibold">${ai.part} — ${ai.cause}</span></div>
          <div>• AI 개선 권고: <span class="text-emerald-300 font-semibold">${ai.recommendation}</span></div>
        </div>
        <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 text-center text-xs">
          <div class="p-2 bg-slate-800 rounded">M-DE: H:${pts.m_de.h} V:${pts.m_de.v} A:${pts.m_de.a}</div>
          <div class="p-2 bg-slate-800 rounded">M-NDE: H:${pts.m_nde.h} V:${pts.m_nde.v} A:${pts.m_nde.a}</div>
          <div class="p-2 bg-slate-800 rounded">P-DE: H:${pts.p_de.h} V:${pts.p_de.v} A:${pts.p_de.a}</div>
          <div class="p-2 bg-slate-800 rounded">P-NDE: H:${pts.p_nde.h} V:${pts.p_nde.v} A:${pts.p_nde.a}</div>
        </div>
      `;
      document.getElementById('detailModal').classList.remove('hidden');
    }
    function closeDetailModal() { document.getElementById('detailModal').classList.add('hidden'); }

    function exportToCSV() {
      const active = equipmentData.filter(e => (currentOffice === 'ALL' || e.office === currentOffice) && (!e.period || e.period === currentPeriod));
      const rows = [["사업소", "측정시기", "설비명", "용량(kW)", "최대RMS", "등급", "AI진단부위", "AI개선권고", "현장요약", "정비대책"]];
      active.forEach(e => {
        const ai = diagnoseByAI(e);
        rows.push([e.office, e.period, e.name, e.power, e.maxRms, e.overallGrade, ai.part, ai.recommendation, e.summary, e.action]);
      });
      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "진동진단대장");
      XLSX.writeFile(wb, `BECO_진동진단_${currentOffice}_${currentPeriod}.xlsx`);
    }

    window.onload = function() {
      loadDataWithServerMerge();
      renderPeriodSelectors();
      renderOfficeTabs();
      updateDashboard();
    };
  </script>
</body>
</html>
"""

# 서버 DB 데이터를 HTML 내부에 안전하게 주입
dashboard_html = (
    dashboard_html
    .replace("__SERVER_EQUIP_JSON__", json.dumps(server_equip_data, ensure_ascii=False))
    .replace("__SERVER_OFFICES_JSON__", json.dumps(server_offices, ensure_ascii=False))
    .replace("__SERVER_PERIODS_JSON__", json.dumps(server_periods, ensure_ascii=False))
    .replace("__INITIAL_OFFICE__", initial_office_val)
)

components.html(dashboard_html, height=1080, scrolling=True)