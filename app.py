import streamlit as st
import streamlit.components.v1 as components

# Streamlit 페이지 기본 환경 설정 (브라우저 풀 와이드 모드)
st.set_page_config(
    page_title="BECO 회전설비 진동진단 통합관리",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 전체 대시보드 애플리케이션 (HTML, Tailwind CSS, Chart.js, SheetJS, LocalStorage 영구저장 연동)
dashboard_html = r"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>부산환경공단 사업소별 회전설비 진동진단 통합관리 시스템 (BECO-VIB)</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Chart.js CDN -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <!-- SheetJS (xlsx) CDN -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- Font Awesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Noto+Sans+KR:wght@300;400;500;700;800&display=swap" rel="stylesheet">

  <script>
    tailwind.config = {
      darkMode: 'class',
      theme: {
        extend: {
          fontFamily: {
            sans: ['"Noto Sans KR"', 'Inter', 'sans-serif'],
          },
          colors: {
            beco: {
              navy: '#0f172a',
              card: '#1e293b',
              accent: '#0284c7',
              cyan: '#06b6d4',
              green: '#10b981',
              yellow: '#f59e0b',
              red: '#ef4444',
            }
          }
        }
      }
    }
  </script>

  <style>
    body {
      background-color: #0b1120;
      color: #f1f5f9;
      font-family: 'Noto Sans KR', 'Inter', sans-serif;
    }
    .custom-scrollbar::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .custom-scrollbar::-webkit-scrollbar-track {
      background: #1e293b;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
      background: #475569;
      border-radius: 4px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {
      background: #64748b;
    }
    .glass-card {
      background: rgba(30, 41, 59, 0.7);
      backdrop-filter: blur(12px);
      border: 1px solid rgba(255, 255, 255, 0.07);
    }
    .glow-cyan {
      box-shadow: 0 0 15px -3px rgba(6, 182, 212, 0.25);
    }
    .grade-badge-A { background: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.4); }
    .grade-badge-B { background: rgba(14, 165, 233, 0.15); color: #38bdf8; border: 1px solid rgba(14, 165, 233, 0.4); }
    .grade-badge-C { background: rgba(245, 158, 11, 0.15); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.4); }
    .grade-badge-D { background: rgba(239, 68, 68, 0.18); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.45); }
  </style>
</head>
<body class="min-h-screen custom-scrollbar text-slate-200">

  <!-- Header -->
  <header class="glass-card sticky top-0 z-40 border-b border-slate-700/60 bg-slate-900/90 px-4 py-3 sm:px-6">
    <div class="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-3">
      <div class="flex items-center gap-3">
        <div class="h-10 w-10 rounded-xl bg-gradient-to-br from-cyan-500 to-blue-600 flex items-center justify-center text-white shadow-lg glow-cyan">
          <i class="fa-solid fa-wave-square text-lg"></i>
        </div>
        <div>
          <div class="flex items-center gap-2">
            <h1 class="text-lg sm:text-xl font-bold tracking-tight text-white flex items-center gap-2">
              BECO-VIB Master
              <span class="text-xs font-semibold px-2 py-0.5 rounded-full bg-cyan-950 text-cyan-400 border border-cyan-800">ISO 20816-3</span>
              <span id="syncStatusBadge" class="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-emerald-950 text-emerald-400 border border-emerald-800"><i class="fa-solid fa-hard-drive mr-1"></i>로컬 저장</span>
            </h1>
          </div>
          <p class="text-xs text-slate-400">부산환경공단 회전체 정밀진단 (사업소 추가/삭제 · 12축 계측 및 추이분석)</p>
        </div>
      </div>

      <!-- Right Header Actions -->
      <div class="flex items-center gap-2 sm:gap-3 w-full sm:w-auto justify-end flex-wrap">
        <!-- Google Sheets Sync Button -->
        <button onclick="openSheetModal()" class="flex items-center gap-1.5 bg-emerald-700 hover:bg-emerald-600 text-white text-xs sm:text-sm font-medium py-2 px-3 rounded-lg transition duration-200 shadow">
          <i class="fa-solid fa-table-cells text-emerald-200"></i>
          <span>구글시트 연동</span>
        </button>
        <!-- Add Period Button -->
        <button onclick="openAddPeriodModal()" class="flex items-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs sm:text-sm font-medium py-2 px-3 rounded-lg border border-slate-700 transition duration-200 shadow">
          <i class="fa-regular fa-calendar-plus text-cyan-400"></i>
          <span>측정주기 추가</span>
        </button>
        <!-- Manage Office Button -->
        <button onclick="openAddOfficeModal()" class="flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-500 text-white text-xs sm:text-sm font-medium py-2 px-3 rounded-lg transition duration-200 shadow">
          <i class="fa-solid fa-building-circle-gear"></i>
          <span>사업소 관리</span>
        </button>
        <!-- Add Equipment Button -->
        <button onclick="openAddModal()" class="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-medium py-2 px-3 rounded-lg transition duration-200 shadow">
          <i class="fa-solid fa-plus"></i>
          <span>설비 직접등록</span>
        </button>
        <!-- Excel/CSV Upload Button -->
        <button onclick="openUploadModal()" class="flex items-center gap-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs sm:text-sm font-medium py-2 px-3 rounded-lg transition duration-200 shadow">
          <i class="fa-solid fa-file-excel"></i>
          <span>데이터 업로드</span>
        </button>
        <!-- Calculator Button -->
        <button onclick="openCalcModal()" class="flex items-center gap-1.5 bg-gradient-to-r from-cyan-600 to-blue-600 hover:from-cyan-500 hover:to-blue-500 text-white text-xs sm:text-sm font-medium py-2 px-3 rounded-lg transition duration-200 shadow">
          <i class="fa-solid fa-calculator"></i>
          <span>기준 판정기</span>
        </button>
      </div>
    </div>

    <!-- Office & Period Switcher Bar -->
    <div class="max-w-7xl mx-auto mt-3 pt-3 border-t border-slate-700/60 flex items-center justify-between flex-wrap gap-2">
      <!-- Office Tabs -->
      <div class="flex items-center gap-2 overflow-x-auto pb-1 sm:pb-0" id="officeTabs">
        <!-- Rendered dynamically -->
      </div>

      <!-- Period Selection & Export Controls -->
      <div class="flex items-center gap-3 text-xs">
        <div class="flex items-center gap-1.5 bg-slate-800/90 px-2.5 py-1 rounded-lg border border-slate-700">
          <span class="text-slate-400"><i class="fa-solid fa-clock-rotate-left text-cyan-400 mr-1"></i>측정 시기:</span>
          <select id="periodFilterSelect" onchange="switchPeriod(this.value)" class="bg-transparent text-cyan-300 font-bold focus:outline-none cursor-pointer">
            <!-- Rendered dynamically -->
          </select>
        </div>

        <button onclick="exportToCSV()" class="bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white px-2.5 py-1 rounded border border-slate-700 transition flex items-center gap-1" title="현재 조회 데이터 엑셀 내보내기">
          <i class="fa-solid fa-file-export text-emerald-400"></i>
          <span>내보내기</span>
        </button>

        <button onclick="resetToInitialData()" class="bg-slate-800 hover:bg-rose-950 text-slate-400 hover:text-rose-400 px-2 py-1 rounded border border-slate-700 transition text-[11px]" title="저장된 데이터 초기화">
          <i class="fa-solid fa-arrow-rotate-left mr-1"></i>초기화
        </button>
      </div>
    </div>
  </header>

  <main class="max-w-7xl mx-auto px-4 sm:px-6 py-6 space-y-6">

    <!-- KPI Metric Cards Grid -->
    <section class="grid grid-cols-2 lg:grid-cols-5 gap-3.5 sm:gap-4">
      <div class="glass-card rounded-xl p-4 border border-slate-700/50 hover:border-slate-600 transition">
        <div class="flex items-center justify-between text-slate-400 text-xs mb-1">
          <span>진단 대상 설비</span>
          <i class="fa-solid fa-cubes text-slate-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-total" class="text-2xl sm:text-3xl font-extrabold text-white">0</span>
          <span class="text-xs text-slate-400">대</span>
        </div>
        <p class="text-[11px] text-slate-400 mt-2 truncate">동력 7.5kW ~ 350kW (12축)</p>
      </div>

      <div class="glass-card rounded-xl p-4 border border-emerald-500/20 bg-emerald-950/10 hover:border-emerald-500/40 transition">
        <div class="flex items-center justify-between text-emerald-400 text-xs mb-1">
          <span>양호 상태 (A·B)</span>
          <i class="fa-solid fa-circle-check text-emerald-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-ab" class="text-2xl sm:text-3xl font-extrabold text-emerald-400">0</span>
          <span class="text-xs text-emerald-300/70">대</span>
        </div>
        <p class="text-[11px] text-emerald-400/80 mt-2 flex items-center gap-1 font-medium">
          <span id="metric-ab-pct">0%</span> 정상 가동 유지
        </p>
      </div>

      <div class="glass-card rounded-xl p-4 border border-amber-500/20 bg-amber-950/10 hover:border-amber-500/40 transition">
        <div class="flex items-center justify-between text-amber-400 text-xs mb-1">
          <span>보수필요 (C등급)</span>
          <i class="fa-solid fa-triangle-exclamation text-amber-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-c" class="text-2xl sm:text-3xl font-extrabold text-amber-400">0</span>
          <span class="text-xs text-amber-300/70">대</span>
        </div>
        <p class="text-[11px] text-amber-400/80 mt-2 flex items-center gap-1">
          <i class="fa-solid fa-arrow-trend-up text-[10px]"></i> 추세관리 및 부분보수
        </p>
      </div>

      <div class="glass-card rounded-xl p-4 border border-red-500/30 bg-red-950/20 hover:border-red-500/50 transition">
        <div class="flex items-center justify-between text-red-400 text-xs mb-1">
          <span>즉시점검 (D등급)</span>
          <i class="fa-solid fa-circle-exclamation text-red-400 animate-pulse"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-d" class="text-2xl sm:text-3xl font-extrabold text-red-400">0</span>
          <span class="text-xs text-red-300/70">대</span>
        </div>
        <p class="text-[11px] text-red-300 mt-2 font-medium">긴급 점검 및 분해정비</p>
      </div>

      <div class="glass-card rounded-xl p-4 border border-slate-700/50 col-span-2 lg:col-span-1">
        <div class="flex items-center justify-between text-slate-400 text-xs mb-1">
          <span>정비/미가동 설비</span>
          <i class="fa-solid fa-wrench text-slate-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-stop" class="text-2xl sm:text-3xl font-extrabold text-slate-300">5</span>
          <span class="text-xs text-slate-400">대</span>
        </div>
        <p class="text-[11px] text-slate-400 mt-2">급수 #5·#6, 순환 #1·#2·#3</p>
      </div>
    </section>

    <!-- Visual Chart & Trend Analysis Panel -->
    <section class="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div class="lg:col-span-2 glass-card rounded-xl p-5 border border-slate-700/60">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-4 border-b border-slate-700/60 gap-3">
          <div>
            <h2 class="text-base font-semibold text-white flex items-center gap-2">
              <i class="fa-solid fa-chart-line text-cyan-400"></i>
              <span id="chartMainTitle">주요 설비 반기별 진동 추세 (Trend)</span>
            </h2>
            <p class="text-xs text-slate-400" id="chartSubTitle">과거 반기 측정치 누적 변화와 ISO 20816-3 경보 기준치 비교</p>
          </div>
          
          <div class="flex items-center gap-1.5 bg-slate-800 p-1 rounded-lg border border-slate-700 text-xs">
            <button id="btnModeTrend" onclick="setChartMode('trend')" class="px-3 py-1 rounded font-bold bg-cyan-600 text-white transition">
              <i class="fa-solid fa-chart-line mr-1"></i>반기별 추이선
            </button>
            <button id="btnModeBar" onclick="setChartMode('bar')" class="px-3 py-1 rounded font-medium text-slate-300 hover:text-white transition">
              <i class="fa-solid fa-chart-simple mr-1"></i>현재 상태 막대
            </button>
          </div>
        </div>

        <div class="pt-4 h-[300px] w-full">
          <canvas id="vibrationChart"></canvas>
        </div>
      </div>

      <!-- ISO Standard Quick Guide -->
      <div class="glass-card rounded-xl p-5 border border-slate-700/60 flex flex-col justify-between">
        <div>
          <div class="flex items-center justify-between pb-3 border-b border-slate-700/60">
            <h3 class="text-sm font-semibold text-white flex items-center gap-2">
              <i class="fa-solid fa-book-bookmark text-amber-400"></i>
              ISO 20816-3 진동 판정 기준표
            </h3>
            <span class="text-[10px] text-slate-400">속도 RMS(mm/s)</span>
          </div>
          
          <div class="mt-4 space-y-3 text-xs">
            <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
              <div class="flex justify-between items-center font-bold text-slate-200">
                <span class="text-cyan-300">Group 1 (&gt; 300 kW)</span>
                <span class="text-[11px] font-normal text-slate-400">중온수펌프, 강제송풍기</span>
              </div>
              <div class="grid grid-cols-3 gap-1 mt-2 text-center text-[11px]">
                <div class="p-1 rounded grade-badge-A font-medium">A/B &lt; 4.5</div>
                <div class="p-1 rounded grade-badge-C font-medium">C &gt; 4.5 (주의)</div>
                <div class="p-1 rounded grade-badge-D font-medium">D &gt; 7.1 (중단)</div>
              </div>
            </div>

            <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
              <div class="flex justify-between items-center font-bold text-slate-200">
                <span class="text-emerald-300">Group 2 (15 ~ 300 kW)</span>
                <span class="text-[11px] font-normal text-slate-400">보일러급수, 소각순환, 히트펌프</span>
              </div>
              <div class="grid grid-cols-3 gap-1 mt-2 text-center text-[11px]">
                <div class="p-1 rounded grade-badge-A font-medium">A/B &lt; 2.8</div>
                <div class="p-1 rounded grade-badge-C font-medium">C &gt; 2.8 (주의)</div>
                <div class="p-1 rounded grade-badge-D font-medium">D &gt; 4.5 (중단)</div>
              </div>
            </div>

            <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
              <div class="flex justify-between items-center font-bold text-slate-200">
                <span class="text-purple-300">Group 4 (&lt; 15 kW)</span>
                <span class="text-[11px] font-normal text-slate-400">1·2차 냉각수, 유체냉각수</span>
              </div>
              <div class="grid grid-cols-3 gap-1 mt-2 text-center text-[11px]">
                <div class="p-1 rounded grade-badge-A font-medium">A/B &lt; 2.8</div>
                <div class="p-1 rounded grade-badge-C font-medium">C &gt; 2.8 (주의)</div>
                <div class="p-1 rounded grade-badge-D font-medium">D &gt; 4.5 (중단)</div>
              </div>
            </div>
          </div>
        </div>

        <div class="mt-4 pt-3 border-t border-slate-700/60 text-[11px] text-slate-400 flex items-center justify-between">
          <span>* 4포인트 계측: M-DE, M-NDE, P-DE, P-NDE</span>
          <span class="text-cyan-400 font-medium">현장 표준시트 연계</span>
        </div>
      </div>
    </section>

    <!-- Main Table Section -->
    <section class="glass-card rounded-xl border border-slate-700/60 overflow-hidden">
      <div class="p-4 sm:p-5 border-b border-slate-700/60 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 class="text-base font-bold text-white flex items-center gap-2">
            <i class="fa-solid fa-list-check text-cyan-400"></i>
            설비별 4개 지점 정밀 진단 대장
          </h2>
          <div class="flex items-center gap-2 mt-0.5">
            <p id="tableSubtitle" class="text-xs text-slate-400">전동기/펌프 부하·반부하 12축 진단 및 정비 계획</p>
            <span class="text-slate-600">|</span>
            <span id="filteredCount" class="text-xs font-mono text-cyan-400 font-semibold">총 0개 설비</span>
          </div>
        </div>

        <!-- Filter Controls -->
        <div class="flex flex-wrap items-center gap-2">
          <button onclick="exportToCSV()" class="bg-emerald-700/80 hover:bg-emerald-600 text-white text-xs font-semibold px-3 py-2 rounded-lg border border-emerald-500/50 transition flex items-center gap-1.5 shadow" title="현재 목록 엑셀 파일(.xlsx)로 내려받기">
            <i class="fa-solid fa-file-excel"></i>
            <span>엑셀 다운로드</span>
          </button>

          <div class="relative">
            <select id="officeFilterDropdown" onchange="syncOfficeFilter(this.value)" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2 focus:ring-2 focus:ring-cyan-500 focus:outline-none pr-8 cursor-pointer">
              <option value="ALL">전체 사업소</option>
            </select>
          </div>

          <div class="relative">
            <select id="groupFilter" onchange="renderTable()" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2 focus:ring-2 focus:ring-cyan-500 focus:outline-none pr-8 cursor-pointer">
              <option value="ALL">전체 용량 구분</option>
              <option value="G1">Group 1 (>300kW)</option>
              <option value="G2">Group 2 (15~300kW)</option>
              <option value="G4">Group 4 (<15kW)</option>
            </select>
          </div>

          <div class="relative">
            <select id="statusFilter" onchange="renderTable()" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2 focus:ring-2 focus:ring-cyan-500 focus:outline-none pr-8 cursor-pointer">
              <option value="ALL">전체 평가 등급</option>
              <option value="D">D (즉시점검)</option>
              <option value="C">C (보수필요)</option>
              <option value="AB">A/B (양호)</option>
            </select>
          </div>

          <div class="relative w-full sm:w-44">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
            <input type="text" id="searchInput" oninput="renderTable()" placeholder="설비명/부위 검색..." class="w-full bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg pl-8 pr-3 py-2 focus:ring-2 focus:ring-cyan-500 focus:outline-none">
          </div>
        </div>
      </div>

      <!-- Table Container -->
      <div class="overflow-x-auto custom-scrollbar">
        <table class="w-full text-left border-collapse text-xs">
          <thead>
            <tr class="bg-slate-800/80 text-slate-400 border-b border-slate-700 text-[11px] uppercase tracking-wider font-semibold">
              <th class="py-3 px-3">사업소</th>
              <th class="py-3 px-3">측정시기</th>
              <th class="py-3 px-4">설비명</th>
              <th class="py-3 px-3">정격용량</th>
              <th class="py-3 px-3 text-center">전동기 Max</th>
              <th class="py-3 px-3 text-center">펌프 Max</th>
              <th class="py-3 px-3 text-center">최대 RMS</th>
              <th class="py-3 px-3 text-center">종합평가</th>
              <th class="py-3 px-4">현장 진단 요약 및 이상 부위</th>
              <th class="py-3 px-4">후속 정비대책</th>
              <th class="py-3 px-3 text-center">관리</th>
            </tr>
          </thead>
          <tbody id="tableBody" class="divide-y divide-slate-800 text-slate-300">
            <!-- Dynamically Rendered -->
          </tbody>
        </table>
      </div>
    </section>
  </main>

  <!-- MODAL: Edit Equipment Modal -->
  <div id="editModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-2xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeEditModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-5 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-cyan-500/20 text-cyan-400 border border-cyan-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-pen-to-square"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">설비 정보 및 계측값 수정</h3>
          <p class="text-xs text-slate-400">설비명, 용량, 12축 계측값을 수정하면 ISO 등급이 즉시 재계산됩니다.</p>
        </div>
      </div>

      <form id="editEquipForm" onsubmit="handleEquipmentEditSubmit(event)" class="space-y-4 text-xs">
        <input type="hidden" id="edit_id">

        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div>
            <label class="block text-slate-300 font-medium mb-1">소속 사업소 <span class="text-rose-400">*</span></label>
            <select id="edit_office" required class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500"></select>
          </div>

          <div>
            <label class="block text-slate-300 font-medium mb-1">측정 시기 <span class="text-rose-400">*</span></label>
            <select id="edit_period" required class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500"></select>
          </div>

          <div>
            <label class="block text-slate-300 font-medium mb-1">설비명 <span class="text-rose-400">*</span></label>
            <input type="text" id="edit_name" required placeholder="설비명 입력" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500 font-bold">
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block text-slate-300 font-medium mb-1">정격 용량 (kW) <span class="text-rose-400">*</span></label>
            <input type="number" id="edit_power" required step="0.1" min="0.1" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500 font-mono font-bold text-cyan-300">
          </div>

          <div>
            <label class="block text-slate-300 font-medium mb-1">기초 지지 조건</label>
            <select id="edit_foundation" onchange="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
              <option value="rigid">강성 기초 (Rigid)</option>
              <option value="flexible">탄성 기초 (Flexible)</option>
            </select>
          </div>
        </div>

        <!-- 4개 포인트(12축) 계측값 입력 그리드 -->
        <div class="p-3 rounded-xl bg-slate-800/80 border border-slate-700 space-y-3">
          <span class="font-bold text-cyan-300 block text-xs"><i class="fa-solid fa-arrows-to-dot mr-1.5"></i>4개 계측 지점 진동속도 RMS (mm/s)</span>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
            <div class="p-2.5 rounded-lg bg-slate-900/60 border border-slate-700">
              <span class="text-[11px] font-bold text-amber-300 block mb-1.5"><i class="fa-solid fa-bolt mr-1"></i>1. 전동기 부하측 (M-DE)</span>
              <div class="grid grid-cols-3 gap-1.5">
                <div>
                  <span class="text-[10px] text-slate-400 block">수평 H(X)</span>
                  <input type="number" id="edit_m_de_h" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">수직 V(Y)</span>
                  <input type="number" id="edit_m_de_v" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">축방향 A(Z)</span>
                  <input type="number" id="edit_m_de_a" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
              </div>
            </div>

            <div class="p-2.5 rounded-lg bg-slate-900/60 border border-slate-700">
              <span class="text-[11px] font-bold text-amber-300 block mb-1.5"><i class="fa-solid fa-bolt mr-1"></i>2. 전동기 반부하측 (M-NDE)</span>
              <div class="grid grid-cols-3 gap-1.5">
                <div>
                  <span class="text-[10px] text-slate-400 block">수평 H(X)</span>
                  <input type="number" id="edit_m_nde_h" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">수직 V(Y)</span>
                  <input type="number" id="edit_m_nde_v" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">축방향 A(Z)</span>
                  <input type="number" id="edit_m_nde_a" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
              </div>
            </div>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div class="p-2.5 rounded-lg bg-slate-900/60 border border-slate-700">
              <span class="text-[11px] font-bold text-cyan-300 block mb-1.5"><i class="fa-solid fa-fan mr-1"></i>3. 펌프 부하측 (P-DE)</span>
              <div class="grid grid-cols-3 gap-1.5">
                <div>
                  <span class="text-[10px] text-slate-400 block">수평 H(X)</span>
                  <input type="number" id="edit_p_de_h" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">수직 V(Y)</span>
                  <input type="number" id="edit_p_de_v" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">축방향 A(Z)</span>
                  <input type="number" id="edit_p_de_a" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
              </div>
            </div>

            <div class="p-2.5 rounded-lg bg-slate-900/60 border border-slate-700">
              <span class="text-[11px] font-bold text-cyan-300 block mb-1.5"><i class="fa-solid fa-fan mr-1"></i>4. 펌프 반부하측 (P-NDE)</span>
              <div class="grid grid-cols-3 gap-1.5">
                <div>
                  <span class="text-[10px] text-slate-400 block">수평 H(X)</span>
                  <input type="number" id="edit_p_nde_h" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">수직 V(Y)</span>
                  <input type="number" id="edit_p_nde_v" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">축방향 A(Z)</span>
                  <input type="number" id="edit_p_nde_a" step="0.01" min="0" oninput="autoCalcEditGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
              </div>
            </div>
          </div>
        </div>

        <div id="editAutoEvalBanner" class="p-3 rounded-xl border bg-slate-800/80 border-slate-700 flex items-center justify-between">
          <div>
            <span class="text-[11px] text-slate-400 block font-mono" id="editEvalGroupText">ISO 분류: Group 2</span>
            <span class="text-xs font-semibold text-slate-200" id="editEvalCriteriaText">기준치: A/B ≤ 2.8 | C ≤ 4.5 | D > 4.5</span>
            <span class="text-[11px] text-cyan-300 block mt-0.5" id="editEvalMaxPointText">최대 진동 부위: -</span>
          </div>
          <div class="text-right">
            <span id="editEvalBadge" class="inline-block px-3 py-1 rounded text-xs font-black grade-badge-C">C등급</span>
          </div>
        </div>

        <div>
          <label class="block text-slate-300 font-medium mb-1">현장 진단 요약 및 이상 부위</label>
          <input type="text" id="edit_summary" placeholder="진단 내용 입력" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
        </div>

        <div>
          <label class="block text-slate-300 font-medium mb-1">후속 정비대책</label>
          <input type="text" id="edit_action" placeholder="조치 대책 입력" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
        </div>

        <div class="pt-3 border-t border-slate-700 flex justify-end gap-2">
          <button type="button" onclick="closeEditModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition">취소</button>
          <button type="submit" class="px-5 py-2 bg-cyan-600 hover:bg-cyan-500 text-white font-bold rounded-lg transition shadow-md">수정사항 저장</button>
        </div>
      </form>
    </div>
  </div>

  <!-- MODAL 1: Add Custom Equipment with 4-Point 12-Axis Input -->
  <div id="addModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-2xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeAddModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-5 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-emerald-500/20 text-emerald-400 border border-emerald-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-plus"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">신규 설비 직접 등록 (4개 포인트 12축)</h3>
          <p class="text-xs text-slate-400">용량(kW) 및 부위별 실측 진동 입력 시 ISO 20816-3 등급과 위험 부위가 자동 산출됩니다.</p>
        </div>
      </div>

      <form id="addEquipForm" onsubmit="handleEquipmentSubmit(event)" class="space-y-4 text-xs">
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div>
            <label class="block text-slate-300 font-medium mb-1">소속 사업소 <span class="text-rose-400">*</span></label>
            <select id="inpOffice" required class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500"></select>
          </div>

          <div>
            <label class="block text-slate-300 font-medium mb-1">측정 시기 <span class="text-rose-400">*</span></label>
            <select id="inpPeriod" required class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500"></select>
          </div>

          <div>
            <label class="block text-slate-300 font-medium mb-1">설비명 <span class="text-rose-400">*</span></label>
            <input type="text" id="inpName" required placeholder="예: #3 응축수펌프" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block text-slate-300 font-medium mb-1">정격 용량 (kW) <span class="text-rose-400">*</span></label>
            <input type="number" id="inpPower" required step="0.1" min="0.1" value="75" oninput="autoCalcGrade()" placeholder="kW 단위" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500 font-mono">
          </div>

          <div>
            <label class="block text-slate-300 font-medium mb-1">기초 지지 조건</label>
            <select id="inpFoundation" onchange="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
              <option value="rigid" selected>강성 기초 (Rigid)</option>
              <option value="flexible">탄성 기초 (Flexible)</option>
            </select>
          </div>
        </div>

        <!-- 4개 포인트(12축) 계측값 입력 그리드 -->
        <div class="p-3 rounded-xl bg-slate-800/80 border border-slate-700 space-y-3">
          <span class="font-bold text-cyan-300 block text-xs"><i class="fa-solid fa-arrows-to-dot mr-1.5"></i>4개 계측 지점 진동속도 RMS (mm/s)</span>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
            <div class="p-2.5 rounded-lg bg-slate-900/60 border border-slate-700">
              <span class="text-[11px] font-bold text-amber-300 block mb-1.5"><i class="fa-solid fa-bolt mr-1"></i>1. 전동기 부하측 (M-DE)</span>
              <div class="grid grid-cols-3 gap-1.5">
                <div>
                  <span class="text-[10px] text-slate-400 block">수평 H(X)</span>
                  <input type="number" id="inp_m_de_h" step="0.01" min="0" value="2.8" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">수직 V(Y)</span>
                  <input type="number" id="inp_m_de_v" step="0.01" min="0" value="1.8" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">축방향 A(Z)</span>
                  <input type="number" id="inp_m_de_a" step="0.01" min="0" value="1.5" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
              </div>
            </div>

            <div class="p-2.5 rounded-lg bg-slate-900/60 border border-slate-700">
              <span class="text-[11px] font-bold text-amber-300 block mb-1.5"><i class="fa-solid fa-bolt mr-1"></i>2. 전동기 반부하측 (M-NDE)</span>
              <div class="grid grid-cols-3 gap-1.5">
                <div>
                  <span class="text-[10px] text-slate-400 block">수평 H(X)</span>
                  <input type="number" id="inp_m_nde_h" step="0.01" min="0" value="1.9" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">수직 V(Y)</span>
                  <input type="number" id="inp_m_nde_v" step="0.01" min="0" value="1.2" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">축방향 A(Z)</span>
                  <input type="number" id="inp_m_nde_a" step="0.01" min="0" value="1.1" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
              </div>
            </div>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div class="p-2.5 rounded-lg bg-slate-900/60 border border-slate-700">
              <span class="text-[11px] font-bold text-cyan-300 block mb-1.5"><i class="fa-solid fa-fan mr-1"></i>3. 펌프 부하측 (P-DE)</span>
              <div class="grid grid-cols-3 gap-1.5">
                <div>
                  <span class="text-[10px] text-slate-400 block">수평 H(X)</span>
                  <input type="number" id="inp_p_de_h" step="0.01" min="0" value="3.4" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">수직 V(Y)</span>
                  <input type="number" id="inp_p_de_v" step="0.01" min="0" value="2.2" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">축방향 A(Z)</span>
                  <input type="number" id="inp_p_de_a" step="0.01" min="0" value="2.0" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
              </div>
            </div>

            <div class="p-2.5 rounded-lg bg-slate-900/60 border border-slate-700">
              <span class="text-[11px] font-bold text-cyan-300 block mb-1.5"><i class="fa-solid fa-fan mr-1"></i>4. 펌프 반부하측 (P-NDE)</span>
              <div class="grid grid-cols-3 gap-1.5">
                <div>
                  <span class="text-[10px] text-slate-400 block">수평 H(X)</span>
                  <input type="number" id="inp_p_nde_h" step="0.01" min="0" value="2.3" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">수직 V(Y)</span>
                  <input type="number" id="inp_p_nde_v" step="0.01" min="0" value="1.7" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
                <div>
                  <span class="text-[10px] text-slate-400 block">축방향 A(Z)</span>
                  <input type="number" id="inp_p_nde_a" step="0.01" min="0" value="1.4" oninput="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded p-1 text-slate-100 font-mono text-center">
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- 실시간 자동 계산 미리보기 배너 -->
        <div id="autoEvalBanner" class="p-3.5 rounded-xl border bg-slate-800/80 border-slate-700 flex items-center justify-between">
          <div>
            <span class="text-[11px] text-slate-400 block font-mono" id="evalGroupText">ISO 분류: Group 2 (15~300kW)</span>
            <span class="text-xs font-semibold text-slate-200" id="evalCriteriaText">기준치: A/B ≤ 2.8 | C ≤ 4.5 | D > 4.5</span>
            <span class="text-[11px] text-cyan-300 block mt-0.5" id="evalMaxPointText">최대 진동 부위: 펌프 부하측 수평 (3.40 mm/s)</span>
          </div>
          <div class="text-right">
            <span id="evalBadge" class="inline-block px-3 py-1 rounded text-xs font-black grade-badge-C">C등급 (보수필요)</span>
          </div>
        </div>

        <div>
          <label class="block text-slate-300 font-medium mb-1">현장 진단 요약 및 이상 부위</label>
          <input type="text" id="inpSummary" placeholder="예: 펌프 부하측 베어링 진동 및 미세 발열 확인" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
        </div>

        <div>
          <label class="block text-slate-300 font-medium mb-1">후속 정비대책</label>
          <input type="text" id="inpAction" placeholder="예: 차기 계획정비 시 베어링 교체 및 축정렬 재실시" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
        </div>

        <div class="pt-3 border-t border-slate-700 flex justify-end gap-2">
          <button type="button" onclick="closeAddModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition">취소</button>
          <button type="submit" class="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-lg transition shadow-md">설비 등록 완료</button>
        </div>
      </form>
    </div>
  </div>

  <!-- MODAL 2: Excel & CSV File Upload with Auto Parsing -->
  <div id="uploadModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-lg rounded-2xl shadow-2xl p-6 relative">
      <button onclick="closeUploadModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-4 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-blue-500/20 text-blue-400 border border-blue-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-file-arrow-up"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">현장 엑셀 파일 스마트 업로드</h3>
          <p class="text-xs text-slate-400">최고값(Max), 등급 계산 없이 X/Y축 데이터만 있는 세로형 엑셀을 그대로 올리세요.</p>
        </div>
      </div>

      <div class="space-y-4 text-xs">
        <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
          <label class="text-slate-300 font-medium">데이터 반영 측정시기:</label>
          <select id="uploadPeriodSelect" class="bg-slate-900 border border-slate-600 rounded px-2.5 py-1 text-cyan-300 font-bold focus:outline-none"></select>
        </div>

        <div class="p-3 rounded-xl bg-slate-800/80 border border-slate-700 flex items-center justify-between">
          <div>
            <span class="font-bold text-slate-200 block">AI 세로형 양식 자동 파싱 지원</span>
            <span class="text-slate-400 text-[11px]">여러 줄로 나뉜 측정위치와 X, Y축 값을 시스템이 스스로 묶어냅니다.</span>
          </div>
        </div>

        <div id="dropZone" onclick="document.getElementById('fileInput').click()" class="border-2 border-dashed border-slate-600 hover:border-cyan-500 rounded-2xl p-8 text-center cursor-pointer transition bg-slate-800/40 hover:bg-slate-800/70">
          <i class="fa-solid fa-cloud-arrow-up text-3xl text-cyan-400 mb-2"></i>
          <p class="text-slate-200 font-bold mb-1">현장 엑셀 파일을 여기로 드래그하세요</p>
          <p class="text-slate-400 text-[11px]">양식 변경 금지 (있는 그대로 업로드)</p>
          <input type="file" id="fileInput" accept=".xlsx, .xls, .csv" onchange="handleFileSelect(event)" class="hidden">
        </div>

        <div id="uploadStatusBox" class="hidden p-3 rounded-lg border text-xs"></div>

        <div class="pt-3 border-t border-slate-700 flex justify-between items-center">
          <span class="text-[11px] text-slate-400">12축 최대치 자동 계산 및 ISO 판정 처리됨</span>
          <button onclick="closeUploadModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition">닫기</button>
        </div>
      </div>
    </div>
  </div>

  <!-- MODAL 3: Add Period Modal -->
  <div id="addPeriodModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-sm rounded-2xl shadow-2xl p-6 relative">
      <button onclick="closeAddPeriodModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-4 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-cyan-500/20 text-cyan-400 border border-cyan-500/40 flex items-center justify-center text-lg">
          <i class="fa-regular fa-calendar-plus"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">신규 측정시기 추가</h3>
          <p class="text-xs text-slate-400">반기별 점검 주기를 등록합니다.</p>
        </div>
      </div>

      <form id="addPeriodForm" onsubmit="handlePeriodSubmit(event)" class="space-y-4 text-xs">
        <div>
          <label class="block text-slate-300 font-medium mb-1">측정 시기 명칭 <span class="text-rose-400">*</span></label>
          <input type="text" id="inpNewPeriodName" required placeholder="예: 2027년 상반기" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500 font-bold">
        </div>

        <div class="pt-3 border-t border-slate-700 flex justify-end gap-2">
          <button type="button" onclick="closeAddPeriodModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition">취소</button>
          <button type="submit" class="px-5 py-2 bg-cyan-600 hover:bg-cyan-500 text-white font-bold rounded-lg transition shadow-md">주기 등록</button>
        </div>
      </form>
    </div>
  </div>

  <!-- MODAL 4: Manage & Add / Delete Office Modal -->
  <div id="addOfficeModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-lg rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeAddOfficeModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-5 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-indigo-500/20 text-indigo-400 border border-indigo-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-building-circle-gear"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">사업소 추가 및 삭제 관리</h3>
          <p class="text-xs text-slate-400">신규 사업소를 추가하거나 기존 사업소를 삭제할 수 있습니다.</p>
        </div>
      </div>

      <!-- Add Office Form -->
      <form id="addOfficeForm" onsubmit="handleOfficeSubmit(event)" class="space-y-3 text-xs pb-4 border-b border-slate-700/80">
        <span class="font-bold text-slate-200 block"><i class="fa-solid fa-plus text-indigo-400 mr-1"></i>새 사업소 등록</span>
        <div>
          <label class="block text-slate-300 font-medium mb-1">사업소 명칭 <span class="text-rose-400">*</span></label>
          <input type="text" id="inpOfficeName" required placeholder="예: 남부사업소, 서부사업소, 기장사업소" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-indigo-500">
        </div>

        <div>
          <label class="block text-slate-300 font-medium mb-1">대표 아이콘 테마</label>
          <div class="grid grid-cols-4 gap-2 text-center text-xs">
            <label class="p-2 rounded-lg bg-slate-800 border border-slate-700 cursor-pointer hover:border-indigo-400 flex flex-col items-center gap-1">
              <input type="radio" name="officeIcon" value="fa-fire-flame-curved" checked class="text-indigo-500">
              <i class="fa-solid fa-fire-flame-curved text-amber-400"></i>
              <span class="text-[10px] text-slate-300">소각/열</span>
            </label>
            <label class="p-2 rounded-lg bg-slate-800 border border-slate-700 cursor-pointer hover:border-indigo-400 flex flex-col items-center gap-1">
              <input type="radio" name="officeIcon" value="fa-water" class="text-indigo-500">
              <i class="fa-solid fa-water text-cyan-400"></i>
              <span class="text-[10px] text-slate-300">하수/물</span>
            </label>
            <label class="p-2 rounded-lg bg-slate-800 border border-slate-700 cursor-pointer hover:border-indigo-400 flex flex-col items-center gap-1">
              <input type="radio" name="officeIcon" value="fa-industry" class="text-indigo-500">
              <i class="fa-solid fa-industry text-blue-400"></i>
              <span class="text-[10px] text-slate-300">위생/공장</span>
            </label>
            <label class="p-2 rounded-lg bg-slate-800 border border-slate-700 cursor-pointer hover:border-indigo-400 flex flex-col items-center gap-1">
              <input type="radio" name="officeIcon" value="fa-solar-panel" class="text-indigo-500">
              <i class="fa-solid fa-solar-panel text-emerald-400"></i>
              <span class="text-[10px] text-slate-300">친환경</span>
            </label>
          </div>
        </div>

        <div class="flex justify-end pt-1">
          <button type="submit" class="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-bold rounded-lg transition shadow-md flex items-center gap-1.5">
            <i class="fa-solid fa-plus"></i>
            <span>사업소 등록</span>
          </button>
        </div>
      </form>

      <!-- Office List & Delete Area -->
      <div class="pt-4 text-xs space-y-2">
        <span class="font-bold text-slate-200 block"><i class="fa-solid fa-list-check text-cyan-400 mr-1"></i>등록된 사업소 목록 (삭제 관리)</span>
        <div id="officeManageList" class="space-y-1.5 max-h-48 overflow-y-auto custom-scrollbar pr-1">
          <!-- Dynamically populated -->
        </div>
      </div>

      <div class="pt-4 mt-4 border-t border-slate-700 flex justify-end">
        <button type="button" onclick="closeAddOfficeModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition text-xs">닫기</button>
      </div>
    </div>
  </div>

  <!-- MODAL 5: Interactive ISO Calculator -->
  <div id="calcModal" class="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-lg rounded-2xl shadow-2xl p-6 relative">
      <button onclick="closeCalcModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-5">
        <div class="h-10 w-10 rounded-xl bg-cyan-500/20 text-cyan-400 border border-cyan-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-gauge-high"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">현장 실시간 진동 판정기</h3>
          <p class="text-xs text-slate-400">ISO 20816-3 기준에 따른 기계 상태 즉시 진단</p>
        </div>
      </div>

      <div class="space-y-4 text-xs">
        <div>
          <label class="block text-slate-300 font-medium mb-1.5">설비 정격 용량 (Group)</label>
          <select id="calcGroup" onchange="runEvaluation()" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
            <option value="G1">Group 1 : 대형 기계 (&gt; 300 kW, 중온수펌프·대형팬)</option>
            <option value="G2" selected>Group 2 : 중형 기계 (15 kW ~ 300 kW, 급수·소각순환)</option>
            <option value="G4">Group 4 : 소형 기계 (&lt; 15 kW, 일반 냉각수펌프)</option>
          </select>
        </div>

        <div>
          <label class="block text-slate-300 font-medium mb-1.5">기초 지지 조건 (Foundation)</label>
          <div class="grid grid-cols-2 gap-2">
            <label class="flex items-center gap-2 p-2.5 rounded-lg bg-slate-800 border border-slate-700 cursor-pointer hover:border-slate-600">
              <input type="radio" name="calcFoundation" value="rigid" checked onchange="runEvaluation()" class="text-cyan-500 focus:ring-0">
              <span class="text-slate-300">강성 기초 (Rigid)</span>
            </label>
            <label class="flex items-center gap-2 p-2.5 rounded-lg bg-slate-800 border border-slate-700 cursor-pointer hover:border-slate-600">
              <input type="radio" name="calcFoundation" value="flexible" onchange="runEvaluation()" class="text-cyan-500 focus:ring-0">
              <span class="text-slate-300">탄성 기초 (Flexible)</span>
            </label>
          </div>
        </div>

        <div>
          <div class="flex justify-between items-center mb-1.5">
            <label class="text-slate-300 font-medium">측정 속도 RMS (mm/s)</label>
            <span id="calcValDisplay" class="font-mono text-cyan-400 font-bold text-sm">3.40 mm/s</span>
          </div>
          <input type="range" id="calcVibSlider" min="0" max="15" step="0.1" value="3.4" oninput="updateSliderVal(this.value)" class="w-full accent-cyan-500 bg-slate-700 h-2 rounded-lg cursor-pointer">
          <div class="flex justify-between text-[10px] text-slate-500 mt-1 font-mono">
            <span>0.0</span>
            <span>2.8</span>
            <span>4.5</span>
            <span>7.1</span>
            <span>11.0</span>
            <span>15.0+</span>
          </div>
        </div>

        <div id="calcResultBox" class="mt-4 p-4 rounded-xl border transition-all duration-300">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <span id="calcGradeBadge" class="text-lg font-black px-2.5 py-0.5 rounded-lg grade-badge-C">Grade C</span>
              <span id="calcGradeTitle" class="text-sm font-bold text-amber-400">보수 필요 (주의 단계)</span>
            </div>
            <span id="calcBoundaryText" class="text-[11px] font-mono text-slate-400">2.8 &lt; RMS &le; 4.5</span>
          </div>
          <p id="calcRecommendation" class="mt-2 text-slate-300 text-[11px] leading-relaxed">
            단기간 운전은 가능하나 결함 진행 중입니다. 주파수 분석(FFT)을 통한 원인 규명 및 차기 정비 시 부품 교체를 권고합니다.
          </p>
        </div>
      </div>

      <div class="mt-5 pt-3 border-t border-slate-700 flex justify-end">
        <button onclick="closeCalcModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold rounded-lg transition">닫기</button>
      </div>
    </div>
  </div>

  <!-- MODAL 6: Detail Equipment Modal -->
  <div id="detailModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-3xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeDetailModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div id="detailModalContent"></div>
    </div>
  </div>

  <!-- MODAL 7: Google Sheets Integration Modal -->
  <div id="sheetModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeSheetModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-5 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-emerald-500/20 text-emerald-400 border border-emerald-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-table-cells"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">구글 시트(Google Sheets) 중앙 DB 연동</h3>
          <p class="text-xs text-slate-400">구글 시트를 클라우드 DB로 사용하여 여러 사용자와 실시간 데이터를 공유합니다.</p>
        </div>
      </div>

      <div class="space-y-4 text-xs">
        <div>
          <label class="block text-slate-300 font-medium mb-1">구글 웹 앱 배포 URL (Google Apps Script Web App URL)</label>
          <input type="url" id="inpSheetUrl" placeholder="https://script.google.com/macros/s/.../exec" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-emerald-500 font-mono text-[11px]">
          <p class="text-[11px] text-slate-400 mt-1">구글 스프레드시트의 [확장 프로그램 > Apps Script]에서 배포된 웹 앱 주소를 붙여넣으세요.</p>
        </div>

        <div class="flex items-center gap-2 pt-2">
          <button onclick="saveSheetUrl()" class="flex-1 bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-2 rounded-lg transition shadow">
            <i class="fa-solid fa-link mr-1"></i>URL 저장 및 연결
          </button>
          <button onclick="disconnectSheet()" class="px-3 py-2 bg-slate-800 hover:bg-rose-950 text-slate-400 hover:text-rose-400 rounded-lg transition border border-slate-700">
            연결 해제
          </button>
        </div>

        <div class="p-3.5 rounded-xl bg-slate-800/80 border border-slate-700 space-y-2">
          <span class="font-bold text-slate-200 block text-xs"><i class="fa-solid fa-arrows-rotate text-cyan-400 mr-1.5"></i>데이터 동기화 작업</span>
          <div class="grid grid-cols-2 gap-2">
            <button onclick="syncFromGoogleSheets()" class="p-2.5 rounded-lg bg-slate-700 hover:bg-slate-600 text-cyan-300 font-bold border border-slate-600 transition flex items-center justify-center gap-1.5">
              <i class="fa-solid fa-cloud-arrow-down"></i>
              <span>구글시트에서 불러오기</span>
            </button>
            <button onclick="syncToGoogleSheets()" class="p-2.5 rounded-lg bg-emerald-800/60 hover:bg-emerald-700 text-emerald-200 font-bold border border-emerald-600/50 transition flex items-center justify-center gap-1.5">
              <i class="fa-solid fa-cloud-arrow-up"></i>
              <span>현재 데이터 시트에 저장</span>
            </button>
          </div>
          <div id="sheetSyncStatusBox" class="hidden p-2.5 rounded border text-[11px] font-medium"></div>
        </div>

        <div class="p-3 rounded-lg bg-slate-950/60 border border-slate-800 text-slate-400 leading-relaxed text-[11px]">
          <p class="font-bold text-slate-300 mb-1"><i class="fa-solid fa-circle-info text-amber-400 mr-1"></i>구글 시트 연동 간단 3단계:</p>
          <ol class="list-decimal list-inside space-y-1">
            <li>구글 드라이브에서 새 스프레드시트를 생성합니다.</li>
            <li>[확장 프로그램] → [Apps Script] 클릭 후 제공된 연동 코드를 붙여넣습니다.</li>
            <li>[배포] → [새 배포] (액세스 권한: 모든 사용자) 후 나오는 URL을 위에 입력합니다.</li>
          </ol>
        </div>
      </div>
    </div>
  </div>

  <!-- JavaScript Application Logic with LocalStorage & Google Sheets Sync -->
  <script>
    let googleSheetAppUrl = localStorage.getItem('becovib_google_sheet_url') || '';
    let periodList = ['2025년 상반기', '2025년 하반기', '2026년 상반기', '2026년 하반기'];
    let currentPeriod = '2026년 상반기';

    const defaultEquipmentData = [
      {
        id: 'fd_fan',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '강제송풍기 (FD Fan)',
        power: 310,
        group: 'G1',
        foundation: 'rigid',
        totalQty: 4,
        unoperated: '-',
        normalList: '#2, #3',
        warningList: '-',
        dangerList: '#1, #4',
        overallGrade: 'D',
        maxRms: 7.3,
        points: {
          m_de:  { h: 6.8, v: 4.1, a: 3.5 },
          m_nde: { h: 3.2, v: 2.1, a: 1.8 },
          p_de:  { h: 7.3, v: 5.2, a: 4.8 },
          p_nde: { h: 4.5, v: 3.1, a: 2.6 }
        },
        axes: {
          motor: { h: 6.8, v: 4.1, a: 3.5 },
          pump:  { h: 7.3, v: 5.2, a: 4.8 }
        },
        history: { '2025년 상반기': 3.8, '2025년 하반기': 4.9, '2026년 상반기': 7.3, '2026년 하반기': null },
        summary: '#1 : 전동기 진동, #4 : 송풍기 부하측 수평진동(7.3) 심각',
        action: '#1, #4 전동기 및 팬 베어링 분해정비 요망',
        detailNotes: '송풍기 부하측(P-DE) 수평(X축) 진동이 7.3mm/s로 D등급 임계치 초과.'
      },
      {
        id: 'boiler_feed',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '보일러 급수펌프',
        power: 95,
        group: 'G2',
        foundation: 'rigid',
        totalQty: 6,
        unoperated: '#5, #6',
        normalList: '#1',
        warningList: '#2, #3, #4',
        dangerList: '-',
        overallGrade: 'C',
        maxRms: 3.9,
        points: {
          m_de:  { h: 3.4, v: 2.1, a: 1.8 },
          m_nde: { h: 1.8, v: 1.2, a: 1.0 },
          p_de:  { h: 3.9, v: 2.8, a: 2.2 },
          p_nde: { h: 2.5, v: 1.9, a: 1.5 }
        },
        axes: {
          motor: { h: 3.4, v: 2.1, a: 1.8 },
          pump:  { h: 3.9, v: 2.8, a: 2.2 }
        },
        history: { '2025년 상반기': 2.1, '2025년 하반기': 2.7, '2026년 상반기': 3.9, '2026년 하반기': null },
        summary: '#2 : 전동기/펌프 진동, #3·#4 : 펌프 부하측 베어링 진동',
        action: '#2, #3, #4 호기 하절기 정비 시 베어링 및 축정렬 보수',
        detailNotes: '펌프 부하측 수평(X축) 3.9mm/s로 C등급(주의). 커플링 엘레먼트 마모 및 축정렬 불량 추정.'
      },
      {
        id: 'hot_water_circ',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '중온수 순환펌프',
        power: 350,
        group: 'G1',
        foundation: 'rigid',
        totalQty: 6,
        unoperated: '#1, #2, #3',
        normalList: '#5, #6',
        warningList: '#4',
        dangerList: '-',
        overallGrade: 'B',
        maxRms: 3.2,
        points: {
          m_de:  { h: 3.2, v: 2.0, a: 1.9 },
          m_nde: { h: 1.9, v: 1.3, a: 1.1 },
          p_de:  { h: 2.6, v: 1.8, a: 1.7 },
          p_nde: { h: 2.1, v: 1.5, a: 1.2 }
        },
        axes: {
          motor: { h: 3.2, v: 2.0, a: 1.9 },
          pump:  { h: 2.6, v: 1.8, a: 1.7 }
        },
        history: { '2025년 상반기': 2.8, '2025년 하반기': 3.0, '2026년 상반기': 3.2, '2026년 하반기': null },
        summary: '#3 : 액츄에이터 제작 중, #4 : 전동기 부하측 진동 및 이음',
        action: '#3호기 밸브 제작 완료 후 교체, #4호기 향후 보수',
        detailNotes: '모터 부하측 수평방향 3.2mm/s로 B등급(양호 상한선).'
      },
      {
        id: 'incin_pump',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '소각장 순환펌프',
        power: 95,
        group: 'G2',
        foundation: 'rigid',
        totalQty: 3,
        unoperated: '-',
        normalList: '-',
        warningList: '#9',
        dangerList: '#7, #8',
        overallGrade: 'D',
        maxRms: 6.2,
        points: {
          m_de:  { h: 5.8, v: 3.9, a: 3.1 },
          m_nde: { h: 2.9, v: 2.1, a: 1.7 },
          p_de:  { h: 6.2, v: 4.4, a: 3.8 },
          p_nde: { h: 3.7, v: 2.5, a: 2.1 }
        },
        axes: {
          motor: { h: 5.8, v: 3.9, a: 3.1 },
          pump:  { h: 6.2, v: 4.4, a: 3.8 }
        },
        history: { '2025년 상반기': 3.2, '2025년 하반기': 4.4, '2026년 상반기': 6.2, '2026년 하반기': null },
        summary: '#7, #8 : 전동기/펌프 부하측 진동 심각, #9 : M/Seal 미세누수',
        action: '#7, #8호기 즉시 분해 점검, #9호기 씰 교체',
        detailNotes: '부하측 모터(5.8) 및 펌프(6.2) 모두 D등급 초과. 베어링 소착 위험.'
      },
      {
        id: 'heat_pump',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '히트펌프 1·2호기 (저온/고온열)',
        power: 55,
        group: 'G2',
        foundation: 'rigid',
        totalQty: 10,
        unoperated: '-',
        normalList: '-',
        warningList: '-',
        dangerList: '저온(#1,#2) / 고온(#1~#3)',
        overallGrade: 'D',
        maxRms: 7.4,
        points: {
          m_de:  { h: 5.2, v: 3.8, a: 4.1 },
          m_nde: { h: 2.6, v: 1.9, a: 1.8 },
          p_de:  { h: 6.9, v: 4.9, a: 7.4 },
          p_nde: { h: 3.4, v: 2.7, a: 2.9 }
        },
        axes: {
          motor: { h: 5.2, v: 3.8, a: 4.1 },
          pump:  { h: 6.9, v: 4.9, a: 7.4 }
        },
        history: { '2025년 상반기': 3.5, '2025년 하반기': 5.2, '2026년 상반기': 7.4, '2026년 하반기': null },
        summary: '저온열 : 메카니칼 씰 누수, 고온열 : 펌프 부하측 축방향(Z축: 7.4) 진동 과다',
        action: '저온열 즉시보수 실시, 고온열 베어링 순차적 분해정비',
        detailNotes: '고온열 펌프 부하측 축방향(Z축) 진동이 7.4mm/s로 D등급 임계치 초과.'
      }
    ];

    let equipmentData = [];

    const defaultOfficeList = [
      { name: '에너지사업소', icon: 'fa-fire-flame-curved', color: 'text-amber-400' },
      { name: '명지사업소', icon: 'fa-industry', color: 'text-cyan-400' },
      { name: '해운대사업소', icon: 'fa-water', color: 'text-blue-400' }
    ];

    let officeList = [];

    function saveToLocalStorage() {
      try {
        localStorage.setItem('becovib_equipment_data', JSON.stringify(equipmentData));
        localStorage.setItem('becovib_office_list', JSON.stringify(officeList));
        localStorage.setItem('becovib_period_list', JSON.stringify(periodList));
        localStorage.setItem('becovib_current_period', currentPeriod);
        if (googleSheetAppUrl) {
          localStorage.setItem('becovib_google_sheet_url', googleSheetAppUrl);
        }
      } catch (err) {
        console.warn('LocalStorage save failed:', err);
      }
    }

    function updateSyncBadge() {
      const badge = document.getElementById('syncStatusBadge');
      if (!badge) return;
      if (googleSheetAppUrl) {
        badge.className = 'text-[10px] font-semibold px-2 py-0.5 rounded-full bg-cyan-950 text-cyan-300 border border-cyan-700 cursor-pointer';
        badge.innerHTML = '<i class="fa-solid fa-cloud mr-1"></i>구글시트 연동중';
      } else {
        badge.className = 'text-[10px] font-semibold px-2 py-0.5 rounded-full bg-emerald-950 text-emerald-400 border border-emerald-800 cursor-pointer';
        badge.innerHTML = '<i class="fa-solid fa-hard-drive mr-1"></i>로컬 저장';
      }
    }

    function openSheetModal() {
      document.getElementById('sheetModal').classList.remove('hidden');
      document.getElementById('inpSheetUrl').value = googleSheetAppUrl;
      const statusBox = document.getElementById('sheetSyncStatusBox');
      statusBox.classList.add('hidden');
    }

    function closeSheetModal() {
      document.getElementById('sheetModal').classList.add('hidden');
    }

    function saveSheetUrl() {
      const url = document.getElementById('inpSheetUrl').value.trim();
      const statusBox = document.getElementById('sheetSyncStatusBox');
      statusBox.classList.remove('hidden');

      if (!url) {
        statusBox.className = 'p-2.5 rounded border text-[11px] bg-red-950/20 border-red-500/40 text-red-400';
        statusBox.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-1"></i>웹 앱 URL을 입력해 주세요.';
        return;
      }

      googleSheetAppUrl = url;
      localStorage.setItem('becovib_google_sheet_url', url);
      updateSyncBadge();

      statusBox.className = 'p-2.5 rounded border text-[11px] bg-emerald-950/20 border-emerald-500/40 text-emerald-400';
      statusBox.innerHTML = '<i class="fa-solid fa-check mr-1"></i>구글 시트 연동 URL이 저장되었습니다.';
    }

    function disconnectSheet() {
      googleSheetAppUrl = '';
      localStorage.removeItem('becovib_google_sheet_url');
      document.getElementById('inpSheetUrl').value = '';
      updateSyncBadge();
      const statusBox = document.getElementById('sheetSyncStatusBox');
      statusBox.classList.remove('hidden');
      statusBox.className = 'p-2.5 rounded border text-[11px] bg-slate-800 border-slate-700 text-slate-300';
      statusBox.innerHTML = '구글 시트 연동이 해제되었습니다. 로컬 저장소 모드로 작동합니다.';
    }

    async function syncFromGoogleSheets() {
      const statusBox = document.getElementById('sheetSyncStatusBox');
      statusBox.classList.remove('hidden');

      if (!googleSheetAppUrl) {
        statusBox.className = 'p-2.5 rounded border text-[11px] bg-amber-950/20 border-amber-500/40 text-amber-400';
        statusBox.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-1"></i>먼저 웹 앱 URL을 입력하고 저장해 주세요.';
        return;
      }

      statusBox.className = 'p-2.5 rounded border text-[11px] bg-cyan-950/20 border-cyan-500/40 text-cyan-300';
      statusBox.innerHTML = '<i class="fa-solid fa-spinner fa-spin mr-1"></i>구글 시트에서 최신 데이터를 가져오는 중...';

      try {
        const response = await fetch(googleSheetAppUrl);
        const result = await response.json();

        if (result.status === 'success' && Array.isArray(result.data) && result.data.length > 0) {
          equipmentData = result.data;
          
          equipmentData.forEach(item => {
            if (item.period && !periodList.includes(item.period)) periodList.push(item.period);
            ensureOfficeRegistered(item.office);
          });

          saveToLocalStorage();
          renderPeriodSelectors();
          renderOfficeTabs();
          updateDashboard();

          statusBox.className = 'p-2.5 rounded border text-[11px] bg-emerald-950/20 border-emerald-500/40 text-emerald-400 font-semibold';
          statusBox.innerHTML = `<i class="fa-solid fa-circle-check mr-1"></i>구글 시트에서 총 ${equipmentData.length}개의 설비 데이터를 성공적으로 동기화했습니다!`;
        } else {
          statusBox.className = 'p-2.5 rounded border text-[11px] bg-amber-950/20 border-amber-500/40 text-amber-300';
          statusBox.innerHTML = '구글 시트가 비어 있거나 올바른 데이터 형식이 아닙니다.';
        }
      } catch (err) {
        console.error(err);
        statusBox.className = 'p-2.5 rounded border text-[11px] bg-red-950/20 border-red-500/40 text-red-400';
        statusBox.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-1"></i>동기화 실패: 웹 앱 권한 설정 또는 URL을 확인해 주세요.';
      }
    }

    async function syncToGoogleSheets() {
      const statusBox = document.getElementById('sheetSyncStatusBox');
      statusBox.classList.remove('hidden');

      if (!googleSheetAppUrl) {
        statusBox.className = 'p-2.5 rounded border text-[11px] bg-amber-950/20 border-amber-500/40 text-amber-400';
        statusBox.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-1"></i>먼저 웹 앱 URL을 입력하고 저장해 주세요.';
        return;
      }

      statusBox.className = 'p-2.5 rounded border text-[11px] bg-cyan-950/20 border-cyan-500/40 text-cyan-300';
      statusBox.innerHTML = '<i class="fa-solid fa-spinner fa-spin mr-1"></i>현재 대장 데이터를 구글 시트에 기록 중...';

      try {
        const payload = JSON.stringify({
          action: 'save',
          data: equipmentData
        });

        await fetch(googleSheetAppUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'application/json' },
          body: payload
        });

        statusBox.className = 'p-2.5 rounded border text-[11px] bg-emerald-950/20 border-emerald-500/40 text-emerald-400 font-semibold';
        statusBox.innerHTML = `<i class="fa-solid fa-circle-check mr-1"></i>구글 스프레드시트에 성공적으로 저장 전송되었습니다!`;
      } catch (err) {
        console.error(err);
        statusBox.className = 'p-2.5 rounded border text-[11px] bg-red-950/20 border-red-500/40 text-red-400';
        statusBox.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-1"></i>저장 실패: 네트워크 또는 Apps Script 설정을 확인해 주세요.';
      }
    }

    function loadFromLocalStorage() {
      try {
        const savedEquip = localStorage.getItem('becovib_equipment_data');
        const savedOffices = localStorage.getItem('becovib_office_list');
        const savedPeriods = localStorage.getItem('becovib_period_list');
        const savedCurPeriod = localStorage.getItem('becovib_current_period');

        if (savedEquip) equipmentData = JSON.parse(savedEquip);
        else equipmentData = JSON.parse(JSON.stringify(defaultEquipmentData));

        if (savedOffices) officeList = JSON.parse(savedOffices);
        else officeList = JSON.parse(JSON.stringify(defaultOfficeList));

        if (savedPeriods) periodList = JSON.parse(savedPeriods);
        if (savedCurPeriod && periodList.includes(savedCurPeriod)) currentPeriod = savedCurPeriod;

      } catch (err) {
        console.warn('LocalStorage load failed, using defaults:', err);
        equipmentData = JSON.parse(JSON.stringify(defaultEquipmentData));
        officeList = JSON.parse(JSON.stringify(defaultOfficeList));
      }
    }

    function resetToInitialData() {
      if (confirm('저장된 설비 및 사업소 데이터를 최초 기본값으로 초기화하시겠습니까?')) {
        localStorage.removeItem('becovib_equipment_data');
        localStorage.removeItem('becovib_office_list');
        localStorage.removeItem('becovib_period_list');
        localStorage.removeItem('becovib_current_period');
        loadFromLocalStorage();
        renderPeriodSelectors();
        renderOfficeTabs();
        updateDashboard();
      }
    }

    let currentOffice = 'ALL';
    let currentChartMode = 'trend';
    let vibrationChart = null;
    let detailChart = null;
    let detailAxesChart = null;

    function renderPeriodSelectors() {
      const mainSel = document.getElementById('periodFilterSelect');
      const modalSel = document.getElementById('inpPeriod');
      const uploadSel = document.getElementById('uploadPeriodSelect');

      let options = '';
      periodList.forEach(p => {
        options += `<option value="${p}" ${p === currentPeriod ? 'selected' : ''}>${p}</option>`;
      });

      if (mainSel) mainSel.innerHTML = options;
      if (modalSel) modalSel.innerHTML = options;
      if (uploadSel) uploadSel.innerHTML = options;
    }

    function switchPeriod(val) {
      currentPeriod = val;
      saveToLocalStorage();
      updateDashboard();
    }

    function openAddPeriodModal() {
      document.getElementById('addPeriodModal').classList.remove('hidden');
      document.getElementById('inpNewPeriodName').focus();
    }
    function closeAddPeriodModal() {
      document.getElementById('addPeriodModal').classList.add('hidden');
    }

    function handlePeriodSubmit(e) {
      e.preventDefault();
      const newP = document.getElementById('inpNewPeriodName').value.trim();
      if (!newP) return;

      if (!periodList.includes(newP)) {
        periodList.push(newP);
      }
      currentPeriod = newP;
      saveToLocalStorage();
      renderPeriodSelectors();
      document.getElementById('addPeriodForm').reset();
      closeAddPeriodModal();
      updateDashboard();
    }

    function renderOfficeTabs() {
      const container = document.getElementById('officeTabs');
      const dropdown = document.getElementById('officeFilterDropdown');
      const selectModal = document.getElementById('inpOffice');

      let tabsHtml = `
        <button onclick="switchOffice('ALL')" class="office-tab px-3 py-1.5 rounded-lg text-xs font-bold transition shrink-0 ${currentOffice === 'ALL' ? 'bg-cyan-600 text-white shadow' : 'bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white'}" data-office="ALL">
          전체 사업소
        </button>
      `;

      officeList.forEach(off => {
        const isSelected = currentOffice === off.name;
        tabsHtml += `
          <div class="inline-flex items-center rounded-lg transition shrink-0 ${isSelected ? 'bg-cyan-600 text-white shadow' : 'bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white'}">
            <button onclick="switchOffice('${off.name}')" class="px-2.5 py-1.5 text-xs font-semibold flex items-center gap-1.5" data-office="${off.name}">
              <i class="fa-solid ${off.icon} ${isSelected ? 'text-white' : off.color}"></i>
              <span>${off.name}</span>
            </button>
            <button onclick="deleteOffice('${off.name}', event)" class="pr-2 pl-1 py-1.5 text-[11px] opacity-60 hover:opacity-100 hover:text-rose-300 transition" title="${off.name} 삭제">
              <i class="fa-solid fa-xmark"></i>
            </button>
          </div>
        `;
      });
      container.innerHTML = tabsHtml;

      let dropdownHtml = `<option value="ALL">전체 사업소</option>`;
      officeList.forEach(off => {
        dropdownHtml += `<option value="${off.name}">${off.name}</option>`;
      });
      dropdown.innerHTML = dropdownHtml;
      dropdown.value = currentOffice;

      if (selectModal) {
        let selectHtml = '';
        officeList.forEach(off => {
          selectHtml += `<option value="${off.name}">${off.name}</option>`;
        });
        selectModal.innerHTML = selectHtml;
      }

      renderOfficeManageList();
    }

    function renderOfficeManageList() {
      const listContainer = document.getElementById('officeManageList');
      if (!listContainer) return;

      if (officeList.length === 0) {
        listContainer.innerHTML = '<p class="text-slate-500 py-2 text-center">등록된 사업소가 없습니다.</p>';
        return;
      }

      listContainer.innerHTML = officeList.map(off => {
        const count = equipmentData.filter(e => e.office === off.name).length;
        return `
          <div class="flex items-center justify-between p-2 rounded-lg bg-slate-800/80 border border-slate-700">
            <div class="flex items-center gap-2">
              <i class="fa-solid ${off.icon} ${off.color}"></i>
              <span class="font-bold text-slate-200">${off.name}</span>
              <span class="text-[10px] text-slate-400 bg-slate-900 px-1.5 py-0.5 rounded border border-slate-700">설비 ${count}대</span>
            </div>
            <button type="button" onclick="deleteOffice('${off.name}')" class="px-2 py-1 rounded bg-rose-950/40 hover:bg-rose-900/80 text-rose-300 hover:text-white border border-rose-800/50 text-[11px] transition flex items-center gap-1">
              <i class="fa-solid fa-trash-can"></i>
              <span>삭제</span>
            </button>
          </div>
        `;
      }).join('');
    }

    function openAddOfficeModal() {
      document.getElementById('addOfficeModal').classList.remove('hidden');
      renderOfficeManageList();
      document.getElementById('inpOfficeName').focus();
    }
    function closeAddOfficeModal() {
      document.getElementById('addOfficeModal').classList.add('hidden');
    }

    function handleOfficeSubmit(e) {
      e.preventDefault();
      const name = document.getElementById('inpOfficeName').value.trim();
      const icon = document.querySelector('input[name="officeIcon"]:checked').value;

      if (!name) return;

      if (officeList.some(o => o.name === name)) {
        alert('이미 등록되어 있는 사업소 명칭입니다.');
        return;
      }

      let color = 'text-cyan-400';
      if (icon === 'fa-fire-flame-curved') color = 'text-amber-400';
      else if (icon === 'fa-water') color = 'text-blue-400';
      else if (icon === 'fa-solar-panel') color = 'text-emerald-400';

      officeList.push({ name, icon, color });
      saveToLocalStorage();
      document.getElementById('addOfficeForm').reset();
      renderOfficeTabs();
      switchOffice(name);
    }

    function deleteOffice(officeName, evt) {
      if (evt) evt.stopPropagation();

      const count = equipmentData.filter(e => e.office === officeName).length;
      let msg = `[${officeName}] 사업소를 삭제하시겠습니까?`;
      if (count > 0) {
        msg = `[${officeName}] 사업소에 등록된 설비 데이터가 총 ${count}건 있습니다.\n사업소를 삭제하면 소속 설비 데이터도 함께 삭제됩니다.\n정말 삭제하시겠습니까?`;
      }

      if (!confirm(msg)) return;

      officeList = officeList.filter(o => o.name !== officeName);
      if (count > 0) {
        equipmentData = equipmentData.filter(e => e.office !== officeName);
      }

      if (currentOffice === officeName) {
        currentOffice = 'ALL';
      }

      saveToLocalStorage();
      renderOfficeTabs();
      updateDashboard();
    }

    function ensureOfficeRegistered(officeName) {
      if (!officeName || officeName === 'ALL') return;
      if (/^\d{4,5}$/.test(officeName)) return;

      if (!officeList.some(o => o.name === officeName)) {
        officeList.push({
          name: officeName,
          icon: 'fa-industry',
          color: 'text-indigo-400'
        });
        saveToLocalStorage();
        renderOfficeTabs();
      }
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

    function switchOffice(officeName) {
      currentOffice = officeName;

      document.querySelectorAll('.office-tab').forEach(tab => {
        if (tab.dataset.office === officeName) {
          tab.className = 'office-tab px-3 py-1.5 rounded-lg text-xs font-bold transition shrink-0 bg-cyan-600 text-white shadow';
        } else {
          tab.className = 'office-tab px-3 py-1.5 rounded-lg text-xs font-semibold transition shrink-0 bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white';
        }
      });

      document.getElementById('officeFilterDropdown').value = officeName;
      document.getElementById('tableSubtitle').innerText = officeName === 'ALL' 
        ? '전체 사업소 회전기계 상태진단 및 조치 대장' 
        : `${officeName} 주요 회전설비 상태진단 및 조치 대장`;

      updateDashboard();
    }

    function syncOfficeFilter(val) {
      switchOffice(val);
    }

    function setChartMode(mode) {
      currentChartMode = mode;
      const btnTrend = document.getElementById('btnModeTrend');
      const btnBar = document.getElementById('btnModeBar');

      if (mode === 'trend') {
        btnTrend.className = 'px-3 py-1 rounded font-bold bg-cyan-600 text-white transition';
        btnBar.className = 'px-3 py-1 rounded font-medium text-slate-300 hover:text-white transition';
        document.getElementById('chartMainTitle').innerText = '주요 설비 반기별 진동 추세선 (Trend)';
        document.getElementById('chartSubTitle').innerText = '과거 반기 측정치 누적 변화와 ISO 20816-3 경보 기준치 비교';
      } else {
        btnBar.className = 'px-3 py-1 rounded font-bold bg-cyan-600 text-white transition';
        btnTrend.className = 'px-3 py-1 rounded font-medium text-slate-300 hover:text-white transition';
        document.getElementById('chartMainTitle').innerText = `${currentPeriod} 설비별 최대 진동치 현황`;
        document.getElementById('chartSubTitle').innerText = '현 주기 대상 설비의 실측 진동값 vs ISO 20816-3 경보 기준 비교';
      }
      renderChart();
    }

    function updateDashboard() {
      const activeData = equipmentData.filter(e => {
        const officeMatch = (currentOffice === 'ALL' || e.office === currentOffice);
        const periodMatch = (!e.period || e.period === currentPeriod);
        return officeMatch && periodMatch;
      });

      const totalUnits = activeData.reduce((acc, cur) => acc + (cur.totalQty || 1), 0);
      const gradeA = activeData.filter(e => e.overallGrade === 'A').length;
      const gradeB = activeData.filter(e => e.overallGrade === 'B').length;
      const gradeC = activeData.filter(e => e.overallGrade === 'C').length;
      const gradeD = activeData.filter(e => e.overallGrade === 'D').length;

      document.getElementById('metric-total').innerText = totalUnits;
      document.getElementById('metric-ab').innerText = (gradeA + gradeB);
      document.getElementById('metric-c').innerText = gradeC;
      document.getElementById('metric-d').innerText = gradeD;

      const pct = activeData.length > 0 ? Math.round(((gradeA + gradeB) / activeData.length) * 100) : 0;
      document.getElementById('metric-ab-pct').innerText = pct + '%';

      renderTable();
      renderChart();
    }

    function renderChart() {
      const ctx = document.getElementById('vibrationChart').getContext('2d');
      const activeData = equipmentData.filter(e => {
        const officeMatch = (currentOffice === 'ALL' || e.office === currentOffice);
        const periodMatch = (!e.period || e.period === currentPeriod);
        return officeMatch && periodMatch;
      });

      if (vibrationChart) vibrationChart.destroy();

      if (currentChartMode === 'trend') {
        const targetEquipList = (activeData.length > 0 ? activeData : equipmentData).slice(0, 5);
        const colorPalette = ['#ef4444', '#f59e0b', '#06b6d4', '#10b981', '#8b5cf6'];

        const datasets = targetEquipList.map((eq, idx) => {
          const trendData = periodList.map(p => {
            if (eq.history && eq.history[p] !== undefined) return eq.history[p];
            return eq.period === p ? eq.maxRms : null;
          });

          return {
            label: `${eq.name} (${eq.office.replace('사업소','')})`,
            data: trendData,
            borderColor: colorPalette[idx % colorPalette.length],
            backgroundColor: colorPalette[idx % colorPalette.length],
            borderWidth: 2.5,
            pointRadius: 4.5,
            pointHoverRadius: 7,
            tension: 0.3,
            spanGaps: true
          };
        });

        vibrationChart = new Chart(ctx, {
          type: 'line',
          data: {
            labels: periodList,
            datasets: datasets
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: {
                position: 'top',
                labels: { color: '#cbd5e1', font: { size: 11 }, usePointStyle: true }
              },
              tooltip: {
                backgroundColor: '#0f172a',
                titleColor: '#38bdf8',
                bodyColor: '#f1f5f9',
                borderColor: '#334155',
                borderWidth: 1,
                callbacks: {
                  label: function(c) {
                    return `${c.dataset.label}: ${c.parsed.y !== null ? c.parsed.y + ' mm/s' : '미측정'}`;
                  }
                }
              }
            },
            scales: {
              x: {
                grid: { color: 'rgba(51, 65, 85, 0.3)' },
                ticks: { color: '#94a3b8', font: { size: 11 } }
              },
              y: {
                min: 0,
                max: 9,
                grid: { color: 'rgba(51, 65, 85, 0.3)' },
                ticks: {
                  color: '#94a3b8',
                  font: { size: 11 },
                  callback: function(v) { return v + ' mm/s'; }
                }
              }
            }
          }
        });

      } else {
        const labels = activeData.map(e => `${e.name.split(' ')[0]} (${e.office.replace('사업소','')})`);
        const dataValues = activeData.map(e => e.maxRms);
        const backgroundColors = activeData.map(e => {
          if (e.overallGrade === 'D') return 'rgba(239, 68, 68, 0.85)';
          if (e.overallGrade === 'C') return 'rgba(245, 158, 11, 0.85)';
          if (e.overallGrade === 'A') return 'rgba(16, 185, 129, 0.85)';
          return 'rgba(6, 182, 212, 0.85)';
        });

        vibrationChart = new Chart(ctx, {
          type: 'bar',
          data: {
            labels: labels,
            datasets: [{
              label: '최대 실측 RMS (mm/s)',
              data: dataValues,
              backgroundColor: backgroundColors,
              borderRadius: 6,
              borderWidth: 0,
              barThickness: 24,
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: {
                backgroundColor: '#0f172a',
                titleColor: '#38bdf8',
                bodyColor: '#f1f5f9',
                borderColor: '#334155',
                borderWidth: 1,
                callbacks: {
                  label: function(c) {
                    const item = activeData[c.dataIndex];
                    return [
                      `사업소: ${item.office}`,
                      `실측 RMS: ${c.parsed.y} mm/s`,
                      `평가 등급: ${item.overallGrade}등급`,
                      `정격 용량: ${item.power} kW (${item.group})`
                    ];
                  }
                }
              }
            },
            scales: {
              x: {
                grid: { color: 'rgba(51, 65, 85, 0.3)' },
                ticks: { color: '#94a3b8', font: { size: 10 } }
              },
              y: {
                min: 0,
                max: 9,
                grid: { color: 'rgba(51, 65, 85, 0.3)' },
                ticks: {
                  color: '#94a3b8',
                  font: { size: 11 },
                  callback: function(v) { return v + ' mm/s'; }
                }
              }
            }
          }
        });
      }
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
        if (query) {
          const combined = `${item.name} ${item.office} ${item.summary} ${item.action}`.toLowerCase();
          if (!combined.includes(query)) return false;
        }
        return true;
      });

      document.getElementById('filteredCount').innerText = `총 ${filtered.length}개 설비 목록 표시 중`;

      if (filtered.length === 0) {
        tbody.innerHTML = `
          <tr>
            <td colspan="11" class="text-center py-8 text-slate-500">
              <i class="fa-solid fa-filter-circle-xmark text-2xl mb-2 block"></i>
              <span>조건에 부합하는 설비 데이터가 없습니다.</span>
            </td>
          </tr>
        `;
        return;
      }

      tbody.innerHTML = filtered.map(item => {
        let badgeClass = 'grade-badge-B';
        let badgeLabel = 'B (양호)';
        if (item.overallGrade === 'D') { badgeClass = 'grade-badge-D'; badgeLabel = 'D (즉시점검)'; }
        else if (item.overallGrade === 'C') { badgeClass = 'grade-badge-C'; badgeLabel = 'C (보수필요)'; }
        else if (item.overallGrade === 'A') { badgeClass = 'grade-badge-A'; badgeLabel = 'A (최상)'; }

        let officeColor = item.office === '에너지사업소' ? 'text-amber-400 bg-amber-950/30 border-amber-800/40' :
                          (item.office === '명지사업소' ? 'text-cyan-400 bg-cyan-950/30 border-cyan-800/40' : 'text-blue-400 bg-blue-950/30 border-blue-800/40');

        let motorMax = 0, pumpMax = 0;
        if (item.points) {
          motorMax = Math.max(item.points.m_de.h, item.points.m_de.v, item.points.m_de.a, item.points.m_nde.h, item.points.m_nde.v, item.points.m_nde.a);
          pumpMax = Math.max(item.points.p_de.h, item.points.p_de.v, item.points.p_de.a, item.points.p_nde.h, item.points.p_nde.v, item.points.p_nde.a);
        } else if (item.axes) {
          motorMax = Math.max(item.axes.motor.h, item.axes.motor.v, item.axes.motor.a);
          pumpMax = Math.max(item.axes.pump.h, item.axes.pump.v, item.axes.pump.a);
        } else {
          motorMax = item.maxRms; pumpMax = item.maxRms;
        }

        return `
          <tr class="hover:bg-slate-800/50 transition">
            <td class="py-3 px-3">
              <span class="inline-block px-2 py-0.5 rounded text-[10px] font-semibold border ${officeColor}">
                ${item.office}
              </span>
            </td>
            <td class="py-3 px-3 font-mono text-[11px] text-slate-400">${item.period || currentPeriod}</td>
            <td class="py-3 px-4 font-bold text-slate-100 flex items-center gap-2">
              <i class="fa-solid fa-gear text-slate-400"></i>
              ${item.name}
            </td>
            <td class="py-3 px-3 font-mono text-cyan-300 font-semibold">${item.power} kW</td>
            <td class="py-3 px-3 text-center font-mono font-bold text-amber-300">${motorMax.toFixed(2)}</td>
            <td class="py-3 px-3 text-center font-mono font-bold text-cyan-300">${pumpMax.toFixed(2)}</td>
            <td class="py-3 px-3 text-center font-mono font-black ${item.overallGrade === 'D' ? 'text-red-400' : (item.overallGrade === 'C' ? 'text-amber-400' : 'text-emerald-400')}">
              ${parseFloat(item.maxRms).toFixed(2)}
            </td>
            <td class="py-3 px-3 text-center">
              <span class="inline-block px-2.5 py-1 rounded text-[11px] font-bold ${badgeClass}">
                ${badgeLabel}
              </span>
            </td>
            <td class="py-3 px-4 text-slate-300 max-w-xs truncate" title="${item.summary}">
              ${item.summary}
            </td>
            <td class="py-3 px-4 text-slate-400 max-w-xs truncate" title="${item.action}">
              <span class="${item.overallGrade === 'D' ? 'text-red-400 font-medium' : ''}">${item.action}</span>
            </td>
            <td class="py-3 px-3 text-center">
              <div class="flex items-center justify-center gap-1">
                <button onclick="openEditModal('${item.id}')" class="px-2 py-1 rounded bg-slate-800 hover:bg-slate-700 text-amber-300 hover:text-amber-200 text-[11px] font-medium border border-slate-700 transition" title="수정">
                  수정
                </button>
                <button onclick="viewEquipmentDetail('${item.id}')" class="px-2 py-1 rounded bg-slate-800 hover:bg-slate-700 text-cyan-400 text-[11px] font-medium border border-slate-700 transition" title="상세보기">
                  상세
                </button>
                <button onclick="deleteEquipment('${item.id}')" class="px-2 py-1 rounded bg-slate-800 hover:bg-rose-950 text-slate-400 hover:text-rose-400 text-[11px] border border-slate-700 transition" title="삭제">
                  <i class="fa-solid fa-trash-can"></i>
                </button>
              </div>
            </td>
          </tr>
        `;
      }).join('');
    }

    function openEditModal(id) {
      const item = equipmentData.find(e => e.id === id);
      if (!item) return;

      const officeSelect = document.getElementById('edit_office');
      const periodSelect = document.getElementById('edit_period');

      let offOpts = '';
      officeList.forEach(o => {
        offOpts += `<option value="${o.name}" ${o.name === item.office ? 'selected' : ''}>${o.name}</option>`;
      });
      officeSelect.innerHTML = offOpts;

      let perOpts = '';
      periodList.forEach(p => {
        perOpts += `<option value="${p}" ${p === (item.period || currentPeriod) ? 'selected' : ''}>${p}</option>`;
      });
      periodSelect.innerHTML = perOpts;

      document.getElementById('edit_id').value = item.id;
      document.getElementById('edit_name').value = item.name || '';
      document.getElementById('edit_power').value = item.power || 55;
      document.getElementById('edit_foundation').value = item.foundation || 'rigid';

      const pts = item.points || {
        m_de: { h: item.maxRms || 0, v: 0, a: 0 },
        m_nde: { h: 0, v: 0, a: 0 },
        p_de: { h: item.maxRms || 0, v: 0, a: 0 },
        p_nde: { h: 0, v: 0, a: 0 }
      };

      document.getElementById('edit_m_de_h').value = pts.m_de.h;
      document.getElementById('edit_m_de_v').value = pts.m_de.v;
      document.getElementById('edit_m_de_a').value = pts.m_de.a;

      document.getElementById('edit_m_nde_h').value = pts.m_nde.h;
      document.getElementById('edit_m_nde_v').value = pts.m_nde.v;
      document.getElementById('edit_m_nde_a').value = pts.m_nde.a;

      document.getElementById('edit_p_de_h').value = pts.p_de.h;
      document.getElementById('edit_p_de_v').value = pts.p_de.v;
      document.getElementById('edit_p_de_a').value = pts.p_de.a;

      document.getElementById('edit_p_nde_h').value = pts.p_nde.h;
      document.getElementById('edit_p_nde_v').value = pts.p_nde.v;
      document.getElementById('edit_p_nde_a').value = pts.p_nde.a;

      document.getElementById('edit_summary').value = item.summary || '';
      document.getElementById('edit_action').value = item.action || '';

      autoCalcEditGrade();
      document.getElementById('editModal').classList.remove('hidden');
    }

    function closeEditModal() {
      document.getElementById('editModal').classList.add('hidden');
    }

    function autoCalcEditGrade() {
      const power = parseFloat(document.getElementById('edit_power').value) || 0;
      const foundation = document.getElementById('edit_foundation').value;

      const m_de_h = parseFloat(document.getElementById('edit_m_de_h').value) || 0;
      const m_de_v = parseFloat(document.getElementById('edit_m_de_v').value) || 0;
      const m_de_a = parseFloat(document.getElementById('edit_m_de_a').value) || 0;
      const m_nde_h = parseFloat(document.getElementById('edit_m_nde_h').value) || 0;
      const m_nde_v = parseFloat(document.getElementById('edit_m_nde_v').value) || 0;
      const m_nde_a = parseFloat(document.getElementById('edit_m_nde_a').value) || 0;

      const p_de_h = parseFloat(document.getElementById('edit_p_de_h').value) || 0;
      const p_de_v = parseFloat(document.getElementById('edit_p_de_v').value) || 0;
      const p_de_a = parseFloat(document.getElementById('edit_p_de_a').value) || 0;
      const p_nde_h = parseFloat(document.getElementById('edit_p_nde_h').value) || 0;
      const p_nde_v = parseFloat(document.getElementById('edit_p_nde_v').value) || 0;
      const p_nde_a = parseFloat(document.getElementById('edit_p_nde_a').value) || 0;

      const pointMap = [
        { name: '전동기 부하측 수평(X)', val: m_de_h },
        { name: '전동기 부하측 수직(Y)', val: m_de_v },
        { name: '전동기 부하측 축방향(Z)', val: m_de_a },
        { name: '전동기 반부하측 수평(X)', val: m_nde_h },
        { name: '전동기 반부하측 수직(Y)', val: m_nde_v },
        { name: '전동기 반부하측 축방향(Z)', val: m_nde_a },
        { name: '펌프 부하측 수평(X)', val: p_de_h },
        { name: '펌프 부하측 수직(Y)', val: p_de_v },
        { name: '펌프 부하측 축방향(Z)', val: p_de_a },
        { name: '펌프 반부하측 수평(X)', val: p_nde_h },
        { name: '펌프 반부하측 수직(Y)', val: p_nde_v },
        { name: '펌프 반부하측 축방향(Z)', val: p_nde_a }
      ];

      pointMap.sort((a, b) => b.val - a.val);
      const topPoint = pointMap[0];
      const maxRms = topPoint.val;

      const res = calculateIsoGrade(power, maxRms, foundation);
      const badge = document.getElementById('editEvalBadge');
      const groupText = document.getElementById('editEvalGroupText');
      const critText = document.getElementById('editEvalCriteriaText');
      const pointText = document.getElementById('editEvalMaxPointText');

      groupText.innerText = `ISO 분류: ${res.group} (${power > 300 ? '>300kW' : (power < 15 ? '<15kW' : '15~300kW')}) | 기초: ${foundation === 'rigid' ? '강성(Rigid)' : '탄성(Flexible)'}`;

      if (res.group === 'G1') {
        critText.innerText = foundation === 'rigid' ? '기준치: A/B ≤ 4.5 | C ≤ 7.1 | D > 7.1' : '기준치: A/B ≤ 7.1 | C ≤ 11.0 | D > 11.0';
      } else {
        critText.innerText = foundation === 'rigid' ? '기준치: A/B ≤ 2.8 | C ≤ 4.5 | D > 4.5' : '기준치: A/B ≤ 4.5 | C ≤ 7.1 | D > 7.1';
      }

      pointText.innerText = `최대 진동 부위: ${topPoint.name} (${topPoint.val.toFixed(2)} mm/s)`;

      badge.className = 'inline-block px-3 py-1 rounded text-xs font-black';
      if (res.grade === 'A') {
        badge.classList.add('grade-badge-A');
        badge.innerText = 'A등급 (최상)';
      } else if (res.grade === 'B') {
        badge.classList.add('grade-badge-B');
        badge.innerText = 'B등급 (양호)';
      } else if (res.grade === 'C') {
        badge.classList.add('grade-badge-C');
        badge.innerText = 'C등급 (보수필요)';
      } else {
        badge.classList.add('grade-badge-D');
        badge.innerText = 'D등급 (즉시점검)';
      }
    }

    function handleEquipmentEditSubmit(e) {
      e.preventDefault();
      const id = document.getElementById('edit_id').value;
      const item = equipmentData.find(x => x.id === id);
      if (!item) return;

      const office = document.getElementById('edit_office').value;
      const period = document.getElementById('edit_period').value;
      const name = document.getElementById('edit_name').value.trim();
      const power = parseFloat(document.getElementById('edit_power').value);
      const foundation = document.getElementById('edit_foundation').value;

      const m_de_h = parseFloat(document.getElementById('edit_m_de_h').value) || 0;
      const m_de_v = parseFloat(document.getElementById('edit_m_de_v').value) || 0;
      const m_de_a = parseFloat(document.getElementById('edit_m_de_a').value) || 0;
      const m_nde_h = parseFloat(document.getElementById('edit_m_nde_h').value) || 0;
      const m_nde_v = parseFloat(document.getElementById('edit_m_nde_v').value) || 0;
      const m_nde_a = parseFloat(document.getElementById('edit_m_nde_a').value) || 0;

      const p_de_h = parseFloat(document.getElementById('edit_p_de_h').value) || 0;
      const p_de_v = parseFloat(document.getElementById('edit_p_de_v').value) || 0;
      const p_de_a = parseFloat(document.getElementById('edit_p_de_a').value) || 0;
      const p_nde_h = parseFloat(document.getElementById('edit_p_nde_h').value) || 0;
      const p_nde_v = parseFloat(document.getElementById('edit_p_nde_v').value) || 0;
      const p_nde_a = parseFloat(document.getElementById('edit_p_nde_a').value) || 0;

      const maxRms = Math.max(m_de_h, m_de_v, m_de_a, m_nde_h, m_nde_v, m_nde_a, p_de_h, p_de_v, p_de_a, p_nde_h, p_nde_v, p_nde_a);
      const summary = document.getElementById('edit_summary').value.trim() || '측정 완료';
      const action = document.getElementById('edit_action').value.trim() || '현 상태 유지';

      const evalRes = calculateIsoGrade(power, maxRms, foundation);

      item.office = office;
      item.period = period;
      item.name = name;
      item.power = power;
      item.group = evalRes.group;
      item.foundation = foundation;
      item.overallGrade = evalRes.grade;
      item.maxRms = maxRms;
      item.points = {
        m_de:  { h: m_de_h, v: m_de_v, a: m_de_a },
        m_nde: { h: m_nde_h, v: m_nde_v, a: m_nde_a },
        p_de:  { h: p_de_h, v: p_de_v, a: p_de_a },
        p_nde: { h: p_nde_h, v: p_nde_v, a: p_nde_a }
      };
      item.axes = {
        motor: { h: Math.max(m_de_h, m_nde_h), v: Math.max(m_de_v, m_nde_v), a: Math.max(m_de_a, m_nde_a) },
        pump:  { h: Math.max(p_de_h, p_nde_h), v: Math.max(p_de_v, p_nde_v), a: Math.max(p_de_a, p_nde_a) }
      };
      if (!item.history) item.history = {};
      item.history[period] = maxRms;
      item.summary = summary;
      item.action = action;

      saveToLocalStorage();
      closeEditModal();
      updateDashboard();
    }

    function autoCalcGrade() {
      const power = parseFloat(document.getElementById('inpPower').value) || 0;
      const foundation = document.getElementById('inpFoundation').value;

      const m_de_h = parseFloat(document.getElementById('inp_m_de_h').value) || 0;
      const m_de_v = parseFloat(document.getElementById('inp_m_de_v').value) || 0;
      const m_de_a = parseFloat(document.getElementById('inp_m_de_a').value) || 0;

      const m_nde_h = parseFloat(document.getElementById('inp_m_nde_h').value) || 0;
      const m_nde_v = parseFloat(document.getElementById('inp_m_nde_v').value) || 0;
      const m_nde_a = parseFloat(document.getElementById('inp_m_nde_a').value) || 0;

      const p_de_h = parseFloat(document.getElementById('inp_p_de_h').value) || 0;
      const p_de_v = parseFloat(document.getElementById('inp_p_de_v').value) || 0;
      const p_de_a = parseFloat(document.getElementById('inp_p_de_a').value) || 0;

      const p_nde_h = parseFloat(document.getElementById('inp_p_nde_h').value) || 0;
      const p_nde_v = parseFloat(document.getElementById('inp_p_nde_v').value) || 0;
      const p_nde_a = parseFloat(document.getElementById('inp_p_nde_a').value) || 0;

      const pointMap = [
        { name: '전동기 부하측 수평(X)', val: m_de_h },
        { name: '전동기 부하측 수직(Y)', val: m_de_v },
        { name: '전동기 부하측 축방향(Z)', val: m_de_a },
        { name: '전동기 반부하측 수평(X)', val: m_nde_h },
        { name: '전동기 반부하측 수직(Y)', val: m_nde_v },
        { name: '전동기 반부하측 축방향(Z)', val: m_nde_a },
        { name: '펌프 부하측 수평(X)', val: p_de_h },
        { name: '펌프 부하측 수직(Y)', val: p_de_v },
        { name: '펌프 부하측 축방향(Z)', val: p_de_a },
        { name: '펌프 반부하측 수평(X)', val: p_nde_h },
        { name: '펌프 반부하측 수직(Y)', val: p_nde_v },
        { name: '펌프 반부하측 축방향(Z)', val: p_nde_a }
      ];

      pointMap.sort((a, b) => b.val - a.val);
      const topPoint = pointMap[0];
      const maxRms = topPoint.val;

      const res = calculateIsoGrade(power, maxRms, foundation);
      const badge = document.getElementById('evalBadge');
      const groupText = document.getElementById('evalGroupText');
      const critText = document.getElementById('evalCriteriaText');
      const pointText = document.getElementById('evalMaxPointText');

      groupText.innerText = `ISO 분류: ${res.group} (${power > 300 ? '>300kW' : (power < 15 ? '<15kW' : '15~300kW')}) | 기초: ${foundation === 'rigid' ? '강성(Rigid)' : '탄성(Flexible)'}`;

      if (res.group === 'G1') {
        critText.innerText = foundation === 'rigid' ? '기준치: A/B ≤ 4.5 | C ≤ 7.1 | D > 7.1' : '기준치: A/B ≤ 7.1 | C ≤ 11.0 | D > 11.0';
      } else {
        critText.innerText = foundation === 'rigid' ? '기준치: A/B ≤ 2.8 | C ≤ 4.5 | D > 4.5' : '기준치: A/B ≤ 4.5 | C ≤ 7.1 | D > 7.1';
      }

      pointText.innerText = `최대 진동 부위: ${topPoint.name} (${topPoint.val.toFixed(2)} mm/s)`;

      badge.className = 'inline-block px-3 py-1 rounded text-xs font-black';
      if (res.grade === 'A') {
        badge.classList.add('grade-badge-A');
        badge.innerText = 'A등급 (최상)';
      } else if (res.grade === 'B') {
        badge.classList.add('grade-badge-B');
        badge.innerText = 'B등급 (양호)';
      } else if (res.grade === 'C') {
        badge.classList.add('grade-badge-C');
        badge.innerText = 'C등급 (보수필요)';
      } else {
        badge.classList.add('grade-badge-D');
        badge.innerText = 'D등급 (즉시점검)';
      }
    }

    function handleEquipmentSubmit(e) {
      e.preventDefault();
      const office = document.getElementById('inpOffice').value;
      const period = document.getElementById('inpPeriod').value;
      const name = document.getElementById('inpName').value.trim();
      const power = parseFloat(document.getElementById('inpPower').value);
      const foundation = document.getElementById('inpFoundation').value;

      const m_de_h = parseFloat(document.getElementById('inp_m_de_h').value) || 0;
      const m_de_v = parseFloat(document.getElementById('inp_m_de_v').value) || 0;
      const m_de_a = parseFloat(document.getElementById('inp_m_de_a').value) || 0;

      const m_nde_h = parseFloat(document.getElementById('inp_m_nde_h').value) || 0;
      const m_nde_v = parseFloat(document.getElementById('inp_m_nde_v').value) || 0;
      const m_nde_a = parseFloat(document.getElementById('inp_m_nde_a').value) || 0;

      const p_de_h = parseFloat(document.getElementById('inp_p_de_h').value) || 0;
      const p_de_v = parseFloat(document.getElementById('inp_p_de_v').value) || 0;
      const p_de_a = parseFloat(document.getElementById('inp_p_de_a').value) || 0;

      const p_nde_h = parseFloat(document.getElementById('inp_p_nde_h').value) || 0;
      const p_nde_v = parseFloat(document.getElementById('inp_p_nde_v').value) || 0;
      const p_nde_a = parseFloat(document.getElementById('inp_p_nde_a').value) || 0;

      const maxRms = Math.max(m_de_h, m_de_v, m_de_a, m_nde_h, m_nde_v, m_nde_a, p_de_h, p_de_v, p_de_a, p_nde_h, p_nde_v, p_nde_a);
      const summary = document.getElementById('inpSummary').value.trim() || '측정 완료 및 이상 유무 점검';
      const action = document.getElementById('inpAction').value.trim() || '현 상태 모니터링 유지';

      const evalRes = calculateIsoGrade(power, maxRms, foundation);
      ensureOfficeRegistered(office);

      const newId = 'custom_' + Date.now();
      let historyObj = {};
      historyObj[period] = maxRms;

      const newEntry = {
        id: newId,
        office: office,
        period: period,
        name: name,
        power: power,
        group: evalRes.group,
        foundation: foundation,
        totalQty: 1,
        unoperated: '-',
        normalList: evalRes.grade === 'A' || evalRes.grade === 'B' ? '#1' : '-',
        warningList: evalRes.grade === 'C' ? '#1' : '-',
        dangerList: evalRes.grade === 'D' ? '#1' : '-',
        overallGrade: evalRes.grade,
        maxRms: maxRms,
        points: {
          m_de:  { h: m_de_h, v: m_de_v, a: m_de_a },
          m_nde: { h: m_nde_h, v: m_nde_v, a: m_nde_a },
          p_de:  { h: p_de_h, v: p_de_v, a: p_de_a },
          p_nde: { h: p_nde_h, v: p_nde_v, a: p_nde_a }
        },
        axes: {
          motor: { h: Math.max(m_de_h, m_nde_h), v: Math.max(m_de_v, m_nde_v), a: Math.max(m_de_a, m_nde_a) },
          pump:  { h: Math.max(p_de_h, p_nde_h), v: Math.max(p_de_v, p_nde_v), a: Math.max(p_de_a, p_nde_a) }
        },
        history: historyObj,
        summary: summary,
        action: action,
        detailNotes: '신규 등록 설비 (12축 정밀 계측)'
      };

      equipmentData.unshift(newEntry);
      saveToLocalStorage();
      closeAddModal();
      document.getElementById('addEquipForm').reset();

      if (currentPeriod !== period) switchPeriod(period);
      else if (currentOffice !== 'ALL' && currentOffice !== office) switchOffice(office);
      else updateDashboard();
    }

    function deleteEquipment(id) {
      if (!confirm('해당 설비 계측 데이터를 삭제하시겠습니까?')) return;
      equipmentData = equipmentData.filter(e => e.id !== id);
      saveToLocalStorage();
      updateDashboard();
    }

    function openAddModal() {
      document.getElementById('addModal').classList.remove('hidden');
      autoCalcGrade();
    }
    function closeAddModal() {
      document.getElementById('addModal').classList.add('hidden');
    }

    function openCalcModal() {
      document.getElementById('calcModal').classList.remove('hidden');
      runEvaluation();
    }
    function closeCalcModal() {
      document.getElementById('calcModal').classList.add('hidden');
    }

    function openUploadModal() {
      document.getElementById('uploadModal').classList.remove('hidden');
      document.getElementById('uploadStatusBox').classList.add('hidden');
    }
    function closeUploadModal() {
      document.getElementById('uploadModal').classList.add('hidden');
    }

    function openDetailModal() {
      document.getElementById('detailModal').classList.remove('hidden');
    }
    function closeDetailModal() {
      document.getElementById('detailModal').classList.add('hidden');
    }

    function viewEquipmentDetail(id) {
      const item = equipmentData.find(e => e.id === id);
      if (!item) return;

      const detailContainer = document.getElementById('detailModalContent');
      let badgeClass = 'grade-badge-B';
      let badgeLabel = 'B (양호)';
      if (item.overallGrade === 'D') { badgeClass = 'grade-badge-D'; badgeLabel = 'D (즉시점검)'; }
      else if (item.overallGrade === 'C') { badgeClass = 'grade-badge-C'; badgeLabel = 'C (보수필요)'; }
      else if (item.overallGrade === 'A') { badgeClass = 'grade-badge-A'; badgeLabel = 'A (최상)'; }

      let points = item.points || {
        m_de:  { h: item.axes?.motor.h || item.maxRms, v: item.axes?.motor.v || 0, a: item.axes?.motor.a || 0 },
        m_nde: { h: (item.axes?.motor.h || item.maxRms)*0.6, v: 0, a: 0 },
        p_de:  { h: item.axes?.pump.h || item.maxRms, v: item.axes?.pump.v || 0, a: item.axes?.pump.a || 0 },
        p_nde: { h: (item.axes?.pump.h || item.maxRms)*0.6, v: 0, a: 0 }
      };

      detailContainer.innerHTML = `
        <div class="flex items-center justify-between pb-3 border-b border-slate-700">
          <div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-slate-800 text-cyan-400 border border-slate-700">${item.office}</span>
              <span class="text-xs font-mono text-slate-400">${item.period || currentPeriod}</span>
            </div>
            <h3 class="text-lg font-bold text-white mt-1">${item.name}</h3>
          </div>
          <span class="px-3 py-1 rounded text-xs font-black ${badgeClass}">${badgeLabel}</span>
        </div>

        <div class="mt-4 p-3 bg-slate-800/80 rounded-xl border border-slate-700">
          <span class="text-xs font-bold text-slate-200 block mb-2"><i class="fa-solid fa-crosshairs text-cyan-400 mr-1.5"></i>4개 지점(DE/NDE) 12축 실측 RMS (mm/s)</span>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 text-center text-xs">
            <div class="bg-slate-900/80 p-2 rounded border border-slate-700">
              <span class="text-amber-400 font-semibold block text-[11px]">전동기 부하(M-DE)</span>
              <span class="font-mono text-slate-300 text-[10px]">H:${points.m_de.h} V:${points.m_de.v} A:${points.m_de.a}</span>
            </div>
            <div class="bg-slate-900/80 p-2 rounded border border-slate-700">
              <span class="text-amber-400 font-semibold block text-[11px]">전동기 반부하(M-NDE)</span>
              <span class="font-mono text-slate-300 text-[10px]">H:${points.m_nde.h} V:${points.m_nde.v} A:${points.m_nde.a}</span>
            </div>
            <div class="bg-slate-900/80 p-2 rounded border border-slate-700">
              <span class="text-cyan-400 font-semibold block text-[11px]">펌프 부하(P-DE)</span>
              <span class="font-mono text-slate-300 text-[10px]">H:${points.p_de.h} V:${points.p_de.v} A:${points.p_de.a}</span>
            </div>
            <div class="bg-slate-900/80 p-2 rounded border border-slate-700">
              <span class="text-cyan-400 font-semibold block text-[11px]">펌프 반부하(P-NDE)</span>
              <span class="font-mono text-slate-300 text-[10px]">H:${points.p_nde.h} V:${points.p_nde.v} A:${points.p_nde.a}</span>
            </div>
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 my-4">
          <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700">
            <span class="text-xs font-semibold text-slate-300 block mb-2"><i class="fa-solid fa-arrows-up-down-left-right text-cyan-400 mr-1.5"></i>모터 vs 펌프 방향별(H/V/A) 비교</span>
            <div class="h-44 w-full">
              <canvas id="detailAxesChart"></canvas>
            </div>
          </div>
          <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700">
            <span class="text-xs font-semibold text-slate-300 block mb-2"><i class="fa-solid fa-chart-line text-emerald-400 mr-1.5"></i>반기별 진동 추세 (Trend)</span>
            <div class="h-44 w-full">
              <canvas id="detailTrendChart"></canvas>
            </div>
          </div>
        </div>

        <div class="space-y-2 text-xs">
          <div>
            <h4 class="font-semibold text-slate-300 mb-1"><i class="fa-solid fa-stethoscope text-cyan-400 mr-1.5"></i>현장 진단 요약</h4>
            <p class="p-2 rounded bg-slate-800 text-slate-300 border border-slate-700/80">${item.summary}</p>
          </div>
          <div>
            <h4 class="font-semibold text-slate-300 mb-1"><i class="fa-solid fa-screwdriver-wrench text-amber-400 mr-1.5"></i>후속 정비 대책</h4>
            <p class="p-2 rounded bg-slate-800 text-slate-300 border border-slate-700/80">${item.action}</p>
          </div>
        </div>

        <div class="pt-4 mt-4 border-t border-slate-700 flex justify-end">
          <button onclick="closeDetailModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg font-medium transition text-xs">닫기</button>
        </div>
      `;

      openDetailModal();

      setTimeout(() => {
        const ctxAxes = document.getElementById('detailAxesChart').getContext('2d');
        if (detailAxesChart) detailAxesChart.destroy();
        detailAxesChart = new Chart(ctxAxes, {
          type: 'bar',
          data: {
            labels: ['수평 H(X)', '수직 V(Y)', '축방향 A(Z)'],
            datasets: [
              {
                label: '전동기',
                data: [
                  Math.max(points.m_de.h, points.m_nde.h),
                  Math.max(points.m_de.v, points.m_nde.v),
                  Math.max(points.m_de.a, points.m_nde.a)
                ],
                backgroundColor: 'rgba(245, 158, 11, 0.85)',
                borderRadius: 4
              },
              {
                label: '펌프',
                data: [
                  Math.max(points.p_de.h, points.p_nde.h),
                  Math.max(points.p_de.v, points.p_nde.v),
                  Math.max(points.p_de.a, points.p_nde.a)
                ],
                backgroundColor: 'rgba(6, 182, 212, 0.85)',
                borderRadius: 4
              }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#cbd5e1', font: { size: 10 } } } },
            scales: {
              x: { ticks: { color: '#94a3b8', font: { size: 10 } } },
              y: { ticks: { color: '#94a3b8', font: { size: 10 } }, beginAtZero: true }
            }
          }
        });

        const ctxTrend = document.getElementById('detailTrendChart').getContext('2d');
        if (detailChart) detailChart.destroy();
        const trendVals = periodList.map(p => (item.history && item.history[p] !== undefined) ? item.history[p] : (item.period === p ? item.maxRms : null));

        detailChart = new Chart(ctxTrend, {
          type: 'line',
          data: {
            labels: periodList,
            datasets: [{
              label: '최대 RMS (mm/s)',
              data: trendVals,
              borderColor: '#10b981',
              backgroundColor: '#10b981',
              borderWidth: 2,
              pointRadius: 4,
              spanGaps: true
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
              x: { ticks: { color: '#94a3b8', font: { size: 10 } } },
              y: { ticks: { color: '#94a3b8', font: { size: 10 } }, beginAtZero: true }
            }
          }
        });
      }, 50);
    }

    function updateSliderVal(val) {
      document.getElementById('calcValDisplay').innerText = parseFloat(val).toFixed(2) + ' mm/s';
      runEvaluation();
    }

    function runEvaluation() {
      const group = document.getElementById('calcGroup').value;
      const foundation = document.querySelector('input[name="calcFoundation"]:checked').value;
      const rms = parseFloat(document.getElementById('calcVibSlider').value);

      const fakePower = group === 'G1' ? 350 : (group === 'G4' ? 10 : 75);
      const res = calculateIsoGrade(fakePower, rms, foundation);

      const badge = document.getElementById('calcGradeBadge');
      const title = document.getElementById('calcGradeTitle');
      const box = document.getElementById('calcResultBox');
      const rec = document.getElementById('calcRecommendation');
      const bound = document.getElementById('calcBoundaryText');

      badge.className = 'text-lg font-black px-2.5 py-0.5 rounded-lg';

      if (res.grade === 'A') {
        badge.classList.add('grade-badge-A');
        badge.innerText = 'Grade A';
        title.className = 'text-sm font-bold text-emerald-400';
        title.innerText = '양호 (최상 신품 상태)';
        box.className = 'mt-4 p-4 rounded-xl border transition-all duration-300 bg-emerald-950/20 border-emerald-500/40';
        rec.innerText = '진동이 매우 안정적이며 제한 없이 장기 연속 운전이 가능합니다.';
      } else if (res.grade === 'B') {
        badge.classList.add('grade-badge-B');
        badge.innerText = 'Grade B';
        title.className = 'text-sm font-bold text-cyan-400';
        title.innerText = '양호 (정상 연속 가동)';
        box.className = 'mt-4 p-4 rounded-xl border transition-all duration-300 bg-cyan-950/20 border-cyan-500/40';
        rec.innerText = '기계가 정상 운전 허용 범위 내에 있습니다. 정기 점검 주기를 유지하십시오.';
      } else if (res.grade === 'C') {
        badge.classList.add('grade-badge-C');
        badge.innerText = 'Grade C';
        title.className = 'text-sm font-bold text-amber-400';
        title.innerText = '보수 필요 (주의 단계)';
        box.className = 'mt-4 p-4 rounded-xl border transition-all duration-300 bg-amber-950/20 border-amber-500/40';
        rec.innerText = '단기간 운전은 가능하나 결함 진행 중입니다. 주파수 분석(FFT)을 통한 원인 규명 및 차기 정비 시 부품 교체를 권고합니다.';
      } else {
        badge.classList.add('grade-badge-D');
        badge.innerText = 'Grade D';
        title.className = 'text-sm font-bold text-red-400';
        title.innerText = '즉시 점검 (운전 중단 권고)';
        box.className = 'mt-4 p-4 rounded-xl border transition-all duration-300 bg-red-950/30 border-red-500/50';
        rec.innerText = '설비 손상 및 소착, 돌발 정지 위험 수준입니다. 즉시 기기를 정지하고 분해정비를 실시하십시오.';
      }

      if (group === 'G1') {
        bound.innerText = foundation === 'rigid' ? 'A/B: ≤4.5 | C: ≤7.1 | D: >7.1' : 'A/B: ≤7.1 | C: ≤11.0 | D: >11.0';
      } else {
        bound.innerText = foundation === 'rigid' ? 'A/B: ≤2.8 | C: ≤4.5 | D: >4.5' : 'A/B: ≤4.5 | C: ≤7.1 | D: >7.1';
      }
    }

    function handleFileSelect(e) {
      const file = e.target.files[0];
      if (!file) return;
      processDataFile(file);
    }

    function processDataFile(file) {
      const statusBox = document.getElementById('uploadStatusBox');
      const targetPeriod = document.getElementById('uploadPeriodSelect').value;
      statusBox.classList.remove('hidden');
      statusBox.innerHTML = `<span class="text-cyan-400 font-semibold"><i class="fa-solid fa-spinner fa-spin mr-1"></i>스마트 AI 파싱 중: ${file.name}</span>`;

      const reader = new FileReader();

      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          
          let targetSheetName = workbook.SheetNames[0];
          for (let name of workbook.SheetNames) {
            if (name.includes('진동') || name.includes('측정')) {
              targetSheetName = name;
              break;
            }
          }

          const worksheet = workbook.Sheets[targetSheetName];
          // 병합셀(세로형) 처리를 위해 raw 데이터로 읽음
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

          let parsedList = [];
          let currentEquip = null;
          let currentOffice = '에너지사업소'; // 기본값 (상황에 따라 동적할당 가능)

          for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length < 5) continue; // 데이터가 없는 행 스킵

            let dateStr = String(row[0]).trim();
            let equipName = String(row[1]).trim();
            
            // 헤더 행 무시
            if (equipName === '설비명' || String(row[3]).includes('측정위치')) continue;

            // 새 설비명이 등장하면 (병합셀 시작점)
            if (equipName !== '') {
              if (currentEquip) parsedList.push(currentEquip); // 이전 설비 저장
              
              currentEquip = {
                name: equipName,
                date: dateStr || targetPeriod,
                power: parseFloat(row[2]) || 55, // 전동기(kW)
                points: {
                  m_de: { h: 0, v: 0, a: 0 }, m_nde: { h: 0, v: 0, a: 0 },
                  p_de: { h: 0, v: 0, a: 0 }, p_nde: { h: 0, v: 0, a: 0 }
                },
                summary: '',
                action: ''
              };
            }

            if (!currentEquip) continue;

            // 세로형 데이터 매핑 로직 (측정위치, 축, rms, 이상원인, 우선순위)
            let pointStr = String(row[3]).trim().replace(/\s/g, ''); // 예: "1(전동기부하측)"
            let axisStr = String(row[4]).toLowerCase().trim(); // 예: "x"
            let rmsVal = parseFloat(row[5]);
            let causeStr = String(row[8] || '').trim();
            let actionStr = String(row[9] || '').trim();

            if (!isNaN(rmsVal) && rmsVal > 0) { // 빈칸이나 0은 무시하고 유효한 수치만 처리
              let targetPoint = null;
              if (pointStr.includes('전동기') && pointStr.includes('반부하')) targetPoint = currentEquip.points.m_nde;
              else if (pointStr.includes('전동기')) targetPoint = currentEquip.points.m_de;
              else if ((pointStr.includes('펌프') || pointStr.includes('송풍기')) && pointStr.includes('반부하')) targetPoint = currentEquip.points.p_nde;
              else if (pointStr.includes('펌프') || pointStr.includes('송풍기')) targetPoint = currentEquip.points.p_de;

              if (targetPoint) {
                if (axisStr === 'x' || axisStr === 'h') targetPoint.h = Math.max(targetPoint.h, rmsVal);
                if (axisStr === 'y' || axisStr === 'v') targetPoint.v = Math.max(targetPoint.v, rmsVal);
                if (axisStr === 'z' || axisStr === 'a') targetPoint.a = Math.max(targetPoint.a, rmsVal);
              }
            }

            // 이상 원인 및 조치 내용 누적 (중복 방지)
            if (causeStr && causeStr !== '-' && !currentEquip.summary.includes(causeStr)) {
              currentEquip.summary += (currentEquip.summary ? ', ' : '') + causeStr;
            }
            if (actionStr && actionStr !== '양호' && !currentEquip.action.includes(actionStr)) {
              currentEquip.action += (currentEquip.action ? ', ' : '') + actionStr;
            }
          }
          
          if (currentEquip) parsedList.push(currentEquip); // 마지막 설비 저장

          if (parsedList.length === 0) {
            statusBox.className = 'p-3 rounded-lg border text-xs bg-amber-950/20 border-amber-500/40 text-amber-400';
            statusBox.innerHTML = `<i class="fa-solid fa-triangle-exclamation mr-1"></i>읽어들일 데이터가 없습니다.`;
            return;
          }

          // 추출된 설비들을 시스템에 등록
          parsedList.forEach((eq, idx) => {
            let p = eq.points;
            let maxRms = Math.max(p.m_de.h, p.m_de.v, p.m_de.a, p.m_nde.h, p.m_nde.v, p.m_nde.a, p.p_de.h, p.p_de.v, p.p_de.a, p.p_nde.h, p.p_nde.v, p.p_nde.a);
            if (maxRms === 0) maxRms = 1.0; // 값이 아예 없으면 기본값 처리

            const evalRes = calculateIsoGrade(eq.power, maxRms, 'rigid');
            
            let historyObj = {};
            historyObj[targetPeriod] = maxRms;

            equipmentData.unshift({
              id: 'smart_' + Date.now() + '_' + idx,
              period: targetPeriod,
              office: currentOffice, // 필요 시 엑셀 내에서 파싱 가능
              name: eq.name,
              power: eq.power,
              group: evalRes.group,
              foundation: 'rigid',
              totalQty: 1,
              overallGrade: evalRes.grade,
              maxRms: maxRms,
              points: eq.points,
              axes: {
                motor: { h: Math.max(p.m_de.h, p.m_nde.h), v: Math.max(p.m_de.v, p.m_nde.v), a: Math.max(p.m_de.a, p.m_nde.a) },
                pump:  { h: Math.max(p.p_de.h, p.p_nde.h), v: Math.max(p.p_de.v, p.p_nde.v), a: Math.max(p.p_de.a, p.p_nde.a) }
              },
              history: historyObj,
              summary: eq.summary || '측정 완료 (특이사항 없음)',
              action: eq.action || '추세 관리',
              detailNotes: '현장 세로형 다중행 엑셀 자동 분석(AI 파싱)으로 등록됨'
            });
          });

          if (!periodList.includes(targetPeriod)) periodList.push(targetPeriod);
          currentPeriod = targetPeriod;
          saveToLocalStorage();
          renderPeriodSelectors();

          statusBox.className = 'p-3 rounded-lg border text-xs bg-emerald-950/20 border-emerald-500/40 text-emerald-400 font-medium';
          statusBox.innerHTML = `<i class="fa-solid fa-circle-check mr-1"></i>현장 양식 완벽 해독! 총 <strong>${parsedList.length}건</strong>의 설비(12축)가 정상 등록되었습니다!`;

          updateDashboard();
          setTimeout(() => { closeUploadModal(); }, 2500);

        } catch (err) {
          console.error(err);
          statusBox.className = 'p-3 rounded-lg border text-xs bg-red-950/20 border-red-500/40 text-red-400';
          statusBox.innerHTML = `<i class="fa-solid fa-triangle-exclamation mr-1"></i>파일 해독 중 오류가 발생했습니다.`;
        }
      };

      reader.readAsArrayBuffer(file);
    }

    function downloadSampleTemplate() {
      const templateData = [
        [
          "사업소", "설비명", "용량kW",
          "전동기_부하측_H(X)", "전동기_부하측_V(Y)", "전동기_부하측_A(Z)",
          "전동기_반부하측_H(X)", "전동기_반부하측_V(Y)", "전동기_반부하측_A(Z)",
          "펌프_부하측_H(X)", "펌프_부하측_V(Y)", "펌프_부하측_A(Z)",
          "펌프_반부하측_H(X)", "펌프_반부하측_V(Y)", "펌프_반부하측_A(Z)",
          "진단요약", "조치계획", "측정시기"
        ],
        [
          "에너지사업소", "강제송풍기 #1", 310,
          6.8, 4.1, 3.5,
          3.2, 2.1, 1.8,
          7.3, 5.2, 4.8,
          4.5, 3.1, 2.6,
          "송풍기 부하측 진동 심각", "베어링 분해정비 요망", "2026년 상반기"
        ],
        [
          "에너지사업소", "보일러 급수펌프 #2", 95,
          3.4, 2.1, 1.8,
          1.8, 1.2, 1.0,
          3.9, 2.8, 2.2,
          2.5, 1.9, 1.5,
          "펌프 부하측 베어링 진동", "하절기 오버홀 시 교체", "2026년 상반기"
        ],
        [
          "에너지사업소", "중온수 순환펌프 #4", 350,
          3.2, 2.0, 1.9,
          1.9, 1.3, 1.1,
          2.6, 1.8, 1.7,
          2.1, 1.5, 1.2,
          "전동기 부하측 미세 이음", "추세 감시 유지", "2026년 상반기"
        ]
      ];

      const ws = XLSX.utils.aoa_to_sheet(templateData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "진동측정결과");
      XLSX.writeFile(wb, "회전설비_진동측정결과_표준양식.xlsx");
    }

    function exportToCSV() {
      const activeData = equipmentData.filter(e => {
        const officeMatch = (currentOffice === 'ALL' || e.office === currentOffice);
        const periodMatch = (!e.period || e.period === currentPeriod);
        return officeMatch && periodMatch;
      });

      const rows = [
        [
          "사업소", "측정시기", "설비명", "용량(kW)", "기계그룹", "최대RMS(mm/s)", "종합등급",
          "M-DE(H)", "M-DE(V)", "M-DE(A)", "M-NDE(H)", "M-NDE(V)", "M-NDE(A)",
          "P-DE(H)", "P-DE(V)", "P-DE(A)", "P-NDE(H)", "P-NDE(V)", "P-NDE(A)",
          "진단요약", "후속대책"
        ]
      ];

      activeData.forEach(e => {
        const pts = e.points || {
          m_de: {h:'',v:'',a:''}, m_nde: {h:'',v:'',a:''},
          p_de: {h:'',v:'',a:''}, p_nde: {h:'',v:'',a:''}
        };

        rows.push([
          e.office, e.period || currentPeriod, e.name, e.power, e.group, e.maxRms, e.overallGrade,
          pts.m_de.h, pts.m_de.v, pts.m_de.a,
          pts.m_nde.h, pts.m_nde.v, pts.m_nde.a,
          pts.p_de.h, pts.p_de.v, pts.p_de.a,
          pts.p_nde.h, pts.p_nde.v, pts.p_nde.a,
          e.summary, e.action
        ]);
      });

      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "진동측정결과");
      XLSX.writeFile(wb, `회전설비_진동진단결과_${currentOffice}_${currentPeriod}.xlsx`);
    }

    const dropZone = document.getElementById('dropZone');
    ['dragenter', 'dragover'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropZone.classList.add('border-cyan-400', 'bg-slate-800');
      }, false);
    });
    ['dragleave', 'drop'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropZone.classList.remove('border-cyan-400', 'bg-slate-800');
      }, false);
    });
    dropZone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files.length) processDataFile(files[0]);
    });

    window.onload = function() {
      loadFromLocalStorage();
      renderPeriodSelectors();
      renderOfficeTabs();
      updateDashboard();
      updateSyncBadge();

      if (googleSheetAppUrl) {
        syncFromGoogleSheets();
      }
    };
  </script>
</body>
</html>
"""

# Streamlit Components를 통한 렌더링 (height: 1050px, 스크롤 가능)
components.html(dashboard_html, height=1050, scrolling=True)