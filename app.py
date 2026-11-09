import streamlit as st
import streamlit.components.v1 as components

# Streamlit 기본 레이아웃을 전체 너비(wide)로 설정
st.set_page_config(
    page_title="부산환경공단 회전설비 진동진단 통합관리 시스템 (BECO-VIB)",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Streamlit 기본 여백 제거
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container {
        padding: 0rem !important;
        max-width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)

# BECO-VIB 시스템 전체 HTML/CSS/JS 코드
DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>부산환경공단 사업소별 회전설비 진동진단 통합관리 시스템 (BECO-VIB)</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Chart.js CDN -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <!-- SheetJS (xlsx) CDN for Excel & CSV File Parsing -->
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
            </h1>
          </div>
          <p class="text-xs text-slate-400">부산환경공단 에너지사업소 회전체 상태진단 및 예지보전 플랫폼</p>
        </div>
      </div>

      <!-- Right Header Actions -->
      <div class="flex items-center gap-3 w-full sm:w-auto justify-end">
        <div class="text-right hidden md:block">
          <p class="text-xs text-slate-400">데이터 규격: ISO 20816-3 (신뢰도 평가)</p>
          <p class="text-xs font-mono text-cyan-400"><i class="fa-solid fa-server mr-1"></i>다사업소 통합 DB</p>
        </div>
        <!-- Add Office Button -->
        <button onclick="openAddOfficeModal()" class="flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-500 text-white text-xs sm:text-sm font-medium py-2 px-3 rounded-lg transition duration-200 shadow">
          <i class="fa-solid fa-building-circle-plus"></i>
          <span>사업소 추가</span>
        </button>
        <!-- Add Period Button -->
        <button onclick="openAddPeriodModal()" class="flex items-center gap-1.5 bg-purple-600 hover:bg-purple-500 text-white text-xs sm:text-sm font-medium py-2 px-3 rounded-lg transition duration-200 shadow">
          <i class="fa-solid fa-calendar-plus"></i>
          <span>측정주기 추가</span>
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

    <!-- Office / Plant Switcher Bar & Period Selector -->
    <div class="max-w-7xl mx-auto mt-3 pt-3 border-t border-slate-700/60 flex items-center justify-between flex-wrap gap-2">
      <div class="flex items-center gap-2 overflow-x-auto pb-1 sm:pb-0" id="officeTabs">
        <!-- Rendered dynamically by renderOfficeTabs() -->
      </div>

      <div class="flex items-center gap-3 text-xs text-slate-400">
        <!-- Period Selector -->
        <div class="flex items-center gap-1.5 bg-slate-800 px-2.5 py-1 rounded-lg border border-slate-700">
          <i class="fa-regular fa-calendar-check text-cyan-400"></i>
          <span class="text-slate-300 font-medium">측정 시기:</span>
          <select id="periodFilterDropdown" onchange="switchPeriod(this.value)" class="bg-slate-900 border border-slate-700 text-cyan-300 text-xs font-bold rounded px-2 py-0.5 focus:outline-none">
            <!-- Dynamically Rendered -->
          </select>
        </div>

        <span id="currentOfficeLabel" class="font-bold text-white bg-slate-800 px-2.5 py-1 rounded border border-slate-700">전체 사업소 현황</span>
        <button onclick="exportToCSV()" class="hover:text-cyan-400 transition ml-1" title="현재 조회 데이터 CSV 내보내기">
          <i class="fa-solid fa-file-export mr-1"></i>내보내기
        </button>
      </div>
    </div>
  </header>

  <main class="max-w-7xl mx-auto px-4 sm:px-6 py-6 space-y-6">

    <!-- KPI Metric Cards Grid -->
    <section class="grid grid-cols-2 lg:grid-cols-5 gap-3.5 sm:gap-4">
      <!-- Total Equipment -->
      <div class="glass-card rounded-xl p-4 border border-slate-700/50 hover:border-slate-600 transition">
        <div class="flex items-center justify-between text-slate-400 text-xs mb-1">
          <span>진단 대상 설비</span>
          <i class="fa-solid fa-cubes text-slate-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-total" class="text-2xl sm:text-3xl font-extrabold text-white">28</span>
          <span class="text-xs text-slate-400">대</span>
        </div>
        <p class="text-[11px] text-slate-400 mt-2 truncate">동력 7.5kW ~ 350kW</p>
      </div>

      <!-- Grade A/B Normal -->
      <div class="glass-card rounded-xl p-4 border border-emerald-500/20 bg-emerald-950/10 hover:border-emerald-500/40 transition">
        <div class="flex items-center justify-between text-emerald-400 text-xs mb-1">
          <span>양호 상태 (A·B)</span>
          <i class="fa-solid fa-circle-check text-emerald-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-ab" class="text-2xl sm:text-3xl font-extrabold text-emerald-400">11</span>
          <span class="text-xs text-emerald-300/70">대</span>
        </div>
        <p class="text-[11px] text-emerald-400/80 mt-2 flex items-center gap-1 font-medium">
          <span id="metric-ab-pct">39.3%</span> 정상 가동 유지
        </p>
      </div>

      <!-- Grade C Warning -->
      <div class="glass-card rounded-xl p-4 border border-amber-500/20 bg-amber-950/10 hover:border-amber-500/40 transition">
        <div class="flex items-center justify-between text-amber-400 text-xs mb-1">
          <span>보수필요 (C등급)</span>
          <i class="fa-solid fa-triangle-exclamation text-amber-400"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-c" class="text-2xl sm:text-3xl font-extrabold text-amber-400">7</span>
          <span class="text-xs text-amber-300/70">대</span>
        </div>
        <p class="text-[11px] text-amber-400/80 mt-2 flex items-center gap-1">
          <i class="fa-solid fa-arrow-trend-up text-[10px]"></i> 추세관리 및 부분보수
        </p>
      </div>

      <!-- Grade D Immediate -->
      <div class="glass-card rounded-xl p-4 border border-red-500/30 bg-red-950/20 hover:border-red-500/50 transition">
        <div class="flex items-center justify-between text-red-400 text-xs mb-1">
          <span>즉시점검 (D등급)</span>
          <i class="fa-solid fa-circle-exclamation text-red-400 animate-pulse"></i>
        </div>
        <div class="flex items-baseline gap-2">
          <span id="metric-d" class="text-2xl sm:text-3xl font-extrabold text-red-400">10</span>
          <span class="text-xs text-red-300/70">대</span>
        </div>
        <p class="text-[11px] text-red-300 mt-2 font-medium">
          긴급 점검 및 분해정비
        </p>
      </div>

      <!-- Unoperated / Testing -->
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

    <!-- ISO Criteria & Visual Chart Panel -->
    <section class="grid grid-cols-1 lg:grid-cols-3 gap-6">
      
      <!-- Chart Column (2 cols) with Mode Tabs (Current Bar vs Historical Trend Line) -->
      <div class="lg:col-span-2 glass-card rounded-xl p-5 border border-slate-700/60 flex flex-col justify-between">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-4 border-b border-slate-700/60 gap-3">
          <div>
            <div class="flex items-center gap-2">
              <h2 class="text-base font-semibold text-white flex items-center gap-2" id="chartMainTitle">
                <i class="fa-solid fa-chart-line text-cyan-400"></i>
                설비 진동치 상태 및 시계열 추이 분석
              </h2>
            </div>
            <p class="text-xs text-slate-400" id="chartSubTitle">측정 주기별 진동 추세를 선 그래프로 관찰하여 이상 징후 조기 포착</p>
          </div>
          
          <!-- View Toggle Buttons -->
          <div class="flex items-center gap-2 bg-slate-900/90 p-1 rounded-lg border border-slate-700 text-xs">
            <button id="btnViewTrend" onclick="setChartViewMode('trend')" class="px-2.5 py-1 rounded-md font-bold bg-cyan-600 text-white transition flex items-center gap-1.5 shadow">
              <i class="fa-solid fa-arrow-trend-up text-[11px]"></i>
              <span>반기별 추이선 (Trend)</span>
            </button>
            <button id="btnViewBar" onclick="setChartViewMode('bar')" class="px-2.5 py-1 rounded-md text-slate-400 hover:text-white transition flex items-center gap-1.5">
              <i class="fa-solid fa-chart-simple text-[11px]"></i>
              <span>현재 주기 비교 (Bar)</span>
            </button>
          </div>
        </div>

        <!-- Trend Controls (Only visible in Trend mode) -->
        <div id="trendFilterBar" class="py-2.5 flex items-center justify-between flex-wrap gap-2 text-xs border-b border-slate-800">
          <div class="flex items-center gap-2">
            <span class="text-slate-400 font-medium"><i class="fa-solid fa-filter text-cyan-400 mr-1"></i>추이선 대상 설비:</span>
            <select id="trendEquipSelect" onchange="renderChart()" class="bg-slate-800 border border-slate-700 text-slate-200 rounded px-2.5 py-1 focus:outline-none focus:ring-1 focus:ring-cyan-500 cursor-pointer">
              <option value="ALL">상위 주요 위험/주의 설비 전체 비교 (Top 5)</option>
              <!-- Filled dynamically with equipment list -->
            </select>
          </div>
          <div class="flex items-center gap-3 text-[11px] text-slate-400">
            <span class="flex items-center gap-1"><span class="w-3 h-0.5 bg-red-400 border border-dashed border-red-500 inline-block"></span> D(즉시점검 4.5/7.1)</span>
            <span class="flex items-center gap-1"><span class="w-3 h-0.5 bg-amber-400 border border-dashed border-amber-500 inline-block"></span> C(보수필요 2.8/4.5)</span>
          </div>
        </div>

        <div class="pt-4 h-[290px] w-full">
          <canvas id="vibrationChart"></canvas>
        </div>
      </div>

      <!-- ISO Standard Quick Guide (1 col) -->
      <div class="glass-card rounded-xl p-5 border border-slate-700/60 flex flex-col justify-between">
        <div>
          <div class="flex items-center justify-between pb-3 border-b border-slate-700/60">
            <h3 class="text-sm font-semibold text-white flex items-center gap-2">
              <i class="fa-solid fa-book-bookmark text-amber-400"></i>
              ISO 20816-3 진동 판정 기준표
            </h3>
            <span class="text-[10px] text-slate-400">진동속도 RMS(mm/s)</span>
          </div>
          
          <div class="mt-4 space-y-3 text-xs">
            <!-- Group 1 -->
            <div class="p-2.5 rounded-lg bg-slate-800/80 border border-slate-700">
              <div class="flex justify-between items-center font-bold text-slate-200">
                <span class="text-cyan-300">Group 1 (&gt; 300 kW)</span>
                <span class="text-[11px] font-normal text-slate-400">중온수펌프, 강제송풍기 등</span>
              </div>
              <div class="grid grid-cols-3 gap-1 mt-2 text-center text-[11px]">
                <div class="p-1 rounded grade-badge-A font-medium">A/B &lt; 4.5</div>
                <div class="p-1 rounded grade-badge-C font-medium">C &gt; 4.5 (주의)</div>
                <div class="p-1 rounded grade-badge-D font-medium">D &gt; 7.1 (중단)</div>
              </div>
            </div>

            <!-- Group 2 -->
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

            <!-- Group 4 -->
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
          <span>* 측정원칙: 수평(H), 수직(V), 축방향(A)</span>
          <span class="text-cyan-400 font-medium">에너지사업소 기술표준</span>
        </div>
      </div>
    </section>

    <!-- Main Table Section -->
    <section class="glass-card rounded-xl border border-slate-700/60 overflow-hidden">
      <!-- Table Header & Controls -->
      <div class="p-4 sm:p-5 border-b border-slate-700/60 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 class="text-base font-bold text-white flex items-center gap-2">
            <i class="fa-solid fa-list-check text-cyan-400"></i>
            설비별 세부 진단 결과 및 정비 계획
          </h2>
          <div class="flex items-center gap-2 mt-0.5">
            <p id="tableSubtitle" class="text-xs text-slate-400">전체 사업소 회전기계 상태진단 및 조치 대장</p>
            <span class="text-slate-600">|</span>
            <span id="filteredCount" class="text-xs font-mono text-cyan-400 font-semibold">총 0개 설비</span>
          </div>
        </div>

        <!-- Filter Controls -->
        <div class="flex flex-wrap items-center gap-2">
          <!-- Office Secondary Filter -->
          <div class="relative">
            <select id="officeFilterDropdown" onchange="syncOfficeFilter(this.value)" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2 focus:ring-2 focus:ring-cyan-500 focus:outline-none pr-8 cursor-pointer">
              <option value="ALL">전체 사업소</option>
              <option value="에너지사업소">에너지사업소</option>
              <option value="명지사업소">명지사업소</option>
              <option value="해운대사업소">해운대사업소</option>
            </select>
          </div>

          <!-- Group Filter -->
          <div class="relative">
            <select id="groupFilter" onchange="renderTable()" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2 focus:ring-2 focus:ring-cyan-500 focus:outline-none pr-8 cursor-pointer">
              <option value="ALL">전체 용량 구분</option>
              <option value="G1">Group 1 (>300kW)</option>
              <option value="G2">Group 2 (15~300kW)</option>
              <option value="G4">Group 4 (<15kW)</option>
            </select>
          </div>

          <!-- Status Filter -->
          <div class="relative">
            <select id="statusFilter" onchange="renderTable()" class="bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg px-3 py-2 focus:ring-2 focus:ring-cyan-500 focus:outline-none pr-8 cursor-pointer">
              <option value="ALL">전체 평가 등급</option>
              <option value="D">D (즉시점검)</option>
              <option value="C">C (보수필요)</option>
              <option value="AB">A/B (양호)</option>
            </select>
          </div>

          <!-- Search Input -->
          <div class="relative w-full sm:w-44">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
            <input type="text" id="searchInput" oninput="renderTable()" placeholder="설비명/내용 검색..." class="w-full bg-slate-800 border border-slate-600 text-slate-200 text-xs rounded-lg pl-8 pr-3 py-2 focus:ring-2 focus:ring-cyan-500 focus:outline-none">
          </div>
        </div>
      </div>

      <!-- Table Container -->
      <div class="overflow-x-auto custom-scrollbar">
        <table class="w-full text-left border-collapse text-xs">
          <thead>
            <tr class="bg-slate-800/80 text-slate-400 border-b border-slate-700 text-[11px] uppercase tracking-wider font-semibold">
              <th class="py-3 px-3">측정시기</th>
              <th class="py-3 px-3">사업소</th>
              <th class="py-3 px-4">설비명</th>
              <th class="py-3 px-3">정격용량</th>
              <th class="py-3 px-3">기계구분</th>
              <th class="py-3 px-3 text-center">측정 RMS</th>
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

  <!-- MODAL 1: Add Custom Equipment with Period & Realtime ISO Evaluation -->
  <div id="addModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-xl rounded-2xl shadow-2xl p-6 relative max-h-[92vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeAddModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-5 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-emerald-500/20 text-emerald-400 border border-emerald-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-plus"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">신규 설비 직접 등록</h3>
          <p class="text-xs text-slate-400">측정 시기 및 용량(kW) 입력 시 ISO 20816-3 등급이 자동 산출됩니다.</p>
        </div>
      </div>

      <form id="addEquipForm" onsubmit="handleEquipmentSubmit(event)" class="space-y-4 text-xs">
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <!-- 측정 시기 선택 -->
          <div>
            <label class="block text-slate-300 font-medium mb-1">측정 시기 <span class="text-rose-400">*</span></label>
            <select id="inpPeriod" required class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-cyan-300 font-bold focus:outline-none focus:ring-2 focus:ring-cyan-500">
              <!-- Dynamically Rendered -->
            </select>
          </div>

          <!-- 사업소 선택 -->
          <div>
            <label class="block text-slate-300 font-medium mb-1">소속 사업소 <span class="text-rose-400">*</span></label>
            <select id="inpOffice" required class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
              <option value="에너지사업소">에너지사업소</option>
              <option value="명지사업소">명지사업소</option>
              <option value="해운대사업소">해운대사업소</option>
              <option value="남부사업소">남부사업소</option>
            </select>
          </div>

          <!-- 설비명 -->
          <div>
            <label class="block text-slate-300 font-medium mb-1">설비명 <span class="text-rose-400">*</span></label>
            <input type="text" id="inpName" required placeholder="예: #3 응축수 회수펌프" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <!-- 용량(kW) -->
          <div>
            <label class="block text-slate-300 font-medium mb-1">정격 용량 (kW) <span class="text-rose-400">*</span></label>
            <input type="number" id="inpPower" required step="0.1" min="0.1" value="75" oninput="autoCalcGrade()" placeholder="kW 단위" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500 font-mono">
          </div>

          <!-- 기초 조건 -->
          <div>
            <label class="block text-slate-300 font-medium mb-1">기초 조건 (Foundation)</label>
            <select id="inpFoundation" onchange="autoCalcGrade()" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
              <option value="rigid" selected>강성 기초 (Rigid)</option>
              <option value="flexible">탄성 기초 (Flexible)</option>
            </select>
          </div>

          <!-- 측정 속도 RMS -->
          <div>
            <label class="block text-slate-300 font-medium mb-1">진동 실측치 RMS (mm/s) <span class="text-rose-400">*</span></label>
            <input type="number" id="inpRms" required step="0.01" min="0" value="3.10" oninput="autoCalcGrade()" placeholder="mm/s" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500 font-mono font-bold text-cyan-400">
          </div>
        </div>

        <!-- 실시간 자동 계산 미리보기 배너 -->
        <div id="autoEvalBanner" class="p-3.5 rounded-xl border bg-slate-800/80 border-slate-700 flex items-center justify-between">
          <div>
            <span class="text-[11px] text-slate-400 block font-mono" id="evalGroupText">ISO 분류: Group 2 (15~300kW)</span>
            <span class="text-xs font-semibold text-slate-200" id="evalCriteriaText">기준치: A/B ≤ 2.8 | C ≤ 4.5 | D > 4.5</span>
          </div>
          <div class="text-right">
            <span id="evalBadge" class="inline-block px-3 py-1 rounded text-xs font-black grade-badge-C">C등급 (보수필요)</span>
          </div>
        </div>

        <!-- 진단 요약 및 이상 부위 -->
        <div>
          <label class="block text-slate-300 font-medium mb-1">현장 진단 요약 및 이상 부위</label>
          <input type="text" id="inpSummary" placeholder="예: 전동기 구동측 베어링 진동 및 미세 발열 확인" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
        </div>

        <!-- 후속 정비 대책 -->
        <div>
          <label class="block text-slate-300 font-medium mb-1">후속 정비대책</label>
          <input type="text" id="inpAction" placeholder="예: 차기 계획정비 시 베어링 교체 및 축정렬 재실시" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
        </div>

        <!-- 상세 소견 -->
        <div>
          <label class="block text-slate-300 font-medium mb-1">엔지니어 상세 소견 (선택)</label>
          <textarea id="inpNotes" rows="2" placeholder="주파수(FFT) 성분, 회전수, 최근 보수 이력 등을 입력" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500"></textarea>
        </div>

        <div class="pt-3 border-t border-slate-700 flex justify-end gap-2">
          <button type="button" onclick="closeAddModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition">취소</button>
          <button type="submit" class="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-lg transition shadow-md">설비 등록 완료</button>
        </div>
      </form>
    </div>
  </div>

  <!-- MODAL: Add New Measurement Period -->
  <div id="addPeriodModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-md rounded-2xl shadow-2xl p-6 relative">
      <button onclick="closeAddPeriodModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-5 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-purple-500/20 text-purple-400 border border-purple-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-calendar-plus"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">신규 측정 주기 추가</h3>
          <p class="text-xs text-slate-400">데이터를 누적 관리할 신규 반기/분기 주기를 등록합니다.</p>
        </div>
      </div>

      <form id="addPeriodForm" onsubmit="handlePeriodSubmit(event)" class="space-y-4 text-xs">
        <div>
          <label class="block text-slate-300 font-medium mb-1">측정 주기 명칭 <span class="text-rose-400">*</span></label>
          <input type="text" id="inpPeriodName" required placeholder="예: 2026년 하반기, 2027년 상반기" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-purple-500">
        </div>
        <p class="text-[11px] text-slate-400 leading-relaxed">
          * 주기를 등록하면 상단 시기 선택 탭, 설비 등록창, 엑셀 템플릿에 자동 반영되어 반기별 시계열 데이터가 누적됩니다.
        </p>

        <div class="pt-3 border-t border-slate-700 flex justify-end gap-2">
          <button type="button" onclick="closeAddPeriodModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition">취소</button>
          <button type="submit" class="px-5 py-2 bg-purple-600 hover:bg-purple-500 text-white font-bold rounded-lg transition shadow-md">측정주기 생성</button>
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
          <h3 class="text-base font-bold text-white">진동 측정 데이터 업로드</h3>
          <p class="text-xs text-slate-400">엑셀(.xlsx, .xls) 및 CSV 파일을 자동 인식하여 일괄 등록합니다.</p>
        </div>
      </div>

      <div class="space-y-4 text-xs">
        <!-- Target Period Selector for Upload -->
        <div class="p-3 rounded-xl bg-slate-800/80 border border-slate-700 flex items-center justify-between">
          <span class="font-medium text-slate-300">업로드 대상 측정시기:</span>
          <select id="uploadPeriodSelect" class="bg-slate-900 border border-slate-600 rounded px-2.5 py-1 text-cyan-300 font-bold focus:outline-none">
            <!-- Dynamically populated -->
          </select>
        </div>

        <!-- Template Download Guide -->
        <div class="p-3 rounded-xl bg-slate-800/80 border border-slate-700 flex items-center justify-between">
          <div>
            <span class="font-bold text-slate-200 block">표준 업로드 양식 템플릿</span>
            <span class="text-slate-400 text-[11px]">열 구성: 사업소, 설비명, 용량kW, 실측RMS, 진단요약, 조치계획, 측정시기(선택)</span>
          </div>
          <button onclick="downloadSampleTemplate()" class="px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-cyan-300 rounded-lg font-medium border border-slate-600 transition flex items-center gap-1.5">
            <i class="fa-solid fa-download"></i>
            <span>양식 다운로드</span>
          </button>
        </div>

        <!-- Drag & Drop Area -->
        <div id="dropZone" onclick="document.getElementById('fileInput').click()" class="border-2 border-dashed border-slate-600 hover:border-cyan-500 rounded-2xl p-8 text-center cursor-pointer transition bg-slate-800/40 hover:bg-slate-800/70">
          <i class="fa-solid fa-cloud-arrow-up text-3xl text-cyan-400 mb-2"></i>
          <p class="text-slate-200 font-bold mb-1">엑셀 또는 CSV 파일을 여기로 드래그하세요</p>
          <p class="text-slate-400 text-[11px]">또는 클릭하여 내 PC에서 파일 선택 (.xlsx, .xls, .csv)</p>
          <input type="file" id="fileInput" accept=".xlsx, .xls, .csv" onchange="handleFileSelect(event)" class="hidden">
        </div>

        <!-- Upload Status & Preview -->
        <div id="uploadStatusBox" class="hidden p-3 rounded-lg border text-xs"></div>

        <div class="pt-3 border-t border-slate-700 flex justify-between items-center">
          <span class="text-[11px] text-slate-400">kW값에 따라 ISO 등급 자동 산출</span>
          <button onclick="closeUploadModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition">닫기</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Interactive ISO 20816-3 Calculator Modal -->
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
        <!-- Machine Group -->
        <div>
          <label class="block text-slate-300 font-medium mb-1.5">설비 정격 용량 (Group)</label>
          <select id="calcGroup" onchange="runEvaluation()" class="w-full bg-slate-800 border border-slate-700 rounded-lg p-2.5 text-slate-200 focus:outline-none focus:ring-2 focus:ring-cyan-500">
            <option value="G1">Group 1 : 대형 기계 (&gt; 300 kW, 중온수펌프·대형팬)</option>
            <option value="G2" selected>Group 2 : 중형 기계 (15 kW ~ 300 kW, 급수·소각순환)</option>
            <option value="G4">Group 4 : 소형 기계 (&lt; 15 kW, 일반 냉각수펌프)</option>
          </select>
        </div>

        <!-- Foundation Type -->
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

        <!-- Vibration RMS Input -->
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

        <!-- Diagnostic Result Box -->
        <div id="calcResultBox" class="mt-4 p-4 rounded-xl border transition-all duration-300">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <span id="calcGradeBadge" class="text-lg font-black px-2.5 py-0.5 rounded-lg grade-badge-C">Grade C</span>
              <span id="calcGradeTitle" class="text-sm font-bold text-amber-400">보수 필요 (주의 단계)</span>
            </div>
            <span id="calcBoundaryText" class="text-[11px] font-mono text-slate-400">2.8 &lt; RMS &le; 4.5</span>
          </div>
          <p id="calcRecommendation" class="mt-2 text-slate-300 text-[11px] leading-relaxed">
            단기간 운전은 가능하나 설비 손상이 우려되는 수준입니다. 주파수 분석(FFT)을 통한 원인 규명(축정렬, 베어링) 및 차기 계획정비 시 부품 교체를 권고합니다.
          </p>
        </div>
      </div>

      <div class="mt-5 pt-3 border-t border-slate-700 flex justify-end">
        <button onclick="closeCalcModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold rounded-lg transition">닫기</button>
      </div>
    </div>
  </div>

  <!-- MODAL 3: Add Custom Office (Branch) -->
  <div id="addOfficeModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-md rounded-2xl shadow-2xl p-6 relative">
      <button onclick="closeAddOfficeModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div class="flex items-center gap-3 mb-5 border-b border-slate-700 pb-3">
        <div class="h-10 w-10 rounded-xl bg-indigo-500/20 text-indigo-400 border border-indigo-500/40 flex items-center justify-center text-lg">
          <i class="fa-solid fa-building-circle-plus"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-white">신규 사업소 추가</h3>
          <p class="text-xs text-slate-400">추가된 사업소는 탭 및 설비 등록창에 즉시 반영됩니다.</p>
        </div>
      </div>

      <form id="addOfficeForm" onsubmit="handleOfficeSubmit(event)" class="space-y-4 text-xs">
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

        <div class="pt-3 border-t border-slate-700 flex justify-end gap-2">
          <button type="button" onclick="closeAddOfficeModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg font-medium transition">취소</button>
          <button type="submit" class="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-bold rounded-lg transition shadow-md">사업소 등록</button>
        </div>
      </form>
    </div>
  </div>

  <!-- Detail Equipment Modal with Historical Trend Line -->
  <div id="detailModal" class="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 hidden">
    <div class="glass-card bg-slate-900 border border-slate-700 w-full max-w-2xl rounded-2xl shadow-2xl p-6 relative max-h-[90vh] overflow-y-auto custom-scrollbar">
      <button onclick="closeDetailModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white transition">
        <i class="fa-solid fa-xmark text-lg"></i>
      </button>

      <div id="detailModalContent">
        <!-- Rendered by JS -->
      </div>
    </div>
  </div>

  <!-- Scripts -->
  <script>
    // Registered Measurement Periods (Ordered Chronologically)
    let periodList = [
      '2025년 상반기',
      '2025년 하반기',
      '2026년 상반기',
      '2026년 하반기'
    ];

    let currentPeriod = '2026년 상반기'; // Default Period
    let chartViewMode = 'trend'; // 'trend' (시계열 추이선) or 'bar' (현재 주기 막대)

    // Master Multi-Office Equipment Dataset with Historical Time-Series
    let equipmentData = [
      // 1. 에너지사업소
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
        history: { '2025년 상반기': 3.8, '2025년 하반기': 4.9, '2026년 상반기': 7.3, '2026년 하반기': null },
        summary: '#1 : 전동기 진동, #4 : 송풍기 진동 심각',
        action: '#1, #4 전동기 및 팬 베어링 분해정비 요망',
        detailNotes: '310kW 대용량 보일러 연소용 송풍기로 연속 운전 중. #1호기는 구동측 베어링 진동, #4호기는 임펠러 불평형(Unbalance) 의심.'
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
        history: { '2025년 상반기': 2.1, '2025년 하반기': 2.7, '2026년 상반기': 3.9, '2026년 하반기': null },
        summary: '#2 : 전동기/펌프 진동, #3·#4 : 펌프 베어링 진동',
        action: '#2, #3, #4 호기 하절기 정비 시 베어링 및 축정렬 보수',
        detailNotes: '#5, #6호기는 대기 상태. #2호기는 커플링 엘레먼트 마모 추정, #3·#4호기는 펌프 베어링 고주파 진동 발생.'
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
        history: { '2025년 상반기': 2.8, '2025년 하반기': 3.0, '2026년 상반기': 3.2, '2026년 하반기': null },
        summary: '#3 : 액츄에이터 제작 중, #4 : 전동기 진동 및 이음',
        action: '#3호기 밸브 제작 완료 후 교체, #4호기 향후 보수',
        detailNotes: '지역난방 열공급 주설비(350kW). #4호기는 전동기 구동측에서 미세 금속 이음 발생하여 정기 추세 관리 중.'
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
        history: { '2025년 상반기': 3.2, '2025년 하반기': 4.4, '2026년 상반기': 6.2, '2026년 하반기': null },
        summary: '#7, #8 : 전동기/펌프 진동 심각, #9 : M/Seal 미세누수',
        action: '#7, #8호기 즉시 분해 점검, #9호기 씰 교체',
        detailNotes: '#7, #8호기는 ISO Group 2 Rigid 기준 D등급(4.5mm/s 초과)으로 베어링 소착 위험. 예비 펌프 즉시 전환 요망.'
      },
      {
        id: 'cool_1st',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '1차 냉각수펌프',
        power: 11,
        group: 'G4',
        foundation: 'rigid',
        totalQty: 3,
        unoperated: '-',
        normalList: '-',
        warningList: '#1',
        dangerList: '#2, #3',
        overallGrade: 'D',
        maxRms: 4.8,
        history: { '2025년 상반기': 1.9, '2025년 하반기': 3.1, '2026년 상반기': 4.8, '2026년 하반기': null },
        summary: '#2 : 전동기 진동 하자보수 요청, #1·#3 : 베어링 진동',
        action: '#2호기 시공사 하자보수 재확인, #1·#3호기 부품 교체',
        detailNotes: '11kW 소형 설비군. #2호기는 시운전 중 이상진동으로 하자보수 접수 완료. #3호기는 펌프측 베어링 손상 징후.'
      },
      {
        id: 'cool_2nd',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '2차 냉각수펌프',
        power: 11,
        group: 'G4',
        foundation: 'rigid',
        totalQty: 3,
        unoperated: '-',
        normalList: '#4, #6',
        warningList: '#5',
        dangerList: '-',
        overallGrade: 'B',
        maxRms: 2.4,
        history: { '2025년 상반기': 2.1, '2025년 하반기': 2.2, '2026년 상반기': 2.4, '2026년 하반기': null },
        summary: '#5 : 펌프측 베어링 진동 미세 증가',
        action: '#5호기 향후 계획정비 시 베어링 교체',
        detailNotes: '#4, #6호기는 1.8mm/s 이하로 안정적. #5호기는 정상 상한선 부근으로 오일 보충 및 감시.'
      },
      {
        id: 'fluid_cool',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '유체 냉각수펌프',
        power: 7.5,
        group: 'G4',
        foundation: 'rigid',
        totalQty: 3,
        unoperated: '-',
        normalList: '#8, #9',
        warningList: '#7',
        dangerList: '-',
        overallGrade: 'B',
        maxRms: 2.6,
        history: { '2025년 상반기': 2.0, '2025년 하반기': 2.3, '2026년 상반기': 2.6, '2026년 하반기': null },
        summary: '#7 : 펌프측 진동으로 향후 보수 요망',
        action: '현재 운전 이상 없음, 분기별 정기점검 유지',
        detailNotes: '전반적인 상태 양호. #7호기는 베어링 하우징 수평방향 2.6mm/s로 기준치 이내.'
      },
      {
        id: 'water_treat',
        office: '에너지사업소',
        period: '2026년 상반기',
        name: '수처리펌프',
        power: 18.5,
        group: 'G2',
        foundation: 'rigid',
        totalQty: 2,
        unoperated: '-',
        normalList: '#1, #2',
        warningList: '-',
        dangerList: '-',
        overallGrade: 'A',
        maxRms: 1.3,
        history: { '2025년 상반기': 1.1, '2025년 하반기': 1.2, '2026년 상반기': 1.3, '2026년 하반기': null },
        summary: '#1, #2 전 설비 이상 없음 (최상)',
        action: '현 상태 유지 및 반기별 점검',
        detailNotes: '전동기 및 펌프 진동 속도 1.3mm/s 이하로 최상의 A등급 유지.'
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
        history: { '2025년 상반기': 3.5, '2025년 하반기': 5.2, '2026년 상반기': 7.4, '2026년 하반기': null },
        summary: '저온열 : 메카니칼 씰 누수, 고온열 : 베어링 진동 과다',
        action: '저온열 즉시보수 실시, 고온열 베어링 순차적 분해정비',
        detailNotes: '총 10대. 저온열은 온수 누출 수반되어 부식 위험 존재, 고온열은 축방향 진동치가 7.4mm/s로 D등급 임계치 초과.'
      },

      // 2. 명지사업소
      {
        id: 'mj_pump_1',
        office: '명지사업소',
        period: '2026년 상반기',
        name: '하수유입 1차 펌프 #1',
        power: 220,
        group: 'G2',
        foundation: 'rigid',
        totalQty: 1,
        unoperated: '-',
        normalList: '#1',
        warningList: '-',
        dangerList: '-',
        overallGrade: 'B',
        maxRms: 2.1,
        history: { '2025년 상반기': 1.8, '2025년 하반기': 2.0, '2026년 상반기': 2.1, '2026년 하반기': null },
        summary: '전반적인 진동 상태 양호 및 발열 없음',
        action: '분기별 윤활유 보충 및 추세 관리',
        detailNotes: '명지사업소 유입동 메인 펌프. 진동 2.1mm/s로 양호(B등급) 범위 내 안정 운전 중.'
      },
      {
        id: 'mj_blower_2',
        office: '명지사업소',
        period: '2026년 상반기',
        name: '생물반응조 송풍기 #2',
        power: 400,
        group: 'G1',
        foundation: 'flexible',
        totalQty: 1,
        unoperated: '-',
        normalList: '-',
        warningList: '#2',
        dangerList: '-',
        overallGrade: 'C',
        maxRms: 8.2,
        history: { '2025년 상반기': 4.5, '2025년 하반기': 6.8, '2026년 상반기': 8.2, '2026년 하반기': null },
        summary: '토출측 커플링 진동 및 미세 불평형 확인',
        action: '차기 가동 중지 시 축정렬 레이저 재조정',
        detailNotes: '400kW 고압 터보 송풍기(탄성 기초). 기준치 7.1mm/s를 상회(8.2mm/s)하여 C등급(보수필요) 판정.'
      },

      // 3. 해운대사업소
      {
        id: 'hud_sludge_pump',
        office: '해운대사업소',
        period: '2026년 상반기',
        name: '소화슬러지 이송펌프 #1',
        power: 37,
        group: 'G2',
        foundation: 'rigid',
        totalQty: 1,
        unoperated: '-',
        normalList: '#1',
        warningList: '-',
        dangerList: '-',
        overallGrade: 'A',
        maxRms: 1.1,
        history: { '2025년 상반기': 2.5, '2025년 하반기': 2.9, '2026년 상반기': 1.1, '2026년 하반기': null },
        summary: '오버홀 정비 완료 후 신품 상태 유지',
        action: '정상 운전 유지',
        detailNotes: '최근 베어링 및 씰 전체 교체 완료로 진동치 1.1mm/s 최상 상태.'
      }
    ];

    // Registered Offices List
    let officeList = [
      { name: '에너지사업소', icon: 'fa-fire-flame-curved', color: 'text-amber-400' },
      { name: '명지사업소', icon: 'fa-industry', color: 'text-cyan-400' },
      { name: '해운대사업소', icon: 'fa-water', color: 'text-blue-400' }
    ];

    let currentOffice = 'ALL';
    let vibrationChart = null;
    let detailChart = null;

    // Render Period Dropdowns
    function renderPeriodSelectors() {
      const topDropdown = document.getElementById('periodFilterDropdown');
      const modalSelect = document.getElementById('inpPeriod');
      const uploadSelect = document.getElementById('uploadPeriodSelect');

      let options = periodList.map(p => `<option value="${p}" ${p === currentPeriod ? 'selected' : ''}>${p}</option>`).join('');

      if (topDropdown) topDropdown.innerHTML = options;
      if (modalSelect) modalSelect.innerHTML = options;
      if (uploadSelect) uploadSelect.innerHTML = options;
    }

    function switchPeriod(p) {
      currentPeriod = p;
      updateDashboard();
    }

    function openAddPeriodModal() {
      document.getElementById('addPeriodModal').classList.remove('hidden');
      document.getElementById('inpPeriodName').focus();
    }
    function closeAddPeriodModal() {
      document.getElementById('addPeriodModal').classList.add('hidden');
    }

    function handlePeriodSubmit(e) {
      e.preventDefault();
      const pName = document.getElementById('inpPeriodName').value.trim();
      if (!pName) return;

      if (!periodList.includes(pName)) {
        periodList.push(pName);
        currentPeriod = pName;
        renderPeriodSelectors();
        updateDashboard();
      }

      document.getElementById('addPeriodForm').reset();
      closeAddPeriodModal();
    }

    // Toggle Chart View (Trend vs Bar)
    function setChartViewMode(mode) {
      chartViewMode = mode;
      const btnTrend = document.getElementById('btnViewTrend');
      const btnBar = document.getElementById('btnViewBar');
      const trendFilter = document.getElementById('trendFilterBar');
      const mainTitle = document.getElementById('chartMainTitle');
      const subTitle = document.getElementById('chartSubTitle');

      if (mode === 'trend') {
        btnTrend.className = 'px-2.5 py-1 rounded-md font-bold bg-cyan-600 text-white transition flex items-center gap-1.5 shadow';
        btnBar.className = 'px-2.5 py-1 rounded-md text-slate-400 hover:text-white transition flex items-center gap-1.5';
        trendFilter.classList.remove('hidden');
        mainTitle.innerHTML = '<i class="fa-solid fa-chart-line text-cyan-400"></i> 설비 진동치 상태 및 반기별 추이선 (Trend)';
        subTitle.innerText = '반기별 시계열 데이터 누적 분석을 통한 이상 징후 조기 포착';
      } else {
        btnBar.className = 'px-2.5 py-1 rounded-md font-bold bg-cyan-600 text-white transition flex items-center gap-1.5 shadow';
        btnTrend.className = 'px-2.5 py-1 rounded-md text-slate-400 hover:text-white transition flex items-center gap-1.5';
        trendFilter.classList.add('hidden');
        mainTitle.innerHTML = '<i class="fa-solid fa-chart-simple text-cyan-400"></i> 현재 주기 설비별 RMS 및 ISO 경보 기준 비교';
        subTitle.innerText = `${currentPeriod} 기준 실측치와 경보(C), 위험(D) 임계선`;
      }
      renderChart();
    }

    // Render Office Tabs Dynamically
    function renderOfficeTabs() {
      const container = document.getElementById('officeTabs');
      const dropdown = document.getElementById('officeFilterDropdown');
      const selectModal = document.getElementById('inpOffice');

      let tabsHtml = `
        <button onclick="switchOffice('ALL')" class="office-tab px-3 py-1.5 rounded-lg text-xs font-bold transition ${currentOffice === 'ALL' ? 'bg-cyan-600 text-white shadow' : 'bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white'}" data-office="ALL">
          전체 사업소
        </button>
      `;

      officeList.forEach(off => {
        const isSelected = currentOffice === off.name;
        tabsHtml += `
          <button onclick="switchOffice('${off.name}')" class="office-tab px-3 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 ${isSelected ? 'bg-cyan-600 text-white font-bold shadow' : 'bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white'}" data-office="${off.name}">
            <i class="fa-solid ${off.icon} ${isSelected ? 'text-white' : off.color}"></i>
            <span>${off.name}</span>
          </button>
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
    }

    function openAddOfficeModal() {
      document.getElementById('addOfficeModal').classList.remove('hidden');
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
      if (officeList.some(o => o.name === name)) return;

      let color = 'text-cyan-400';
      if (icon === 'fa-fire-flame-curved') color = 'text-amber-400';
      else if (icon === 'fa-water') color = 'text-blue-400';
      else if (icon === 'fa-solar-panel') color = 'text-emerald-400';

      officeList.push({ name, icon, color });
      document.getElementById('addOfficeForm').reset();
      closeAddOfficeModal();

      renderOfficeTabs();
      switchOffice(name);
    }

    function ensureOfficeRegistered(officeName) {
      if (!officeName || officeName === 'ALL') return;
      if (!officeList.some(o => o.name === officeName)) {
        officeList.push({
          name: officeName,
          icon: 'fa-industry',
          color: 'text-indigo-400'
        });
        renderOfficeTabs();
      }
    }

    // ISO Evaluation Helper Function
    function calculateIsoGrade(powerKw, rmsValue, foundation = 'rigid') {
      let group = 'G2';
      if (powerKw > 300) group = 'G1';
      else if (powerKw < 15) group = 'G4';

      let grade = 'A';
      const rms = parseFloat(rmsValue);

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
          tab.className = 'office-tab px-3 py-1.5 rounded-lg text-xs font-bold transition bg-cyan-600 text-white shadow';
        } else {
          tab.className = 'office-tab px-3 py-1.5 rounded-lg text-xs font-semibold transition bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white';
        }
      });

      document.getElementById('currentOfficeLabel').innerText = officeName === 'ALL' ? '전체 사업소 현황' : `${officeName} 현황`;
      document.getElementById('officeFilterDropdown').value = officeName;
      document.getElementById('tableSubtitle').innerText = officeName === 'ALL' 
        ? '전체 사업소 회전기계 상태진단 및 조치 대장' 
        : `${officeName} 주요 회전설비 상태진단 및 조치 대장`;

      updateDashboard();
    }

    function syncOfficeFilter(val) {
      switchOffice(val);
    }

    // Update Trend Selector in Chart
    function updateTrendEquipDropdown(activeData) {
      const select = document.getElementById('trendEquipSelect');
      const curVal = select.value;

      let html = `<option value="ALL">상위 주요 위험/주의 설비 전체 비교 (Top 5)</option>`;
      activeData.forEach(item => {
        html += `<option value="${item.id}">${item.name} (${item.office})</option>`;
      });
      select.innerHTML = html;
      if (activeData.some(e => e.id === curVal)) {
        select.value = curVal;
      }
    }

    // Recalculate Metrics and Charts
    function updateDashboard() {
      // Filter by Office
      const officeFiltered = currentOffice === 'ALL' 
        ? equipmentData 
        : equipmentData.filter(e => e.office === currentOffice);

      // Filter by Period for Current Snapshot
      const activeData = officeFiltered.filter(e => (!e.period || e.period === currentPeriod));

      // Total count calculation
      const totalUnits = activeData.reduce((acc, cur) => acc + (cur.totalQty || 1), 0);
      const gradeA = activeData.filter(e => e.overallGrade === 'A').length;
      const gradeB = activeData.filter(e => e.overallGrade === 'B').length;
      const gradeC = activeData.filter(e => e.overallGrade === 'C').length;
      const gradeD = activeData.filter(e => e.overallGrade === 'D').length;

      document.getElementById('metric-total').innerText = totalUnits;
      document.getElementById('metric-ab').innerText = (gradeA + gradeB);
      document.getElementById('metric-c').innerText = gradeC;
      document.getElementById('metric-d').innerText = gradeD;

      const pct = totalUnits > 0 ? Math.round(((gradeA + gradeB) / activeData.length) * 100) : 0;
      document.getElementById('metric-ab-pct').innerText = pct + '%';

      updateTrendEquipDropdown(officeFiltered);
      renderTable();
      renderChart();
    }

    // Chart.js Rendering (Supports Trend Line & Current Snapshot Bar)
    function renderChart() {
      const ctx = document.getElementById('vibrationChart').getContext('2d');
      const officeFiltered = currentOffice === 'ALL' 
        ? equipmentData 
        : equipmentData.filter(e => e.office === currentOffice);

      if (vibrationChart) {
        vibrationChart.destroy();
      }

      if (chartViewMode === 'trend') {
        // LINE CHART: Historical Trend across Periods
        const selectedId = document.getElementById('trendEquipSelect') ? document.getElementById('trendEquipSelect').value : 'ALL';
        
        let targetItems = [];
        if (selectedId === 'ALL') {
          // Sort by maxRms descending, pick top 5
          targetItems = [...officeFiltered].sort((a, b) => b.maxRms - a.maxRms).slice(0, 5);
        } else {
          const found = officeFiltered.find(e => e.id === selectedId);
          if (found) targetItems = [found];
        }

        const colors = [
          '#ef4444', '#f59e0b', '#06b6d4', '#10b981', '#8b5cf6', '#ec4899'
        ];

        const datasets = targetItems.map((item, idx) => {
          const color = colors[idx % colors.length];
          const dataPoints = periodList.map(p => {
            if (item.history && item.history[p] !== undefined) {
              return item.history[p];
            }
            // Fallback: If matching period, return maxRms
            return item.period === p ? item.maxRms : null;
          });

          return {
            label: `${item.name} (${item.office.replace('사업소','')})`,
            data: dataPoints,
            borderColor: color,
            backgroundColor: color,
            tension: 0.35,
            borderWidth: 2.5,
            pointRadius: 5,
            pointHoverRadius: 7,
            spanGaps: true
          };
        });

        // Add ISO Warning Baseline (2.8 mm/s) & Danger Baseline (4.5 mm/s)
        datasets.push({
          label: 'ISO 주의 기준선 (Group2: 2.8)',
          data: periodList.map(() => 2.8),
          borderColor: 'rgba(245, 158, 11, 0.6)',
          borderDash: [5, 5],
          borderWidth: 1.5,
          pointRadius: 0,
          fill: false
        });

        datasets.push({
          label: 'ISO 정비 기준선 (Group2: 4.5)',
          data: periodList.map(() => 4.5),
          borderColor: 'rgba(239, 68, 68, 0.6)',
          borderDash: [4, 4],
          borderWidth: 1.5,
          pointRadius: 0,
          fill: false
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
            interaction: { mode: 'index', intersect: false },
            plugins: {
              legend: {
                position: 'top',
                labels: { color: '#94a3b8', font: { size: 10 }, boxWidth: 12, padding: 8 }
              },
              tooltip: {
                backgroundColor: '#0f172a',
                titleColor: '#38bdf8',
                bodyColor: '#f1f5f9',
                borderColor: '#334155',
                borderWidth: 1,
                padding: 10,
                callbacks: {
                  label: function(ctx) {
                    if (ctx.parsed.y === null) return null;
                    return ` ${ctx.dataset.label}: ${ctx.parsed.y.toFixed(2)} mm/s`;
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
        // BAR CHART: Snapshot of the selected period
        const activeData = officeFiltered.filter(e => (!e.period || e.period === currentPeriod));
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
              label: '실측 RMS (mm/s)',
              data: dataValues,
              backgroundColor: backgroundColors,
              borderRadius: 6,
              borderWidth: 0,
              barThickness: 22,
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
                padding: 10,
                callbacks: {
                  label: function(context) {
                    const item = activeData[context.dataIndex];
                    return [
                      `측정 시기: ${item.period || currentPeriod}`,
                      `사업소: ${item.office}`,
                      `실측 RMS: ${context.parsed.y} mm/s`,
                      `평가 등급: ${item.overallGrade}등급`
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

    // Render Table based on active filters & Period
    function renderTable() {
      const gFilter = document.getElementById('groupFilter').value;
      const sFilter = document.getElementById('statusFilter').value;
      const query = document.getElementById('searchInput').value.trim().toLowerCase();
      const tbody = document.getElementById('tableBody');

      let filtered = equipmentData.filter(item => {
        // Period Match
        if (item.period && item.period !== currentPeriod) return false;
        // Office Match
        if (currentOffice !== 'ALL' && item.office !== currentOffice) return false;
        // Group Match
        if (gFilter !== 'ALL' && item.group !== gFilter) return false;
        // Status Match
        if (sFilter === 'D' && item.overallGrade !== 'D') return false;
        if (sFilter === 'C' && item.overallGrade !== 'C') return false;
        if (sFilter === 'AB' && (item.overallGrade !== 'A' && item.overallGrade !== 'B')) return false;
        // Search Query
        if (query) {
          const combined = `${item.name} ${item.office} ${item.summary} ${item.action}`.toLowerCase();
          if (!combined.includes(query)) return false;
        }
        return true;
      });

      document.getElementById('filteredCount').innerText = `총 ${filtered.length}개 설비 목록 표시 중 (${currentPeriod})`;

      if (filtered.length === 0) {
        tbody.innerHTML = `
          <tr>
            <td colspan="10" class="text-center py-8 text-slate-500">
              <i class="fa-solid fa-filter-circle-xmark text-2xl mb-2 block"></i>
              <span>선택된 시기(${currentPeriod}) 및 검색 조건에 부합하는 설비 데이터가 없습니다.</span>
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

        return `
          <tr class="hover:bg-slate-800/50 transition">
            <td class="py-3 px-3">
              <span class="inline-block px-2 py-0.5 rounded text-[10px] font-mono font-medium bg-slate-800 text-slate-300 border border-slate-700">
                ${item.period || currentPeriod}
              </span>
            </td>
            <td class="py-3 px-3">
              <span class="inline-block px-2 py-0.5 rounded text-[10px] font-semibold border ${officeColor}">
                ${item.office}
              </span>
            </td>
            <td class="py-3 px-4 font-bold text-slate-100 flex items-center gap-2">
              <i class="fa-solid fa-gear text-slate-400"></i>
              ${item.name}
            </td>
            <td class="py-3 px-3 font-mono text-cyan-300 font-semibold">${item.power} kW</td>
            <td class="py-3 px-3">
              <span class="px-2 py-0.5 rounded text-[10px] font-mono bg-slate-800 text-slate-300 border border-slate-700">${item.group}</span>
            </td>
            <td class="py-3 px-3 text-center font-mono font-bold ${item.overallGrade === 'D' ? 'text-red-400' : (item.overallGrade === 'C' ? 'text-amber-400' : 'text-emerald-400')}">
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
              <div class="flex items-center justify-center gap-1.5">
                <button onclick="viewEquipmentDetail('${item.id}')" class="px-2 py-1 rounded bg-slate-800 hover:bg-slate-700 text-cyan-400 text-[11px] font-medium border border-slate-700 transition" title="상세 및 추이선">
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

    // Live Evaluation in Add Modal
    function autoCalcGrade() {
      const power = parseFloat(document.getElementById('inpPower').value) || 0;
      const foundation = document.getElementById('inpFoundation').value;
      const rms = parseFloat(document.getElementById('inpRms').value) || 0;

      const res = calculateIsoGrade(power, rms, foundation);
      const banner = document.getElementById('autoEvalBanner');
      const badge = document.getElementById('evalBadge');
      const groupText = document.getElementById('evalGroupText');
      const critText = document.getElementById('evalCriteriaText');

      groupText.innerText = `ISO 분류: ${res.group} (${power > 300 ? '>300kW' : (power < 15 ? '<15kW' : '15~300kW')}) | 기초: ${foundation === 'rigid' ? '강성(Rigid)' : '탄성(Flexible)'}`;

      if (res.group === 'G1') {
        critText.innerText = foundation === 'rigid' ? '기준치: A/B ≤ 4.5 | C ≤ 7.1 | D > 7.1' : '기준치: A/B ≤ 7.1 | C ≤ 11.0 | D > 11.0';
      } else {
        critText.innerText = foundation === 'rigid' ? '기준치: A/B ≤ 2.8 | C ≤ 4.5 | D > 4.5' : '기준치: A/B ≤ 4.5 | C ≤ 7.1 | D > 7.1';
      }

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

    // Add Equipment Form Submit
    function handleEquipmentSubmit(e) {
      e.preventDefault();
      const period = document.getElementById('inpPeriod').value;
      const office = document.getElementById('inpOffice').value;
      const name = document.getElementById('inpName').value.trim();
      const power = parseFloat(document.getElementById('inpPower').value);
      const foundation = document.getElementById('inpFoundation').value;
      const rms = parseFloat(document.getElementById('inpRms').value);
      const summary = document.getElementById('inpSummary').value.trim() || '측정 완료 및 이상 유무 점검';
      const action = document.getElementById('inpAction').value.trim() || '현 상태 모니터링 유지';
      const notes = document.getElementById('inpNotes').value.trim() || '신규 현장 계측 등록 건.';

      const evalRes = calculateIsoGrade(power, rms, foundation);
      ensureOfficeRegistered(office);

      // Check if equipment already exists to append history, or create new
      const existing = equipmentData.find(e => e.name === name && e.office === office);
      let historyObj = {};
      if (existing && existing.history) {
        historyObj = { ...existing.history };
      }
      historyObj[period] = rms;

      const newId = 'custom_' + Date.now();
      const newEntry = {
        id: newId,
        period: period,
        office: office,
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
        maxRms: rms,
        history: historyObj,
        summary: summary,
        action: action,
        detailNotes: notes
      };

      equipmentData.unshift(newEntry);
      closeAddModal();
      document.getElementById('addEquipForm').reset();

      currentPeriod = period;
      renderPeriodSelectors();
      if (currentOffice !== 'ALL' && currentOffice !== office) {
        switchOffice(office);
      } else {
        updateDashboard();
      }
    }

    // Delete Equipment
    function deleteEquipment(id) {
      equipmentData = equipmentData.filter(e => e.id !== id);
      updateDashboard();
    }

    // Modal Control Functions
    function openAddModal() {
      document.getElementById('addModal').classList.remove('hidden');
      renderPeriodSelectors();
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
      renderPeriodSelectors();
    }
    function closeUploadModal() {
      document.getElementById('uploadModal').classList.add('hidden');
    }

    function openDetailModal() {
      document.getElementById('detailModal').classList.remove('hidden');
    }
    function closeDetailModal() {
      document.getElementById('detailModal').classList.add('hidden');
      if (detailChart) {
        detailChart.destroy();
        detailChart = null;
      }
    }

    // View Equipment Detail Modal with Dedicated Line Trend Chart
    function viewEquipmentDetail(id) {
      const item = equipmentData.find(e => e.id === id);
      if (!item) return;

      const detailContainer = document.getElementById('detailModalContent');
      let badgeClass = 'grade-badge-B';
      let badgeLabel = 'B (양호)';
      if (item.overallGrade === 'D') { badgeClass = 'grade-badge-D'; badgeLabel = 'D (즉시점검)'; }
      else if (item.overallGrade === 'C') { badgeClass = 'grade-badge-C'; badgeLabel = 'C (보수필요)'; }
      else if (item.overallGrade === 'A') { badgeClass = 'grade-badge-A'; badgeLabel = 'A (최상)'; }

      detailContainer.innerHTML = `
        <div class="flex items-center justify-between pb-3 border-b border-slate-700">
          <div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-slate-800 text-cyan-400 border border-slate-700">${item.office}</span>
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-slate-800 text-purple-300 border border-slate-700">${item.period || currentPeriod}</span>
            </div>
            <h3 class="text-lg font-bold text-white mt-1">${item.name}</h3>
          </div>
          <span class="px-3 py-1 rounded text-xs font-black ${badgeClass}">${badgeLabel}</span>
        </div>

        <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 my-4 text-xs">
          <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
            <span class="text-slate-400 block mb-0.5">정격 용량</span>
            <span class="font-bold font-mono text-cyan-300 text-sm">${item.power} kW</span>
          </div>
          <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
            <span class="text-slate-400 block mb-0.5">ISO 그룹</span>
            <span class="font-bold font-mono text-slate-200 text-sm">${item.group}</span>
          </div>
          <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
            <span class="text-slate-400 block mb-0.5">기초 조건</span>
            <span class="font-bold text-slate-200 text-sm">${item.foundation === 'rigid' ? '강성(Rigid)' : '탄성(Flexible)'}</span>
          </div>
          <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
            <span class="text-slate-400 block mb-0.5">최대 실측치</span>
            <span class="font-bold font-mono ${item.overallGrade === 'D' ? 'text-red-400' : 'text-cyan-400'} text-sm">${parseFloat(item.maxRms).toFixed(2)} mm/s</span>
          </div>
        </div>

        <!-- Historical Trend Chart inside Modal -->
        <div class="p-3 bg-slate-800/60 rounded-xl border border-slate-700 my-4">
          <div class="flex justify-between items-center mb-2">
            <span class="text-xs font-bold text-slate-200 flex items-center gap-1.5">
              <i class="fa-solid fa-chart-line text-cyan-400"></i>
              반기별 진동 추세 추이선 (Trend Line)
            </span>
            <span class="text-[11px] text-slate-400">기준: RMS mm/s</span>
          </div>
          <div class="h-44 w-full">
            <canvas id="detailTrendChart"></canvas>
          </div>
        </div>

        <div class="space-y-3 text-xs">
          <div>
            <h4 class="font-semibold text-slate-300 mb-1"><i class="fa-solid fa-stethoscope text-cyan-400 mr-1.5"></i>현장 진단 요약</h4>
            <p class="p-2.5 rounded-lg bg-slate-800 text-slate-300 border border-slate-700/80">${item.summary}</p>
          </div>
          <div>
            <h4 class="font-semibold text-slate-300 mb-1"><i class="fa-solid fa-screwdriver-wrench text-amber-400 mr-1.5"></i>후속 정비 대책</h4>
            <p class="p-2.5 rounded-lg bg-slate-800 text-slate-300 border border-slate-700/80">${item.action}</p>
          </div>
          <div>
            <h4 class="font-semibold text-slate-300 mb-1"><i class="fa-solid fa-clipboard-list text-emerald-400 mr-1.5"></i>엔지니어 상세 소견</h4>
            <p class="p-2.5 rounded-lg bg-slate-800 text-slate-400 leading-relaxed border border-slate-700/80">${item.detailNotes || '등록된 추가 상세 소견이 없습니다.'}</p>
          </div>
        </div>

        <div class="pt-4 mt-4 border-t border-slate-700 flex justify-end">
          <button onclick="closeDetailModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg font-medium transition text-xs">닫기</button>
        </div>
      `;

      openDetailModal();

      // Render Individual Trend Line
      setTimeout(() => {
        const dCtx = document.getElementById('detailTrendChart').getContext('2d');
        const dData = periodList.map(p => {
          if (item.history && item.history[p] !== undefined) return item.history[p];
          return item.period === p ? item.maxRms : null;
        });

        if (detailChart) detailChart.destroy();
        detailChart = new Chart(dCtx, {
          type: 'line',
          data: {
            labels: periodList,
            datasets: [
              {
                label: item.name,
                data: dData,
                borderColor: '#06b6d4',
                backgroundColor: 'rgba(6, 182, 212, 0.15)',
                borderWidth: 3,
                pointRadius: 6,
                fill: true,
                tension: 0.3
              },
              {
                label: 'C등급 경보(2.8)',
                data: periodList.map(() => 2.8),
                borderColor: 'rgba(245, 158, 11, 0.7)',
                borderDash: [4, 4],
                borderWidth: 1.5,
                pointRadius: 0
              },
              {
                label: 'D등급 위험(4.5)',
                data: periodList.map(() => 4.5),
                borderColor: 'rgba(239, 68, 68, 0.7)',
                borderDash: [4, 4],
                borderWidth: 1.5,
                pointRadius: 0
              }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
              x: { grid: { color: 'rgba(51, 65, 85, 0.3)' }, ticks: { color: '#94a3b8' } },
              y: { min: 0, max: 9, grid: { color: 'rgba(51, 65, 85, 0.3)' }, ticks: { color: '#94a3b8' } }
            }
          }
        });
      }, 50);
    }

    // Interactive Calculator Logic
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
        rec.innerText = '단기간 운전은 가능하나 결함이 진행 중입니다. 주파수 분석(축정렬/베어링) 및 차기 계획정비 시 부품 교체를 권고합니다.';
      } else {
        badge.classList.add('grade-badge-D');
        badge.innerText = 'Grade D';
        title.className = 'text-sm font-bold text-red-400';
        title.innerText = '즉시 점검 (운전 중단 권고)';
        box.className = 'mt-4 p-4 rounded-xl border transition-all duration-300 bg-red-950/30 border-red-500/50';
        rec.innerText = '설비 손상 및 소착, 돌발 정지 위험 수준입니다. 즉시 기기를 정지하고 정밀 점검 및 분해정비를 실시하십시오.';
      }

      if (group === 'G1') {
        bound.innerText = foundation === 'rigid' ? 'A/B: ≤4.5 | C: ≤7.1 | D: >7.1' : 'A/B: ≤7.1 | C: ≤11.0 | D: >11.0';
      } else {
        bound.innerText = foundation === 'rigid' ? 'A/B: ≤2.8 | C: ≤4.5 | D: >4.5' : 'A/B: ≤4.5 | C: ≤7.1 | D: >7.1';
      }
    }

    // Excel & CSV File Upload Processing
    function handleFileSelect(e) {
      const file = e.target.files[0];
      if (!file) return;
      processDataFile(file);
    }

    function processDataFile(file) {
      const statusBox = document.getElementById('uploadStatusBox');
      const targetPeriod = document.getElementById('uploadPeriodSelect').value;
      statusBox.classList.remove('hidden');
      statusBox.innerHTML = `<span class="text-cyan-400 font-semibold"><i class="fa-solid fa-spinner fa-spin mr-1"></i>파일 파싱 중: ${file.name}</span>`;

      const reader = new FileReader();

      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          if (jsonData.length < 2) {
            statusBox.innerHTML = `<span class="text-red-400">오류: 데이터 행이 존재하지 않는 빈 파일입니다.</span>`;
            return;
          }

          let addedCount = 0;
          // Expected Header: 사업소, 설비명, 용량kW, 실측RMS, 진단요약, 조치계획, 측정시기
          for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length < 3) continue;

            const office = (row[0] || '에너지사업소').toString().trim();
            const name = (row[1] || '미지정 회전설비').toString().trim();
            const power = parseFloat(row[2]) || 55;
            const rms = parseFloat(row[3]) || 2.5;
            const summary = (row[4] || '엑셀 일괄 업로드 설비').toString().trim();
            const action = (row[5] || '추세 관리 및 정기점검').toString().trim();
            const period = (row[6] || targetPeriod).toString().trim();

            if (!periodList.includes(period)) {
              periodList.push(period);
            }

            ensureOfficeRegistered(office);
            const evalRes = calculateIsoGrade(power, rms, 'rigid');

            let historyObj = {};
            historyObj[period] = rms;

            equipmentData.unshift({
              id: 'imp_' + Date.now() + '_' + i,
              period: period,
              office: office,
              name: name,
              power: power,
              group: evalRes.group,
              foundation: 'rigid',
              totalQty: 1,
              unoperated: '-',
              normalList: evalRes.grade === 'A' || evalRes.grade === 'B' ? '#1' : '-',
              warningList: evalRes.grade === 'C' ? '#1' : '-',
              dangerList: evalRes.grade === 'D' ? '#1' : '-',
              overallGrade: evalRes.grade,
              maxRms: rms,
              history: historyObj,
              summary: summary,
              action: action,
              detailNotes: `엑셀 파일(${file.name}) 일괄 업로드 등록 건`
            });
            addedCount++;
          }

          currentPeriod = targetPeriod;
          renderPeriodSelectors();

          statusBox.className = 'p-3 rounded-lg border text-xs bg-emerald-950/20 border-emerald-500/40 text-emerald-400 font-medium';
          statusBox.innerHTML = `<i class="fa-solid fa-circle-check mr-1"></i>총 <strong>${addedCount}건</strong>의 설비가 (${targetPeriod}) 주기로 성공적으로 등록되었습니다!`;

          updateDashboard();
          setTimeout(() => { closeUploadModal(); }, 1500);

        } catch (err) {
          console.error(err);
          statusBox.className = 'p-3 rounded-lg border text-xs bg-red-950/20 border-red-500/40 text-red-400';
          statusBox.innerHTML = `<i class="fa-solid fa-triangle-exclamation mr-1"></i>파일 파싱 실패: 형식을 확인해주세요.`;
        }
      };

      reader.readAsArrayBuffer(file);
    }

    // Download Sample CSV / Excel Template with Period column
    function downloadSampleTemplate() {
      const templateData = [
        ["사업소", "설비명", "용량kW", "실측RMS", "진단요약", "조치계획", "측정시기"],
        ["에너지사업소", "#5 응축수펌프", 45, 2.3, "베어링 수평 진동 안정", "정상 운전 유지", "2026년 상반기"],
        ["명지사업소", "소화가스 압축기 #1", 160, 4.8, "고주파 진동 및 커플링 이음", "커플링 엘레먼트 교체 요망", "2026년 상반기"],
        ["해운대사업소", "탈수여액 펌프 #2", 15, 5.2, "모터 구동측 과다 진동", "긴급 정밀점검 필요", "2026년 상반기"]
      ];

      const ws = XLSX.utils.aoa_to_sheet(templateData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "진단데이터_양식");
      XLSX.writeFile(wb, "회전설비_진동측정_업로드양식.xlsx");
    }

    // Export Currently Displayed List to CSV
    function exportToCSV() {
      const target = currentOffice === 'ALL' ? equipmentData : equipmentData.filter(e => e.office === currentOffice);
      const rows = [
        ["측정시기", "사업소", "설비명", "용량(kW)", "기계그룹", "실측RMS(mm/s)", "종합등급", "진단요약", "조치대책"]
      ];

      target.forEach(e => {
        rows.push([e.period || currentPeriod, e.office, e.name, e.power, e.group, e.maxRms, e.overallGrade, e.summary, e.action]);
      });

      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "진단결과");
      XLSX.writeFile(wb, `회전설비_진동진단결과_${currentOffice}_${currentPeriod}.xlsx`);
    }

    // Drag and Drop Event Handlers
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

    // Window Init
    window.onload = function() {
      renderPeriodSelectors();
      renderOfficeTabs();
      updateDashboard();
      setChartViewMode('trend');
    };
  </script>
</body>
</html>"""

# Streamlit 환경 안에서 HTML 대시보드를 iframe 형태로 온전히 구동
components.html(DASHBOARD_HTML, height=1400, scrolling=True)