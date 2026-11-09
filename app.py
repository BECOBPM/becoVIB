# ... existing code ...
        <!-- Template Download Guide -->
        <div class="p-3 rounded-xl bg-slate-800/80 border border-slate-700 flex items-center justify-between">
          <div>
            <span class="font-bold text-slate-200 block">진동측정결과 시트 표준 양식 (4개 포인트 12축)</span>
            <span class="text-slate-400 text-[11px]">열 구성: 사업소, 설비명, 용량kW, 전동기부하(H/V/A), 전동기반부하(H/V/A), 펌프부하(H/V/A), 펌프반부하(H/V/A), 진단요약, 조치계획, 측정시기</span>
          </div>
          <button onclick="downloadSampleTemplate()" class="px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-cyan-300 rounded-lg font-medium border border-slate-600 transition flex items-center gap-1.5">
            <i class="fa-solid fa-download"></i>
            <span>양식 다운로드</span>
          </button>
        </div>
# ... existing code ...
    // Master Multi-Office Equipment Dataset with 4 Points (DE/NDE) & 12 Axes Measurements
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

    let officeList = [
      { name: '에너지사업소', icon: 'fa-fire-flame-curved', color: 'text-amber-400' },
      { name: '명지사업소', icon: 'fa-industry', color: 'text-cyan-400' },
      { name: '해운대사업소', icon: 'fa-water', color: 'text-blue-400' }
    ];

    let currentOffice = 'ALL';
    let vibrationChart = null;
    let detailChart = null;
    let detailAxesChart = null;
# ... existing code ...
    // Excel & CSV File Upload Processing with '진동측정결과' Sheet Auto Detection
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
          
          // Auto detect '진동측정결과' sheet or use first sheet
          let targetSheetName = workbook.SheetNames[0];
          for (let name of workbook.SheetNames) {
            if (name.includes('진동측정') || name.includes('진동') || name.includes('측정결과')) {
              targetSheetName = name;
              break;
            }
          }

          const worksheet = workbook.Sheets[targetSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          if (jsonData.length < 2) {
            statusBox.innerHTML = `<span class="text-red-400">오류: [${targetSheetName}] 시트에 데이터 행이 존재하지 않습니다.</span>`;
            return;
          }

          // Locate header row (look for 설비명 or 구동 or 펌프)
          let headerIdx = 0;
          for (let r = 0; r < Math.min(jsonData.length, 10); r++) {
            const rowStr = (jsonData[r] || []).join(' ');
            if (rowStr.includes('설비명') || rowStr.includes('전동기') || rowStr.includes('모터') || rowStr.includes('측정')) {
              headerIdx = r;
              break;
            }
          }

          let addedCount = 0;

          for (let i = headerIdx + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length < 3) continue;

            // Check if row is sub-header
            const firstCell = String(row[0] || '').trim();
            const secondCell = String(row[1] || '').trim();
            if (firstCell.includes('구분') || firstCell.includes('사업소') || secondCell.includes('설비명')) continue;

            let office = '에너지사업소';
            let name = '미지정 설비';
            let power = 55;
            let period = targetPeriod;
            let summary = '';
            let action = '';

            let m_de_h = 0, m_de_v = 0, m_de_a = 0;
            let m_nde_h = 0, m_nde_v = 0, m_nde_a = 0;
            let p_de_h = 0, p_de_v = 0, p_de_a = 0;
            let p_nde_h = 0, p_nde_v = 0, p_nde_a = 0;
            let maxRms = 0;

            if (row.length >= 15) {
              // Full 4-Point (12 Axis) format from '진동측정결과'
              // [사업소, 설비명, 용량kW, M-DE(H,V,A), M-NDE(H,V,A), P-DE(H,V,A), P-NDE(H,V,A), 진단요약, 조치계획, 측정시기]
              office = String(row[0] || '에너지사업소').trim();
              name = String(row[1] || '미지정 설비').trim();
              power = parseFloat(row[2]) || 55;

              m_de_h = parseFloat(row[3]) || 0;
              m_de_v = parseFloat(row[4]) || 0;
              m_de_a = parseFloat(row[5]) || 0;

              m_nde_h = parseFloat(row[6]) || 0;
              m_nde_v = parseFloat(row[7]) || 0;
              m_nde_a = parseFloat(row[8]) || 0;

              p_de_h = parseFloat(row[9]) || 0;
              p_de_v = parseFloat(row[10]) || 0;
              p_de_a = parseFloat(row[11]) || 0;

              p_nde_h = parseFloat(row[12]) || 0;
              p_nde_v = parseFloat(row[13]) || 0;
              p_nde_a = parseFloat(row[14]) || 0;

              summary = (row[15] || '진동측정결과 시트 등록').toString().trim();
              action = (row[16] || '추세 관리').toString().trim();
              if (row[17]) period = String(row[17]).trim();

              maxRms = Math.max(m_de_h, m_de_v, m_de_a, m_nde_h, m_nde_v, m_nde_a, p_de_h, p_de_v, p_de_a, p_nde_h, p_nde_v, p_nde_a);
            } else if (row.length >= 9) {
              // Standard 3-Axis format
              office = String(row[0] || '에너지사업소').trim();
              name = String(row[1] || '미지정 설비').trim();
              power = parseFloat(row[2]) || 55;

              m_de_h = parseFloat(row[3]) || 0;
              m_de_v = parseFloat(row[4]) || 0;
              m_de_a = parseFloat(row[5]) || 0;
              p_de_h = parseFloat(row[6]) || 0;
              p_de_v = parseFloat(row[7]) || 0;
              p_de_a = parseFloat(row[8]) || 0;

              summary = (row[9] || '3축 계측 업로드').toString().trim();
              action = (row[10] || '추세 관리').toString().trim();
              if (row[11]) period = String(row[11]).trim();

              maxRms = Math.max(m_de_h, m_de_v, m_de_a, p_de_h, p_de_v, p_de_a);
            } else {
              // Simple RMS format
              office = String(row[0] || '에너지사업소').trim();
              name = String(row[1] || '미지정 설비').trim();
              power = parseFloat(row[2]) || 55;
              maxRms = parseFloat(row[3]) || 2.5;
              m_de_h = maxRms; p_de_h = maxRms;
              summary = (row[4] || '일괄 업로드').toString().trim();
              action = (row[5] || '추세 관리').toString().trim();
              if (row[6]) period = String(row[6]).trim();
            }

            if (!periodList.includes(period)) periodList.push(period);
            ensureOfficeRegistered(office);
            const evalRes = calculateIsoGrade(power, maxRms, 'rigid');

            let historyObj = {};
            historyObj[period] = maxRms;

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
              detailNotes: `시트 [${targetSheetName}]에서 파싱 등록된 설비`
            });
            addedCount++;
          }

          currentPeriod = targetPeriod;
          renderPeriodSelectors();

          statusBox.className = 'p-3 rounded-lg border text-xs bg-emerald-950/20 border-emerald-500/40 text-emerald-400 font-medium';
          statusBox.innerHTML = `<i class="fa-solid fa-circle-check mr-1"></i>[<strong>${targetSheetName}</strong>] 시트에서 총 <strong>${addedCount}건</strong>이 성공적으로 추출 및 판정되었습니다!`;

          updateDashboard();
          setTimeout(() => { closeUploadModal(); }, 1500);

        } catch (err) {
          console.error(err);
          statusBox.className = 'p-3 rounded-lg border text-xs bg-red-950/20 border-red-500/40 text-red-400';
          statusBox.innerHTML = `<i class="fa-solid fa-triangle-exclamation mr-1"></i>시트 파싱 실패: 파일 서식을 확인해 주세요.`;
        }
      };

      reader.readAsArrayBuffer(file);
    }

    // Download Sample Excel Template Matching '진동측정결과' Sheet (4 Points 12 Axes)
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

    // Export Currently Displayed List to CSV with 4 Points
# ... existing code ...