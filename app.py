import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule

wb = openpyxl.Workbook()

# ==========================================
# 1. 진동측정기준 안내 시트 (첫 번째 탭)
# ==========================================
ws_guide = wb.active
ws_guide.title = "진동측정기준"
ws_guide.views.sheetView[0].showGridLines = True

# Title
ws_guide.merge_cells("A1:F1")
ws_guide["A1"] = "■ ISO 10816-3 진동 평가 기준표 (전동기 및 회전기기)"
ws_guide["A1"].font = Font(name="맑은 고딕", size=14, bold=True, color="1F4E78")

# Table Headers
headers_guide = ["Machine Group", "설비 구분 / 용량", "기초 구분", "A (양호)", "B (장기운전)", "C (보수필요)", "D (즉시점검)"]
ws_guide.append([]) # Blank row 2
ws_guide.append(headers_guide) # Row 3

guide_data = [
    ["Group 1", "대형 전동기/설비 (P > 300 kW)", "강성 (Rigid)", "≤ 2.3 mm/s", "≤ 4.5 mm/s", "≤ 7.1 mm/s", "> 7.1 mm/s"],
    ["Group 1", "대형 전동기/설비 (P > 300 kW)", "유연 (Flexible)", "≤ 3.5 mm/s", "≤ 7.1 mm/s", "≤ 11.0 mm/s", "> 11.0 mm/s"],
    ["Group 2", "중형 전동기/설비 (15 kW < P ≤ 300 kW)", "강성 (Rigid)", "≤ 1.4 mm/s", "≤ 2.8 mm/s", "≤ 4.5 mm/s", "> 4.5 mm/s"],
    ["Group 2", "중형 전동기/설비 (15 kW < P ≤ 300 kW)", "유연 (Flexible)", "≤ 2.3 mm/s", "≤ 4.5 mm/s", "≤ 7.1 mm/s", "> 7.1 mm/s"],
    ["Group 3", "펌프 (P > 15 kW, 분리형 드라이버)", "강성 (Rigid)", "≤ 1.4 mm/s", "≤ 2.8 mm/s", "≤ 4.5 mm/s", "> 4.5 mm/s"],
    ["Group 3", "펌프 (P > 15 kW, 분리형 드라이버)", "유연 (Flexible)", "≤ 2.3 mm/s", "≤ 4.5 mm/s", "≤ 7.1 mm/s", "> 7.1 mm/s"],
    ["Group 4", "펌프 (P > 15 kW, 일체형 드라이버)", "강성 (Rigid)", "≤ 0.71 mm/s", "≤ 1.8 mm/s", "≤ 2.8 mm/s", "> 2.8 mm/s"],
    ["Group 4", "펌프 (P > 15 kW, 일체형 드라이버)", "유연 (Flexible)", "≤ 1.4 mm/s", "≤ 2.8 mm/s", "≤ 4.5 mm/s", "> 4.5 mm/s"],
]

for row in guide_data:
    ws_guide.append(row)

# Style Guide Table
thin_border = Border(left=Side(style='thin', color='D9D9D9'),
                     right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'),
                     bottom=Side(style='thin', color='D9D9D9'))

for row in ws_guide.iter_rows(min_row=3, max_row=11, min_col=1, max_col=7):
    for cell in row:
        cell.font = Font(name="맑은 고딕", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        if cell.row == 3:
            cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")

# ==========================================
# 2. 진동 측정 데이터 입력 시트 (두 번째 탭)
# ==========================================
ws_data = wb.create_sheet(title="진동측정결과")
ws_data.views.sheetView[0].showGridLines = True

headers_data = ["측정일자", "설비명", "전동기(kW)", "측정위치", "방향", "진동속도 (rms)", "부하상태(%)", "판정 (A~D)", "이상 원인", "정비 우선순위"]
ws_data.append(headers_data)

# Header Styling
for cell in ws_data[1]:
    cell.fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    cell.font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center")

# 샘플 데이터 5줄 삽입 (자동 수식 내장)
sample_rows = [
    ["2026-04-02", "FD Fan #1", 310, "1(전동기)", "X", 12.3, 30],
    ["2026-04-02", "FD Fan #1", 310, "1(전동기)", "Y", 2.8, 30],
    ["2026-04-02", "1차 냉각수 #1", 11, "2(펌프)", "Z", 4.0, 100],
    ["2026-04-02", "보일러 급수펌프 #1", 95, "1(전동기)", "X", 1.1, 100],
    ["2026-04-02", "보일러 급수펌프 #1", 95, "1(전동기)", "Y", 4.8, 100]
]

for idx, r in enumerate(sample_rows, start=2):
    # H열(판정): 자동 판정 수식
    formula_grade = f'=IF(ISBLANK(F{idx}),"",IF(C{idx}>300,IF(F{idx}<=2.3,"A",IF(F{idx}<=4.5,"B",IF(F{idx}<=7.1,"C","D"))),IF(F{idx}<=1.4,"A",IF(F{idx}<=2.8,"B",IF(F{idx}<=4.5,"C","D")))))'
    # J열(정비 우선순위): 자동 연동 수식
    formula_priority = f'=IF(H{idx}="A","양호",IF(H{idx}="B","양호",IF(H{idx}="C","보수필요",IF(H{idx}="D","즉시점검",""))))'
    
    row_data = r + [formula_grade, "", formula_priority]
    ws_data.append(row_data)

# 데이터 셀 서식 및 alignment
for row in ws_data.iter_rows(min_row=2, max_row=100, min_col=1, max_col=10):
    for cell in row:
        cell.font = Font(name="맑은 고딕", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")

# 조건부 서식 (H열 판정에 따라 색상 다르게)
fill_d = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # 빨강
fill_c = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid") # 주황/노랑
fill_ab = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # 녹색

ws_data.conditional_formatting.add("H2:H100", CellIsRule(operator='equal', formula=['"D"'], fill=fill_d))
ws_data.conditional_formatting.add("H2:H100", CellIsRule(operator='equal', formula=['"C"'], fill=fill_c))
ws_data.conditional_formatting.add("H2:H100", CellIsRule(operator='equal', formula=['"B"'], fill=fill_ab))
ws_data.conditional_formatting.add("H2:H100", CellIsRule(operator='equal', formula=['"A"'], fill=fill_ab))

# 저장
wb.save("설비점검_진동측정_표준양식.xlsx")
print("엑셀 샘플 파일 생성이 완료되었습니다.")