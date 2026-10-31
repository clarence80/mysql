import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    """셀의 배경색을 설정하는 함수"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """셀의 상하좌우 여백(Padding)을 설정하는 함수 (단위: dxa)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table):
    """표에 얇은 회색 테두리를 적용하는 함수"""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        '<w:tblBorders %s>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        '</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

def format_cell(cell, text, bold=False, font_size=10, align=WD_ALIGN_PARAGRAPH.CENTER, bg_color=None):
    """셀 내부 텍스트의 폰트, 정렬, 배경색을 한 번에 설정하는 함수"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.font.name = '맑은 고딕'
    run._r.get_or_add_rPr().get_or_add_rXml().set(qn('w:eastAsia'), '맑은 고딕')
    run.font.size = Pt(font_size)
    run.bold = bold
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
    if bg_color:
        set_cell_background(cell, bg_color)

# 1. 문서 생성 및 기본 여백 설정
doc = docx.Document()
sections = doc.sections
for section in sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

# --- 표지 (1페이지) ---
p_top = doc.add_paragraph()
p_top.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p_top.add_run("단락\t\t\t\t\t\t\t\t\t파일").font.size = Pt(9)

for _ in range(5): doc.add_paragraph() # 상단 공백

p_title1 = doc.add_paragraph()
p_title1.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p_title1.add_run("KB금융그룹")
run.font.name = '맑은 고딕'
run.font.size = Pt(26)
run.bold = True

p_title2 = doc.add_paragraph()
p_title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p_title2.add_run("보고서 자동화 시스템")
run.font.name = '맑은 고딕'
run.font.size = Pt(18)
run.bold = True

p_title3 = doc.add_paragraph()
p_title3.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p_title3.add_run("09월 서비스카탈로그 서비스 수준 보고서")
run.font.name = '맑은 고딕'
run.font.size = Pt(16)
run.bold = True

for _ in range(8): doc.add_paragraph() # 중간 공백

p_date = doc.add_paragraph()
p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p_date.add_run("2026.09.01 ~ 2026.09.30")
run.font.size = Pt(11)

for _ in range(2): doc.add_paragraph()

p_logo1 = doc.add_paragraph()
p_logo1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_logo1.add_run("KB금융그룹").bold = True

p_logo2 = doc.add_paragraph()
p_logo2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_logo2.add_run("KB데이타시스템").bold = True

# --- 본문 페이지로 이동 ---
doc.add_page_break()

# 헤더용 단락
p_header = doc.add_paragraph()
run_h = p_header.add_run("KB금융그룹\t\t\t\t\t\tSLA(Service Level Agreement) 방안")
run_h.font.size = Pt(10)
p_sub = doc.add_paragraph()
run_s = p_sub.add_run("서비스 수준 보고서")
run_s.font.size = Pt(14)
run_s.bold = True

# 타이틀 스타일 헬퍼
def add_section_title(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = '맑은 고딕'
    run.font.size = Pt(11)
    run.bold = True

# --- 1. 기본 정보 ---
add_section_title("1. 기본 정보")
table1 = doc.add_table(rows=3, cols=4)
table1.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(table1)

t1_data = [
    [("보고서명", True), ("서비스카탈로그(SU03 DBA 기술지원)\n서비스 수준 보고서", False), ("서비스 기간", True), ("2026년 09월 01일 ~ 2026년 09월 30일", False)],
    [("담당부서 (시스템명)", True), ("인프라본부/보고서자동화시스템팀/P리전\n통합MSP", False), ("작성자(연락처)", True), ("홍길동(010-XXXX-XXXX)", False)],
    [("공급사업자", True), ("주식회사 케이비데이타시스템", False), ("제출일자", True), ("2026년 10월 25일", False)]
]
for r_idx, row in enumerate(table1.rows):
    for c_idx, cell in enumerate(row.cells):
        text, is_header = t1_data[r_idx][c_idx]
        bg = "F2F2F2" if is_header else None
        format_cell(cell, text, bold=is_header, bg_color=bg, align=WD_ALIGN_PARAGRAPH.CENTER if is_header else WD_ALIGN_PARAGRAPH.LEFT)

# --- 2. 평가 항목 ---
add_section_title("2. 평가 항목")
table2 = doc.add_table(rows=7, cols=8)
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(table2)

headers2 = ["업무 영역", "측정 지표", "점수", "운영 기준", "측정 주기", "평가 기준\n목표 수준", "(09)월 실적\n측정결과", "평가점수"]
for c_idx, text in enumerate(headers2):
    format_cell(table2.rows[0].cells[c_idx], text, bold=True, bg_color="E6E6E6")

t2_rows = [
    ["서비스카탈로그\n관리\n(100)", "가동률", "25", "비율", "월간", "25", "100%", "25"],
    ["서비스카탈로그\n관리\n(100)", "장애 발생 여부", "25", "여부", "월간", "25", "장애 미발생", "25"],
    ["서비스카탈로그\n관리\n(100)", "장애/알람 적기\n미준수", "25", "건수", "연간", "25", "0건", "25"],
    ["서비스카탈로그\n관리\n(100)", "SR적기처리율", "25", "비율", "월간", "21", "100%", "25"],
]
for r_idx, row_data in enumerate(t2_rows):
    for c_idx, text in enumerate(row_data):
        format_cell(table2.rows[r_idx+1].cells[c_idx], text)

# 합계 행
format_cell(table2.rows[5].cells[0], "합계", bold=True)
format_cell(table2.rows[5].cells[1], "", bold=False)
format_cell(table2.rows[5].cells[2], "100", bold=True)
format_cell(table2.rows[5].cells[3], "", bold=False)
format_cell(table2.rows[5].cells[4], "", bold=False)
format_cell(table2.rows[5].cells[5], "96", bold=True)
format_cell(table2.rows[5].cells[6], "", bold=False)
format_cell(table2.rows[5].cells[7], "100", bold=True)

# 종합평가 행 병합 및 작성
summary_cell_start = table2.rows[6].cells[0]
for c_idx in range(1, 8):
    summary_cell_start.merge(table2.rows[6].cells[c_idx])
format_cell(summary_cell_start, "종합평가: 100 점 (S 등급)", bold=True, bg_color="FFF2CC")

# 세로 병합 처리 (업무 영역 및 합계 일부)
table2.rows[1].cells[0].merge(table2.rows[2].cells[0]).merge(table2.rows[3].cells[0]).merge(table2.rows[4].cells[0])
table2.rows[5].cells[0].merge(table2.rows[5].cells[1])

# --- 3. 추가 의견 ---
add_section_title("3. 추가 의견")
table3 = doc.add_table(rows=2, cols=1)
table3.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(table3)
format_cell(table3.rows[0].cells[0], "내용", bold=True, bg_color="E6E6E6")
format_cell(table3.rows[1].cells[0], "\n\n", align=WD_ALIGN_PARAGRAPH.LEFT) # 빈 공간 확보

# --- 4. 상세 내용 ---
add_section_title("4. 상세 내용")
table4 = doc.add_table(rows=5, cols=4)
table4.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(table4)

headers4 = ["업무 영역", "측정 지표", "내용", "관련 문서"]
for c_idx, text in enumerate(headers4):
    format_cell(table4.rows[0].cells[c_idx], text, bold=True, bg_color="E6E6E6")

t4_rows = [
    ["서비스카탈로그 관리", "가동률", "DBA 기술지원 서비스로 인한 비가동시간은 0분(minute)으로 가동률(%)은\n100.00%입니다.", "해당없음"],
    ["서비스카탈로그 관리", "장애 발생 여부", "장애 미발생", "해당없음"],
    ["서비스카탈로그 관리", "장애/알람 적기\n미준수", "0건", "해당없음"],
    ["서비스카탈로그 관리", "SR적기처리율", "접수 기준 기한내 SR처리율(%)은 100.00% 입니다.\n09월 납기윤 미준수건은 0 건 입니다.", "해당없음"]
]
for r_idx, row_data in enumerate(t4_rows):
    for c_idx, text in enumerate(row_data):
        align = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 2 else WD_ALIGN_PARAGRAPH.CENTER
        format_cell(table4.rows[r_idx+1].cells[c_idx], text, align=align)

# 상세 내용의 '업무 영역' 세로 병합
table4.rows[1].cells[0].merge(table4.rows[2].cells[0]).merge(table4.rows[3].cells[0]).merge(table4.rows[4].cells[0])

# 하단 페이지 번호용 여백 공백 제거 및 문서 저장
doc.save("서비스_수준_보고서.docx")
print("성공적으로 워드 문서('서비스_수준_보고서.docx')가 생성되었습니다.")
