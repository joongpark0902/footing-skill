# -*- coding: utf-8 -*-
"""FDD 예시용 가상 자료 생성기 — 보고서 draft(pptx) + 실사조서(xlsx).

    python make_sample.py

회사·숫자는 전부 지어낸 것이다(㈜한결식품(가상)). 보고서에는 풋팅이 잡아야 할 오류를 일부러 심어 두었다.
어떤 오류를 심었는지는 아래 PLANTED 와 README 에 적어 두었으니, 산출물과 대조해 볼 수 있다.
"""
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
CO = "㈜한결식품(가상)"
PERIODS = ["FY2023", "FY2024", "FY2025", "2026 1H"]

# ──────────────────────────────────────────────── 조서(진짜 값, 백만원)
IS = [  # (계정, [FY2023, FY2024, FY2025, 2026 1H])
    ("매출", [18420, 21350, 24560, 12980]),
    ("매출원가", [12890, 14730, 16920, 8950]),
    ("매출총이익", [5530, 6620, 7640, 4030]),
    ("판매비와관리비", [3980, 4510, 5120, 2760]),
    ("영업이익", [1550, 2110, 2520, 1270]),
    ("감가상각비·무형자산상각비", [620, 690, 760, 410]),
    ("EBITDA", [2170, 2800, 3280, 1680]),
]
QOE = [  # (조정 항목, 값, 근거)
    ("보고 EBITDA", [2170, 2800, 3280, 1680], "IS_요약"),
    ("(+) 일회성 소송합의금", [0, 180, 0, 0], "FY2024 ○○유통 소송 합의 — 영업외 성격, 판관비에 계상"),
    ("(−) 특수관계자 임차료 시가 초과분", [-60, -60, -60, -30], "대표이사 소유 창고 임차료 월 15백만 vs 시가 10백만"),
    ("(−) 정부보조금 수익", [0, 0, -120, 0], "FY2025 스마트공장 보조금, 영업수익에 포함 — 비경상"),
    ("(−) 재고평가손실 환입", [0, 0, -90, 0], "FY2025 전기 평가손실 환입분 — 비경상"),
    ("조정 EBITDA", [2110, 2920, 3010, 1650], "합계"),
]
NWC = [  # 기말 잔액
    ("매출채권", [2310, 2640, 3050, 3180]),
    ("재고자산", [1480, 1620, 1810, 1900]),
    ("매입채무", [-1150, -1290, -1470, -1520]),
    ("기타유동부채", [-120, -140, -180, -190]),
    ("순운전자본", [2520, 2830, 3210, 3370]),
]
DEBT = [  # FY2025 말
    ("단기차입금", 2100, "○○은행 운전자금, 만기 2026-06"),
    ("장기차입금", 1800, "△△은행 시설자금, 만기 2028-12"),
    ("리스부채", 1240, "본사 사옥 리스 (K-IFRS 1116)"),
    ("(−) 현금및현금성자산", -1690, "은행조회서 대사 완료"),
    ("순차입금", 3450, "리스부채 포함"),
]
CUST = [  # FY2025 매출처
    ("A유통", 9820, "확인완료", ""),
    ("B마트", 6140, "확인필요", "반품 조건(90일 미판매분 반품) 계약서 미입수 — 매출 인식 시점 검토 필요"),
    ("C식자재", 3680, "확인완료", ""),
    ("기타", 4920, "확인완료", "거래처 42개"),
    ("합계", 24560, "", ""),
]

# ──────────────────────────────────────────────── 보고서에 심은 오류
PLANTED = """
E1 S 숫자대사   요약 손익 FY2025 매출 24,650 (조서 24,560 — 숫자 뒤바뀜)
E2 S 합계검산   조정 EBITDA FY2024 합계 2,900 (항목 합 2,920)
E3 A 서술검산   "FY2025 매출은 전년 대비 17.2% 성장" (24,560/21,350−1 = 15.0%)
E4 S 결론·근거  순차입금 표가 리스부채를 빼고 2,210 — 본문은 "리스부채 포함 순차입금" 이라고 씀 (조서 3,450)
E5 A 결론·근거  조서 매출처별 시트에 B마트 '확인필요'(반품 조건 미입수)가 남아 있는데 보고서는 "특이사항 없음"
E6 B 합계검산   순운전자본 FY2025 합계 3,120 (항목 합 3,210)
E7 C 용어       "매출 총이익 / 매출총이익", "순운전 자본 / 순운전자본" 혼용
E8 B 누락       순운전자본 표에 단위·기준일 표기 없음, 순운전자본 정의(산식) 없음
"""


def _font(run, size=11, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "맑은 고딕"
    if color:
        run.font.rgb = RGBColor.from_string(color)


def add_table(slide, rows, left, top, width, height, col_widths=None, header_fill="1F3864", num_cols=None):
    nr, nc = len(rows), len(rows[0])
    shape = slide.shapes.add_table(nr, nc, Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            tbl.columns[i].width = Inches(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = "" if val is None else (format(val, ",") if isinstance(val, (int, float)) else str(val))
            p = cell.text_frame.paragraphs[0]
            p.alignment = 3 if (c > 0 and (num_cols is None or c in num_cols)) else 1   # RIGHT / LEFT
            for run in p.runs:
                _font(run, 10, bold=(r == 0), color="FFFFFF" if r == 0 else None)
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(header_fill)
    return tbl


def add_text(slide, text, left, top, width, height, size=11, bold=False, color=None):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = line
        _font(run, size, bold, color)
    return tb


def title(slide, text):
    add_text(slide, text, 0.5, 0.3, 12, 0.7, size=20, bold=True, color="1F3864")


def build_pptx(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    s = prs.slides.add_slide(blank)
    add_text(s, "재무실사 보고서 (Draft)", 0.8, 2.3, 11, 1, size=32, bold=True, color="1F3864")
    add_text(s, CO + "\n실사 기준일: 2026년 6월 30일\n실사 대상기간: FY2023 ~ 2026년 상반기", 0.8, 3.5, 11, 1.5, size=16)
    add_text(s, "본 자료의 회사·숫자는 풋팅 스킬 예시를 위해 지어낸 것입니다.", 0.8, 6.5, 11, 0.5, size=11, color="808080")

    s = prs.slides.add_slide(blank)
    title(s, "1. 요약 (Executive Summary)")
    add_text(s, (
        "• FY2025 매출은 전년 대비 17.2% 성장한 24,560백만원이며, 조정 EBITDA 는 3,010백만원입니다.\n"
        "• 주요 거래처(A유통·B마트·C식자재) 매출 인식에 특이사항은 없습니다.\n"
        "• 실사 기준일 현재 리스부채 포함 순차입금은 2,210백만원입니다.\n"
        "• 순운전 자본은 매출 증가에 따라 FY2023 2,520백만원에서 FY2025 3,120백만원으로 증가했습니다."
    ), 0.6, 1.2, 12, 3, size=13)

    s = prs.slides.add_slide(blank)
    title(s, "2. 요약 손익")
    add_text(s, "(단위: 백만원)", 10.5, 1.0, 2.5, 0.4, size=10)
    rows = [["구분"] + PERIODS]
    for name, v in IS:
        rows.append([name] + list(v))
    rows[3][0] = "매출 총이익"
    rows[1][3] = 24650                                     # E1
    add_table(s, rows, 0.6, 1.4, 12, 3.2, col_widths=[4, 2, 2, 2, 2])
    add_text(s, "• 매출은 B2B 식자재 유통 확대로 3개년 연평균 15.5% 성장\n• 매출총이익률 FY2023 30.0% → FY2025 31.1%",
             0.6, 4.9, 12, 1.2, size=12)

    s = prs.slides.add_slide(blank)
    title(s, "3. 이익의 질 (조정 EBITDA)")
    add_text(s, "(단위: 백만원)", 10.5, 1.0, 2.5, 0.4, size=10)
    rows = [["구분"] + PERIODS]
    for name, v, _ in QOE:
        rows.append([name] + list(v))
    rows[-1][2] = 2900                                     # E2
    add_table(s, rows, 0.6, 1.4, 12, 3.2, col_widths=[5, 1.75, 1.75, 1.75, 1.75])
    add_text(s, "• 일회성 소송합의금(FY2024)과 비경상 수익(정부보조금·재고평가손실 환입)을 제거\n"
                "• 대표이사 소유 창고 임차료는 시가(월 10백만원) 초과분을 차감", 0.6, 4.9, 12, 1.2, size=12)

    s = prs.slides.add_slide(blank)
    title(s, "4. 순운전자본")
    rows = [["구분"] + PERIODS]                           # E8: 단위·기준일 없음
    for name, v in NWC:
        rows.append([name] + list(v))
    rows[-1][3] = 3120                                     # E6
    add_table(s, rows, 0.6, 1.4, 12, 2.6, col_widths=[4, 2, 2, 2, 2])
    add_text(s, "• 매출채권 회전일수 FY2023 46일 → FY2025 45일로 안정적\n• 재고자산은 원재료 선매입으로 증가 추세",
             0.6, 4.3, 12, 1.2, size=12)

    s = prs.slides.add_slide(blank)
    title(s, "5. 순차입금 (2025년 12월 31일 현재)")
    add_text(s, "(단위: 백만원)", 10.5, 1.0, 2.5, 0.4, size=10)
    rows = [["구분", "금액", "비고"]]
    for name, v, memo in DEBT:
        if name == "리스부채":
            continue                                       # E4: 리스부채 누락
        rows.append([name, 2210 if name == "순차입금" else v, memo if name != "순차입금" else "리스부채 포함"])
    add_table(s, rows, 0.6, 1.4, 12, 2.4, col_widths=[4, 2, 6], num_cols={1})
    add_text(s, "• 단기차입금은 2026년 6월 만기 도래 — 차환 조건 확인 필요", 0.6, 4.1, 12, 0.8, size=12)

    s = prs.slides.add_slide(blank)
    title(s, "6. 매출처별 매출 (FY2025)")
    add_text(s, "(단위: 백만원)", 10.5, 1.0, 2.5, 0.4, size=10)
    rows = [["거래처", "매출", "비중"]]
    for name, v, _, _ in CUST:
        rows.append([name, v, "%.1f%%" % (v / 24560 * 100)])
    add_table(s, rows, 0.6, 1.4, 8, 2.6, col_widths=[3, 2.5, 2.5], num_cols={1, 2})
    add_text(s, "• 상위 3개 거래처 비중 80.0%\n• 주요 거래처 매출 인식에 특이사항 없음 (E5: 조서는 B마트 확인필요)".replace(" (E5: 조서는 B마트 확인필요)", ""),
             0.6, 4.3, 12, 1.2, size=12)
    prs.save(path)


# ──────────────────────────────────────────────── 조서 xlsx
F = "맑은 고딕"
HDR = PatternFill("solid", fgColor="DCE6F1")
FLAG = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
BD = Border(left=thin, right=thin, top=thin, bottom=thin)


def sheet(wb, name, header, rows, widths, note=None):
    ws = wb.create_sheet(name)
    ws.append([name + " — " + CO + " (단위: 백만원)"])
    ws["A1"].font = Font(name=F, size=11, bold=True)
    if note:
        ws.append([note])
        ws["A2"].font = Font(name=F, size=9, color="808080")
    ws.append([])
    ws.append(header)
    hr = ws.max_row
    for c in range(1, len(header) + 1):
        cell = ws.cell(hr, c)
        cell.font, cell.fill, cell.border = Font(name=F, size=10, bold=True), HDR, BD
        cell.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append(row)
        r = ws.max_row
        for c in range(1, len(header) + 1):
            cell = ws.cell(r, c)
            cell.font, cell.border = Font(name=F, size=10), BD
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0;(#,##0)"
            if cell.value == "확인필요":
                cell.fill = FLAG
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    return ws


def build_xlsx(path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheet(wb, "IS_요약", ["계정"] + PERIODS + ["출처"],
          [[n] + v + ["계정별원장 집계"] for n, v in IS], [26, 12, 12, 12, 12, 20],
          note="총계정원장 → 계정별 집계. 2026 1H 는 미감사 가결산.")
    sheet(wb, "QoE조정", ["조정 항목"] + PERIODS + ["근거"],
          [[n] + v + [b] for n, v, b in QOE], [34, 12, 12, 12, 12, 60])
    sheet(wb, "NWC", ["항목"] + PERIODS + ["출처"],
          [[n] + v + ["기말 시산표"] for n, v in NWC], [20, 12, 12, 12, 12, 16],
          note="기준일: 각 회계기간 말일 (2026 1H 는 2026-06-30). 순운전자본 = 매출채권 + 재고자산 − 매입채무 − 기타유동부채.")
    sheet(wb, "순차입금", ["항목", "FY2025말", "비고"],
          [[n, v, m] for n, v, m in DEBT], [26, 14, 44],
          note="차입금 잔액은 금융기관 조회서, 리스부채는 리스 명세와 대사.")
    sheet(wb, "매출처별", ["거래처", "FY2025 매출", "확인상태", "비고"],
          [[n, v, st, memo] for n, v, st, memo in CUST], [16, 14, 12, 70],
          note="매출처별 집계는 매출원장. 확인상태 '확인필요' 는 미결 항목.")
    wb.save(path)


if __name__ == "__main__":
    p1 = os.path.join(HERE, "FDD보고서_한결식품(가상)_draft.pptx")
    p2 = os.path.join(HERE, "wp_한결식품(가상)_실사조서.xlsx")
    build_pptx(p1)
    build_xlsx(p2)
    print("생성:", p1)
    print("생성:", p2)
    print(PLANTED)
