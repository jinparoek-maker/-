# -*- coding: utf-8 -*-
"""발표 원고 Word 파일 생성 (약 1분 축약본; 슬라이드 노트의 전체 원고에서 문장만 줄임)"""
import ast
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn

src = open("make_slide.py", encoding="utf-8").read()
ns = {}
for node in ast.parse(src).body:  # make_slide.py를 실행하지 않고 상수만 읽기
    if isinstance(node, ast.Assign) and getattr(node.targets[0], "id", "") in ("NOTES", "TITLE"):
        ns[node.targets[0].id] = ast.literal_eval(node.value)

FONT = "Noto Sans KR"  # 없으면 "Malgun Gothic"
OUT = "05_글로벌밸류체인_해외진출_최홍욱_발표원고.docx"


def font(run, size, bold=False, color="2D3748"):
    run.font.name = FONT
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)


SCRIPT = """제 주제는 중국 자동차 산업의 글로벌 밸류체인과 해외 진출입니다.
2025년 중국의 자동차 수출은 709만 8천 대로 3년 연속 세계 1위였고, 올해는 1~9월에만 820만 대를 넘었습니다.
다만 수출 시장의 순위는 크게 바뀌었습니다. 멕시코는 올해 1월 관세를 50%로 올린 뒤 6위로 내려갔고, 브라질도 7월부터 관세를 35%로 올렸습니다.
이런 관세에 대응해 중국 기업들은 현지 공장을 늘리고 있습니다. CATL은 9월 헝가리에서 100GWh 규모 배터리 공장을 가동했고, BYD는 브라질 공장에서 9개월 만에 10만 대를 생산했습니다.
정리하면, 수출은 계속 늘고 있고, 관세가 높은 시장에서는 현지 생산이 빠르게 늘고 있습니다. 감사합니다."""

d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, side, Cm(2.5))

p = d.add_paragraph()
font(p.add_run("발표 원고 – 05. 글로벌 밸류체인 해외 진출"), 18, True, "14213D")
p = d.add_paragraph()
font(p.add_run("모빌리티 & 미래 자동차 · 최홍욱 · 약 1분"), 11, False, "5A6472")
p = d.add_paragraph()
font(p.add_run("슬라이드: " + ns["TITLE"]), 11, False, "5A6472")
p.paragraph_format.space_after = Pt(18)

for line in SCRIPT.split("\n"):
    p = d.add_paragraph()
    p.paragraph_format.line_spacing = 1.6
    p.paragraph_format.space_after = Pt(10)
    font(p.add_run(line), 12)

d.save(OUT)
print("saved", OUT)
