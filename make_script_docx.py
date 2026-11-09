# -*- coding: utf-8 -*-
"""발표 원고 Word 파일 생성 (텍스트는 make_slide.py의 NOTES와 동일)"""
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


d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, side, Cm(2.5))

p = d.add_paragraph()
font(p.add_run("발표 원고 – 05. 글로벌 밸류체인 해외 진출"), 18, True, "14213D")
p = d.add_paragraph()
font(p.add_run("모빌리티 & 미래 자동차 · 최홍욱"), 11, False, "5A6472")
p = d.add_paragraph()
font(p.add_run("슬라이드: " + ns["TITLE"]), 11, False, "5A6472")
p.paragraph_format.space_after = Pt(18)

for line in ns["NOTES"].split("\n"):
    p = d.add_paragraph()
    p.paragraph_format.line_spacing = 1.6
    p.paragraph_format.space_after = Pt(10)
    font(p.add_run(line), 12)

d.save(OUT)
print("saved", OUT)
