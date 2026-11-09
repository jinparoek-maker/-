# -*- coding: utf-8 -*-
"""05_글로벌밸류체인_해외진출_최홍욱.pptx 생성 스크립트 (python-pptx)"""
import re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
from PIL import ImageFont

OUT = "05_글로벌밸류체인_해외진출_최홍욱.pptx"

BG, NAVY, RED = "F7F6F2", "14213D", "B3122E"
BODY, SRC, YELLOW = "2D3748", "5A6472", "F2C14E"
KR, SC = "Noto Sans KR", "Noto Sans SC"   # 없으면 "Malgun Gothic", "Microsoft YaHei"로 교체

# 텍스트 폭 측정용 (렌더 환경의 Noto Sans CJK)
_FONT_BOLD = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"


def text_width_in(text, size_pt, path=_FONT_BOLD):
    f = ImageFont.truetype(path, int(size_pt * 10), index=1)  # index 1 = KR
    return f.getlength(text) / 10 / 72


def rgb(h):
    return RGBColor.from_string(h)


HAN = re.compile(r"[一-鿿㐀-䶿]")
HANGUL = re.compile(r"[가-힯ᄀ-ᇿ㄰-㆏]")


def split_scripts(text):
    """한자 구간 → SC 폰트, 그 외 → KR 폰트로 나눈 (text, font) 목록"""
    out, cur, cur_font = [], "", None
    for ch in text:
        if HAN.match(ch):
            f = SC
        elif HANGUL.match(ch):
            f = KR
        else:
            f = cur_font or KR
        if f != cur_font and cur:
            out.append((cur, cur_font))
            cur = ""
        cur_font = f
        cur += ch
    if cur:
        out.append((cur, cur_font))
    return out


def style_run(run, size, color, bold=False, font=KR):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", "zh-CN" if font == SC else "ko-KR")
    rPr.set("altLang", "en-US")
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)


def add_box(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, name=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    return tb, tf


def add_runs(p, parts, size, color, bold=False):
    """parts: [(text, color|None, bold|None)]; 한자는 자동으로 SC 폰트"""
    for text, c, b in parts:
        for seg, font in split_scripts(text):
            r = p.add_run()
            r.text = seg
            style_run(r, size, c or color, bold if b is None else b, font)


def set_bullet(p, color=NAVY, marL=0.17, space_after=5):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Emu(Inches(marL))))
    pPr.set("indent", str(-Emu(Inches(marL))))
    for tag in ("a:buClr", "a:buFont", "a:buChar", "a:buNone"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(buClr, qn("a:srgbClr")).set("val", color)
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", "•")
    p.space_after = Pt(space_after)


def rect(slide, x, y, w, h, color, name=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(color)
    s.line.fill.background()
    s.shadow.inherit = False
    if name:
        s.name = name
    return s


# ---------------------------------------------------------------- 내용
KICKER = "모빌리티 & 미래 자동차 · 05. 글로벌 밸류체인 해외 진출 · 2025–2026 최신 동향 · 최홍욱"
TITLE = "중국 자동차 산업, 수출 확대와 함께 현지 생산으로 무게중심 이동"

COLUMNS = [
    {
        "title": "① 수출 규모: 3년 연속 세계 1위",
        "num": "709.8만 대",
        "label": "2025년 수출, +21.1%",
        "inset": 0.15,
        "bullets": [
            "2025년 신에너지차(NEV) 수출 261.5만 대로 전년 대비 2배",
            "2026년 1~9월 820.5만 대(+65.8%). 6월부터 4개월 연속 월 100만 대 돌파, NEV 비중 50% 이상",
            "중국자동차공업협회(CAAM): 2026년 연간 1,000만 대 돌파 전망",
            "BYD 수출 105.4만 대(+140%), 동시에 해외 공장 확대",
        ],
    },
    {
        "title": "② 시장 재편과 무역장벽",
        "num": "50%",
        "label": "멕시코의 중국산 완성차 관세(2026.1~)",
        "bullets": [
            "멕시코: 2025년 최대 수출국(62.5만 대) → 관세 20%→50% 인상 후 2026 상반기 6위",
            "러시아: 2025년 폐차세 인상 등으로 −46% → 2026 상반기 1위 복귀",
            "브라질: 2026 상반기 NEV 최대 시장(약 30만 대), 7월부터 전동화차 수입관세 35%",
            "EU: 2026.1 상계관세(최대 35.3%)를 ‘최저가격 약속’으로 대체하는 지침 발표 (첫 승인은 폭스바겐 안후이 1개 차종)",
        ],
    },
    {
        "title": "③ 관세 대응: 생산기지·배터리 현지화",
        "num": "100GWh",
        "label": "CATL 헝가리 공장",
        "bullets": [
            "CATL 데브레첸 셀 공장 2026.9 가동(73.4억 유로, 中 기업 해외 최대 배터리 기지). 독일·헝가리·스페인 합산 유럽 생산능력 (계획 포함) 160GWh 이상",
            "BYD: 브라질 공장 2025.10 가동 후 9개월 만에 10만 대 생산, 헝가리 공장은 2026년 4분기 조립 개시 예정",
            "브라질에서만 최소 8개 중국 브랜드가 현지 생산 추진",
        ],
    },
]

IMPLICATION_HEAD = "시사점:"
IMPLICATION_BODY = (" 멕시코·브라질·EU가 관세를 올리자 BYD·CATL 등은 현지 공장으로 대응했다. "
                    "수출은 여전히 늘고 있어, 당분간 ‘수출+현지 생산’ 병행이 이어질 전망이다.")

SOURCE = ("출처: 中国汽车工业协会(CAAM) 발표 – 界面新闻(2026.1.14), 每日经济新闻(2026.10.12) · "
          "时代财经(2026.1.16) · 乘联分会 崔东树 분석 – 中国汽车报(2026.8.3), 第一财经 · "
          "汽车之家(2025.12) · 凤凰网汽车(2026) · Caixin Global(2026.1.13) · CnEVPost(2026.2.11) · "
          "IT之家/Reuters(2026.6.10) · 网易汽车(2026) · 观察者网·科技日报(2026.9.29) · 36氪(2026.11.2)")

NOTES = """제 주제는 중국 자동차 산업의 글로벌 밸류체인과 해외 진출입니다.
먼저 규모를 보면, 2025년 중국의 자동차 수출은 709만 8천 대로 3년 연속 세계 1위였습니다. 신에너지차 수출은 261만 5천 대로 1년 만에 두 배가 됐습니다. 올해는 1~9월에만 820만 대를 넘었고, 6월부터는 매달 100만 대 이상 수출하고 있습니다. 중국자동차공업협회는 올해 연간 1,000만 대를 넘을 것으로 봅니다.
다만 수출 시장의 순위는 크게 바뀌었습니다. 작년 최대 수출국이던 멕시코는 올해 1월 관세를 50%로 올린 뒤 6위로 내려갔고, 러시아가 다시 1위가 됐습니다. 브라질은 신에너지차 최대 시장이 됐지만 7월부터 관세를 35%로 올렸습니다. EU는 고율 관세 대신 최저가격 약속을 받는 길을 열었는데, 아직 승인된 사례는 폭스바겐 안후이의 1개 차종입니다.
이런 관세에 대응해 중국 기업들은 현지 공장을 늘리고 있습니다. CATL은 9월 헝가리에서 100GWh 규모 배터리 공장을 가동했고, BYD는 브라질 공장에서 9개월 만에 10만 대를 생산했습니다.
정리하면, 수출은 계속 늘고 있고, 동시에 관세가 높은 시장에서는 현지 생산이 빠르게 늘고 있습니다. 감사합니다."""

# ---------------------------------------------------------------- 레이아웃
SW, SH = 13.333, 7.5
MX = 0.48
CW_ALL = SW - 2 * MX
GAP = 0.32
CW = (CW_ALL - 2 * GAP) / 3

Y_KICKER, Y_TITLE = 0.38, 0.68
Y_COL = 1.42
NUM_PT, BODY_PT = 28, 13
Y_BAR, H_BAR = 5.86, 0.66
Y_SRC = 6.66

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.background.fill.solid()
slide.background.fill.fore_color.rgb = rgb(BG)

# 소제목 + 제목
_, tf = add_box(slide, MX, Y_KICKER, CW_ALL, 0.26, name="Kicker")
add_runs(tf.paragraphs[0], [(KICKER, None, None)], 12, RED, True)
_, tf = add_box(slide, MX, Y_TITLE, CW_ALL, 0.5, name="Title")
add_runs(tf.paragraphs[0], [(TITLE, None, None)], 26, NAVY, True)

# 3단
for i, col in enumerate(COLUMNS):
    x = MX + i * (CW + GAP)
    rect(slide, x, Y_COL, CW, 0.07, NAVY, name=f"Col{i+1} Rule")
    _, tf = add_box(slide, x, Y_COL + 0.17, CW, 0.32, name=f"Col{i+1} Title")
    add_runs(tf.paragraphs[0], [(col["title"], None, None)], 15, NAVY, True)

    y_num = Y_COL + 0.58
    num_w = text_width_in(col["num"], NUM_PT) + 0.25
    _, tf = add_box(slide, x, y_num, num_w, 0.55, anchor=MSO_ANCHOR.BOTTOM, name=f"Col{i+1} Number")
    add_runs(tf.paragraphs[0], [(col["num"], None, None)], NUM_PT, RED, True)
    lx = x + num_w + 0.1
    _, tf = add_box(slide, lx, y_num + 0.05, x + CW - lx, 0.45, anchor=MSO_ANCHOR.BOTTOM,
                    name=f"Col{i+1} Number Label")
    add_runs(tf.paragraphs[0], [(col["label"], None, None)], 12, SRC, False)

    y_b = y_num + 0.72
    bw = CW - col.get("inset", 0)  # 줄바꿈 위치 미세 조정용
    _, tf = add_box(slide, x, y_b, bw, Y_BAR - 0.18 - y_b, name=f"Col{i+1} Bullets")
    for j, b in enumerate(col["bullets"]):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.line_spacing = 1.12
        set_bullet(p, space_after=6)
        add_runs(p, [(b, None, None)], BODY_PT, BODY, False)

# 시사점 바
bar = rect(slide, MX, Y_BAR, CW_ALL, H_BAR, NAVY, name="Implication Bar")
tf = bar.text_frame
tf.word_wrap = True
tf.auto_size = None
tf.margin_left = tf.margin_right = Inches(0.25)
tf.margin_top = tf.margin_bottom = Inches(0.06)
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
add_runs(p, [(IMPLICATION_HEAD, YELLOW, True), (IMPLICATION_BODY, BG, False)], 13.5, BG)

# 출처
_, tf = add_box(slide, MX, Y_SRC, CW_ALL, SH - 0.3 - Y_SRC, name="Sources")
p = tf.paragraphs[0]
p.line_spacing = 1.1
add_runs(p, [(SOURCE, None, None)], 10, SRC, False)
for r in p.runs:  # 자간 -0.2pt: 출처명이 줄 끝에서 끊기지 않도록
    r._r.get_or_add_rPr().set("spc", "-20")

# 발표자 노트
slide.notes_slide.notes_text_frame.text = NOTES

prs.save(OUT)
print("saved", OUT)
