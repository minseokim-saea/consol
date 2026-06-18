# -*- coding: utf-8 -*-
"""Google OTP(2단계 인증) 등록 안내 워드 문서 생성"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2D, 0x5A, 0xA0)
GRAY = RGBColor(0x6C, 0x75, 0x7D)
DARK = RGBColor(0x21, 0x25, 0x29)
KFONT = "맑은 고딕"

doc = Document()

# ---- 기본 스타일/폰트 ----
def set_font(run, size=11, bold=False, color=DARK, font=KFONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), font)
    rfonts.set(qn('w:hAnsi'), font)
    rfonts.set(qn('w:eastAsia'), font)

normal = doc.styles['Normal']
normal.font.name = KFONT
normal.font.size = Pt(11)
normal.element.rPr.rFonts.set(qn('w:eastAsia'), KFONT)

# 페이지 여백
sec = doc.sections[0]
sec.top_margin = Cm(2.2); sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.3); sec.right_margin = Cm(2.3)

def shade(cell, hex_color):
    tcpr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), hex_color)
    tcpr.append(sh)

def set_borders(cell, color="CCCCCC", sz="4", style="single"):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top','left','bottom','right'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), style); e.set(qn('w:sz'), sz)
        e.set(qn('w:space'), '0'); e.set(qn('w:color'), color)
        borders.append(e)
    tcpr.append(borders)

def heading(text, size=15, color=NAVY, space_before=14, space_after=6, bar=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    set_font(r, size=size, bold=True, color=color)
    if bar:
        # 하단 테두리(밑줄 바)
        ppr = p._p.get_or_add_pPr()
        pbdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '12')
        bottom.set(qn('w:space'), '4'); bottom.set(qn('w:color'), '2D5AA0')
        pbdr.append(bottom); ppr.append(pbdr)
    return p

def para(runs, align=None, space_after=4, space_before=0, indent=None):
    """runs: list of (text, dict) tuples"""
    p = doc.add_paragraph()
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if indent: p.paragraph_format.left_indent = Cm(indent)
    for text, kw in runs:
        r = p.add_run(text)
        set_font(r, **kw)
    return p

def screenshot_box(label, desc):
    """캡쳐 자리 표시 박스"""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Cm(15.5)
    shade(cell, "F4F7FB")
    set_borders(cell, color="9DB8DA", sz="6", style="dashed")
    # 첫 줄
    p0 = cell.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(14)
    p0.paragraph_format.space_after = Pt(2)
    r0 = p0.add_run(label)
    set_font(r0, size=11, bold=True, color=NAVY)
    # 둘째 줄
    p1 = cell.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_after = Pt(14)
    r1 = p1.add_run(desc)
    set_font(r1, size=9.5, bold=False, color=GRAY)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def note_box(lines, fill="FFF8E1", border="E0C200"):
    """안내/주의 박스"""
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    cell.width = Cm(15.5)
    shade(cell, fill)
    set_borders(cell, color=border, sz="4", style="single")
    for i, (text, kw) in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_before = Pt(4 if i == 0 else 1)
        p.paragraph_format.space_after = Pt(4 if i == len(lines)-1 else 1)
        r = p.add_run(text)
        set_font(r, **kw)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ===================== 표지 제목 =====================
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_after = Pt(2)
set_font(title.add_run("Google OTP(2단계 인증) 등록 안내"), size=22, bold=True, color=NAVY)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.paragraph_format.space_after = Pt(10)
set_font(sub.add_run("연결 재무보고 통합 시스템"), size=12, bold=False, color=BLUE)

# 인트로
para([
    ("연결 재무보고 통합 시스템은 보안을 위해 ", {}),
    ("2단계 인증(OTP)", {"bold": True, "color": NAVY}),
    ("을 사용합니다. 아이디·비밀번호 외에, 휴대폰 앱에 표시되는 ", {}),
    ("6자리 숫자", {"bold": True, "color": NAVY}),
    ("를 한 번 더 입력해야 로그인됩니다.", {}),
], space_after=4)
para([
    ("최초 1회만 등록", {"bold": True, "color": NAVY}),
    ("하면 되고, 이후에는 로그인할 때 코드만 입력하면 됩니다.", {}),
], space_after=6)

note_box([
    ("소요 시간: 약 2~3분   |   준비물: 스마트폰", {"size": 10.5, "bold": True, "color": NAVY}),
], fill="EAF1FB", border="9DB8DA")

# ===================== 1단계 =====================
heading("1단계.  휴대폰에 인증 앱 설치하기")
para([("스마트폰에서 아래 앱 중 하나를 설치하세요. ", {}),
      ("(Google Authenticator 권장)", {"bold": True, "color": NAVY})], space_after=6)

# 표
t = doc.add_table(rows=3, cols=2)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
widths = [Cm(4.5), Cm(11.0)]
rows_data = [
    ("휴대폰 종류", "설치 방법", True),
    ("아이폰(iPhone)", "App Store 에서  「Google Authenticator」  검색 후 설치", False),
    ("안드로이드", "Play 스토어 에서  「Google Authenticator」  검색 후 설치", False),
]
for ri, (c1, c2, head) in enumerate(rows_data):
    for ci, txt in enumerate((c1, c2)):
        cell = t.cell(ri, ci)
        cell.width = widths[ci]
        set_borders(cell, color="C9D6E5", sz="4")
        if head:
            shade(cell, "1F3864")
        elif ri % 2 == 0:
            shade(cell, "F4F7FB")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(3)
        if ci == 0:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        set_font(r, size=10.5, bold=head or ci == 0,
                 color=RGBColor(0xFF,0xFF,0xFF) if head else DARK)
doc.add_paragraph().paragraph_format.space_after = Pt(2)

note_box([
    ("Microsoft Authenticator, Authy 앱도 똑같이 사용할 수 있습니다.",
     {"size": 10, "color": GRAY}),
], fill="F4F7FB", border="C9D6E5")

screenshot_box("📷 캡쳐 ①", "앱스토어에서 Google Authenticator 를 검색한 화면")

# ===================== 2단계 =====================
heading("2단계.  시스템에서 등록 화면 열기")
para([
    ("PC에서 시스템에 ", {}),
    ("아이디·비밀번호로 로그인", {"bold": True, "color": NAVY}),
    ("하면, 아직 OTP를 등록하지 않은 경우 아래와 같은 ", {}),
    ("「2단계 인증(OTP) 등록」", {"bold": True, "color": NAVY}),
    (" 화면이 자동으로 나타납니다.", {}),
], space_after=6)
note_box([
    ("“보안 정책에 따라 OTP 등록 후 시스템을 이용할 수 있습니다” 라는 안내가 보이면 정상입니다.",
     {"size": 10, "color": DARK}),
    ("등록을 마쳐야 다음 화면으로 넘어갑니다.", {"size": 10, "color": GRAY}),
])
screenshot_box("📷 캡쳐 ②", "시스템의 「2단계 인증(OTP) 등록」 화면 — QR 코드가 보이는 화면")

# ===================== 3단계 =====================
heading("3단계.  앱으로 QR 코드 스캔하기")
steps3 = [
    ("①  휴대폰에서 ", "Google Authenticator 앱", " 을 엽니다."),
    ("②  화면 아래(또는 오른쪽 아래)의 ", "＋ (코드 추가 / QR 코드 스캔)", " 버튼을 누릅니다."),
    ("③  ", "「QR 코드 스캔」", " 을 선택합니다."),
    ("④  휴대폰 카메라로 PC 화면의 ", "QR 코드", " 를 비춥니다."),
]
for a, b, c in steps3:
    para([(a, {}), (b, {"bold": True, "color": NAVY}), (c, {})], space_after=3, indent=0.3)
para([
    ("→ 스캔되면 앱에 ", {}),
    ("“연결재무보고시스템”", {"bold": True, "color": NAVY}),
    (" 항목이 생기고, ", {}),
    ("6자리 숫자", {"bold": True, "color": NAVY}),
    ("가 표시됩니다.", {}),
], space_before=4, space_after=6, indent=0.3)

screenshot_box("📷 캡쳐 ③", "앱에서 ＋ 버튼 → 「QR 코드 스캔」 을 누르는 화면")
screenshot_box("📷 캡쳐 ④", "스캔 후 6자리 코드가 표시된 앱 화면")

note_box([
    ("QR 코드가 스캔되지 않을 때", {"size": 10.5, "bold": True, "color": RGBColor(0x8A,0x6D,0x00)}),
    ("등록 화면의 QR 코드 아래에 있는 키(영문·숫자 문자열)를 앱에서 직접 입력해도 됩니다.",
     {"size": 10, "color": DARK}),
    ("앱에서  ＋ → 「설정 키 입력」  을 선택한 뒤, 계정 이름은 아무거나, 키 칸에 그 문자열을 입력하세요.",
     {"size": 10, "color": DARK}),
])

# ===================== 4단계 =====================
heading("4단계.  6자리 코드 입력해서 등록 완료")
steps4 = [
    ("①  앱에 표시된 ", "6자리 숫자", " 를 확인합니다."),
    ("②  PC 등록 화면의 입력칸에 그 ", "6자리 숫자", " 를 입력합니다."),
    ("③  ", "[등록 완료]", " 버튼을 누릅니다."),
]
for a, b, c in steps4:
    para([(a, {}), (b, {"bold": True, "color": NAVY}), (c, {})], space_after=3, indent=0.3)
para([
    ("✓ ", {"bold": True, "color": RGBColor(0x1E,0x7E,0x34)}),
    ("“등록 완료”되면 메인 화면으로 들어갑니다. 이제 등록이 끝났습니다!",
     {"bold": True, "color": RGBColor(0x1E,0x7E,0x34)}),
], space_before=4, space_after=6, indent=0.3)

screenshot_box("📷 캡쳐 ⑤", "6자리 코드를 입력하고 [등록 완료] 를 누르는 화면")

note_box([
    ("코드는 30초마다 자동으로 바뀝니다.", {"size": 10.5, "bold": True, "color": RGBColor(0x8A,0x6D,0x00)}),
    ("입력 중 숫자가 바뀌면, 앱에 보이는 최신 6자리로 다시 입력하세요.", {"size": 10, "color": DARK}),
])

# ===================== 다음 로그인부터 =====================
heading("다음 로그인부터는?")
para([
    ("앞으로 로그인할 때마다 아이디·비밀번호를 입력한 뒤, ", {}),
    ("「2단계 인증」", {"bold": True, "color": NAVY}),
    (" 화면이 나오면 앱을 열어 그때그때 표시되는 ", {}),
    ("6자리 코드", {"bold": True, "color": NAVY}),
    ("를 입력하면 됩니다. ", {}),
    ("(재등록 불필요)", {"color": GRAY}),
], space_after=6)
screenshot_box("📷 캡쳐 ⑥", "로그인 시 6자리 코드를 입력하는 「2단계 인증」 화면")

# ===================== FAQ =====================
heading("자주 묻는 질문")
faqs = [
    ("Q.  휴대폰을 바꿨어요 / 앱을 지웠어요 / 코드가 계속 틀려요",
     "관리자에게 “2단계 인증(2FA) 초기화”를 요청하세요. 초기화 후 1~4단계를 다시 등록하면 됩니다."),
    ("Q.  코드를 입력해도 “올바르지 않습니다”라고 나와요",
     "휴대폰 시간이 자동(네트워크 시간)으로 맞춰져 있는지 확인하세요. 시간이 어긋나면 코드가 틀릴 수 있습니다. "
     "그리고 현재 앱에 떠 있는 최신 6자리인지 다시 확인하세요."),
    ("Q.  로그인 코드 입력 화면이 사라졌어요",
     "코드 입력은 5분 이내에 마쳐야 합니다. 시간이 지나면 로그인부터 다시 하세요."),
]
for q, a in faqs:
    para([(q, {"bold": True, "color": NAVY, "size": 10.5})], space_before=5, space_after=2)
    para([("    " + a, {"size": 10.5, "color": DARK})], space_after=3)

# 푸터
doc.add_paragraph().paragraph_format.space_after = Pt(6)
foot = doc.add_paragraph()
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(foot.add_run("문의: 시스템 관리자"), size=9, color=GRAY)

out = r"C:\패키지프로그램\Google_OTP_등록안내.docx"
doc.save(out)
print("saved:", out)
