"""
경영진 보고자료 — PPT 를 인쇄한 듯한 가로형(A4) PDF.

페이지마다 함수 하나(_page_xxx)를 두고 PAGES 순서대로 그린다.
지금은 표지만 있다.
"""
import io
import os
import re
from datetime import datetime
from pathlib import Path

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

BASE_DIR = Path(__file__).resolve().parent
ASSET_DIR = BASE_DIR / 'static' / 'report'
COVER_BG = ASSET_DIR / 'cover_bg.png'           # 표지 배경 (원본 표지에서 글자만 지움)
LOGO_COLOR = ASSET_DIR / 'logo_color.png'        # 표지 하단 로고 — 투명 배경
LOGO_WHITE = ASSET_DIR / 'logo_white.png'        # 같은 로고, 흰색 (어두운 바탕용)
WAVE_WHITE = ASSET_DIR / 'wave_white.png'        # 표지 물결 무늬, 흰색 반투명 (어두운 바탕용)

PAGE_W, PAGE_H = landscape(A4)          # 841.9 x 595.3 pt

NAVY = HexColor('#012161')
BRAND = HexColor('#1679BE')             # 로고 파랑
MUTED = HexColor('#8A9BBD')
SUB = HexColor('#5B677A')
RULE = HexColor('#DCE3EE')

REPORTER = '글로벌세아 재무팀'
TITLE = '글로벌세아 연결재무보고'

# 보고서 구성 — 목차와 본문 장 순서의 기준. (키, 제목, 설명)
SECTIONS = [
    ('summary', '연결 실적 요약', '자산·부채·차입금, EBITDA·이자보상배율 등 핵심지표'),
    ('affiliates', '주요 계열사별 실적', '계열사별 매출·영업이익·당기순이익·EBITDA·차입금·현금'),
    ('fs', '연결 재무제표', '재무상태표·손익계산서·현금흐름표 (중분류)'),
    ('holding', '지주비율 현황', '자회사 주식가액 ÷ 자산총계, 지주회사 요건 점검'),
    ('trend', '실적 추이', '분기별 매출·이익·차입금 추이'),
]


def _section(key):
    """(장 번호, 제목) — SECTIONS 순서가 곧 장 번호."""
    for i, (k, title, _d) in enumerate(SECTIONS, start=1):
        if k == key:
            return i, title
    raise KeyError(key)


# ── 글꼴 ──────────────────────────────────────────────────────────────────
# 서버 OS 에 따라 있는 한글 글꼴을 쓴다. 없으면 PDF 표준 한글 CID 글꼴로 대신한다.
_FONT_CANDIDATES = {
    'regular': ['malgun.ttf', 'NanumGothic.ttf', 'NanumBarunGothic.ttf'],
    'bold': ['malgunbd.ttf', 'NanumGothicBold.ttf', 'NanumBarunGothicBold.ttf'],
}
_FONT_DIRS = [
    Path(os.environ.get('WINDIR', r'C:\Windows')) / 'Fonts',
    Path('/usr/share/fonts/truetype/nanum'),
    Path('/usr/share/fonts/nanum'),
    Path.home() / '.fonts',
]
_fonts = None


def _find_font(names):
    for d in _FONT_DIRS:
        for n in names:
            p = d / n
            if p.is_file():
                return p
    return None


def fonts():
    """(regular, bold) 글꼴 이름 — 처음 호출 때 한 번 등록한다."""
    global _fonts
    if _fonts:
        return _fonts
    reg, bold = _find_font(_FONT_CANDIDATES['regular']), _find_font(_FONT_CANDIDATES['bold'])
    if reg:
        pdfmetrics.registerFont(TTFont('KR', str(reg)))
        if bold:
            pdfmetrics.registerFont(TTFont('KR-Bold', str(bold)))
        _fonts = ('KR', 'KR-Bold' if bold else 'KR')
    else:
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont('HYGothic-Medium'))
        _fonts = ('HYGothic-Medium', 'HYGothic-Medium')
    return _fonts


# ── 공통 ──────────────────────────────────────────────────────────────────
def period_label(period):
    """'2026-3Q' → '2026년 3분기'"""
    m = re.match(r'^(\d{4})-(\d)Q$', str(period or ''))
    return f'{m.group(1)}년 {m.group(2)}분기' if m else str(period or '')


def _draw_cover_image(c, path):
    """배경 이미지를 페이지에 꽉 차게(넘치는 쪽은 가운데 기준으로 잘라) 그린다."""
    img = ImageReader(str(path))
    iw, ih = img.getSize()
    scale = max(PAGE_W / iw, PAGE_H / ih)
    w, h = iw * scale, ih * scale
    c.drawImage(img, (PAGE_W - w) / 2, (PAGE_H - h) / 2, w, h)


# ── 페이지 ────────────────────────────────────────────────────────────────
def _page_cover(c, ctx):
    regular, bold = fonts()
    _draw_cover_image(c, COVER_BG)
    cx = PAGE_W / 2

    # 결산기간
    c.setFillColor(NAVY)
    c.setFont(regular, 30)
    c.drawCentredString(cx, PAGE_H * 0.635, period_label(ctx['period']))

    # 제목
    c.setFont(bold, 48)
    c.drawCentredString(cx, PAGE_H * 0.505, TITLE)

    # 작성일 (마감 전이면 잠정치 표시)
    date_line = ctx['generated_at'].strftime('%Y. %m. %d')
    if ctx.get('provisional'):
        date_line += '  ·  잠정치'
    c.setFillColor(MUTED)
    c.setFont(regular, 16)
    c.drawCentredString(cx, PAGE_H * 0.335, date_line)

    # 보고 주체 — 오른쪽 하단 (하단 로고의 밑선에 맞춤)
    c.setFillColor(NAVY)
    c.setFont(bold, 14)
    c.drawRightString(PAGE_W - 48, 50, f'- {REPORTER}')


def _page_number(c, n, color=MUTED):
    regular, _ = fonts()
    c.setFillColor(color)
    c.setFont(regular, 9)
    c.drawRightString(PAGE_W - 40, 26, str(n))


def _page_toc(c, ctx):
    regular, bold = fonts()

    # 왼쪽 남색 패널 — 표지의 물결 무늬를 흐리게 얹는다
    PW = 280
    c.setFillColor(NAVY)
    c.rect(0, 0, PW, PAGE_H, stroke=0, fill=1)
    wave = ImageReader(str(WAVE_WHITE))
    ww, wh = wave.getSize()
    s = PAGE_H / wh
    c.saveState()
    p = c.beginPath()
    p.rect(0, 0, PW, PAGE_H)
    c.clipPath(p, stroke=0, fill=0)
    c.setFillAlpha(0.35)
    c.drawImage(wave, PW - ww * s, 0, ww * s, PAGE_H, mask='auto')
    c.restoreState()

    c.setFillColor(HexColor('#9CC9F0'))
    c.setFont(bold, 11)
    c.drawString(48, PAGE_H - 110, 'C O N T E N T S')
    c.setFillColor(HexColor('#FFFFFF'))
    c.setFont(bold, 38)
    c.drawString(46, PAGE_H - 160, '목차')
    c.setStrokeColor(HexColor('#9CC9F0'))
    c.setLineWidth(1.5)
    c.line(48, PAGE_H - 182, 88, PAGE_H - 182)

    c.setFillColor(HexColor('#C9D8EE'))
    c.setFont(regular, 11)
    c.drawString(48, 92, f'{period_label(ctx["period"])}')
    c.drawString(48, 74, TITLE)
    logo = ImageReader(str(LOGO_WHITE))
    lw, lh = logo.getSize()
    c.drawImage(logo, 48, 34, 92, 92 * lh / lw, mask='auto')

    # 오른쪽 — 장 목록
    x0, x1 = PW + 70, PAGE_W - 70
    top, bottom = PAGE_H - 92, 92
    step = (top - bottom) / len(SECTIONS)
    for i, (_key, title, desc) in enumerate(SECTIONS):
        y = top - step * i                      # 이 칸의 윗선
        mid = y - step / 2
        c.setFillColor(BRAND)
        c.setFont(bold, 26)
        c.drawString(x0, mid - 9, f'{i + 1:02d}')
        c.setFillColor(NAVY)
        c.setFont(bold, 18)
        c.drawString(x0 + 62, mid + 2, title)
        c.setFillColor(SUB)
        c.setFont(regular, 11)
        c.drawString(x0 + 62, mid - 16, desc)
        if i < len(SECTIONS) - 1:
            c.setStrokeColor(RULE)
            c.setLineWidth(0.75)
            c.line(x0, y - step, x1, y - step)

    _page_number(c, ctx['page_no'])


# ── 본문 장 공통 틀 ────────────────────────────────────────────────────────
M = 40                                   # 좌우 여백
BODY_TOP = PAGE_H - 92                   # 제목 아래 본문 시작
BODY_BOTTOM = 54                         # 바닥글 위
UP = HexColor('#D0453A')                 # 증가 (국내 관례: 빨강 ▲)
DOWN = HexColor('#1F6FD1')               # 감소 (파랑 ▼)
ICE = HexColor('#EEF3FA')
TINT = HexColor('#F5F9FE')


def _frame(c, ctx, section_key, unit=None):
    """본문 장 머리글(장 번호·제목·단위)과 바닥글(로고·보고서명·쪽 번호)."""
    regular, bold = fonts()
    section_no, title = _section(section_key)
    y = PAGE_H - 62
    c.setFillColor(BRAND)
    c.setFont(bold, 20)
    c.drawString(M, y, f'{section_no:02d}')
    c.setFillColor(NAVY)
    c.drawString(M + 36, y, title)
    if ctx.get('provisional'):
        tw = pdfmetrics.stringWidth(title, bold, 20)
        _pill(c, M + 36 + tw + 10, y - 1, '잠정치', HexColor('#FDF1E6'), HexColor('#C0661E'))
    if unit:
        c.setFillColor(MUTED)
        c.setFont(regular, 9.5)
        c.drawRightString(PAGE_W - M, y, f'(단위: {unit})')
    c.setStrokeColor(RULE)
    c.setLineWidth(0.75)
    c.line(M, PAGE_H - 76, PAGE_W - M, PAGE_H - 76)
    c.setStrokeColor(NAVY)
    c.setLineWidth(2)
    c.line(M, PAGE_H - 76, M + 64, PAGE_H - 76)

    # 바닥글
    c.setStrokeColor(RULE)
    c.setLineWidth(0.5)
    c.line(M, 40, PAGE_W - M, 40)
    logo = ImageReader(str(LOGO_COLOR))
    lw, lh = logo.getSize()
    c.drawImage(logo, M, 18, 58, 58 * lh / lw, mask='auto')
    c.setFillColor(MUTED)
    c.setFont(regular, 8.5)
    c.drawRightString(PAGE_W - M - 22, 24, f'{TITLE}  |  {period_label(ctx["period"])}')
    c.drawRightString(PAGE_W - M, 24, str(ctx['page_no']))


def _pill(c, x, y, text, fill, color, size=8.5):
    regular, bold = fonts()
    w = pdfmetrics.stringWidth(text, bold, size) + 12
    c.setFillColor(fill)
    c.roundRect(x, y - 3.5, w, size + 7, (size + 7) / 2, stroke=0, fill=1)
    c.setFillColor(color)
    c.setFont(bold, size)
    c.drawString(x + 6, y, text)
    return w


def _no_data(c, text):
    regular, _ = fonts()
    c.setFillColor(SUB)
    c.setFont(regular, 14)
    c.drawCentredString(PAGE_W / 2, (BODY_TOP + BODY_BOTTOM) / 2, text)


# ── 숫자 표기 ─────────────────────────────────────────────────────────────
def _fmt(v, kind):
    if v is None:
        return '–'
    if kind == 'amt':
        e = round(v / 1e8)
        return f'({-e:,.0f})' if e < 0 else f'{e:,.0f}'      # 음수는 괄호 (계열사별 실적 엑셀과 같음)
    if kind == 'pct':
        return f'{v:,.1f}%'
    return f'{v:,.1f}배'


def _delta(cur, prior, kind):
    """(증감 문자열, 증감률 문자열, 부호) — 비율 지표는 %p·배 차이만."""
    if cur is None or prior is None:
        return '–', '–', 0
    d = cur - prior
    if kind == 'amt':
        d_r = round(d / 1e8)
        sign = (d_r > 0) - (d_r < 0)
        txt = f'{abs(d_r):,.0f}'
        rate = f'{d / abs(prior) * 100:+.1f}%' if prior else '–'
    else:
        d_r = round(d, 1)
        sign = (d_r > 0) - (d_r < 0)
        txt = f'{abs(d_r):,.1f}' + ('%p' if kind == 'pct' else '배')
        rate = '–'
    if sign == 0:
        return '0', rate if kind == 'amt' else '–', 0
    return ('▲ ' if sign > 0 else '▼ ') + txt, rate, sign


# ── 01 연결 실적 요약 ─────────────────────────────────────────────────────
SUMMARY_ROWS = [
    ('재무상태', None, None),
    ('자산총계', 'assets', 'amt'),
    ('부채총계', 'liab', 'amt'),
    ('총차입금', 'debt', 'amt'),
    ('Net Debt', 'net_debt', 'amt'),
    ('자본총계', 'equity', 'amt'),
    ('부채비율', 'debt_ratio', 'pct'),
    ('수익성 · 상환능력', None, None),
    ('EBITDA', 'ebitda', 'amt'),
    ('Net Debt / EBITDA', 'nd_ebitda', 'x'),
    ('영업이익', 'op', 'amt'),
    ('이자비용', 'interest', 'amt'),
    ('이자보상배율', 'icr', 'x'),
]
SUMMARY_KEY_RATIOS = [
    ('부채비율', 'debt_ratio', 'pct', '부채총계 ÷ 자본총계'),
    ('Net Debt / EBITDA', 'nd_ebitda', 'x', 'Net Debt ÷ EBITDA'),
    ('이자보상배율', 'icr', 'x', '영업이익 ÷ 이자비용'),
]


def _page_summary(c, ctx):
    regular, bold = fonts()
    _frame(c, ctx, 'summary', unit='억원')
    s = (ctx.get('data') or {}).get('summary') or {}
    cur, prior = s.get('current'), s.get('prior') or {}
    if s.get('error') or not cur:
        _no_data(c, s.get('error') or f'{period_label(ctx["period"])} 연결 자료가 없습니다.')
        return

    q, year, py = s['quarter'], s['year'], s['prior_year']
    cur_sub = '연환산 · 분기말' if s.get('annualized') else '연간 · 기말'

    # ── 표 ──
    x0 = M
    widths = [150, 92, 92, 84, 62]
    xs = [x0]
    for w in widths:
        xs.append(xs[-1] + w)
    tw = xs[-1] - x0
    hh, rh = 36, 27.5
    top = BODY_TOP + 4

    c.setFillColor(NAVY)
    c.rect(x0, top - hh, tw, hh, stroke=0, fill=1)
    heads = [('구분', None), (f'{py}년', '연간 · 기말'), (f'{year}년 {q}Q', cur_sub), ('증감', None), ('증감률', None)]
    for i, (h1, h2) in enumerate(heads):
        c.setFillColor(HexColor('#FFFFFF'))
        cx = (xs[i] + xs[i + 1]) / 2
        if i == 0:
            c.setFont(bold, 10.5)
            c.drawString(xs[0] + 12, top - hh / 2 - 4, h1)
        elif h2:
            c.setFont(bold, 10.5)
            c.drawCentredString(cx, top - 15, h1)
            c.setFillColor(HexColor('#B9CBE6'))
            c.setFont(regular, 7.5)
            c.drawCentredString(cx, top - 27, h2)
        else:
            c.setFont(bold, 10.5)
            c.drawCentredString(cx, top - hh / 2 - 4, h1)

    y = top - hh
    ratio_keys = {k for _, k, _, _ in SUMMARY_KEY_RATIOS}
    for label, key, kind in SUMMARY_ROWS:
        if key is None:                                   # 묶음 머리
            gh = 20
            c.setFillColor(ICE)
            c.rect(x0, y - gh, tw, gh, stroke=0, fill=1)
            c.setFillColor(NAVY)
            c.setFont(bold, 8.5)
            c.drawString(x0 + 12, y - gh + 6.5, label)
            y -= gh
            continue
        if key in ratio_keys:
            c.setFillColor(TINT)
            c.rect(x0, y - rh, tw, rh, stroke=0, fill=1)
        base = y - rh / 2 - 3.5
        c.setFillColor(NAVY)
        c.setFont(bold if key in ratio_keys else regular, 10.5)
        c.drawString(x0 + 12, base, label)
        c.setFillColor(SUB)
        c.setFont(regular, 10.5)
        c.drawRightString(xs[2] - 12, base, _fmt(prior.get(key), kind))
        c.setFillColor(NAVY)
        c.setFont(bold, 10.5)
        c.drawRightString(xs[3] - 12, base, _fmt(cur.get(key), kind))
        d_txt, rate, sign = _delta(cur.get(key), prior.get(key), kind)
        c.setFillColor(UP if sign > 0 else DOWN if sign < 0 else SUB)
        c.setFont(regular, 9.5)
        c.drawRightString(xs[4] - 12, base, d_txt)
        c.setFillColor(SUB)
        c.drawRightString(xs[5] - 10, base, rate)
        y -= rh
        c.setStrokeColor(RULE)
        c.setLineWidth(0.5)
        c.line(x0, y, x0 + tw, y)
    c.setStrokeColor(NAVY)
    c.setLineWidth(1)
    c.line(x0, y, x0 + tw, y)

    # 설명 주석은 두지 않고, 자료가 빠졌을 때만 알린다
    notes = []
    if not s.get('prior_source'):
        notes.append(f'{py}년 연결 자료가 없어 전년 값을 표시하지 못했습니다.')
    c.setFillColor(MUTED)
    c.setFont(regular, 8)
    ny = y - 20
    for n in notes:
        c.drawString(x0, ny, '※ ' + n)
        ny -= 12.5

    # ── 오른쪽: 핵심 비율 카드 (표와 위아래를 맞춤) ──
    cx0 = x0 + tw + 22
    cw = PAGE_W - M - cx0
    gap = 10
    ch = (top - y - gap * 2) / 3
    for i, (label, key, kind, formula) in enumerate(SUMMARY_KEY_RATIOS):
        ct = top - i * (ch + gap)
        c.setFillColor(HexColor('#FFFFFF'))
        c.setStrokeColor(RULE)
        c.setLineWidth(0.75)
        c.roundRect(cx0, ct - ch, cw, ch, 8, stroke=1, fill=1)
        px = cx0 + 16
        c.setFillColor(NAVY)
        c.setFont(bold, 11.5)
        c.drawString(px, ct - 24, label)
        c.setFillColor(MUTED)
        c.setFont(regular, 8)
        c.drawRightString(cx0 + cw - 16, ct - 24, formula)

        cv, pv = cur.get(key), prior.get(key)
        c.setFillColor(NAVY)
        c.setFont(bold, 26)
        big = _fmt(cv, kind)
        c.drawString(px, ct - 56, big)
        d_txt, _, sign = _delta(cv, pv, kind)
        if d_txt not in ('–',):
            bw = pdfmetrics.stringWidth(big, bold, 26)
            _pill(c, px + bw + 10, ct - 52, f'전년 대비 {d_txt}',
                  HexColor('#FBEAE8') if sign > 0 else HexColor('#E8F0FB') if sign < 0 else ICE,
                  UP if sign > 0 else DOWN if sign < 0 else SUB)

        # 전년 / 당해 막대
        vals = [(f'{py}', pv, HexColor('#C9D3E3')), (f'{year}.{q}Q', cv, BRAND)]
        vmax = max([abs(v) for _, v, _ in vals if v is not None] or [0])
        bx, bw_max = px + 44, cw - 32 - 44 - 50
        for j, (lab, v, col) in enumerate(vals):
            by = ct - ch + 30 - j * 16
            c.setFillColor(SUB)
            c.setFont(regular, 8)
            c.drawString(px, by + 1, lab)
            c.setFillColor(HexColor('#F1F4F8'))
            c.rect(bx, by - 1, bw_max, 9, stroke=0, fill=1)
            if v is not None and vmax:
                c.setFillColor(col)
                c.rect(bx, by - 1, max(bw_max * abs(v) / vmax, 1.5), 9, stroke=0, fill=1)
            c.setFillColor(NAVY if j else SUB)
            c.setFont(bold if j else regular, 8)
            c.drawRightString(cx0 + cw - 16, by + 1, _fmt(v, kind))


# ── 02 주요 계열사별 실적 ─────────────────────────────────────────────────
def _short_period(p):
    """'2026-1Q' → '26.1Q'"""
    m = re.match(r'^(\d{4})-(\d)Q$', str(p or ''))
    return f"{m.group(1)[2:]}.{m.group(2)}Q" if m else str(p or '')


def _page_affiliates(c, ctx):
    regular, bold = fonts()
    _frame(c, ctx, 'affiliates', unit='억원')
    a = (ctx.get('data') or {}).get('affiliates') or {}
    cur = a.get('current')
    if a.get('error') or not cur:
        _no_data(c, a.get('error') or f'{period_label(ctx["period"])} 계열사별 실적 자료가 없습니다.')
        return
    prior = a.get('prior') or {}
    pp, cp = a.get('prior_period'), ctx['period']
    metrics = cur['metrics']
    cols = cur['columns']                         # 표의 행 = 계열사

    x0 = M
    name_w = 88
    vw = (PAGE_W - 2 * M - name_w) / (len(metrics) * 2)
    top = BODY_TOP + 4
    h1, h2 = 21, 17
    rh = 21.5

    # 머리글 — 지표 / (작년 4Q · 해당분기)
    tw = PAGE_W - 2 * M
    c.setFillColor(NAVY)
    c.rect(x0, top - h1 - h2, tw, h1 + h2, stroke=0, fill=1)
    c.setFillColor(HexColor('#FFFFFF'))
    c.setFont(bold, 10)
    c.drawString(x0 + 10, top - (h1 + h2) / 2 - 3.5, '계열사')
    for i, m in enumerate(metrics):
        gx = x0 + name_w + i * 2 * vw
        c.setFillColor(HexColor('#FFFFFF'))
        c.setFont(bold, 10)
        c.drawCentredString(gx + vw, top - 14.5, m['label'])
        c.setFillColor(HexColor('#B9CBE6'))
        c.setFont(regular, 8)
        c.drawCentredString(gx + vw / 2, top - h1 - 12, _short_period(pp) if pp else '전년 4Q')
        c.setFillColor(HexColor('#FFFFFF'))
        c.setFont(bold, 8)
        c.drawCentredString(gx + vw * 1.5, top - h1 - 12, _short_period(cp))
        c.setStrokeColor(HexColor('#3A5689'))
        c.setLineWidth(0.5)
        c.line(gx + 6, top - h1, gx + 2 * vw - 6, top - h1)
        if i:
            c.line(gx, top - h1 - h2 + 3, gx, top - 3)

    y = top - h1 - h2
    body_top = y
    for col in cols:
        is_total, is_etc = col['key'] == 'total', col['key'] == 'etc'
        if is_total:
            c.setFillColor(ICE)
            c.rect(x0, y - rh, tw, rh, stroke=0, fill=1)
            c.setStrokeColor(NAVY)
            c.setLineWidth(0.8)
            c.line(x0, y, x0 + tw, y)
        base = y - rh / 2 - 3.2
        c.setFillColor(SUB if is_etc else NAVY)
        c.setFont(bold if is_total else regular, 9.5)
        c.drawString(x0 + 10, base, '합계 (연결)' if is_total else col['label'])
        cv = cur['values'].get(col['key']) or {}
        pv = ((prior.get('values') or {}).get(col['key'])) if prior else None
        for i, m in enumerate(metrics):
            gx = x0 + name_w + i * 2 * vw
            c.setFillColor(MUTED if is_etc else SUB)
            c.setFont(bold if is_total else regular, 8.5)
            c.drawRightString(gx + vw - 7, base, _fmt(pv.get(m['key']) if pv else None, 'amt'))
            c.setFillColor(SUB if is_etc else NAVY)
            c.setFont(bold, 8.5)
            c.drawRightString(gx + 2 * vw - 7, base, _fmt(cv.get(m['key']), 'amt'))
        y -= rh
        if not is_total:
            c.setStrokeColor(RULE)
            c.setLineWidth(0.4)
            c.line(x0, y, x0 + tw, y)
    c.setStrokeColor(NAVY)
    c.setLineWidth(1)
    c.line(x0, y, x0 + tw, y)
    # 지표 묶음 구분선 + 해당분기 열 옅은 배경은 선 위에 겹치지 않게 얇게
    c.setStrokeColor(RULE)
    c.setLineWidth(0.6)
    for i in range(1, len(metrics)):
        gx = x0 + name_w + i * 2 * vw
        c.line(gx, body_top, gx, y)

    # 설명 주석은 두지 않고, 자료가 빠졌을 때만 알린다
    notes = []
    if not prior:
        notes.append(f'{a.get("prior_note") or "전년 4Q 자료가 없어 표시하지 못했습니다."}')
    missing = [f"{col['label']}({', '.join(col['missing'])})" for col in cols if col.get('missing')]
    if missing:
        notes.append('패키지 미업로드: ' + ' · '.join(missing))
    c.setFillColor(MUTED)
    c.setFont(regular, 7.5)
    ny = y - 15
    for n in notes:
        c.drawString(x0, ny, '※ ' + n)
        ny -= 11


# ── 03 연결 재무제표 ──────────────────────────────────────────────────────
_FS_TOTALS = ('자산총계', '부채총계', '자본총계', '부채 및 자본합계', '영업이익', '당기순이익', '기말의현금')


def _fs_col_labels(stmt, period, fs):
    m = re.match(r'^(\d{4})-(\d)Q$', str(period or ''))
    y, q = (int(m.group(1)), int(m.group(2))) if m else (0, 0)
    cur = f"{str(y)[2:]}.{q}Q" + ('말' if stmt == 'BS' else ' 누적')
    if not fs[stmt].get('has_prior'):
        return cur, '전기'
    if stmt == 'BS':
        return cur, f"{str(y - 1)[2:]}년말"
    if fs.get('prior_mode') == 'same':
        return cur, f"{str(y - 1)[2:]}.{q}Q 누적"
    return cur, f"{str(y - 1)[2:]}년" + ('' if q == 4 else ' 월할')


def _fs_panel(c, x, w, top, rh, stmt, period, fs):
    regular, bold = fonts()
    p = fs[stmt]
    c.setFillColor(NAVY)
    c.rect(x, top - 22, w, 22, stroke=0, fill=1)
    c.setFillColor(HexColor('#FFFFFF'))
    c.setFont(bold, 10.5)
    c.drawString(x + 10, top - 15, p['title'])

    nw = 54                                           # 숫자 열 너비
    r_pri, r_cur = x + w - 8, x + w - 8 - nw
    cur_lab, pri_lab = _fs_col_labels(stmt, period, fs)
    y = top - 22
    c.setFillColor(ICE)
    c.rect(x, y - 17, w, 17, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont(bold, 8)
    c.drawString(x + 10, y - 11.5, '과목')
    c.drawRightString(r_cur, y - 11.5, cur_lab)
    c.setFillColor(SUB)
    c.setFont(regular, 8)
    c.drawRightString(r_pri, y - 11.5, pri_lab)
    y -= 17

    for r in p['rows']:
        name = r['name']
        numbered = bool(re.match(r'^\s*\d+\s*\.', name))          # CF 1.~4.
        sub_level = r['level'] >= 1 or numbered
        is_total = r['level'] == 0 and not numbered and any(t in name for t in _FS_TOTALS)
        if r['kind'] == 'header':
            c.setFillColor(BRAND)
            c.setFont(bold, 8.5)
            c.drawString(x + 10, y - rh / 2 - 3, name)
            y -= rh
            continue
        if is_total:
            c.setFillColor(TINT)
            c.rect(x, y - rh, w, rh, stroke=0, fill=1)
        base = y - rh / 2 - 3
        strong = r['kind'] == 'subtotal' and not sub_level
        c.setFillColor(SUB if sub_level else NAVY)
        nx = x + (20 if sub_level else 10)
        font = bold if strong else regular
        size = 8.5                                     # 긴 과목명은 숫자 칸을 넘지 않게 줄인다
        while size > 6.5 and pdfmetrics.stringWidth(name, font, size) > r_cur - nw + 6 - nx:
            size -= 0.25
        c.setFont(font, size)
        c.drawString(nx, base, name)
        c.setFillColor(SUB if sub_level else NAVY)
        c.setFont(bold if strong else regular, 8.5)
        c.drawRightString(r_cur, base, _fmt(r['cur'], 'amt'))
        c.setFillColor(MUTED if sub_level else SUB)
        c.setFont(regular, 8.5)
        c.drawRightString(r_pri, base, _fmt(r['pri'], 'amt') if p.get('has_prior') else '–')
        y -= rh
        c.setStrokeColor(RULE)
        c.setLineWidth(0.4 if not is_total else 0.7)
        c.line(x, y, x + w, y)
    c.setStrokeColor(NAVY)
    c.setLineWidth(1)
    c.line(x, y, x + w, y)


def _page_fs(c, ctx):
    _frame(c, ctx, 'fs', unit='억원')
    fs = (ctx.get('data') or {}).get('fs') or {}
    if fs.get('error') or not fs.get('BS'):
        _no_data(c, fs.get('error') or f'{period_label(ctx["period"])} 연결 재무제표 자료가 없습니다.')
        return
    top = BODY_TOP + 4
    n = max(len(fs[s]['rows']) for s in ('BS', 'PL', 'CF'))
    rh = min(16.5, (top - 39 - BODY_BOTTOM - 4) / max(n, 1))
    gap = 14
    widths = [236, 236, PAGE_W - 2 * M - 472 - 2 * gap]
    x = M
    for st, w in zip(('BS', 'PL', 'CF'), widths):
        _fs_panel(c, x, w, top, rh, st, ctx['period'], fs)
        x += w + gap


# ── 04 지주비율 현황 ──────────────────────────────────────────────────────
WARN = HexColor('#D9822B')               # 50% 이상 (화면과 같은 주황 계열)
WARN_BG = HexColor('#FDF3E8')


def _eok(v):
    return '–' if v is None else f'{_fmt(v, "amt")}억원'


def _page_holding(c, ctx):
    regular, bold = fonts()
    _frame(c, ctx, 'holding', unit='억원')
    h = (ctx.get('data') or {}).get('holding') or {}
    if h.get('error') or 'ratio' not in h:
        _no_data(c, h.get('error') or f'{period_label(ctx["period"])} 지주비율 자료가 없습니다.')
        return
    ratio, th = h.get('ratio'), h.get('threshold') or 50.0
    over = ratio is not None and ratio >= th
    tone = WARN if over else BRAND
    top = BODY_TOP + 4
    W = PAGE_W - 2 * M

    # ── 핵심지표 + 50% 기준 막대 ──
    ch = 128
    c.setFillColor(HexColor('#FFFFFF'))
    c.setStrokeColor(RULE)
    c.setLineWidth(0.75)
    c.roundRect(M, top - ch, W, ch, 8, stroke=1, fill=1)
    kpis = [('자회사 주식가액 합계', _eok(h.get('equity_total')), '지분법적용투자주식', NAVY),
            ('자산총계', _eok(h.get('total_assets')), f'{h.get("group") or ""} 패키지', NAVY),
            ('지주비율', '–' if ratio is None else f'{ratio:.2f}%',
             '자산총계 없음' if ratio is None else (f'{th:.0f}% 이상 — 지주회사 요건에 해당' if over
                                                  else f'{th:.0f}% 미만'), tone)]
    kw = W / 3
    for i, (lab, val, sub, col) in enumerate(kpis):
        kx = M + 24 + i * kw
        if i:
            c.setStrokeColor(RULE)
            c.line(M + i * kw, top - 18, M + i * kw, top - 74)
        c.setFillColor(SUB)
        c.setFont(regular, 9.5)
        c.drawString(kx, top - 26, lab)
        c.setFillColor(col)
        c.setFont(bold, 24)
        c.drawString(kx, top - 54, val)
        c.setFillColor(col if i == 2 else MUTED)
        c.setFont(bold if i == 2 else regular, 8.5)
        c.drawString(kx, top - 70, sub)
    gx, gw, gy = M + 24, W - 48, top - 100
    c.setFillColor(HexColor('#EEF1F6'))
    c.roundRect(gx, gy, gw, 10, 5, stroke=0, fill=1)
    if ratio is not None:
        c.setFillColor(tone)
        c.roundRect(gx, gy, max(gw * min(ratio, 100) / 100, 10), 10, 5, stroke=0, fill=1)
    mx = gx + gw * th / 100
    c.setStrokeColor(NAVY)
    c.setLineWidth(1.5)
    c.line(mx, gy - 4, mx, gy + 14)
    c.setFillColor(MUTED)
    c.setFont(regular, 8)
    c.drawString(gx, gy - 14, '0%')
    c.drawRightString(gx + gw, gy - 14, '100%')
    c.setFillColor(NAVY)
    c.setFont(bold, 8)
    c.drawCentredString(mx, gy - 14, f'{th:.0f}% · 지주회사 요건')

    # ── 자회사별 주식가액 ──
    ty = top - ch - 18
    tw = 460
    cols = [M, M + 110, M + 230, M + 300, M + tw]          # 자회사 | 억원 | 구성비 | 막대
    hh, rh = 24, 27
    c.setFillColor(NAVY)
    c.rect(M, ty - hh, tw, hh, stroke=0, fill=1)
    c.setFillColor(HexColor('#FFFFFF'))
    c.setFont(bold, 9.5)
    c.drawString(cols[0] + 12, ty - 15.5, '자회사')
    c.drawRightString(cols[2] - 12, ty - 15.5, '지분법적용투자주식')
    c.drawRightString(cols[3] - 10, ty - 15.5, '구성비')
    y = ty - hh
    rows = h.get('rows') or []
    smax = max([r.get('share') or 0 for r in rows] or [0])
    for r in rows:
        base = y - rh / 2 - 3.2
        c.setFillColor(NAVY)
        c.setFont(regular, 9.5)
        c.drawString(cols[0] + 12, base, str(r.get('company') or ''))
        c.drawRightString(cols[2] - 12, base, _fmt(r.get('amount'), 'amt'))
        c.setFillColor(SUB)
        c.drawRightString(cols[3] - 10, base, f'{(r.get("share") or 0):.1f}%')
        bw = (cols[4] - cols[3] - 22) * ((r.get('share') or 0) / smax if smax else 0)
        c.setFillColor(HexColor('#F1F4F8'))
        c.rect(cols[3] + 10, y - rh / 2 - 4, cols[4] - cols[3] - 22, 8, stroke=0, fill=1)
        c.setFillColor(BRAND)
        c.rect(cols[3] + 10, y - rh / 2 - 4, max(bw, 0), 8, stroke=0, fill=1)
        y -= rh
        c.setStrokeColor(RULE)
        c.setLineWidth(0.5)
        c.line(M, y, M + tw, y)
    for lab, val, share, fill in (('합계', h.get('equity_total'), '100.0%', ICE),
                                  ('자산총계', h.get('total_assets'), '', None),
                                  ('지주비율', None, None, TINT)):
        if fill:
            c.setFillColor(fill)
            c.rect(M, y - rh, tw, rh, stroke=0, fill=1)
        base = y - rh / 2 - 3.2
        c.setFillColor(NAVY)
        c.setFont(bold, 9.5)
        c.drawString(cols[0] + 12, base, lab)
        if lab == '지주비율':
            c.setFillColor(tone)
            c.drawRightString(cols[2] - 12, base, '–' if ratio is None else f'{ratio:.2f}%')
        else:
            c.drawRightString(cols[2] - 12, base, _fmt(val, 'amt'))
            c.setFillColor(SUB)
            c.setFont(regular, 9.5)
            c.drawRightString(cols[3] - 10, base, share)
        y -= rh
        c.setStrokeColor(RULE)
        c.setLineWidth(0.5)
        c.line(M, y, M + tw, y)
    c.setStrokeColor(NAVY)
    c.setLineWidth(1)
    c.line(M, y, M + tw, y)
    table_bottom = y

    # ── 50% 요건까지 남은 금액 ──
    g = h.get('gap')
    px, pw = M + tw + 22, W - tw - 22
    ph = ty - table_bottom
    c.setFillColor(WARN_BG if over else TINT)
    c.setStrokeColor(WARN if over else RULE)
    c.setLineWidth(0.75)
    c.roundRect(px, ty - ph, pw, ph, 8, stroke=1, fill=1)
    ix = px + 20
    if not g or ratio is None:
        c.setFillColor(SUB)
        c.setFont(regular, 10)
        c.drawString(ix, ty - 30, '자산총계가 없어 계산할 수 없습니다.')
    else:
        a, e = abs(g.get('asset_increase') or 0), abs(g.get('equity_decrease') or 0)
        if over:
            head, sub = f'지주비율이 {th:.0f}%를 넘었습니다', f'{th:.0f}% 아래로 내리려면 둘 중 하나가 필요합니다.'
            items = [('자산총계를 늘린다면', '+ ' + _eok(a), '자산총계 = 주식가액 × 2 (주식가액 고정)'),
                     ('자회사 주식가액을 줄인다면', '- ' + _eok(e), '주식가액 = 자산총계 ÷ 2 (자산총계 고정)')]
        else:
            head, sub = f'지주비율이 {th:.0f}% 아래입니다', f'{th:.0f}%에 닿기까지 남은 여유입니다.'
            items = [('자산총계 감소 여력', '- ' + _eok(a), '자산총계 = 주식가액 × 2 (주식가액 고정)'),
                     ('자회사 주식가액 증가 여력', '+ ' + _eok(e), '주식가액 = 자산총계 ÷ 2 (자산총계 고정)')]
        c.setFillColor(WARN if over else NAVY)
        c.setFont(bold, 13)
        c.drawString(ix, ty - 30, head)
        c.setFillColor(SUB)
        c.setFont(regular, 9)
        c.drawString(ix, ty - 47, sub)
        iy = ty - 66
        step = (ph - 66 - 12) / 2
        for k, (lab, val, how) in enumerate(items):
            y0 = iy - k * step
            if k:
                c.setStrokeColor(WARN if over else RULE)
                c.setLineWidth(0.5)
                c.line(ix, y0 + 6, px + pw - 20, y0 + 6)
            c.setFillColor(SUB)
            c.setFont(regular, 9.5)
            c.drawString(ix, y0 - 12, lab)
            c.setFillColor(WARN if over else NAVY)
            c.setFont(bold, 20)
            c.drawString(ix, y0 - 38, val)
            c.setFillColor(MUTED)
            c.setFont(regular, 8)
            c.drawString(ix, y0 - 54, how)

    notes = list(h.get('notes') or [])
    c.setFillColor(MUTED)
    c.setFont(regular, 7.5)
    ny = table_bottom - 16
    for n in notes:
        c.drawString(M, ny, '※ ' + n)
        ny -= 11


PAGES = [_page_cover, _page_toc, _page_summary, _page_affiliates, _page_fs, _page_holding]


def build_mgmt_report(period, *, provisional=False, generated_at=None, data=None):
    """경영진 보고자료 PDF 를 bytes 로 만든다. data: 장별 숫자 (app.py 에서 계산해 넘김)."""
    ctx = {
        'period': period,
        'provisional': provisional,
        'generated_at': generated_at or datetime.now(),
        'data': data or {},
    }
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
    c.setTitle(f'{TITLE} — {period_label(period)}')
    c.setAuthor(REPORTER)
    c.setSubject('경영진 보고자료')
    for n, draw in enumerate(PAGES, start=1):
        ctx['page_no'] = n
        draw(c, ctx)
        c.showPage()
    c.save()
    return buf.getvalue()
