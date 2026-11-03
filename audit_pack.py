"""
감사인 제출자료 — 연결정산표·현금정산표·주석합산 엑셀을 한 워크북으로 묶는다.

각 자료는 기존 생성기(consol_engine / cf_engine / note_aggregate)가 만든 워크북을
그대로 받아 시트 단위로 복사한다. 맨 앞에는 '목록' 시트를 두고 각 시트로 가는
링크를 걸며, 각 시트 A1 옆에는 목록으로 돌아가는 링크를 단다.
"""
from copy import copy

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import coordinate_from_string

INDEX_SHEET = '목록'

_TITLE_FONT = Font(bold=True, size=16, name='맑은 고딕', color='1F3864')
_HDR_FONT = Font(bold=True, color='FFFFFF', name='맑은 고딕', size=10)
_HDR_FILL = PatternFill('solid', start_color='1F3864')
_SEC_FONT = Font(bold=True, color='1F3864', name='맑은 고딕', size=10)
_SEC_FILL = PatternFill('solid', start_color='EBF0FA')
_DATA_FONT = Font(name='맑은 고딕', size=10)
_LINK_FONT = Font(name='맑은 고딕', size=10, color='0563C1', underline='single')
_ERR_FONT = Font(name='맑은 고딕', size=10, color='C0392B')
_NOTE_FONT = Font(name='맑은 고딕', size=9, color='6B7280')
_THIN = Side(border_style='thin', color='BFBFBF')
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_BAD_CHARS = set('[]:*?/\\')


def safe_sheet_title(name, used):
    """엑셀 시트명 규칙(31자, 금지문자)에 맞추고 중복이면 번호를 붙인다."""
    base = ''.join('_' if ch in _BAD_CHARS else ch for ch in str(name)).strip("'") or 'Sheet'
    base = base[:31]
    title, n = base, 2
    while title.lower() in used:
        suffix = f'({n})'
        title = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(title.lower())
    return title


def copy_sheet(src, dst, row_offset=0):
    """src 워크시트의 값·서식·열너비·행높이·병합·틀고정을 dst 로 복사.
    워크북이 달라 스타일 객체는 개별 복사해야 한다.
    row_offset 만큼 아래로 내려 쓴다 (수식 참조도 같이 이동)."""
    for row in src.iter_rows():
        for c in row:
            if c.value is None and not c.has_style:
                continue
            v = c.value
            if row_offset and isinstance(v, str) and v.startswith('='):
                dest = f'{c.column_letter}{c.row + row_offset}'
                v = Translator(v, origin=c.coordinate).translate_formula(dest)
            d = dst.cell(row=c.row + row_offset, column=c.column, value=v)
            if c.has_style:
                d.font = copy(c.font)
                d.fill = copy(c.fill)
                d.border = copy(c.border)
                d.alignment = copy(c.alignment)
                d.number_format = c.number_format
                d.protection = copy(c.protection)
    for key, dim in src.column_dimensions.items():
        dd = dst.column_dimensions[key]
        dd.width = dim.width
        dd.hidden = dim.hidden
        dd.outline_level = dim.outline_level
    for idx, dim in src.row_dimensions.items():
        if dim.height is not None:
            dst.row_dimensions[idx + row_offset].height = dim.height
        if dim.hidden:
            dst.row_dimensions[idx + row_offset].hidden = True
    for rng in src.merged_cells.ranges:
        rng = copy(rng)
        rng.shift(0, row_offset)
        dst.merge_cells(rng.coord)
    if src.freeze_panes:
        col, row = coordinate_from_string(src.freeze_panes)
        dst.freeze_panes = f'{col}{row + row_offset}'
    elif row_offset:
        dst.freeze_panes = f'A{1 + row_offset}'   # 복귀 링크 행은 항상 보이게
    if src.sheet_properties.tabColor is not None:
        dst.sheet_properties.tabColor = copy(src.sheet_properties.tabColor)


def build_audit_pack(period, parts, output_path, provisional=False):
    """parts: [{'category', 'label', 'title', 'desc', 'wb': Workbook | None,
                'sheets': [이름,...] | None, 'error': str | None}]
    label 은 목록·시트명, title 은 시트 1행에 쓰는 자료 이름(생략 시 label).
    category 순서대로 목록에 쓰고, 각 워크북의 시트를 '라벨' 기반 이름으로 복사한다.
    sheets 를 주면 그 시트만(그 순서로), 생략하면 전부 복사."""
    out = Workbook()
    idx = out.active
    idx.title = INDEX_SHEET
    used = {INDEX_SHEET.lower()}

    entries = []   # (category, label, desc, sheet_title | None, error | None)
    for p in parts:
        wb = p.get('wb')
        if wb is None:
            entries.append((p['category'], p['label'], p.get('desc') or '', None,
                            p.get('error') or '생성 실패'))
            continue
        names = p.get('sheets') or wb.sheetnames
        multi = len(names) > 1
        for nm in names:
            if nm not in wb.sheetnames:
                continue
            src = wb[nm]
            # 주석합산처럼 시트가 여러 개면 원래 시트명을 쓰고, 하나면 라벨을 쓴다.
            title = safe_sheet_title(nm if multi else p['label'], used)
            dst = out.create_sheet(title)
            copy_sheet(src, dst, row_offset=1)
            # 1행: 목록 복귀 링크 + 자료 이름
            back = dst.cell(1, 1, '◀ 목록')
            back.hyperlink = f"#'{INDEX_SHEET}'!A1"
            back.font = _LINK_FONT
            dst.cell(1, 2, f"{p['category']} — {nm if multi else p.get('title') or p['label']}").font = _SEC_FONT
            desc = p.get('desc') or ''
            if multi and not desc and isinstance(src['A1'].value, str):
                desc = src['A1'].value   # 주석합산 시트는 A1 이 제목
            entries.append((p['category'], p['label'] if not multi else nm,
                            desc, title, None))

    # ── 목록 시트 ─────────────────────────────────────────────
    idx.column_dimensions['A'].width = 6
    idx.column_dimensions['B'].width = 16
    idx.column_dimensions['C'].width = 40
    idx.column_dimensions['D'].width = 60
    idx.cell(1, 1, '감사인 제출자료').font = _TITLE_FONT
    idx.cell(2, 1, f'결산기간: {period}').font = _DATA_FONT
    r = 3
    if provisional:
        idx.cell(r, 1, '※ 결산 마감 전 잠정 수치입니다. 마감 전까지 변동될 수 있습니다.').font = _ERR_FONT
        r += 1
    idx.cell(r, 1, '시트명을 클릭하면 해당 시트로 이동합니다. '
                   '각 시트 왼쪽 위의 [◀ 목록]을 누르면 이 화면으로 돌아옵니다.').font = _NOTE_FONT
    r += 2

    for ci, h in enumerate(['No', '구분', '시트', '설명'], 1):
        c = idx.cell(r, ci, h)
        c.font = _HDR_FONT; c.fill = _HDR_FILL; c.border = _BORDER
        c.alignment = Alignment(horizontal='center', vertical='center')
    idx.freeze_panes = idx.cell(r + 1, 1).coordinate
    r += 1

    no = 0
    last_cat = None
    for cat, label, desc, title, err in entries:
        if cat != last_cat:
            for ci in range(1, 5):
                c = idx.cell(r, ci, cat if ci == 2 else None)
                c.font = _SEC_FONT; c.fill = _SEC_FILL; c.border = _BORDER
            r += 1
            last_cat = cat
        no += 1
        idx.cell(r, 1, no)
        idx.cell(r, 2, '')
        c = idx.cell(r, 3, label)
        if title:
            c.hyperlink = f"#'{title}'!A1"
            c.font = _LINK_FONT
            idx.cell(r, 4, desc).font = _DATA_FONT
        else:
            c.font = _DATA_FONT
            idx.cell(r, 4, f'생성 실패 — {err}').font = _ERR_FONT
        for ci in range(1, 5):
            idx.cell(r, ci).border = _BORDER
        idx.cell(r, 1).alignment = Alignment(horizontal='center')
        idx.cell(r, 1).font = _DATA_FONT
        r += 1

    out.save(str(output_path))
    return [e for e in entries if e[3]], [e for e in entries if not e[3]]
