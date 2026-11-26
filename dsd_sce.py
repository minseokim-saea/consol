"""연결자본변동표 당기 블록 채우기 — 보고용 재무제표(BS·PL)로 계산.

시스템에는 연결 자본변동표가 없으므로 다음과 같이 만든다 (금액 단위 원).
  · 당기초 행      : 틀 그대로 (이월했다면 작년 당기말 = 새 당기초)
  · 당기순이익 행  : 손익계산서 지배기업지분순이익(이익잉여금 열) · 비지배지분순이익(비지배지분 열)
  · 기타포괄손익 행: 재무상태표 해당 줄의 기초 대비 증감 (기타포괄손익누계액 열)
                    기초 = 틀 DSD 재무상태표 전기 칸 값
  · 그 밖의 변동 행: '-' 로 비우고, 마지막 변동 행에 '당기말 − 당기초 − 위 행들' 차액을 모은다
  · 당기말 행      : 재무상태표 자본 구성 (자본금·자본잉여금·자본조정·기타포괄손익누계액·이익잉여금·비지배지분)
  · 총계 열        : 행마다 다른 열의 합
결과 형식은 dsd_builder.plan 과 같다.
"""
from __future__ import annotations

import re

from dsd_engine import DSDDocument, fmt_amount, norm, parse_amount
from dsd_fs import find_statement_tables, fs_key

_MARK = re.compile(r'\((전기초|전기말|당기초|당기말)\)')

# (열 머리글에 들어 있는 말, 재무상태표 줄 이름)
_COLS = [('자본금', 'Ⅰ.자본금'), ('자본잉여금', 'Ⅱ.자본잉여금'), ('자본조정', 'Ⅲ.자본조정'),
         ('기타포괄손익누계액', 'Ⅳ.기타포괄손익누계액'), ('이익잉여금', 'Ⅴ.이익잉여금'),
         ('비지배지분', '비지배지분')]

# 자본변동표 행 이름에 든 말 → 재무상태표 기타포괄손익 세부 줄
_OCI = [('매도가능', ['매도가능금융자산평가']),
        ('지분법자본변동', ['지분법자본변동', '부의지분법자본변동']),
        ('해외사업환산', ['해외사업환산손익']),
        ('재평가', ['재평가차익']),
        ('파생상품', ['파생상품평가손익'])]


def _won(v) -> int:
    """원 단위 사사오입 — 행·열 합계가 표시값 그대로 맞도록 모든 칸을 정수로 다룬다."""
    v = float(v or 0)
    return int(abs(v) + 0.5) * (1 if v >= 0 else -1)


def find_sce(doc: DSDDocument):
    for t in doc.tables:
        marks = {}
        for i, r in enumerate(t.rows):
            first = next((c for c in r if c.col == 0), None)
            m = _MARK.search(first.text) if first is not None else None
            if m:
                marks[m.group(1)] = i
        if {'당기초', '당기말'} <= set(marks) and marks['당기초'] < marks['당기말']:
            return t, marks['당기초'], marks['당기말']
    return None, None, None


def _sys_value(rows, name):
    k = norm(name)
    for x in rows:
        if norm(x['name']) == k:
            return x.get('cur') or 0
    k = fs_key(name)
    for x in rows:
        if fs_key(x['name']) == k:
            return x.get('cur') or 0
    return None


def _bs_prior(doc, name):
    """틀 DSD 재무상태표 전기 칸 값 (기타포괄손익 세부 줄의 기초)."""
    t = find_statement_tables(doc).get('BS')
    if t is None:
        return None
    pcols = [c for c, h in t.col_headers().items() if '(전)' in h]
    if not pcols:
        return None
    k = fs_key(name)
    for r in t.body_rows():
        first = next((c for c in r if c.col == 0), None)
        if first is not None and fs_key(first.text) == k:
            for pc in pcols:
                cell = t.cell_at(r, pc)
                if cell is not None and cell.text.strip():
                    return parse_amount(cell.text) or 0
            return 0
    return None


def plan_sce(doc: DSDDocument, fs_rows: dict) -> list[dict]:
    t, c0, c1 = find_sce(doc)
    bs, pl = (fs_rows or {}).get('BS') or [], (fs_rows or {}).get('PL') or []
    if t is None or not bs:
        return []
    heads = t.col_headers()
    colmap, total_col = {}, None
    for c, h in heads.items():
        if c == 0:
            continue
        if '총계' in h or h.endswith('합계'):
            total_col = c
            continue
        for key, bs_name in _COLS:
            if key in h and key not in colmap.values():
                colmap[c] = key
                break
    bs_of = dict(_COLS)
    base = {'table_id': 'SCE', 'table_name': '연결자본변동표', 'raw': None, 'review': False,
            'sheet': '보고용 재무제표 (BS·PL 로 계산)', 'dsd_table': t.index}
    out = []

    def err(row, msg):
        out.append(dict(base, row=row, col='', old='', new='', status='error', message=msg, expr=''))

    closing, opening = {}, {}
    for c, key in colmap.items():
        v = _sys_value(bs, bs_of[key])
        if v is None:
            err('', f'재무상태표에서 {bs_of[key]} 줄을 찾지 못함')
            return out
        closing[c] = _won(v)
        oc = t.cell_at(t.rows[c0], c)
        opening[c] = _won(parse_amount(oc.text) if oc is not None else 0)

    moves = list(range(c0 + 1, c1))
    label = lambda i: next((x.text for x in t.rows[i] if x.col == 0), '')
    values = {i: {c: 0 for c in colmap} for i in moves}
    how = {i: '' for i in moves}
    residual_row = None
    for i in moves:
        lab = norm(label(i))
        if '당기순이익' in lab or '당기순손익' in lab:
            for c, key in colmap.items():
                if key == '이익잉여금':
                    values[i][c] = _won(_sys_value(pl, '지배기업지분순이익'))
                elif key == '비지배지분':
                    values[i][c] = _won(_sys_value(pl, '비지배지분순이익'))
            how[i] = '손익계산서 지배기업지분순이익 · 비지배지분순이익'
            continue
        oci = next((names for kw, names in _OCI if kw in lab), None)
        if oci:
            oci_col = next((c for c, k in colmap.items() if k == '기타포괄손익누계액'), None)
            cur = sum(_won(_sys_value(bs, n)) for n in oci)
            pri = sum(_won(_bs_prior(doc, n)) for n in oci)
            if oci_col is not None:
                values[i][oci_col] = cur - pri
            how[i] = f"재무상태표 {' + '.join(oci)} 증감"
            continue
        residual_row = i              # 마지막 '그 밖의 변동' 행에 차액을 모은다
    if residual_row is None:
        err('', '차액을 넣을 변동 행이 없음 (당기순이익·기타포괄손익 외 행이 하나는 있어야 함)')
        return out
    for i in moves:
        if i != residual_row and not how[i]:
            how[i] = '시스템에 없는 항목 — 나머지 변동 행에 합쳐짐'
    for c in colmap:
        assigned = sum(values[i][c] for i in moves if i != residual_row)
        values[residual_row][c] = closing[c] - opening[c] - assigned
    how[residual_row] = '당기말 − 당기초 − 위 행들 (나머지 변동)'
    values[c1] = dict(closing)
    how[c1] = '재무상태표 자본 구성'

    for i in moves + [c1]:
        row = t.rows[i]
        cells = [(c, values[i][c]) for c in colmap]
        if total_col is not None:
            cells.append((total_col, sum(values[i].values())))
        for c, v in cells:
            cell = t.cell_at(row, c)
            if cell is None:
                continue
            new = fmt_amount(v, 0)
            o, n = parse_amount(cell.text), parse_amount(new)
            out.append(dict(base, row=label(i), col=heads.get(c, ''), old=cell.text, new=new, raw=v,
                            status='ok' if (o is not None and o == n) else 'changed',
                            message='', expr=how[i], _cell=cell))
    # 당기말 총계가 재무상태표 자본총계와 다르면 알림
    tot = _sys_value(bs, '자본총계')
    if tot is not None and abs(sum(closing.values()) - tot) > 1:
        err(label(c1), f'자본 구성 합계 {fmt_amount(sum(closing.values()))} 와 재무상태표 자본총계 '
                       f'{fmt_amount(tot)} 가 다름')
    return out
