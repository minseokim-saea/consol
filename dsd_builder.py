"""주석합산 결과 → DART 감사보고서 DSD 채우기.

흐름
  1. 주석합산 전체 워크북(build_all_in_one_excel 결과)을 시트별 표로 읽는다 (NoteSheets).
  2. dsd_mapping.json 의 규칙대로 각 DSD 주석 표 칸의 값을 계산한다.
  3. 작년(또는 이전) DSD 를 틀로 삼아 해당 칸 텍스트만 바꿔 새 DSD 로 저장한다.

매핑 규칙 (dsd_mapping.json)
  tables[]:
    id       — 표시용 번호 (예: "11.2")
    name     — 표시용 이름
    anchor   — DSD 에서 표 바로 앞 문단에 들어 있는 문구 (공백 무시 부분일치)
    nth      — 같은 앵커 뒤 몇 번째 표인지 (기본 1)
    caption  — 단위표 문구로 추가 구분 (예: "(당기)") — 선택
    sheet    — 주석합산 시트명 (셀별로 덮어쓸 수 있음)
    where    — 표 공통 필터 (셀별 where 와 AND)
    unit     — 나눌 단위 (기본 1000 = 천원)
    cells[]:
      row    — DSD 행 첫 칸 문구 (공백·(*1) 무시 일치)
      col    — DSD 열 머리글 (다단 머리글은 '당기말/유동' 처럼 / 로 연결, 끝부분만 써도 됨)
      expr   — 계산식. [열이름] = 데이터 행 합계, {요약라벨} = 시트 상단 요약값,
               숫자·+ - * / ( ) 사용 가능
      where  — 필터. '열=값' / '열~포함' / '열!=값', 여러 값은 '|', 여러 조건은 ';'
      sheet, unit, decimals — 선택
"""
from __future__ import annotations

import ast
import json
import operator
import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from dsd_engine import DSDDocument, fmt_amount, norm, parse_amount

MAPPING_FILE = Path(__file__).with_name('dsd_mapping.json')

_HDR_COLOR = '1F3864'
_TOTAL_COLOR = 'FFE699'


def _fill_rgb(cell):
    try:
        c = cell.fill.fgColor.rgb if cell.fill and cell.fill.fill_type else None
        return (c or '')[-6:].upper() if isinstance(c, str) else ''
    except Exception:
        return ''


def _key(s) -> str:
    return re.sub(r'\s+', '', str(s or ''))


@dataclass
class SheetData:
    name: str
    headers: list = field(default_factory=list)
    rows: list = field(default_factory=list)        # list[dict 정규화헤더 -> 값]
    summary: dict = field(default_factory=dict)     # 상단 '라벨 → 값'


class NoteSheets:
    """주석합산 전체 워크북을 시트별 (머리글, 데이터행, 요약값) 으로 읽는다."""

    def __init__(self, path):
        wb = load_workbook(str(path), data_only=True)
        self.sheets: dict[str, SheetData] = {}
        for ws in wb.worksheets:
            self.sheets[_key(ws.title)] = self._read(ws)
        wb.close()

    @staticmethod
    def _read(ws) -> SheetData:
        sd = SheetData(ws.title)
        grid = list(ws.iter_rows())
        hdr_idx = None
        for i, row in enumerate(grid):
            if sum(1 for c in row if c.value is not None and _fill_rgb(c) == _HDR_COLOR) >= 2:
                hdr_idx = i
                break
        if hdr_idx is None:  # 서식이 없으면 마지막 연속 블록의 첫 행을 머리글로
            blocks, start = [], None
            for i, row in enumerate(grid):
                filled = any(c.value is not None for c in row)
                if filled and start is None:
                    start = i
                if not filled and start is not None:
                    blocks.append(start); start = None
            if start is not None:
                blocks.append(start)
            hdr_idx = blocks[-1] if len(blocks) > 1 else None
        top = grid[:hdr_idx] if hdr_idx is not None else grid
        for row in top:
            vals = [c.value for c in row]
            if vals and isinstance(vals[0], str) and len(vals) > 1 and vals[1] is not None:
                sd.summary.setdefault(_key(vals[0]), vals[1])
        if hdr_idx is None:
            return sd
        sd.headers = [str(c.value).strip() if c.value is not None else '' for c in grid[hdr_idx]]
        keys = [_key(h) for h in sd.headers]
        for row in grid[hdr_idx + 1:]:
            vals = [c.value for c in row]
            if not any(v is not None for v in vals):
                continue
            first = row[0]
            if _fill_rgb(first) == _TOTAL_COLOR or str(first.value or '').strip() in ('전체 합계', '합계', '총계'):
                continue
            sd.rows.append({k: v for k, v in zip(keys, vals) if k})
        return sd

    def get(self, name) -> SheetData | None:
        k = _key(name)
        if k in self.sheets:
            return self.sheets[k]
        # 엑셀 시트명은 31자 제한으로 잘릴 수 있으므로 앞부분 일치도 허용
        for sk, sd in self.sheets.items():
            if sk.startswith(k[:31]) or k.startswith(sk):
                return sd
        return None


# ─── 필터 / 계산식 ───────────────────────────────────────────────────────────

def _match_where(row: dict, where: str) -> bool:
    for cond in [c.strip() for c in (where or '').split(';') if c.strip()]:
        m = re.match(r'^(.+?)(!=|=|~)(.*)$', cond)
        if not m:
            raise ValueError(f'필터 형식 오류: {cond}')
        col, op, vals = _key(m.group(1)), m.group(2), [_key(v) for v in m.group(3).split('|')]
        if col not in row:
            raise KeyError(f'필터 열 없음: {m.group(1)}')
        cell = _key(row.get(col))
        if op == '=' and cell not in vals:
            return False
        if op == '!=' and cell in vals:
            return False
        if op == '~' and not any(v in cell for v in vals):
            return False
    return True


_OPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
        ast.Div: operator.truediv, ast.USub: operator.neg, ast.UAdd: operator.pos}


def _safe_eval(expr: str, env: dict) -> float:
    def ev(n):
        if isinstance(n, ast.Expression):
            return ev(n.body)
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)):
            return n.value
        if isinstance(n, ast.Name):
            return env[n.id]
        if isinstance(n, ast.BinOp) and type(n.op) in _OPS:
            return _OPS[type(n.op)](ev(n.left), ev(n.right))
        if isinstance(n, ast.UnaryOp) and type(n.op) in _OPS:
            return _OPS[type(n.op)](ev(n.operand))
        raise ValueError('허용되지 않는 계산식')
    return ev(ast.parse(expr, mode='eval'))


def evaluate(sheets: NoteSheets, sheet_name: str, expr: str, where: str = '') -> float:
    sd = sheets.get(sheet_name)
    if sd is None:
        raise KeyError(f'시트 없음: {sheet_name}')
    env, i = {}, [0]

    def var(v):
        i[0] += 1
        name = f'_v{i[0]}'
        env[name] = v
        return name

    def col_sum(m):
        col = _key(m.group(1))
        if sd.rows and col not in sd.rows[0] and not any(col in r for r in sd.rows):
            raise KeyError(f'열 없음: {m.group(1)} (시트 {sd.name})')
        if not sd.rows and col not in [_key(h) for h in sd.headers]:
            raise KeyError(f'열 없음: {m.group(1)} (시트 {sd.name})')
        total = 0.0
        for r in sd.rows:
            if where and not _match_where(r, where):
                continue
            v = r.get(col)
            if isinstance(v, (int, float)):
                total += v
        return var(total)

    def summ(m):
        k = _key(m.group(1))
        if k not in sd.summary:
            raise KeyError(f'요약값 없음: {m.group(1)} (시트 {sd.name})')
        v = sd.summary[k]
        return var(float(v) if isinstance(v, (int, float)) else 0.0)

    py = re.sub(r'\[([^\]]+)\]', col_sum, expr)
    py = re.sub(r'\{([^}]+)\}', summ, py)
    return float(_safe_eval(py, env))


# ─── 매핑 적용 ──────────────────────────────────────────────────────────────

def load_mapping(path=None) -> dict:
    with open(path or MAPPING_FILE, encoding='utf-8') as f:
        return json.load(f)


def plan(doc: DSDDocument, sheets: NoteSheets, mapping: dict) -> list[dict]:
    """칸별 결과 목록. status: ok(같음) / changed(바뀜) / error."""
    out = []
    default_unit = mapping.get('unit', 1000)
    for t in mapping.get('tables', []):
        base = {'table_id': t.get('id', ''), 'table_name': t.get('name', '')}
        tbl = doc.find_table(t['anchor'], int(t.get('nth', 1)), t.get('caption', ''))
        for c in t.get('cells', []):
            item = dict(base, row=c['row'], col=c['col'], old='', new='', raw=None,
                        status='error', message='', review=bool(c.get('review')),
                        sheet=c.get('sheet') or t.get('sheet', ''), expr=c['expr'])
            out.append(item)
            if tbl is None:
                item['message'] = 'DSD 에서 표를 찾지 못함'
                continue
            item['dsd_table'] = tbl.index
            r = tbl.find_row(c['row'], int(c.get('row_nth', 1)))
            ci = tbl.find_col(c['col'])
            if r is None or ci is None:
                item['message'] = ('행' if r is None else '열') + ' 을(를) 찾지 못함'
                continue
            cell = tbl.cell_at(r, ci)
            if cell is None:
                item['message'] = '병합된 칸이라 쓸 수 없음'
                continue
            item['old'] = cell.text
            try:
                where = ';'.join(x for x in (t.get('where', ''), c.get('where', '')) if x)
                raw = evaluate(sheets, c.get('sheet') or t.get('sheet', ''), c['expr'], where)
            except Exception as e:  # noqa: BLE001 — 사용자에게 사유 표시
                item['message'] = str(e)
                continue
            unit = c.get('unit', t.get('unit', default_unit)) or 1
            val = raw / unit
            item['raw'] = raw
            item['new'] = fmt_amount(val, int(c.get('decimals', t.get('decimals', 0))))
            item['_cell'] = cell
            old_v, new_v = parse_amount(item['old']), parse_amount(item['new'])
            item['status'] = 'ok' if old_v is not None and new_v is not None and old_v == new_v else 'changed'
    return out


def apply(doc: DSDDocument, results: list[dict], selected=None) -> int:
    """selected: 적용할 결과 순번 집합. None 이면 확인필요(review) 칸을 뺀 전부."""
    n = 0
    for i, it in enumerate(results):
        if selected is None:
            if it.get('review'):
                continue
        elif i not in selected:
            continue
        cell = it.get('_cell')
        # 값이 같은 칸(ok)은 원문(정렬용 공백 등)을 그대로 둔다
        if cell is not None and it['status'] == 'changed':
            if parse_amount(cell.text) != parse_amount(it['new']):
                doc.set_cell(cell, it['new'])
            if it.get('_rename'):
                lab_cell, new_label = it['_rename']
                doc.set_cell(lab_cell, new_label)
            n += 1
    return n


def public(results: list[dict]) -> list[dict]:
    """JSON 응답용 (내부 셀 객체 제거)."""
    return [{k: v for k, v in it.items() if not k.startswith('_')} for it in results]
