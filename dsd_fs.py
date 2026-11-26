"""보고용 재무제표(BS/PL/CF) → DSD 본문 재무제표 당기 칸 채우기.

DSD 재무제표 표: [과목 | 당기(세부) | 당기(합계) | 전기(세부) | 전기(합계)], 금액 단위 원.
시스템 보고용 재무제표(_compute_report_fs 의 rows: name·kind·cur)를 줄 이름으로 맞춘다.
  1. dsd_fs_mapping.json 에 적힌 짝 (DSD 과목 → 시스템 줄 이름 [여러 줄이면 합계], 부호)
     in = 같은 이름 줄이 여럿일 때 그 구분 줄 아래 것, rename = DSD 과목명도 바꿈
  2. 없으면 이름 자동 매칭 — 번호(Ⅰ. 1. (1))·주석 표시·띄어쓰기·'의'·(손실) 무시
값을 넣을 칸은 틀 DSD 에서 그 줄 금액이 있던 칸(세부/합계)을 따른다.
결과 형식은 dsd_builder.plan 과 같아 미리보기·선택·적용을 함께 쓴다.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from dsd_engine import DSDDocument, fmt_amount, norm, parse_amount

FS_MAPPING_FILE = Path(__file__).with_name('dsd_fs_mapping.json')
STATEMENTS = (('BS', '재무상태표'), ('PL', '손익계산서'), ('CF', '현금흐름표'))


def fs_key(s: str) -> str:
    s = norm(s)
    s = re.sub(r'\(주석[^)]*\)', '', s)
    s = re.sub(r'^(?:(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+|[IVXL]+|\d+|[가-하])\.|\(\d+\)|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)', '', s)
    s = s.replace('(손실)', '').replace('(수익)', '').replace('의', '')
    return s


def load_fs_mapping(path=None) -> dict:
    p = Path(path or FS_MAPPING_FILE)
    if not p.exists():
        return {}
    with open(p, encoding='utf-8') as f:
        return json.load(f)


def find_statement_tables(doc: DSDDocument) -> dict:
    """{'BS': Table, ...} — 제목 문단에 표 이름이 있고 머리글에 '(당)' 이 있는 표."""
    out = {}
    for t in doc.tables:
        heads = '/'.join(t.col_headers().values())
        if '(당)' not in heads or len(t.rows) < 5:
            continue
        a = norm(t.anchor)
        for code, name in STATEMENTS:
            if name in a and code not in out:
                out[code] = t
    return out


def _cur_cols(t):
    heads = t.col_headers()
    cols = [c for c, h in heads.items() if '(당)' in h]
    return cols[:2] if len(cols) >= 2 else (cols + cols)[:2]


def plan_fs(doc: DSDDocument, report: dict, mapping: dict | None = None) -> list[dict]:
    """report: {'BS': [rows], 'PL': [...], 'CF': [...]} (rows: name, kind, cur)."""
    mapping = mapping if mapping is not None else load_fs_mapping()
    tables = find_statement_tables(doc)
    out = []
    for code, title in STATEMENTS:
        t = tables.get(code)
        rows = report.get(code) or []
        if t is None or not rows:
            continue
        sub_col, tot_col = _cur_cols(t)
        by_name: dict[str, list[int]] = {}
        for i, x in enumerate(rows):
            by_name.setdefault(norm(x['name']), []).append(i)
            by_name.setdefault('~' + fs_key(x['name']), []).append(i)
        rules = {}
        for m in (mapping.get(code) or []):
            rules[(norm(m['dsd']), int(m.get('nth', 1)))] = m
        used, ptr, seen = set(), 0, {}
        for r in t.body_rows():
            first = next((c for c in r if c.col == 0), None)
            if first is None or not first.text.strip():
                continue
            lab = first.text
            seen[norm(lab)] = seen.get(norm(lab), 0) + 1
            nth = seen[norm(lab)]
            sub, tot = t.cell_at(r, sub_col), t.cell_at(r, tot_col)
            item = {'table_id': code, 'table_name': '연결' + title, 'row': lab, 'col': '당기',
                    'old': '', 'new': '', 'raw': None, 'status': 'error', 'message': '',
                    'review': False, 'sheet': '보고용 재무제표', 'expr': '', 'dsd_table': t.index}
            rule = rules.get((norm(lab), nth))
            if rule is not None and rule.get('skip'):
                continue
            if not any(c.text.strip() for c in (sub, tot) if c is not None):
                continue                  # '자 산' 같은 구분 제목 줄 (틀에 금액 없음)
            if rule is not None:
                names = rule['sys'] if isinstance(rule['sys'], list) else [rule['sys']]
                # in: 같은 이름 줄이 여러 개일 때 이 구분 줄 아래의 것만 (예: 투자활동 현금유입액)
                lo = -1
                if rule.get('in'):
                    hs = by_name.get(norm(rule['in']), [])
                    lo = hs[0] if hs else len(rows)
                idx = []
                for nm in names:
                    pool = [i for i in by_name.get(norm(nm), []) if i > lo]
                    cand = [i for i in pool if i not in used] or pool
                    if not cand:
                        idx = None
                        break
                    idx.append(cand[0])
                item['expr'] = ' + '.join(names) + (' × -1' if rule.get('sign') == -1 else '')                     + (f" ({rule['in']} 아래)" if rule.get('in') else '')
                if idx is None:
                    item['message'] = f'매핑한 시스템 줄을 찾지 못함: {names}'
                    out.append(item)
                    continue
            else:
                cand = [i for i in by_name.get('~' + fs_key(lab), []) if i not in used]
                if not cand:
                    has_num = any(c.text.strip() for c in (sub, tot) if c is not None)
                    if not has_num:
                        continue          # '자 산' 같은 구분 제목 줄
                    item['message'] = '시스템 보고용 재무제표에서 같은 이름의 줄을 찾지 못함 — 매핑 필요'
                    out.append(item)
                    continue
                nxt = [i for i in cand if i >= ptr]
                idx = [nxt[0] if nxt else cand[0]]
                item['expr'] = rows[idx[0]]['name']
            for i in idx:
                used.add(i)
            ptr = max(idx) + 1
            val = sum((rows[i].get('cur') or 0) for i in idx) * (rule.get('sign', 1) if rule else 1)
            # 틀에서 금액이 있던 칸, 없으면 시스템 줄 종류로 결정
            if tot is not None and tot.text.strip() and (sub is None or not sub.text.strip()):
                cell = tot
            elif sub is not None and sub.text.strip():
                cell = sub
            else:
                kind = rows[idx[0]].get('kind')
                cell = tot if (kind == 'subtotal' and tot is not None) else (sub or tot)
            if cell is None:
                item['message'] = '쓸 칸이 없음'
                out.append(item)
                continue
            item['col'] = '당기 합계' if cell is tot and tot is not sub else '당기'
            item['old'] = cell.text
            item['raw'] = val
            item['new'] = fmt_amount(val, 0)
            item['_cell'] = cell
            o, n = parse_amount(item['old']), parse_amount(item['new'])
            item['status'] = 'ok' if (o is not None and n is not None and o == n) else 'changed'
            if rule is not None and rule.get('rename') and norm(lab) != norm(rule['rename']):
                item['_rename'] = (first, rule['rename'])
                item['row'] = f"{lab} → {rule['rename']}"
                item['status'] = 'changed'
            out.append(item)
    return out + plan_levels(doc)


def plan_levels(doc: DSDDocument) -> list[dict]:
    """세부 금액 칸에 값이 있는데 들여쓰기 단계(ALEVEL)가 0 인 줄 — DART 편집기에서 안 보인다.
    바로 앞 세부 줄과 같은 단계로 맞춘다."""
    out = []
    for code, title in STATEMENTS:
        t = find_statement_tables(doc).get(code)
        if t is None:
            continue
        sub_col, tot_col = _cur_cols(t)
        prev = None
        for r in t.body_rows():
            first = next((c for c in r if c.col == 0), None)
            sub, tot = t.cell_at(r, sub_col), t.cell_at(r, tot_col)
            if first is None:
                continue
            is_detail = sub is not None and sub.text.strip() and not (tot is not None and tot.text.strip())
            lv = first.attrs.get('ALEVEL')
            if is_detail and lv == '0' and prev not in (None, '0'):
                span = doc.cell_tag_attr_span(first, 'ALEVEL')
                if span:
                    out.append({'table_id': code, 'table_name': '연결' + title, 'row': first.text,
                                'col': '들여쓰기 단계', 'old': lv, 'new': prev, 'raw': None,
                                'status': 'changed', 'review': False, 'sheet': 'DSD 서식',
                                'expr': '', 'dsd_table': t.index,
                                'message': 'DART 편집기에서 안 보이던 줄 — 앞 세부 줄과 같은 단계로 맞춤',
                                '_span': (span[0], span[1], prev)})
            if is_detail and lv not in (None, '0'):
                prev = lv
    return out
