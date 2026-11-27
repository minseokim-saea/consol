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

from html import escape as _esc

from dsd_engine import DSDDocument, fmt_amount, norm, parse_amount


def html_escape(s):
    return _esc(str(s), quote=False)

FS_MAPPING_FILE = Path(__file__).with_name('dsd_fs_mapping.json')
STATEMENTS = (('BS', '재무상태표'), ('PL', '손익계산서'), ('CF', '현금흐름표'))


def fs_key(s: str) -> str:
    s = norm(s)
    s = re.sub(r'\(주석[^)]*\)', '', s)
    s = re.sub(r'^(?:(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+|[IVXL]+|\d+|[가-하])\.|\(\d+\)|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)', '', s)
    s = s.lstrip('.')
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
        matched_row = {}                  # 시스템 줄 번호 → DSD 행 번호 (줄 추가 위치용)
        for ri, r in enumerate(t.rows):
            if ri < t.header_rows:
                continue
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
                item['expr'] = (' + '.join(names) + (' × -1' if rule.get('sign') == -1 else '')
                                + (f" ÷ {float(rule['div']):,.0f}" if rule.get('div') else '')
                                + (f" ({rule['in']} 아래)" if rule.get('in') else ''))
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
                matched_row[i] = ri
            ptr = max(idx) + 1
            val = sum((rows[i].get('cur') or 0) for i in idx) * (rule.get('sign', 1) if rule else 1)
            if rule is not None and rule.get('div'):
                q = val / float(rule['div'])          # 주당 금액 — 원 단위 사사오입
                val = int(abs(q) + 0.5) * (1 if q >= 0 else -1)
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
        before = (mapping.get('insert_before') or {}).get(code) or {}
        labels = (mapping.get('insert_label') or {}).get(code) or {}
        out += _plan_inserts(doc, t, code, title, rows, used, matched_row, sub_col, tot_col, before, labels)
    if report.get('SUBS'):
        out += plan_subs(doc, report['SUBS'], mapping.get('subsidiaries') or {})
    return out


def _co_key(s: str) -> str:
    """회사명 비교용 — (주)·(*1) 표시, 띄어쓰기·문장부호, 대소문자 무시."""
    s = re.sub(r'\(\*\d+\)|\(주\)|주식회사|\(미제출\)', '', str(s or ''))
    return re.sub(r'[\s.,()·\-]', '', s).lower()


def _unit_of(t) -> int:
    cap = norm(t.caption)
    for k, u in (('백만원', 1000000), ('천원', 1000)):
        if k in cap:
            return u
    return 1


def plan_subs(doc: DSDDocument, subs: dict, conf: dict) -> list[dict]:
    """주석 '주요 종속기업의 요약재무현황' 표 — 회사별 자산·부채·매출·당기순손익을 연결정산표 회사 열 값으로.

    subs : {연결정산표 회사 열 이름: {'1000000': 값, ...}} (원 단위)
    conf : dsd_fs_mapping.json 의 subsidiaries — anchor(표 앞 문단), cols({DSD 열: 계정코드}),
           companies({DSD 회사명: 연결정산표 회사 열 이름})."""
    anchor = conf.get('anchor') or '주요 종속기업의 요약재무현황'
    cols = conf.get('cols') or {}
    by_co = {_co_key(k): v for k, v in subs.items()}
    names = {_co_key(k): v for k, v in (conf.get('companies') or {}).items()}
    tname = '주요 종속기업 요약재무현황'
    akey = norm(anchor).replace(' ', '')
    t = next((x for x in doc.tables
              if akey in norm(x.anchor).replace(' ', '') and len(x.rows) > 1
              and any(h in x.col_headers().values() for h in cols)), None)
    if t is None:
        return [{'table_id': '주석1', 'table_name': tname, 'row': '', 'col': '', 'old': '', 'new': '',
                 'raw': None, 'status': 'error', 'review': False, 'sheet': '연결정산표', 'expr': '',
                 'message': f'DSD 에서 “{anchor}” 표를 찾지 못함'}]
    unit = _unit_of(t)
    heads = t.col_headers()
    out = []
    for r in t.body_rows():
        first = next((c for c in r if c.col == 0), None)
        lab = first.text.strip() if first is not None else ''
        if not lab:
            continue
        co = names.get(_co_key(lab))
        vals = by_co.get(_co_key(co)) if co else None
        for ci, h in heads.items():
            if ci == 0 or h not in cols:
                continue
            cell = t.cell_at(r, ci)
            item = {'table_id': '주석1', 'table_name': tname, 'row': lab, 'col': h,
                    'old': cell.text if cell is not None else '', 'new': '', 'raw': None,
                    'status': 'error', 'review': False, 'sheet': '연결정산표', 'dsd_table': t.index,
                    'expr': f'{co or "?"} · {h}({cols[h]})', 'message': ''}
            out.append(item)
            if cell is None:
                item['message'] = '병합된 칸이라 쓸 수 없음'
                continue
            if co is None:
                item['message'] = 'dsd_fs_mapping.json subsidiaries.companies 에 이 회사가 없음'
                continue
            if vals is None:
                item['message'] = f'연결정산표에 “{co}” 열이 없음'
                continue
            raw = float(vals.get(cols[h]) or 0)
            item['raw'] = raw
            item['new'] = fmt_amount(raw / unit, 0)
            item['_cell'] = cell
            o, n = parse_amount(item['old']), parse_amount(item['new'])
            item['status'] = 'ok' if (o is not None and n is not None and o == n) else 'changed'
    return out


def _dsd_row_index(t, label):
    """틀 DSD 표에서 과목명이 label 과 같은 첫 행 번호 (주석 표시·번호·띄어쓰기 무시)."""
    key = fs_key(label)
    for ri, r in enumerate(t.rows):
        if ri < t.header_rows:
            continue
        lab = next((c.text for c in r if c.col == 0), '')
        if lab.strip() and fs_key(lab) == key:
            return ri
    return None


def _plan_inserts(doc, t, code, title, rows, used, matched_row, sub_col, tot_col, before=None, labels=None):
    """시스템에 금액이 있는데 틀 DSD 에 없는 세부 줄 → 앞 계정 줄 아래에 새 줄로 추가.
    빠지면 소계·합계가 맞지 않으므로 기본 적용. 이미 쓰인 소계의 하위 줄(예: 매출원가 세부)은 제외.
    before = {시스템 줄 이름: DSD 과목명} — 그 과목 바로 위에 추가 (dsd_fs_mapping.json 의 insert_before).
    labels = {시스템 줄 이름: DSD 에 쓸 과목명} — 새 줄 과목명을 바꿔 쓴다 (insert_label)."""
    before = {norm(k): v for k, v in (before or {}).items()}
    labels = {norm(k): v for k, v in (labels or {}).items()}
    skip = set(used)
    for s in used:
        # 소계만 DSD 한 줄로 쓰이고 하위 줄은 하나도 안 쓰였으면 (예: 매출원가) 하위 줄은 추가하지 않는다
        if rows[s].get('kind') == 'subtotal':
            lv = rows[s].get('level', 0)
            kids = []
            for k in range(s + 1, len(rows)):
                if rows[k].get('level', 0) <= lv:
                    break
                kids.append(k)
            if kids and not any(k in used for k in kids):
                skip.update(kids)
    pri_cols = [c for c, h in t.col_headers().items() if '(전)' in h]
    pri_sub = pri_cols[0] if pri_cols else None

    def is_detail(ri):
        r = t.rows[ri]
        s, tt = t.cell_at(r, sub_col), t.cell_at(r, tot_col)
        return s is not None and s.text.strip() and not (tt is not None and tt.text.strip())

    out, pending = [], {}
    for k, x in enumerate(rows):
        if k in skip or x.get('kind') != 'detail' or abs(x.get('cur') or 0) < 0.5:
            continue
        target = _dsd_row_index(t, before[norm(x['name'])]) if norm(x['name']) in before else None
        if target is not None and target - 1 >= t.header_rows:
            anchor = target - 1           # 지정한 과목 바로 위 = 그 앞 행 아래
        else:
            prev = [i for i in matched_row if i < k]
            if not prev:
                continue
            anchor = matched_row[max(prev)]
        tpl = anchor if is_detail(anchor) else next(
            (j for j in range(anchor + 1, len(t.rows)) if is_detail(j)), None)
        if tpl is None:
            continue
        a, b = t.row_spans[tpl]
        raw = doc.raw_span(a, b)
        cols = [c.col for c in t.rows[tpl]]
        pri = x.get('pri') or 0
        label = labels.get(norm(x['name']), x['name'])
        content = {0: html_escape(label), sub_col: fmt_amount(x['cur'], 0)}
        if pri_sub is not None:
            content[pri_sub] = fmt_amount(pri, 0)
        it = iter(cols)
        new_row = re.sub(r'(<T[DE][^>]*>)(.*?)(</T[DE]>)',
                         lambda m: m.group(1) + content.get(next(it), '　') + m.group(3), raw, flags=re.S)
        new_row = re.sub(r'ACODE="[^"]*"', 'ACODE=""', new_row)
        pos = t.row_spans[anchor][1]
        nxt = t.row_spans[anchor + 1][0] if anchor + 1 < len(t.row_spans) else pos
        sep = doc.text[pos:nxt] if nxt > pos else '\n'
        anchor_lab = next((c.text for c in t.rows[anchor] if c.col == 0), '')
        out.append({'table_id': code, 'table_name': '연결' + title,
                    'row': x['name'] if label == x['name'] else f"{x['name']} → {label}", 'col': '줄 추가',
                    'old': '', 'new': fmt_amount(x['cur'], 0), 'raw': x['cur'], 'status': 'changed',
                    'review': False, 'sheet': '보고용 재무제표', 'dsd_table': t.index,
                    'expr': f'“{anchor_lab}” 아래에 추가',
                    'message': '틀 DSD 에 없는 줄 — 추가합니다 (전기 칸: 시스템 전기 값, 없으면 -)',
                    '_insert': (pos, sep + new_row)})
    return out

