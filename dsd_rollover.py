"""DSD 당기 → 전기 이월.

작년 DSD 를 틀로 쓸 때, 작년의 당기 숫자가 올해의 전기 숫자가 된다.
  · 열 이월   : 한 표 안에 '당기'·'전기'(또는 당기말·전기말) 열이 짝으로 있으면 행마다 당기 칸 → 전기 칸 복사.
                다단 머리글은 '당기말/유동' ↔ '전기말/유동' 처럼 나머지 경로가 같은 열끼리 짝짓는다.
  · 표 이월   : (당기) 표와 (전기) 표가 따로 있으면 당기 표 전체를 전기 표 자리에 복사 (행·열 구성까지).
짝을 못 찾은 열이 있는 표는 건드리지 않고 '수동 확인'으로 돌려준다.
이월은 숫자 채우기(dsd_builder.plan) 전에 실행한다.
"""
from __future__ import annotations

import re

from dsd_engine import DSDDocument, Table, norm

_CUR = {'당기': '전기', '당기말': '전기말', '당분기': '전분기', '당분기말': '전분기말'}
_PRI = set(_CUR.values())


def _marker(t: Table) -> str | None:
    """나뉜 표의 당기/전기 표시 — 단위 캡션 우선, 없으면 문단 끝 '(1) 당기' 류."""
    cap = norm(t.caption)
    m = re.match(r'^\(?(당기말?|전기말?)\)?', cap)
    if m:
        return 'cur' if m.group(1).startswith('당') else 'pri'
    m = re.search(r'\(\d\)(당기말?|전기말?)$', norm(t.anchor))
    if m:
        return 'cur' if m.group(1).startswith('당') else 'pri'
    return None


def _split_pairs(doc: DSDDocument):
    pairs, last_cur = [], None
    for t in doc.tables:
        if len(t.rows) < 2:
            continue
        mk = _marker(t)
        if mk == 'cur':
            last_cur = t
        elif mk == 'pri':
            if last_cur is not None and t.index - last_cur.index <= 6:
                pairs.append((last_cur, t))
            last_cur = None
    return pairs


def _col_pairs(t: Table):
    """[(당기열, 전기열)], 짝 없는 당기/전기 열 목록."""
    heads = t.col_headers()
    by_path: dict[str, list[int]] = {}
    for c, h in heads.items():
        by_path.setdefault(h, []).append(c)
    pairs, orphans, used = [], [], set()
    for c, h in heads.items():
        if c == 0:
            continue
        segs = h.split('/')
        idx = [i for i, s in enumerate(segs) if s in _CUR]
        if not idx:
            continue
        i = idx[0]
        target = '/'.join(segs[:i] + [_CUR[segs[i]]] + segs[i + 1:])
        cands = [x for x in by_path.get(target, []) if x not in used]
        if cands:
            used.add(cands[0])
            pairs.append((c, cands[0]))
        else:
            orphans.append(h)
    for c, h in heads.items():
        if c and c not in used and any(s in _PRI for s in h.split('/')):
            orphans.append(h)
    return pairs, orphans


def rollover(doc: DSDDocument) -> dict:
    """이월을 doc 에 적용하고 요약을 돌려준다.

    반환: {'tables': [{table, anchor, mode, cells}], 'manual': [{table, anchor, reason}],
           'current_cells': [(table, cell)]}   ← 이월 뒤 새 숫자로 채워야 할 당기 칸
    """
    out = {'tables': [], 'manual': [], 'current_cells': []}
    split = _split_pairs(doc)
    in_split = {id(x) for p in split for x in p}

    for cur, pri in split:
        doc.replace_table(pri, cur)
        out['tables'].append({'table': pri.index, 'anchor': pri.anchor, 'caption': pri.caption,
                              'mode': '표 이월', 'cells': sum(len(r) for r in cur.body_rows())})
        for r in cur.body_rows():
            for cell in r:
                if cell.col > 0 and cell.text.strip():
                    out['current_cells'].append((cur, cell))

    for t in doc.tables:
        if len(t.rows) < 2 or id(t) in in_split:
            continue
        pairs, orphans = _col_pairs(t)
        if orphans and not pairs:
            out['manual'].append({'table': t.index, 'anchor': t.anchor,
                                  'reason': '짝이 맞지 않는 열: ' + ', '.join(orphans)})
            continue
        if not pairs:
            continue
        if orphans:
            out['manual'].append({'table': t.index, 'anchor': t.anchor,
                                  'reason': '일부 열만 이월 — 짝이 맞지 않는 열: ' + ', '.join(orphans)})
        n = 0
        for r in t.body_rows():
            for cc, pc in pairs:
                src, dst = t.cell_at(r, cc), t.cell_at(r, pc)
                if src is None or dst is None or src.colspan != dst.colspan:
                    continue
                doc.set_raw(dst, doc.get_raw(src))
                n += 1
                if src.text.strip():
                    out['current_cells'].append((t, src))
        out['tables'].append({'table': t.index, 'anchor': t.anchor, 'caption': t.caption,
                              'mode': '열 이월', 'cells': n})
    out['tables'].sort(key=lambda x: x['table'])
    return out


def unrefreshed(roll: dict, results: list[dict], selected=None) -> list[dict]:
    """이월 뒤 당기 칸 중 숫자 채우기 대상이 아닌 칸 — 사람이 직접 고쳐야 한다."""
    covered = set()
    for i, it in enumerate(results):
        if it.get('_cell') is None or it['status'] == 'error':
            continue
        if selected is None and it.get('review'):
            continue
        if selected is not None and i not in selected:
            continue
        covered.add(id(it['_cell']))
    out = []
    for t, cell in roll['current_cells']:
        if id(cell) in covered:
            continue
        row = next((c.text for c in t.rows[cell.row] if c.col == 0), '')
        out.append({'table': t.index, 'anchor': t.anchor, 'caption': t.caption,
                    'row': row, 'col': t.col_headers().get(cell.col, ''), 'text': cell.text})
    return out
