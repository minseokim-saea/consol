"""DSD 기수·회계기간 표시 변경 (예: 제 38 기 → 제 39 기, 2025년 → 2026년).

바꾸는 곳 — '제 N 기' 가 들어 있는 표(표지·재무제표 머리글·주석 머리글)의 칸만:
  · 제 N 기 / 제N기 / 제 N(당) 기말 / 제 N(전) 기   → 기수 + shift
  · 그 표 안 날짜의 연도 (YYYY년, YYYY.MM.DD, YYYY-MM-DD) → 연도 + shift
  · 표지의 회계기간 코드 AUNITVALUE (PERIODFROM2 / PERIODTO2)
주석 본문 문장 속 날짜(취득일·만기일 등)는 사건 날짜라 건드리지 않는다.
"""
from __future__ import annotations

import re

from dsd_engine import DSDDocument

_KISU = re.compile(r'제(\s*)(\d+)(\s*)(\((?:당|전)\))?(\s*)기')
_PERIOD_ATTR = re.compile(r'(AUNIT="PERIOD(?:FROM|TO)\d*"[^>]*?AUNITVALUE=")(\d{8})(")')
_YEAR_PATS = [re.compile(r'(?<!\d)(\d{4})(\s*년)'), re.compile(r'(?<!\d)(\d{4})([.\-]\d{1,2}[.\-]\d{1,2})')]


def template_year(doc: DSDDocument) -> int | None:
    """틀 DSD 의 당기 연도 — 표지 회계기간 끝(PERIODTO) 코드, 없으면 표지 칸의 가장 늦은 연도."""
    m = re.search(r'AUNIT="PERIODTO\d*"[^>]*?AUNITVALUE="(\d{4})\d{4}"', doc.text)
    if m:
        return int(m.group(1))
    years = [int(y) for t in _kisu_tables(doc) for r in t.rows for c in r
             for y in re.findall(r'(?<!\d)(20\d\d)\s*년', c.text)]
    return max(years) if years else None


def _kisu_tables(doc):
    return [t for t in doc.tables if any(_KISU.search(c.text) for r in t.rows for c in r)]


def bump(doc: DSDDocument, shift: int) -> dict:
    """기수·연도를 shift 만큼 바꾼다. 반환: {'cells': n, 'attrs': n, 'kisu': (old, new)}."""
    out = {'cells': 0, 'attrs': 0, 'kisu': None}
    if not shift:
        return out
    nums = [int(m.group(2)) for t in _kisu_tables(doc) for r in t.rows for c in r
            for m in _KISU.finditer(c.text)]
    if nums:
        out['kisu'] = (max(nums), max(nums) + shift)

    def kisu_sub(m):
        return f'제{m.group(1)}{int(m.group(2)) + shift}{m.group(3)}{m.group(4) or ""}{m.group(5)}기'

    def year_sub(m):
        return f'{int(m.group(1)) + shift}{m.group(2)}'

    for t in _kisu_tables(doc):
        for r in t.rows:
            for c in r:
                raw = doc.get_raw(c)
                # 태그 밖 텍스트에만 적용 (속성값은 그대로)
                parts = re.split(r'(<[^>]+>)', raw)
                new_parts = []
                for p in parts:
                    if p.startswith('<'):
                        new_parts.append(p)
                        continue
                    p = _KISU.sub(kisu_sub, p)
                    for pat in _YEAR_PATS:
                        p = pat.sub(year_sub, p)
                    new_parts.append(p)
                new = ''.join(new_parts)
                if new != raw:
                    doc.set_raw(c, new)
                    out['cells'] += 1
    for m in _PERIOD_ATTR.finditer(doc.text):
        v = m.group(2)
        doc.replace_span(m.start(2), m.end(2), f'{int(v[:4]) + shift}{v[4:]}')
        out['attrs'] += 1
    return out
