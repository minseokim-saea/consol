"""DART DSD 파일(감사보고서) 읽기/쓰기 엔진.

DSD = zip(contents.xml, meta.xml). contents.xml 의 서식을 그대로 보존하기 위해
XML 파서로 재직렬화하지 않고, 표 셀의 '안쪽 텍스트 구간'만 원문에서 교체한다.
"""
from __future__ import annotations

import html
import re
import zipfile
from dataclasses import dataclass, field

CELL_TAGS = ("TD", "TE", "TH", "TU")
_RE_BLOCK = re.compile(r"<P(?:\s[^>]*)?>.*?</P>|<TITLE(?:\s[^>]*)?>.*?</TITLE>|<TABLE(?:\s[^>]*)?>.*?</TABLE>", re.S)
_RE_TR = re.compile(r"<TR(?:\s[^>]*)?>(.*?)</TR>", re.S)
_RE_CELL = re.compile(r"<(TD|TE|TH|TU)(\s[^>]*)?>(.*?)</\1>", re.S)
_RE_ATTR = re.compile(r'(\w[\w-]*)="([^"]*)"')
_RE_TAG = re.compile(r"<[^>]+>")


def plain(s: str) -> str:
    """셀/문단 XML 조각 -> 사람이 읽는 텍스트."""
    s = _RE_TAG.sub("", s)
    s = html.unescape(s).replace("&cr;", " ")
    return re.sub(r"\s+", " ", s).strip()


def norm(s: str) -> str:
    """비교용 정규화: 공백 제거, 주석표시 (*1) 제거."""
    s = plain(s) if "<" in s or "&" in s else s
    s = re.sub(r"\(\*\d*\)", "", s)
    return re.sub(r"\s+", "", s)


@dataclass
class Cell:
    tag: str
    attrs: dict
    inner_start: int  # contents 문자열 내 안쪽 텍스트 시작 위치
    inner_end: int
    text: str
    row: int = 0
    col: int = 0  # 격자 기준 열 위치
    colspan: int = 1
    rowspan: int = 1
    header: bool = False


@dataclass
class Table:
    index: int  # 문서 전체 표 순번 (1부터)
    start: int
    end: int
    anchor: str  # 표 앞의 가장 가까운 문단(제목) 텍스트
    caption: str  # 바로 앞 단위표 텍스트 (예: "(당기) (단위: 천원)")
    rows: list = field(default_factory=list)  # list[list[Cell]]
    ncols: int = 0
    header_rows: int = 0
    row_spans: list = field(default_factory=list)  # 행별 <TR>…</TR> 원문 위치

    # ---- 열/행 탐색 ----
    def col_headers(self) -> dict:
        """격자 열 -> 헤더 경로 문자열 (예: '당기말/유동')."""
        paths = {c: [] for c in range(self.ncols)}
        for r in range(self.header_rows):
            for cell in self.rows[r]:
                t = norm(cell.text)
                for c in range(cell.col, cell.col + cell.colspan):
                    if t and (not paths[c] or paths[c][-1] != t):
                        paths[c].append(t)
        return {c: "/".join(p) for c, p in paths.items()}

    def find_col(self, header: str) -> int | None:
        want = norm(header)
        heads = self.col_headers()
        for c, h in heads.items():  # 완전일치
            if h == want:
                return c
        for c, h in heads.items():  # 경로 끝 일치 (예: '유동' 대신 '당기말/유동')
            if h.endswith("/" + want) or h.split("/")[-1] == want:
                return c
        return None

    def body_rows(self):
        return self.rows[self.header_rows:]

    def find_row(self, label: str, nth: int = 1):
        want = norm(label)
        hit = 0
        for r in self.body_rows():
            first = next((c for c in r if c.col == 0), None)
            if first is not None and norm(first.text) == want:
                hit += 1
                if hit == nth:
                    return r
        return None

    def cell_at(self, row: list, col: int):
        for c in row:
            if c.col == col:
                return c
        return None

    def labels(self):
        out = []
        for r in self.body_rows():
            first = next((c for c in r if c.col == 0), None)
            out.append(first.text if first is not None else "")
        return out


def _parse_cell_text(inner: str) -> str:
    return plain(inner)


class DSDDocument:
    def __init__(self, path: str):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.members = {i.filename: z.read(i.filename) for i in z.infolist()}
            self.order = [i.filename for i in z.infolist()]
        raw = self.members["contents.xml"]
        self.bom = raw.startswith(b"\xef\xbb\xbf")
        self.text = raw.decode("utf-8-sig")
        self._edits: dict[tuple[int, int], str] = {}
        self.tables: list[Table] = []
        self._parse()

    # ---- 파싱 ----
    def _parse(self):
        t = self.text
        last_para = ""
        last_caption = ""
        idx = 0
        for m in _RE_BLOCK.finditer(t):
            s = m.group(0)
            if not s.startswith("<TABLE"):
                txt = plain(s)
                if txt:
                    last_para = txt
                    last_caption = ""
                continue
            idx += 1
            tbl = Table(idx, m.start(), m.end(), last_para, last_caption)
            self._parse_table(tbl, s, m.start())
            # 1행짜리 '(단위: 천원)' 표는 다음 표의 캡션으로 기억
            if len(tbl.rows) == 1 and "단위" in " ".join(c.text for c in tbl.rows[0]):
                last_caption = " ".join(c.text for c in tbl.rows[0] if c.text)
            else:
                last_caption = ""
            self.tables.append(tbl)

    def _parse_table(self, tbl: Table, s: str, base: int):
        thead_end = s.find("</THEAD>")
        occupied: dict[tuple[int, int], bool] = {}
        rows = []
        tbl.row_spans = []
        for ri, trm in enumerate(_RE_TR.finditer(s)):
            tbl.row_spans.append((base + trm.start(), base + trm.end()))
            in_head = thead_end != -1 and trm.start() < thead_end
            col = 0
            row = []
            for cm in _RE_CELL.finditer(trm.group(1)):
                while occupied.get((ri, col)):
                    col += 1
                attrs = dict(_RE_ATTR.findall(cm.group(2) or ""))
                cs = int(attrs.get("COLSPAN", 1) or 1)
                rs = int(attrs.get("ROWSPAN", 1) or 1)
                inner_start = base + trm.start(1) + cm.start(3)
                inner_end = base + trm.start(1) + cm.end(3)
                cell = Cell(cm.group(1), attrs, inner_start, inner_end, _parse_cell_text(cm.group(3)),
                            ri, col, cs, rs, in_head or cm.group(1) == "TH")
                for dr in range(rs):
                    for dc in range(cs):
                        occupied[(ri + dr, col + dc)] = True
                row.append(cell)
                col += cs
            rows.append(row)
        tbl.rows = rows
        tbl.ncols = max([c.col + c.colspan for r in rows for c in r] or [0])
        hr = 0
        for r in rows:
            if r and all(c.header for c in r):
                hr += 1
            else:
                break
        if hr == 0 and len(rows) > 1:
            hr = 1  # THEAD 가 없는 표는 첫 행을 헤더로 간주
        tbl.header_rows = hr

    # ---- 표 찾기 ----
    def find_tables(self, anchor: str):
        """앵커(주석 문단 일부)로 표 후보 목록 반환 (단위 캡션표 제외)."""
        want = norm(anchor)
        return [t for t in self.tables if want in norm(t.anchor) and len(t.rows) > 1]

    def find_table(self, anchor: str, nth: int = 1, caption: str = ""):
        cands = self.find_tables(anchor)
        if caption:
            cands = [t for t in cands if norm(caption) in norm(t.caption)]
        return cands[nth - 1] if 0 < nth <= len(cands) else None

    # ---- 편집 ----
    def set_cell(self, cell: Cell, value: str):
        self._edits[(cell.inner_start, cell.inner_end)] = html.escape(value, quote=False)
        cell.text = value

    def current_text(self, cell: Cell) -> str:
        return cell.text

    def raw_span(self, a: int, b: int) -> str:
        """원문 [a, b) 구간 — 그 안의 편집을 반영한 XML."""
        s = self.text[a:b]
        for (x, y), v in sorted(self._edits.items(), reverse=True):
            if a <= x and y <= b:
                s = s[:x - a] + v + s[y - a:]
        return s

    def get_raw(self, cell: Cell) -> str:
        return self._edits.get((cell.inner_start, cell.inner_end),
                               self.text[cell.inner_start:cell.inner_end])

    def set_raw(self, cell: Cell, raw: str):
        """셀 안쪽 XML 을 그대로 넣는다 (다른 셀 내용 복사용)."""
        self._edits[(cell.inner_start, cell.inner_end)] = raw
        cell.text = plain(raw)

    def replace_span(self, a: int, b: int, raw: str):
        """원문 [a, b) 를 raw 로 바꾼다 (속성값·여러 행 교체용)."""
        self._edits[(a, b)] = raw

    def replace_table(self, dst: Table, src: Table):
        """dst 표 전체를 src 표 XML(편집 반영)로 바꾼다."""
        self._edits[(dst.start, dst.end)] = self.raw_span(src.start, src.end)

    def render(self) -> str:
        """편집을 반영한 contents.xml 텍스트."""
        t = self.text
        for (a, b), v in sorted(self._edits.items(), reverse=True):
            t = t[:a] + v + t[b:]
        return t

    @classmethod
    def derived(cls, doc: 'DSDDocument') -> 'DSDDocument':
        """doc 의 편집을 반영한 상태를 새 문서로 다시 읽는다 (행 구조가 바뀐 뒤 칸을 찾을 때)."""
        new = cls.__new__(cls)
        new.path = doc.path
        new.members = dict(doc.members)
        new.order = list(doc.order)
        new.bom = doc.bom
        new.text = doc.render()
        new._edits = {}
        new.tables = []
        new._parse()
        return new

    def cell_tag_attr_span(self, cell: Cell, attr: str):
        """셀 여는 태그 안 attr 값의 원문 위치 (없으면 None)."""
        a = self.text.rfind('<', 0, cell.inner_start)
        m = re.compile(rf'\b{attr}="([^"]*)"').search(self.text, a, cell.inner_start)
        return (m.start(1), m.end(1)) if m else None

    def save(self, out_path: str):
        t = self.text
        spans = sorted(self._edits)
        for (a1, b1), (a2, b2) in zip(spans, spans[1:]):
            if a2 < b1:
                raise ValueError('겹치는 편집이 있습니다 (같은 표를 두 번 고치려 함).')
        for (a, b), v in sorted(self._edits.items(), reverse=True):
            t = t[:a] + v + t[b:]
        data = t.encode("utf-8")
        if self.bom:
            data = b"\xef\xbb\xbf" + data
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for name in self.order:
                z.writestr(name, data if name == "contents.xml" else self.members[name])


# ---- 숫자 서식 ----
def fmt_amount(v: float | None, decimals: int = 0) -> str:
    """DART 관행: 천단위 콤마, 음수 괄호, 0 은 '-'."""
    if v is None:
        return ""
    q = 10 ** decimals
    r = int(abs(v) * q + 0.5) / q  # 사사오입
    if r == 0:
        return "-"
    s = f"{r:,.{decimals}f}"
    return f"({s})" if v < 0 else s


def parse_amount(s: str) -> float | None:
    s = (s or "").strip()
    if s in ("", "-"):
        return 0.0 if s == "-" else None
    neg = s.startswith("(") and s.endswith(")")
    s2 = re.sub(r"[^\d.\-]", "", s)
    if not s2 or s2 in ("-", "."):
        return None
    try:
        v = float(s2)
    except ValueError:
        return None
    return -abs(v) if neg else v
