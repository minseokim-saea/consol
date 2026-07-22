"""
패키지 검증 모듈.

회사별 패키지 파일을 검사해 특이사항(불일치/누락 등) 회사만 추려서 반환.

검증 항목:
  · verify_wcf_diff():     WCF 시트 J436~O494 범위에서 Diff(N열) != 0 인 행 수집.
  · verify_wcf_accounts(): WCF 시트 J10~O426 범위에서
                           J(비용 Adjustment)에 수익코드(41/44/46), 또는
                           M(수익 Adjustment)에 비용코드(42/43/45/48 + 5xxx) 가
                           잘못 입력된 행 수집.
  · verify_wcf_signs():    WCF 시트 U/X/AB/AE(투자·재무 Cash-in/out 금액) 컬럼에서
                           동일 CF 코드 합계가 음수인 경우 수집.
  · verify_wcf_code_positive(): WCF 시트 P열의 특정 CF 코드(기본 CF2200401 퇴직금의 지급) 행들의
                           R열 금액 합계가 양수인지 검사 (해당 항목은 Cash-out이라 음수가 정상).

성능 메모:
  openpyxl로 .xlsm을 열면 sharedStrings/스타일 등 전체 파싱 비용이 크다 (파일당 수십 초~수 분).
  여기서는 zipfile + xml.etree.iterparse로 필요한 시트 XML만 스트림 파싱하여
  파일당 1~3초 수준으로 처리.
"""

import re
import zipfile
from xml.etree import ElementTree as ET


# J=10, K=11, L=12, M=13, N=14(Diff), O=15(Reason)
WCF_J_COL = 10
WCF_N_COL = 14
WCF_O_COL = 15
WCF_START_ROW = 436
WCF_END_ROW   = 494
WCF_DIFF_THRESHOLD = 1.0

# WCF Adjustment 입력 영역: J10~O426 (D5 COUNTIFS 기준).
# J = Adjustment(expense) CODE, K = 과목, L = 금액
# M = Adjustment(revenue) CODE, N = 과목, O = 금액
WCF_ADJ_START_ROW = 10
WCF_ADJ_END_ROW   = 426

# XML 네임스페이스 (Office Open XML)
_NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_NS_R    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_TAG_ROW = f'{{{_NS_MAIN}}}row'
_TAG_C   = f'{{{_NS_MAIN}}}c'
_TAG_V   = f'{{{_NS_MAIN}}}v'
_TAG_IS  = f'{{{_NS_MAIN}}}is'
_TAG_T   = f'{{{_NS_MAIN}}}t'
_TAG_SI  = f'{{{_NS_MAIN}}}si'
_TAG_SHEET = f'{{{_NS_MAIN}}}sheet'

_CELL_REF_RE = re.compile(r'^([A-Z]+)(\d+)$')


def _col_letters_to_index(letters):
    """엑셀 열 문자 → 1-based 인덱스. A=1, Z=26, AA=27."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _split_cell_ref(ref):
    """'B436' → ('B', 436). 매칭 실패 시 (None, None)."""
    m = _CELL_REF_RE.match(ref or '')
    if not m:
        return None, None
    return m.group(1), int(m.group(2))


def _load_shared_strings(zf):
    """sharedStrings.xml 을 stream 파싱해 인덱스 → 텍스트 배열로 반환."""
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return []
    shared = []
    with zf.open('xl/sharedStrings.xml') as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag != _TAG_SI:
                continue
            # rich text 포함 — 모든 <t> 텍스트 이어붙임
            texts = [t.text or '' for t in elem.iter(_TAG_T)]
            shared.append(''.join(texts))
            elem.clear()
    return shared


def _find_wcf_sheet_path(zf):
    """workbook.xml + rels 를 보고 'WCF' 시트의 XML 파일 경로 반환. 없으면 None."""
    # 1) workbook.xml: 시트 이름 → r:id
    if 'xl/workbook.xml' not in zf.namelist():
        return None
    with zf.open('xl/workbook.xml') as f:
        wb_tree = ET.parse(f)
    wb_root = wb_tree.getroot()
    rid = None
    for sheet in wb_root.iter(_TAG_SHEET):
        if sheet.get('name') == 'WCF':
            rid = sheet.get(f'{{{_NS_R}}}id')
            break
    if not rid:
        return None

    # 2) workbook.xml.rels: r:id → Target
    if 'xl/_rels/workbook.xml.rels' not in zf.namelist():
        return None
    with zf.open('xl/_rels/workbook.xml.rels') as f:
        rels_tree = ET.parse(f)
    target = None
    for rel in rels_tree.getroot():
        if rel.get('Id') == rid:
            target = rel.get('Target')
            break
    if not target:
        return None

    # 3) Target은 'worksheets/sheetN.xml' 같은 상대경로 → xl/ prefix
    return target if target.startswith('xl/') else 'xl/' + target


def _cell_value(c_elem, shared):
    """엑셀 셀(XML) → Python 값.
    숫자는 int/float, 문자열은 str. 비어있거나 못 읽으면 None.
    """
    t = c_elem.get('t')   # 's' | 'inlineStr' | 'str' | 'b' | 'n' | None
    if t == 's':
        v = c_elem.find(_TAG_V)
        if v is not None and v.text is not None:
            try:
                idx = int(v.text)
                if 0 <= idx < len(shared):
                    return shared[idx]
            except ValueError:
                pass
        return None
    if t == 'inlineStr':
        is_el = c_elem.find(_TAG_IS)
        if is_el is not None:
            texts = [t_el.text or '' for t_el in is_el.iter(_TAG_T)]
            return ''.join(texts) or None
        return None
    if t == 'str':
        v = c_elem.find(_TAG_V)
        return v.text if v is not None else None
    # 'n' or None (기본 숫자)
    v = c_elem.find(_TAG_V)
    if v is None or v.text is None:
        return None
    text = v.text
    try:
        f = float(text)
        if f == int(f):
            return int(f)
        return f
    except ValueError:
        return text


def verify_wcf_diff(file_path,
                    start_row=WCF_START_ROW, end_row=WCF_END_ROW,
                    j_col=WCF_J_COL, n_col=WCF_N_COL, o_col=WCF_O_COL,
                    threshold=WCF_DIFF_THRESHOLD):
    """WCF 시트의 J436:O494 범위에서 Diff != 0 인 행을 추출.

    반환:
      {
        'sheet_found': bool,
        'rows': [{'row': r, 'j': ..., 'k': ..., 'l': ..., 'm': ..., 'n': float, 'o': str}],
        'error': str | None,
      }
    """
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        return {'sheet_found': False, 'rows': [], 'error': f'파일 열기 실패: {e}'}

    try:
        sheet_path = _find_wcf_sheet_path(zf)
        if not sheet_path or sheet_path not in zf.namelist():
            return {'sheet_found': False, 'rows': [], 'error': None}

        # O열은 텍스트(Reason)이므로 sharedStrings 필요. J~M도 라벨일 수 있어 같이.
        shared = _load_shared_strings(zf)

        rows = []
        with zf.open(sheet_path) as f:
            for event, elem in ET.iterparse(f, events=('end',)):
                if elem.tag != _TAG_ROW:
                    continue
                r_attr = elem.get('r')
                if not r_attr:
                    elem.clear()
                    continue
                r = int(r_attr)
                if r < start_row:
                    elem.clear()
                    continue
                if r > end_row:
                    elem.clear()
                    break

                # 이 row의 J~O 셀만 추출
                vals = {}   # col_idx → value
                for c in elem.findall(_TAG_C):
                    ref = c.get('r')
                    letters, _ = _split_cell_ref(ref)
                    if letters is None:
                        continue
                    col_idx = _col_letters_to_index(letters)
                    if col_idx < j_col or col_idx > o_col:
                        continue
                    vals[col_idx] = _cell_value(c, shared)
                elem.clear()

                n_val = vals.get(n_col)
                if not isinstance(n_val, (int, float)):
                    continue
                if abs(float(n_val)) < threshold:
                    continue

                rows.append({
                    'row': r,
                    'j': vals.get(j_col),
                    'k': vals.get(j_col + 1),
                    'l': vals.get(j_col + 2),
                    'm': vals.get(j_col + 3),
                    'n': float(n_val),
                    'o': str(vals.get(o_col) or ''),
                })
    finally:
        zf.close()

    return {'sheet_found': True, 'rows': rows, 'error': None}


# ─── WCF 계정검증 ──────────────────────────────────────────────────────────
# 계정코드 분류 (COA 기준 + memory/project_cf_line_prefix_rule):
#   수익(revenue): PL 41xxxxx (매출), 44xxxxx (영업외수익/금융수익), 46xxxxx (중단영업)
#   비용(expense): PL 42xxxxx (매출원가), 43xxxxx (판관비), 45xxxxx (영업외비용),
#                  48xxxxx (법인세), MF 5xxxxxx (제조원가)
_REVENUE_PREFIXES = ('41', '44', '46')
_EXPENSE_PREFIXES_PL = ('42', '43', '45', '48')


def _classify_account(code):
    """계정코드 → 'revenue' | 'expense' | None (분류 불가).

    code 는 int 또는 str. 7자리 PL/MF 코드를 가정.
    """
    if code is None:
        return None
    s = str(code).strip()
    if not s or not s[0].isdigit():
        return None
    if s.startswith('5'):
        return 'expense'    # MF (제조원가)
    if s.startswith('4') and len(s) >= 2:
        prefix2 = s[:2]
        if prefix2 in _REVENUE_PREFIXES:
            return 'revenue'
        if prefix2 in _EXPENSE_PREFIXES_PL:
            return 'expense'
    return None


def verify_wcf_accounts(file_path,
                        start_row=WCF_ADJ_START_ROW, end_row=WCF_ADJ_END_ROW):
    """WCF 시트 J/M 열에 반대 성격 계정이 입력된 행을 추출.

    WCF 시트의 J 컬럼은 Adjustment(expense), M 컬럼은 Adjustment(revenue) 영역.
    따라서:
      · J(=비용 영역)에 수익 prefix(41/44/46) 코드가 들어가 있으면 오류
      · M(=수익 영역)에 비용 prefix(42/43/45/48 + 5xxx) 코드가 들어가 있으면 오류

    반환:
      {
        'sheet_found': bool,
        'rows': [{
            'row': r,
            'side': 'J' | 'M',
            'code': str, 'name': str,    # K(J일 때) 또는 N(M일 때)
            'amount': float | None,      # L(J일 때) 또는 O(M일 때)
            'expected': 'expense' | 'revenue',
            'actual':   'revenue' | 'expense',
        }],
        'error': str | None,
      }
    """
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        return {'sheet_found': False, 'rows': [], 'error': f'파일 열기 실패: {e}'}

    try:
        sheet_path = _find_wcf_sheet_path(zf)
        if not sheet_path or sheet_path not in zf.namelist():
            return {'sheet_found': False, 'rows': [], 'error': None}

        shared = _load_shared_strings(zf)

        # 컬럼 인덱스: J=10, K=11, L=12, M=13, N=14, O=15
        J, K, L, M, N, O = 10, 11, 12, 13, 14, 15

        rows = []
        with zf.open(sheet_path) as f:
            for event, elem in ET.iterparse(f, events=('end',)):
                if elem.tag != _TAG_ROW:
                    continue
                r_attr = elem.get('r')
                if not r_attr:
                    elem.clear()
                    continue
                r = int(r_attr)
                if r < start_row:
                    elem.clear()
                    continue
                if r > end_row:
                    elem.clear()
                    break

                vals = {}
                for c in elem.findall(_TAG_C):
                    ref = c.get('r')
                    letters, _ = _split_cell_ref(ref)
                    if letters is None:
                        continue
                    col_idx = _col_letters_to_index(letters)
                    if col_idx < J or col_idx > O:
                        continue
                    vals[col_idx] = _cell_value(c, shared)
                elem.clear()

                # J(비용) 쪽 검사: 수익코드가 들어가 있으면 오류
                j_code = vals.get(J)
                if j_code is not None:
                    cls = _classify_account(j_code)
                    if cls == 'revenue':
                        amt = vals.get(L)
                        rows.append({
                            'row': r,
                            'side': 'J',
                            'code': str(j_code),
                            'name': str(vals.get(K) or ''),
                            'amount': float(amt) if isinstance(amt, (int, float)) else None,
                            'expected': 'expense',
                            'actual': 'revenue',
                        })

                # M(수익) 쪽 검사: 비용코드가 들어가 있으면 오류
                m_code = vals.get(M)
                if m_code is not None:
                    cls = _classify_account(m_code)
                    if cls == 'expense':
                        amt = vals.get(O)
                        rows.append({
                            'row': r,
                            'side': 'M',
                            'code': str(m_code),
                            'name': str(vals.get(N) or ''),
                            'amount': float(amt) if isinstance(amt, (int, float)) else None,
                            'expected': 'revenue',
                            'actual': 'expense',
                        })
    finally:
        zf.close()

    return {'sheet_found': True, 'rows': rows, 'error': None}


# ─── WCF 부호 검증 ────────────────────────────────────────────────────────
# WCF 시트 컬럼 매핑 (1-based):
#   S=19/T=20/U=21  : 투자활동 Cash-in (S=code, T=과목, U=금액)
#   V=22/W=23/X=24  : 투자활동 Cash-out (V=code, W=과목, X=금액)
#   Z=26/AA=27/AB=28: 재무활동 Cash-in (Z=code, AA=과목, AB=금액)
#   AC=29/AD=30/AE=31: 재무활동 Cash-out (AC=code, AD=과목, AE=금액)
WCF_SIGN_GROUPS = [
    # (label,             code_col, name_col, amt_col)
    ('투자 Cash-in',       19, 20, 21),   # U열
    ('투자 Cash-out',      22, 23, 24),   # X열
    ('재무 Cash-in',       26, 27, 28),   # AB열
    ('재무 Cash-out',      29, 30, 31),   # AE열
]


def verify_wcf_signs(file_path,
                     start_row=WCF_ADJ_START_ROW, end_row=WCF_ADJ_END_ROW):
    """WCF 시트의 U/X/AB/AE 컬럼에서 CF 코드별 합계가 음수인 항목을 추출.

    각 컬럼은 같은 활동/방향(투자 Cash-in, 투자 Cash-out, 재무 Cash-in, 재무 Cash-out)의
    '금액' 입력 컬럼. 정상이라면 모두 양수여야 하므로 합계 음수는 입력 오류.

    반환:
      {
        'sheet_found': bool,
        'items': [{
            'column': 'U' | 'X' | 'AB' | 'AE',
            'label':  '투자 Cash-in' | ...,
            'code':   str,
            'name':   str,
            'sum':    float,   # < 0
            'rows':   [{'row': r, 'amount': float}],   # 해당 코드의 모든 입력행
        }],
        'error': str | None,
      }
    """
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        return {'sheet_found': False, 'items': [], 'error': f'파일 열기 실패: {e}'}

    try:
        sheet_path = _find_wcf_sheet_path(zf)
        if not sheet_path or sheet_path not in zf.namelist():
            return {'sheet_found': False, 'items': [], 'error': None}

        shared = _load_shared_strings(zf)

        # 모든 그룹의 컬럼을 한 번에 수집할 인덱스 셋
        all_cols = set()
        for _, c_col, n_col, a_col in WCF_SIGN_GROUPS:
            all_cols.update([c_col, n_col, a_col])
        min_col, max_col = min(all_cols), max(all_cols)

        # 그룹별 누적: {(group_idx, code) -> {'name': str, 'sum': float, 'rows': [...]}
        groups = [
            {}  # group_idx 별 dict
            for _ in WCF_SIGN_GROUPS
        ]

        with zf.open(sheet_path) as f:
            for event, elem in ET.iterparse(f, events=('end',)):
                if elem.tag != _TAG_ROW:
                    continue
                r_attr = elem.get('r')
                if not r_attr:
                    elem.clear()
                    continue
                r = int(r_attr)
                if r < start_row:
                    elem.clear()
                    continue
                if r > end_row:
                    elem.clear()
                    break

                vals = {}
                for c in elem.findall(_TAG_C):
                    ref = c.get('r')
                    letters, _ = _split_cell_ref(ref)
                    if letters is None:
                        continue
                    col_idx = _col_letters_to_index(letters)
                    if col_idx < min_col or col_idx > max_col:
                        continue
                    if col_idx not in all_cols:
                        continue
                    vals[col_idx] = _cell_value(c, shared)
                elem.clear()

                for gi, (_label, c_col, n_col, a_col) in enumerate(WCF_SIGN_GROUPS):
                    code = vals.get(c_col)
                    amt = vals.get(a_col)
                    if code is None or not isinstance(amt, (int, float)):
                        continue
                    code_s = str(code).strip()
                    if not code_s:
                        continue
                    bucket = groups[gi].setdefault(code_s, {
                        'name': str(vals.get(n_col) or ''),
                        'sum': 0.0,
                        'rows': [],
                    })
                    bucket['sum'] += float(amt)
                    bucket['rows'].append({'row': r, 'amount': float(amt)})
                    # 첫 등장한 빈 name 이후에 채워진 name 이 있으면 갱신
                    if not bucket['name']:
                        nm = vals.get(n_col)
                        if nm:
                            bucket['name'] = str(nm)
    finally:
        zf.close()

    # 음수 합계만 추출
    col_letter = {21: 'U', 24: 'X', 28: 'AB', 31: 'AE'}
    items = []
    for gi, (label, _c_col, _n_col, a_col) in enumerate(WCF_SIGN_GROUPS):
        for code, info in groups[gi].items():
            if info['sum'] < 0:
                items.append({
                    'column': col_letter.get(a_col, ''),
                    'label': label,
                    'code': code,
                    'name': info['name'],
                    'sum': info['sum'],
                    'rows': info['rows'],
                })

    # 정렬: 컬럼순(U→X→AB→AE) → 합계 오름차순(가장 큰 음수 먼저)
    col_rank = {'U': 0, 'X': 1, 'AB': 2, 'AE': 3}
    items.sort(key=lambda x: (col_rank.get(x['column'], 9), x['sum']))

    return {'sheet_found': True, 'items': items, 'error': None}


# ─── WCF 특정 코드 부호 검증 ────────────────────────────────────────────────
# P=16 (Code), Q=17 (과목), R=18 (금액)
# 자산·부채 변동 영역(P/Q/R) — Cash-out 성격 항목은 음수로 입력해야 정상이므로
# 합계 양수면 부호 오입력으로 판정.
WCF_PQR_CODE_COL = 16
WCF_PQR_NAME_COL = 17
WCF_PQR_AMT_COL  = 18

# 검출 대상 기본 코드: CF2200401 (퇴직금의 지급, Cash-out)
WCF_SEVERANCE_CODE = 'CF2200401'


def verify_wcf_code_positive(file_path,
                             target_code=WCF_SEVERANCE_CODE,
                             start_row=WCF_ADJ_START_ROW, end_row=WCF_ADJ_END_ROW,
                             code_col=WCF_PQR_CODE_COL,
                             name_col=WCF_PQR_NAME_COL,
                             amt_col=WCF_PQR_AMT_COL):
    """WCF 시트 P열이 target_code 인 행들의 R열 금액 합계가 양수인지 검사.

    CF2200401(퇴직금의 지급) 등은 Cash-out 성격이라 음수가 정상.
    합계가 0 보다 크면 부호 오입력으로 판정.

    반환:
      {
        'sheet_found': bool,
        'target_code': str,
        'found': bool,           # 해당 코드가 시트에 존재했는지
        'name': str,             # 첫 행의 Q열 과목명
        'sum': float,            # R열 합계
        'rows': [{'row': r, 'amount': float|None}],
        'is_positive': bool,     # sum > 0
        'error': str | None,
      }
    """
    out = {
        'sheet_found': False, 'target_code': target_code, 'found': False,
        'name': '', 'sum': 0.0, 'rows': [], 'is_positive': False, 'error': None,
    }
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        out['error'] = f'파일 열기 실패: {e}'
        return out

    try:
        sheet_path = _find_wcf_sheet_path(zf)
        if not sheet_path or sheet_path not in zf.namelist():
            return out

        shared = _load_shared_strings(zf)
        out['sheet_found'] = True

        rows = []
        name = ''
        total = 0.0
        with zf.open(sheet_path) as f:
            for event, elem in ET.iterparse(f, events=('end',)):
                if elem.tag != _TAG_ROW:
                    continue
                r_attr = elem.get('r')
                if not r_attr:
                    elem.clear()
                    continue
                r = int(r_attr)
                if r < start_row:
                    elem.clear()
                    continue
                if r > end_row:
                    elem.clear()
                    break

                vals = {}
                for c in elem.findall(_TAG_C):
                    ref = c.get('r')
                    letters, _ = _split_cell_ref(ref)
                    if letters is None:
                        continue
                    col_idx = _col_letters_to_index(letters)
                    if col_idx not in (code_col, name_col, amt_col):
                        continue
                    vals[col_idx] = _cell_value(c, shared)
                elem.clear()

                p_val = vals.get(code_col)
                if p_val is None:
                    continue
                if str(p_val).strip() != target_code:
                    continue

                amt = vals.get(amt_col)
                amt_f = float(amt) if isinstance(amt, (int, float)) else None
                rows.append({'row': r, 'amount': amt_f})
                if amt_f is not None:
                    total += amt_f
                if not name:
                    nm = vals.get(name_col)
                    if nm:
                        name = str(nm)
    finally:
        zf.close()

    out['found'] = bool(rows)
    out['name'] = name
    out['sum'] = total
    out['rows'] = rows
    out['is_positive'] = bool(rows) and total > 0
    return out


# ─── 퇴직급여 검증 (L3/L3-1 입력 vs 회계기준) ────────────────────────────────
# Cover!B27 = 회계기준 (예: 'K-IFRS', 'K-GAAP')
#   K-IFRS  → L3-1(확정급여채무) 사용. L3에 입력값 있으면 오입력
#   K-GAAP  → L3(퇴직급여충당부채) 사용. L3-1에 입력값 있으면 오입력
# 입력값 판정: C열(=3) 에 0이 아닌 숫자가 있는 행이 1개 이상
RB_COVER_SHEET = 'Cover'
RB_COVER_STANDARD_ROW = 27
RB_COVER_STANDARD_COL = 2  # B
RB_INPUT_COL = 3           # C
RB_L3_SCAN_END = 50        # L3는 r51 안내문까지
RB_L31_SCAN_END = 72       # L3-1은 r72 안내문까지
RB_LABEL_COL = 2           # B (행 라벨)


def _find_sheet_path_by_name(zf, sheet_name):
    """workbook.xml + rels 에서 임의 시트 경로 반환."""
    if 'xl/workbook.xml' not in zf.namelist():
        return None
    with zf.open('xl/workbook.xml') as f:
        wb_tree = ET.parse(f)
    rid = None
    for sheet in wb_tree.getroot().iter(_TAG_SHEET):
        if sheet.get('name') == sheet_name:
            rid = sheet.get(f'{{{_NS_R}}}id')
            break
    if not rid:
        return None
    if 'xl/_rels/workbook.xml.rels' not in zf.namelist():
        return None
    with zf.open('xl/_rels/workbook.xml.rels') as f:
        rels_tree = ET.parse(f)
    target = None
    for rel in rels_tree.getroot():
        if rel.get('Id') == rid:
            target = rel.get('Target')
            break
    if not target:
        return None
    return target if target.startswith('xl/') else 'xl/' + target


def _read_cover_standard(zf, shared):
    """Cover!B27 값 반환 (str | None)."""
    sheet_path = _find_sheet_path_by_name(zf, RB_COVER_SHEET)
    if not sheet_path or sheet_path not in zf.namelist():
        return None
    with zf.open(sheet_path) as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag != _TAG_ROW:
                continue
            r_attr = elem.get('r')
            if not r_attr:
                elem.clear()
                continue
            r = int(r_attr)
            if r < RB_COVER_STANDARD_ROW:
                elem.clear()
                continue
            if r > RB_COVER_STANDARD_ROW:
                elem.clear()
                break
            for c in elem.findall(_TAG_C):
                letters, _ = _split_cell_ref(c.get('r'))
                if letters is None:
                    continue
                if _col_letters_to_index(letters) != RB_COVER_STANDARD_COL:
                    continue
                val = _cell_value(c, shared)
                elem.clear()
                if val is None:
                    return None
                return str(val).strip()
            elem.clear()
    return None


def _classify_standard(raw):
    """Cover!B27 값 → 'ifrs' | 'gaap' | None.

    'K-IFRS' / 'IFRS' 포함 → ifrs
    'K-GAAP' / 'GAAP'  포함 → gaap
    그 외 (빈 값/미상) → None
    """
    if not raw:
        return None
    s = str(raw).upper().replace('-', '').replace(' ', '')
    if 'IFRS' in s:
        return 'ifrs'
    if 'GAAP' in s:
        return 'gaap'
    return None


def _collect_input_rows(zf, shared, sheet_name, end_row,
                        input_col=RB_INPUT_COL, label_col=RB_LABEL_COL):
    """주어진 시트에서 C열에 0이 아닌 숫자가 있는 행을 수집.

    반환: (sheet_found: bool, rows: [{'row': int, 'label': str, 'value': float}])
    """
    rows = []
    sheet_path = _find_sheet_path_by_name(zf, sheet_name)
    if not sheet_path or sheet_path not in zf.namelist():
        return False, rows
    with zf.open(sheet_path) as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag != _TAG_ROW:
                continue
            r_attr = elem.get('r')
            if not r_attr:
                elem.clear()
                continue
            r = int(r_attr)
            if r > end_row:
                elem.clear()
                break
            vals = {}
            for c in elem.findall(_TAG_C):
                letters, _ = _split_cell_ref(c.get('r'))
                if letters is None:
                    continue
                col_idx = _col_letters_to_index(letters)
                if col_idx not in (label_col, input_col):
                    continue
                vals[col_idx] = _cell_value(c, shared)
            elem.clear()

            v = vals.get(input_col)
            if not isinstance(v, (int, float)):
                continue
            if isinstance(v, bool):
                continue
            if float(v) == 0.0:
                continue
            label = vals.get(label_col)
            rows.append({
                'row': r,
                'label': str(label or '').strip() if label else '',
                'value': float(v),
            })
    return True, rows


def verify_retirement_benefit(file_path):
    """L3/L3-1 시트의 입력값과 회계기준(Cover!B27)의 정합성 검사.

    반환:
      {
        'standard_raw': str | None,    # Cover!B27 원본 값
        'standard': 'ifrs' | 'gaap' | None,
        'cover_found': bool,           # Cover 시트 존재 여부
        'target_sheet': 'L3' | 'L3-1' | None,   # 잘못 입력된 (검사 대상) 시트
        'target_sheet_found': bool,
        'rows': [{'row': int, 'label': str, 'value': float}],
        'is_misplaced': bool,
        'error': str | None,
      }
    """
    out = {
        'standard_raw': None, 'standard': None, 'cover_found': False,
        'target_sheet': None, 'target_sheet_found': False,
        'rows': [], 'is_misplaced': False, 'error': None,
    }
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        out['error'] = f'파일 열기 실패: {e}'
        return out

    try:
        shared = _load_shared_strings(zf)
        raw = _read_cover_standard(zf, shared)
        out['standard_raw'] = raw
        out['cover_found'] = raw is not None or (
            _find_sheet_path_by_name(zf, RB_COVER_SHEET) is not None)
        out['standard'] = _classify_standard(raw)
        if out['standard'] is None:
            # 회계기준 미상 → 검증 skip
            return out

        if out['standard'] == 'ifrs':
            # IFRS 회사는 L3 사용 금지 → L3 입력값 검사
            out['target_sheet'] = 'L3'
            found, rows = _collect_input_rows(
                zf, shared, 'L3', RB_L3_SCAN_END)
        else:
            # K-GAAP 회사는 L3-1 사용 금지 → L3-1 입력값 검사
            out['target_sheet'] = 'L3-1'
            found, rows = _collect_input_rows(
                zf, shared, 'L3-1', RB_L31_SCAN_END)

        out['target_sheet_found'] = found
        out['rows'] = rows
        out['is_misplaced'] = bool(rows)
    finally:
        zf.close()

    return out


# ─── CF3 유동성장기차입금 신규차입 검증 ──────────────────────────────────────
# CF3 시트 "1. 증감내역(Details of Changes)" 표 구조:
#   B = Account title (블록 시작행에 7자리 코드, 다음 행들엔 계정명 텍스트)
#   C = 증감내역 (한글 라벨: 기초금액/신규차입/상환/유동성대체/기말금액 ...)
#   E = 연결범위회사(Affiliates), F = 제3자(The 3rd Party), G = Total
# 유동성장기차입금 블록은 코드 2100201('유동성장기차입금' / 'Current Portion Of
# Long-Term Debt')로 시작하며, 바로 아래 장기차입금(2200101) 블록에도 동일한
# '신규차입' 행이 있으므로 2100201 블록 범위 안의 '신규차입' 행만 대상으로 한다.
# 유동성장기차입금은 장기차입금의 유동성대체분이라 신규차입이 잡히면 이상치.
CF3_SHEET = 'CF3'
CF3_CURRENT_PORTION_CODE = '2100201'   # 유동성장기차입금
CF3_CODE_COL   = 2   # B
CF3_TYPE_COL   = 3   # C (증감내역)
CF3_AFFIL_COL  = 5   # E (연결범위회사)
CF3_THIRD_COL  = 6   # F (제3자)
CF3_TOTAL_COL  = 7   # G (Total)
CF3_NEW_BORROW_LABEL = '신규차입'


def verify_cf3_current_portion_new_borrowing(file_path):
    """CF3 시트 유동성장기차입금(2100201) 블록의 '신규차입' 행 금액을 추출.

    Total(G) 금액이 0이 아니면 이상치(is_flagged=True)로 판정.
    (유동성장기차입금은 장기차입금의 유동성대체분이라 신규차입이 비정상)

    반환:
      {
        'sheet_found': bool,
        'found': bool,           # 유동성장기차입금 블록 + 신규차입 행 발견 여부
        'row': int | None,       # 신규차입 행 번호
        'new_borrow': float,     # G열 Total
        'affiliates': float,     # E열 연결범위회사
        'third_party': float,    # F열 제3자
        'is_flagged': bool,      # new_borrow != 0
        'error': str | None,
      }
    """
    out = {
        'sheet_found': False, 'found': False, 'row': None,
        'new_borrow': 0.0, 'affiliates': 0.0, 'third_party': 0.0,
        'is_flagged': False, 'error': None,
    }
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        out['error'] = f'파일 열기 실패: {e}'
        return out

    try:
        sheet_path = _find_sheet_path_by_name(zf, CF3_SHEET)
        if not sheet_path or sheet_path not in zf.namelist():
            return out

        shared = _load_shared_strings(zf)
        out['sheet_found'] = True

        cols = (CF3_CODE_COL, CF3_TYPE_COL, CF3_AFFIL_COL,
                CF3_THIRD_COL, CF3_TOTAL_COL)
        in_block = False
        with zf.open(sheet_path) as f:
            for event, elem in ET.iterparse(f, events=('end',)):
                if elem.tag != _TAG_ROW:
                    continue
                r_attr = elem.get('r')
                if not r_attr:
                    elem.clear()
                    continue
                r = int(r_attr)

                vals = {}
                for c in elem.findall(_TAG_C):
                    letters, _ = _split_cell_ref(c.get('r'))
                    if letters is None:
                        continue
                    col_idx = _col_letters_to_index(letters)
                    if col_idx not in cols:
                        continue
                    vals[col_idx] = _cell_value(c, shared)
                elem.clear()

                # 블록 경계 추적: B열에 7자리 코드가 나오면 블록 전환
                b = vals.get(CF3_CODE_COL)
                if isinstance(b, (int, float)) and not isinstance(b, bool):
                    code_s = str(int(b))
                    if code_s == CF3_CURRENT_PORTION_CODE:
                        in_block = True
                    elif len(code_s) >= 7:
                        in_block = False

                if not in_block:
                    continue

                c_label = vals.get(CF3_TYPE_COL)
                if not c_label or CF3_NEW_BORROW_LABEL not in str(c_label):
                    continue

                # 유동성장기차입금 블록의 '신규차입' 행 발견
                g = vals.get(CF3_TOTAL_COL)
                e = vals.get(CF3_AFFIL_COL)
                ff = vals.get(CF3_THIRD_COL)
                out['found'] = True
                out['row'] = r
                out['new_borrow'] = float(g) if isinstance(g, (int, float)) else 0.0
                out['affiliates'] = float(e) if isinstance(e, (int, float)) else 0.0
                out['third_party'] = float(ff) if isinstance(ff, (int, float)) else 0.0
                break   # 블록당 신규차입 행은 1개
    finally:
        zf.close()

    out['is_flagged'] = out['found'] and out['new_borrow'] != 0
    return out


# ─── CF1/CF2/CF3 기타증감 내용 입력 검증 ─────────────────────────────────────
# CF1/CF2/CF3 시트는 공통적으로 "2. 기타증감 내용(Details of Other transfer)"
# 섹션을 가진다. 섹션은 B열에 라벨 "2. 기타증감 내용..."으로 시작하고,
# "3. 유동성검증..." 행에서 끝난다(CF2처럼 3번 섹션이 없으면 시트 끝까지).
# 섹션 내부 표 헤더(공통):
#   B = 계정명(Account Name), C = 상대계정명(Counter account name),
#   D = 상대계정코드(Counter account code), E = 금액(Amount), F = 상세내용
# 각 계정 블록 = [코드행(B=7자리 숫자)] + [한글명 라벨행] + [영문명 라벨행]
#                + [입력영역(사용자 기재)] + ['Total' 행]
# 빈 패키지는 입력영역이 비어 있고 Total 행의 E(금액)=0.
#
# "내용이 있다"의 판정: 섹션 내부에서 헤더/라벨/Total/코드 행을 제외한 입력행 중
#   · C(상대계정명) 또는 D(상대계정코드) 또는 F(상세내용)에 텍스트가 있거나
#   · E(금액)에 0이 아닌 숫자가 있으면
# 사용자가 기타증감 표에 무언가 기재한 것으로 본다(금액 없이 설명만 있어도 포함).
CF_OTHER_TRANSFER_SHEETS = ('CF1', 'CF2', 'CF3')
COT_SECTION_START_KW = '기타증감 내용'   # "2. 기타증감 내용(Details of Other transfer)"
COT_SECTION_END_KW   = '유동성검증'      # "3. 유동성검증(...)"
COT_NAME_COL    = 2   # B 계정명
COT_COUNTER_COL = 3   # C 상대계정명
COT_CODE_COL    = 4   # D 상대계정코드
COT_AMT_COL     = 5   # E 금액
COT_DETAIL_COL  = 6   # F 상세내용


def _scan_other_transfer_one_sheet(zf, shared, sheet_name):
    """한 시트(CF1/CF2/CF3)의 '2. 기타증감 내용' 섹션을 스캔.

    반환: (sheet_found, entries)
      entries: [{'row', 'name', 'counter', 'code', 'amount', 'detail'}]
               — 사용자가 기재한 입력행만. 없으면 빈 리스트.
    """
    entries = []
    sheet_path = _find_sheet_path_by_name(zf, sheet_name)
    if not sheet_path or sheet_path not in zf.namelist():
        return False, entries

    cols = (COT_NAME_COL, COT_COUNTER_COL, COT_CODE_COL, COT_AMT_COL, COT_DETAIL_COL)
    in_section = False
    with zf.open(sheet_path) as f:
        for _ev, elem in ET.iterparse(f, events=('end',)):
            if elem.tag != _TAG_ROW:
                continue
            if not elem.get('r'):
                elem.clear()
                continue
            r = int(elem.get('r'))
            vals = {}
            for c in elem.findall(_TAG_C):
                letters, _ = _split_cell_ref(c.get('r'))
                if letters is None:
                    continue
                ci = _col_letters_to_index(letters)
                if ci in cols:
                    vals[ci] = _cell_value(c, shared)
            elem.clear()

            b = vals.get(COT_NAME_COL)
            b_str = str(b).strip() if b is not None else ''

            # 섹션 진입/종료 추적
            if not in_section:
                if COT_SECTION_START_KW in b_str:
                    in_section = True
                continue
            if COT_SECTION_END_KW in b_str:
                break   # 섹션 끝

            # 섹션 내부 — 헤더/라벨/Total/코드/계정명 행은 입력행이 아니다.
            #   · 헤더행: B에 '계정명' 포함
            #   · 블록 코드행: B가 7자리 숫자
            #   · 한글명/영문명 라벨행: B에만 텍스트, C~F는 비어 있음
            #   · Total 행: B == 'Total'
            if '계정명' in b_str:
                continue
            if b_str == 'Total':
                continue

            counter = vals.get(COT_COUNTER_COL)
            code    = vals.get(COT_CODE_COL)
            amt     = vals.get(COT_AMT_COL)
            detail  = vals.get(COT_DETAIL_COL)

            counter_s = str(counter).strip() if counter is not None else ''
            code_s    = str(code).strip()    if code    is not None else ''
            detail_s  = str(detail).strip()  if detail  is not None else ''
            amt_f = float(amt) if isinstance(amt, (int, float)) and not isinstance(amt, bool) else None

            # 입력행 판정: 상대계정명/코드/상세내용 텍스트가 있거나, 금액이 0이 아님
            has_text = bool(counter_s or code_s or detail_s)
            has_amt = amt_f is not None and amt_f != 0
            if not (has_text or has_amt):
                continue

            entries.append({
                'row': r,
                'name': b_str,
                'counter': counter_s,
                'code': code_s,
                'amount': amt_f,
                'detail': detail_s,
            })

    return True, entries


def verify_cf_other_transfer(file_path):
    """CF1/CF2/CF3의 "2. 기타증감 내용" 섹션에 기재된 내용이 있는지 검사.

    세 시트 중 하나라도 입력행이 있으면 이상치(is_flagged=True).

    반환:
      {
        'sheet_found': bool,        # CF1/CF2/CF3 중 하나라도 존재
        'found': bool,              # 기타증감 내용이 1건 이상
        'by_sheet': {sheet: [entries]},   # 시트별 입력행 (있는 것만)
        'entries_count': int,
        'sheets_with_content': [sheet, ...],
        'is_flagged': bool,
        'error': str | None,
      }
    """
    out = {
        'sheet_found': False, 'found': False, 'by_sheet': {},
        'entries_count': 0, 'sheets_with_content': [],
        'is_flagged': False, 'error': None,
    }
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        out['error'] = f'파일 열기 실패: {e}'
        return out

    try:
        shared = _load_shared_strings(zf)
        any_sheet = False
        for sheet_name in CF_OTHER_TRANSFER_SHEETS:
            found, entries = _scan_other_transfer_one_sheet(zf, shared, sheet_name)
            if found:
                any_sheet = True
            if entries:
                out['by_sheet'][sheet_name] = entries
                out['sheets_with_content'].append(sheet_name)
                out['entries_count'] += len(entries)
        out['sheet_found'] = any_sheet
    finally:
        zf.close()

    out['found'] = out['entries_count'] > 0
    out['is_flagged'] = out['found']
    return out


# ─── CF4 / CF4-1 "기타변동 내용" 입력 검증 (유형자산 / 무형자산) ───────────────
# CF4(CF정보-유형자산) 시트의 "6. 기타변동 내용(Other Transfer Details)"과
# CF4-1(CF정보-무형자산) 시트의 "3. 기타변동 내용(Other Transfer Details)"은
# 표 구조가 동일하다. 섹션 라벨은 A열에 "N. 기타변동 내용..."으로 시작하며
# 각 시트의 마지막 섹션이라 시트 끝까지가 섹션 범위다.
# 표 헤더:
#   B = No., C = 계정코드(Account code), D = 계정명(Account title/Name),
#   E = 금액(Amount), F = 상대계정명(Counter account name),
#   G = 상대계정코드(Counter account code)
# 각 항목(No. 1,2,3...)은 B열에 번호가 있는 행으로 시작하고, 그 아래 여러 입력행과
# 마지막에 금액만 채워진 "자동 소계행"(텍스트 없음)이 따라온다.
#
# "내용이 있다"의 판정: 섹션 내부에서 헤더/번호행/소계행을 제외하고,
#   계정코드(C)·계정명(D)·상대계정명(F)·상대계정코드(G) 중 하나라도 텍스트가 있으면
#   사용자가 기타변동 내용을 기재한 것으로 본다.
#   (금액만 채워진 행은 자동 소계라 입력으로 보지 않음 → 중복 집계 방지)
CF4_OT_SECTION_KW   = '기타변동 내용'   # A열 "N. 기타변동 내용(Other Transfer Details)"
CF4_OT_SECTION_COL  = 1   # A 섹션 라벨
CF4_OT_NO_COL       = 2   # B No.
CF4_OT_CODE_COL     = 3   # C 계정코드
CF4_OT_NAME_COL     = 4   # D 계정명
CF4_OT_AMT_COL      = 5   # E 금액
CF4_OT_CNAME_COL    = 6   # F 상대계정명
CF4_OT_CCODE_COL    = 7   # G 상대계정코드


def _verify_asset_other_transfer(file_path, sheet_name):
    """CF4/CF4-1 "기타변동 내용" 섹션에 기재된 입력행이 있는지 검사.

    입력행이 1건 이상이면 이상치(is_flagged=True).

    반환:
      {
        'sheet_found': bool,        # 대상 시트 존재
        'found': bool,              # 기타변동 내용 입력행 1건 이상
        'entries': [               # 사용자 입력행만
            {'row', 'code', 'name', 'amount', 'counter_name', 'counter_code'}
        ],
        'entries_count': int,
        'is_flagged': bool,
        'error': str | None,
      }
    """
    out = {'sheet_found': False, 'found': False, 'entries': [],
           'entries_count': 0, 'is_flagged': False, 'error': None}
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        out['error'] = f'파일 열기 실패: {e}'
        return out

    try:
        sheet_path = _find_sheet_path_by_name(zf, sheet_name)
        if not sheet_path or sheet_path not in zf.namelist():
            return out  # sheet_found=False 유지
        shared = _load_shared_strings(zf)
        out['sheet_found'] = True

        cols = (CF4_OT_SECTION_COL, CF4_OT_NO_COL, CF4_OT_CODE_COL,
                CF4_OT_NAME_COL, CF4_OT_AMT_COL, CF4_OT_CNAME_COL, CF4_OT_CCODE_COL)
        in_section = False
        with zf.open(sheet_path) as f:
            for _ev, elem in ET.iterparse(f, events=('end',)):
                if elem.tag != _TAG_ROW:
                    continue
                if not elem.get('r'):
                    elem.clear()
                    continue
                r = int(elem.get('r'))
                vals = {}
                for c in elem.findall(_TAG_C):
                    letters, _ = _split_cell_ref(c.get('r'))
                    if letters is None:
                        continue
                    ci = _col_letters_to_index(letters)
                    if ci in cols:
                        vals[ci] = _cell_value(c, shared)
                elem.clear()

                a_str = str(vals.get(CF4_OT_SECTION_COL) or '').strip()
                if not in_section:
                    if CF4_OT_SECTION_KW in a_str:
                        in_section = True
                    continue

                # 섹션 내부 — 헤더행 스킵
                #   B='No.'(CF4/CF4-1 공통) 또는 C가 계정코드 헤더('계정코드'/'Account code')
                b = vals.get(CF4_OT_NO_COL)
                b_str = str(b).strip() if b is not None else ''
                code_hdr = str(vals.get(CF4_OT_CODE_COL) or '')
                if b_str == 'No.' or '계정코드' in code_hdr or 'Account code' in code_hdr:
                    continue

                code  = vals.get(CF4_OT_CODE_COL)
                name  = vals.get(CF4_OT_NAME_COL)
                amt   = vals.get(CF4_OT_AMT_COL)
                cname = vals.get(CF4_OT_CNAME_COL)
                ccode = vals.get(CF4_OT_CCODE_COL)

                code_s  = str(code).strip()  if code  is not None else ''
                name_s  = str(name).strip()  if name  is not None else ''
                cname_s = str(cname).strip() if cname is not None else ''
                ccode_s = str(ccode).strip() if ccode is not None else ''
                amt_f = float(amt) if isinstance(amt, (int, float)) and not isinstance(amt, bool) else None

                # 입력행 판정: 텍스트 필드(계정코드/계정명/상대계정명/상대계정코드) 중
                # 하나라도 채워져 있어야 입력. 금액만 있는 행은 자동 소계라 제외.
                has_text = bool(code_s or name_s or cname_s or ccode_s)
                if not has_text:
                    continue

                out['entries'].append({
                    'row': r,
                    'code': code_s,
                    'name': name_s,
                    'amount': amt_f,
                    'counter_name': cname_s,
                    'counter_code': ccode_s,
                })
    finally:
        zf.close()

    out['entries_count'] = len(out['entries'])
    out['found'] = out['entries_count'] > 0
    out['is_flagged'] = out['found']
    return out


def verify_cf4_other_transfer(file_path):
    """CF4 "6. 기타변동 내용" 입력 검증 (유형자산). 반환 형식은
    _verify_asset_other_transfer 참고."""
    return _verify_asset_other_transfer(file_path, 'CF4')


def verify_cf41_other_transfer(file_path):
    """CF4-1 "3. 기타변동 내용" 입력 검증 (무형자산). 반환 형식은
    _verify_asset_other_transfer 참고."""
    return _verify_asset_other_transfer(file_path, 'CF4-1')


# ─── GAAP 이익잉여금(미처분) 롤포워드 + 보험수리적손익 변동 검증 ──────────────
# PY 시트 GAAP Diff 표: M열(13)=코드, N열(14)=전기(PY), O열(15)=당기(CY).
#   대상 코드: 3500103 보험수리적손익 / 3500104 미처분이익잉여금 / 3500105 당기순이익
# ① 롤포워드(하드): 미처분이익잉여금이 전기에서 정상 이월됐는지.
#    회사마다 보험수리적손익(OCI) 처리가 달라 두 방식 중 하나만 맞으면 정상으로 본다:
#      roll1 = 당기 미처분 − (전기 미처분 + 전기 당기순이익)          [OCI 미이월: PSC 형]
#      roll2 = (당기 미처분+보험수리적) − (전기 미처분+보험수리적 + 전기 순이익) [OCI→미처분 이월: PWT 형]
#    → roll1==0 또는 roll2==0 이면 정상. 둘 다 != 0 이면 롤포워드 오류.
# ② 보험수리적손익 변동(검토): 전기 대비 당기 값이 달라지면 표시(의도한 조정인지 확인).
#    OCI 는 당기 정당한 움직임이 있을 수 있어 하드 오류가 아니라 검토 대상.
PY_GAAP_SHEET = 'PY'
_GAAP_RE_TOL = 1.0   # 원 단위 허용오차 (환산 반올림/소수 dust 무시)


def _num_or_zero(x):
    return float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0


def verify_gaap_retained_earnings(file_path):
    """PY 시트 GAAP Diff로 미처분이익잉여금 롤포워드 + 보험수리적손익 변동 검증.

    반환:
      {'sheet_found','found',
       'cur_re','prior_re','prior_ni',
       'roll_diff','roll_diff2','roll_ok',
       'cur_oci','prior_oci','oci_change',
       'is_flagged','severity','error'}
      순수 이익잉여금 = 이익준비금+임의적립금+미처분 (비지배지분 제외)
      roll_diff  = 당기 순수RE − (전기 순수RE + 전기 순이익)              [OCI 미이월형]
      roll_diff2 = (당기 순수RE+보험) − (전기 순수RE+보험 + 전기 순이익)   [OCI→미처분 이월형]
      roll_ok    = roll_diff 또는 roll_diff2 중 하나가 0(허용오차 내)
      severity: 'error'(롤포워드 오류) | 'review'(보험수리적 변동) | None
    """
    out = {'sheet_found': False, 'found': False,
           'cur_re': 0.0, 'prior_re': 0.0, 'prior_ni': 0.0,
           'roll_diff': 0.0, 'roll_diff2': 0.0, 'roll_ok': True,
           'cur_oci': 0.0, 'prior_oci': 0.0, 'oci_change': 0.0,
           'is_flagged': False, 'severity': None, 'error': None}
    try:
        zf = zipfile.ZipFile(file_path)
    except Exception as e:
        out['error'] = f'파일 열기 실패: {e}'
        return out
    try:
        sheet_path = _find_sheet_path_by_name(zf, PY_GAAP_SHEET)
        if not sheet_path or sheet_path not in zf.namelist():
            return out  # sheet_found=False 유지
        shared = _load_shared_strings(zf)
        out['sheet_found'] = True

        prior = {}
        cur = {}
        with zf.open(sheet_path) as f:
            for _ev, elem in ET.iterparse(f, events=('end',)):
                if elem.tag != _TAG_ROW:
                    continue
                vals = {}
                for c in elem.findall(_TAG_C):
                    letters, _ = _split_cell_ref(c.get('r'))
                    if letters is None:
                        continue
                    ci = _col_letters_to_index(letters)
                    if ci in (13, 14, 15):   # M=코드, N=전기, O=당기
                        vals[ci] = _cell_value(c, shared)
                elem.clear()
                code = vals.get(13)
                if not isinstance(code, (int, float)) or isinstance(code, bool):
                    continue
                code_i = int(code)
                # 순수 이익잉여금(이익준비금·임의적립금·미처분) + 보험수리적손익 + 당기순이익
                if code_i in (3500101, 3500102, 3500103, 3500104, 3500105):
                    prior[code_i] = _num_or_zero(vals.get(14))
                    cur[code_i] = _num_or_zero(vals.get(15))

        if 3500104 not in cur:
            return out  # PY GAAP Diff 표 없음 → found=False

        out['found'] = True
        # 순수 이익잉여금 = 이익준비금(3500101) + 임의적립금(3500102) + 미처분(3500104)
        # (적립금 전입 등 순수 이익잉여금 내부 이동을 흡수. 비지배지분 3600101은 제외.)
        out['cur_re'] = cur.get(3500101, 0.0) + cur.get(3500102, 0.0) + cur.get(3500104, 0.0)
        out['prior_re'] = prior.get(3500101, 0.0) + prior.get(3500102, 0.0) + prior.get(3500104, 0.0)
        out['prior_ni'] = prior.get(3500105, 0.0)
        out['cur_oci'] = cur.get(3500103, 0.0)
        out['prior_oci'] = prior.get(3500103, 0.0)
        # ① 미처분 단독 롤포워드 / ② 미처분+보험수리적 롤포워드
        out['roll_diff'] = out['cur_re'] - (out['prior_re'] + out['prior_ni'])
        out['roll_diff2'] = ((out['cur_re'] + out['cur_oci'])
                             - (out['prior_re'] + out['prior_oci'] + out['prior_ni']))
        out['oci_change'] = out['cur_oci'] - out['prior_oci']
    finally:
        zf.close()

    # 두 방식 중 하나라도 0이면 미처분 롤포워드 정상 (회사별 OCI 이월정책 차이 흡수)
    out['roll_ok'] = (abs(out['roll_diff']) < _GAAP_RE_TOL
                      or abs(out['roll_diff2']) < _GAAP_RE_TOL)
    roll_err = out['found'] and not out['roll_ok']
    oci_chg = abs(out['oci_change']) >= _GAAP_RE_TOL
    out['is_flagged'] = roll_err or oci_chg
    out['severity'] = 'error' if roll_err else ('review' if oci_chg else None)
    return out
