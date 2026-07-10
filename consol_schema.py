"""
연결정산 스키마 / 데이터 영속 계층
- consol_template.json : 314개 계정 + 소계/공식 행 (태림연결FS에서 추출한 표준 양식)
- consol_groups.json   : 연결그룹 정의 {groups: [{id, name, companies: []}]}
- consol_journals.json : 분개 데이터 {key(group_id__period): {entries: [...], updated_at, updated_by}}
"""
from __future__ import annotations
import json
import uuid
from pathlib import Path
from datetime import datetime
from filelock import FileLock

TEMPLATE_FILE = Path('consol_template.json')
GROUPS_FILE   = Path('consol_groups.json')
JOURNALS_FILE = Path('consol_journals.json')
PRIOR_FILE    = Path('prior_consolidated.json')   # 전년 연결값 수기 입력 데이터

_groups_lock   = FileLock(str(GROUPS_FILE) + '.lock', timeout=10)
_journals_lock = FileLock(str(JOURNALS_FILE) + '.lock', timeout=10)
_prior_lock    = FileLock(str(PRIOR_FILE) + '.lock', timeout=10)


# ─── 템플릿 ──────────────────────────────────────────────────────────────────

_template_cache = None


def load_template():
    """consol_template.json 로드 (캐시)."""
    global _template_cache
    if _template_cache is not None:
        return _template_cache
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f'템플릿 파일이 없습니다: {TEMPLATE_FILE.resolve()}')
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        _template_cache = json.load(f)
    return _template_cache


def template_codes():
    """detail 행의 코드 목록 (중복 제거, 순서 보존)."""
    seen = set()
    out = []
    for row in load_template()['rows']:
        if row.get('kind') == 'detail':
            c = row.get('code')
            if c and c not in seen:
                seen.add(c)
                out.append(c)
    return out


# ─── 그룹 CRUD ───────────────────────────────────────────────────────────────

def _atomic_write_json(path: Path, data):
    tmp = path.with_suffix(path.suffix + '.tmp')
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp.replace(path)


def load_groups():
    if not GROUPS_FILE.exists():
        return {'groups': []}
    with _groups_lock:
        with open(GROUPS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)


def save_groups(data):
    with _groups_lock:
        _atomic_write_json(GROUPS_FILE, data)


def list_groups():
    return load_groups().get('groups', [])


def get_group(group_id):
    for g in list_groups():
        if g['id'] == group_id:
            return g
    return None


def upsert_group(name: str, companies: list[str], group_id: str | None = None,
                 included_groups: list[str] | None = None):
    """그룹 신규/수정. group_id 미지정 시 신규.

    included_groups: 이 그룹이 sub-consolidation으로 포함하는 다른 그룹 ID 목록.
                     해당 그룹의 연결실행 결과(최종 P열)를 1개 컬럼으로 끌어옴.
    """
    data = load_groups()
    groups = data.get('groups', [])
    name = (name or '').strip()
    companies = [c.strip() for c in companies if c and c.strip()]
    included_groups = [str(gid).strip() for gid in (included_groups or []) if str(gid).strip()]
    if not name:
        raise ValueError('그룹명은 비워둘 수 없습니다.')

    # 자기 자신 포함 방지
    if group_id and group_id in included_groups:
        raise ValueError('그룹은 자기 자신을 포함할 수 없습니다.')

    # 포함 그룹 존재 여부 + 순환참조 검증
    existing_ids = {g['id'] for g in groups}
    for inc_id in included_groups:
        if inc_id not in existing_ids:
            raise ValueError(f'존재하지 않는 포함 그룹: {inc_id}')
    if group_id and _has_cycle(groups, group_id, included_groups):
        raise ValueError('순환 참조가 발생합니다 (이 그룹을 포함하는 다른 그룹이 있습니다).')

    if group_id:
        for g in groups:
            if g['id'] == group_id:
                g['name'] = name
                g['companies'] = companies
                g['included_groups'] = included_groups
                g['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                save_groups({'groups': groups})
                return g
        raise ValueError(f'존재하지 않는 그룹: {group_id}')

    new_id = uuid.uuid4().hex[:8]
    g = {
        'id': new_id,
        'name': name,
        'companies': companies,
        'included_groups': included_groups,
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }
    groups.append(g)
    save_groups({'groups': groups})
    return g


def _has_cycle(groups: list[dict], group_id: str, new_includes: list[str]) -> bool:
    """group_id가 new_includes를 포함하도록 변경 시 순환 참조가 발생하는지 검사.

    DFS: new_includes에서 시작해 includes 그래프를 따라가며 group_id에 도달하면 사이클.
    """
    by_id = {g['id']: g for g in groups}
    stack = list(new_includes)
    visited = set()
    while stack:
        cur = stack.pop()
        if cur == group_id:
            return True
        if cur in visited:
            continue
        visited.add(cur)
        g = by_id.get(cur)
        if g:
            stack.extend(g.get('included_groups', []) or [])
    return False


def get_descendants(group_id: str) -> list[str]:
    """group_id가 (직간접) 포함하는 모든 하위 그룹 ID 목록 (DFS, 자기 자신 제외)."""
    by_id = {g['id']: g for g in list_groups()}
    out = []
    seen = set()
    stack = list((by_id.get(group_id) or {}).get('included_groups', []) or [])
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        out.append(cur)
        g = by_id.get(cur)
        if g:
            stack.extend(g.get('included_groups', []) or [])
    return out


def delete_group(group_id):
    data = load_groups()
    data['groups'] = [g for g in data.get('groups', []) if g['id'] != group_id]
    save_groups(data)


# ─── 회사 기간(소속 시작·종료) 필터 ─────────────────────────────────────────
# 그룹 JSON에 선택적 `company_periods` 필드를 두어, 특정 회사가 어느 기간 동안만
# 해당 그룹에 속했는지 표현한다.
#
# 예) "company_periods": { "㈜전주원파워(연결)": {"until": "2025-4Q"} }
#     → 2025-4Q 이전(포함)까지는 그룹 멤버이고, 2026-1Q 이후엔 제외된다.
#
# 키는 그룹 `companies`의 회사명과 정확히 일치해야 한다. `since` / `until` 중
# 일부만 지정 가능 (둘 다 미지정이면 영구 멤버 = 기본 동작과 동일).

def _period_key(period: str) -> int | None:
    """'YYYY-NQ' → 정렬용 정수 키. 예: '2026-2Q' → 20262. 파싱 실패 시 None."""
    if not period or not isinstance(period, str):
        return None
    try:
        y, q = period.split('-', 1)
        return int(y) * 10 + int(q.rstrip('Q').rstrip('q'))
    except (ValueError, AttributeError):
        return None


# ─── 회사 마스터 '적용 시작 분기(since)' 연동 ────────────────────────────────
# 회사 마스터(company_master.json)에서 회사별 since를 읽어 연결 멤버십에도 반영.
# → 신생 자회사는 마스터 한 곳만 설정하면 마감현황·연결 모두 그 분기부터 포함된다.
COMPANY_MASTER_FILE = Path('company_master.json')
_master_since_cache = {'mtime': None, 'map': {}}


def _norm_co_name(s) -> str:
    """회사명 정규화 — 비단어문자 제거 + casefold (app._norm_company_name과 동일 규칙)."""
    import re as _re
    return _re.sub(r'[\W_]+', '', str(s or '').casefold(), flags=_re.UNICODE)


def _master_since_map() -> dict:
    """{정규화된 회사명: since_period_key(int)} — since가 지정된 회사만. mtime 캐시."""
    try:
        mtime = COMPANY_MASTER_FILE.stat().st_mtime
    except OSError:
        return {}
    if _master_since_cache['mtime'] == mtime:
        return _master_since_cache['map']
    out = {}
    try:
        with open(COMPANY_MASTER_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f) or {}
        for c in (data.get('companies') or []):
            sk = _period_key((c.get('since') or '').strip())
            if sk is not None:
                out[_norm_co_name(c.get('name'))] = sk
    except Exception:
        out = {}
    _master_since_cache['mtime'] = mtime
    _master_since_cache['map'] = out
    return out


def effective_companies(group: dict, period: str) -> list[str]:
    """`group['companies']` 중 해당 `period`에 멤버였던 회사만 반환.

    포함 조건 (모두 만족해야 함):
      1) 그룹별 company_periods 의 since/until 범위 안 (미지정 시 영구 멤버)
      2) 회사 마스터의 '적용 시작 분기(since)' 이후 (미지정 시 제한 없음)
    """
    cfg = (group or {}).get('company_periods') or {}
    full = list((group or {}).get('companies') or [])
    if not period:
        return full
    pk = _period_key(period)
    if pk is None:
        return full
    master_since = _master_since_map()
    out = []
    for c in full:
        # 1) 그룹별 기간 제한
        sub = cfg.get(c) or {}
        since_k = _period_key(sub.get('since')) if sub.get('since') else None
        until_k = _period_key(sub.get('until')) if sub.get('until') else None
        if since_k is not None and pk < since_k:
            continue
        if until_k is not None and pk > until_k:
            continue
        # 2) 회사 마스터 적용 시작 분기
        mk = master_since.get(_norm_co_name(c))
        if mk is not None and pk < mk:
            continue
        out.append(c)
    return out


# ─── 분개 영속 ───────────────────────────────────────────────────────────────

def _journal_key(group_id, period):
    return f'{group_id}__{period}'


def load_journals():
    if not JOURNALS_FILE.exists():
        return {}
    with _journals_lock:
        with open(JOURNALS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)


def save_journals(data):
    with _journals_lock:
        _atomic_write_json(JOURNALS_FILE, data)


def _migrate_record(rec):
    """단일 분개 레코드를 신 스키마(adjustment_entries / intercompany_entries)로 마이그레이션."""
    if not isinstance(rec, dict):
        return rec
    if 'adjustment_entries' in rec or 'intercompany_entries' in rec:
        # 신 스키마 — 누락 키 보강
        rec.setdefault('adjustment_entries', [])
        rec.setdefault('intercompany_entries', [])
        return rec
    # 구 스키마: entries → adjustment_entries 로 통째 이동 (보수적 분류)
    rec['adjustment_entries'] = rec.pop('entries', [])
    rec['intercompany_entries'] = []
    return rec


def get_journal(group_id, period):
    rec = load_journals().get(_journal_key(group_id, period))
    if rec is None:
        return None
    return _migrate_record(rec)


def set_journal(group_id, period, adjustment_entries: list = None,
                intercompany_entries: list = None, username: str = ''):
    """분개 두 묶음을 저장.

    adjustment_entries  : 연결조정 (자본/투자 상계, 영업권, 지분법 등)
    intercompany_entries: 내부거래 (매출-매출원가, 채권채무, 미실현 등)
    """
    data = load_journals()
    data[_journal_key(group_id, period)] = {
        'group_id': group_id,
        'period': period,
        'adjustment_entries': adjustment_entries or [],
        'intercompany_entries': intercompany_entries or [],
        'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'updated_by': username,
    }
    save_journals(data)
    return data[_journal_key(group_id, period)]


def set_journal_partial(group_id, period, journal_type: str,
                        entries: list, username: str = ''):
    """한 종류(연결조정 or 내부거래)만 갱신 (다른 한쪽 보존)."""
    if journal_type not in ('adjustment', 'intercompany'):
        raise ValueError(f'잘못된 분개 유형: {journal_type}')
    existing = get_journal(group_id, period) or {}
    adj = existing.get('adjustment_entries', [])
    inter = existing.get('intercompany_entries', [])
    if journal_type == 'adjustment':
        adj = entries
    else:
        inter = entries
    return set_journal(group_id, period, adj, inter, username)


def delete_journal(group_id, period, journal_type: str = None):
    """journal_type 지정 시 해당 종류만 삭제, 미지정 시 전체 삭제."""
    data = load_journals()
    key = _journal_key(group_id, period)
    if key not in data:
        return
    if journal_type is None:
        del data[key]
    else:
        rec = _migrate_record(data[key])
        if journal_type == 'adjustment':
            rec['adjustment_entries'] = []
        elif journal_type == 'intercompany':
            rec['intercompany_entries'] = []
        rec['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data[key] = rec
    save_journals(data)


# ─── 전년 연결값 (수기 입력) ─────────────────────────────────────────────────
# 실무상 전년도 연결을 재실행할 수 없어, 그룹별로 핵심지표 전년값을 수기 저장한다.
# 구조 (모든 금액은 원 단위 raw value):
# {
#   "<group_id>": {
#     "bs_by_year": {                     # 전년말 기준 BS 값
#       "2025": {"자산총계": 0, "부채총계": 0, "자본총계": 0,
#                "차입금총계": 0, "은행차입금": 0}
#     },
#     "pl_by_period": {                   # 전년 분기별 YTD 누계 PL 값
#       "2025-1Q": {"매출액": 0, "매출원가": 0, "매출총이익": 0,
#                   "영업이익": 0, "당기순이익": 0},
#       "2025-2Q": {...}, "2025-3Q": {...}, "2025-4Q": {...}
#     },
#     "updated_at": "...", "updated_by": "..."
#   }
# }

BS_FIELDS = ['자산총계', '부채총계', '자본총계', '차입금총계', '은행차입금']
PL_FIELDS = ['매출액', '영업이익', '당기순이익']


def load_prior():
    if not PRIOR_FILE.exists():
        return {}
    with _prior_lock:
        try:
            with open(PRIOR_FILE, 'r', encoding='utf-8') as f:
                return json.load(f) or {}
        except (json.JSONDecodeError, OSError):
            return {}


def save_prior(data):
    with _prior_lock:
        with open(PRIOR_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


def get_prior(group_id):
    return load_prior().get(group_id) or {
        'bs_by_year': {}, 'pl_by_period': {},
    }


def set_prior(group_id, bs_by_year: dict, pl_by_period: dict, username: str = ''):
    """그룹의 전년 연결값을 통째로 교체 저장."""
    data = load_prior()
    # 0/공란은 그대로 저장 (null/None은 미입력으로 간주 후 KPI에서 폴백)
    data[group_id] = {
        'bs_by_year': bs_by_year or {},
        'pl_by_period': pl_by_period or {},
        'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'updated_by': username,
    }
    save_prior(data)
    return data[group_id]


def prior_year_of(period: str) -> str | None:
    """'2026-1Q' → '2025'. 파싱 실패시 None."""
    if not period or '-' not in period:
        return None
    try:
        return str(int(period.split('-', 1)[0]) - 1)
    except ValueError:
        return None


def prior_period_of(period: str) -> str | None:
    """'2026-1Q' → '2025-1Q'. 파싱 실패시 None."""
    if not period or '-' not in period:
        return None
    parts = period.split('-', 1)
    try:
        return f'{int(parts[0]) - 1}-{parts[1]}'
    except ValueError:
        return None
