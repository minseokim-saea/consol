"""
cf_mapping_v2.json 초안 생성기 — COA 기반.

구조:
  cf_lines : [{cf_code, name, section, section_sign, type}]    # 매핑 가능한 CF 행 목록
  coa      : [{code, name, section(BS/IS), sign(D/C),
               adj_cf_code,   adj_sign,    # 연결조정 분개 라우팅
               inter_cf_code, inter_sign}] # 내부거래 분개 라우팅

특수 CF 코드:
  _NI_     — 당기순이익 행에 흡수 (별도 CF 라인 없음)
  _NONE_   — 매핑 없음 (현금/자본 등 CF에 영향 없는 계정)

기존 cf_mapping_draft.json 의 매핑을 우선 시드로 사용하고,
부족한 부분은 코드 prefix 휴리스틱으로 채움.
"""
import json
import openpyxl
from pathlib import Path

ROOT = Path(r'C:/패키지프로그램')
SUM_PATH = r'C:/연결시스템(26.1분기)/25.4분기/태림합산CF_25YE_시스템용.xlsx'
DRAFT_PATH = ROOT / 'cf_mapping_draft.json'
TEMPLATE_PATH = ROOT / 'consol_template.json'
OUT_PATH = ROOT / 'cf_mapping_v2_draft.json'


def load_cf_lines():
    """합산 시트에서 CF 라인 추출 (cf_mapping_draft.json 의 sections 활용)."""
    with open(DRAFT_PATH, encoding='utf-8') as f:
        draft = json.load(f)
    cf_lines = []
    for sec in draft['sections']:
        roman = sec['roman']
        sub = sec.get('sub') or ''
        section_label = f'{roman[:1]}-{sub.split(".")[0]}' if roman and sub else roman
        for row in sec.get('rows', []):
            cf_lines.append({
                'cf_code': row['cf_code'],
                'name': row['name'],
                'section': section_label,                # 'Ⅰ-2', 'Ⅰ-3', 'Ⅰ-4', 'Ⅱ-1', etc.
                'section_full': f'{roman} > {sub}',
                'section_sign': sec.get('sign_rule', '+'),
                'type': row.get('type') or '',           # asset/liability for WC
            })
    # 특수 옵션
    cf_lines.append({'cf_code': '_NI_', 'name': '당기순이익에 흡수 (별도 라인 없음)',
                     'section': 'NI', 'section_full': '특수', 'section_sign': '+',
                     'type': 'special'})
    cf_lines.append({'cf_code': '_NONE_', 'name': '매핑 없음 (CF 미반영)',
                     'section': 'NONE', 'section_full': '특수', 'section_sign': '',
                     'type': 'special'})
    return cf_lines


def load_coa_list():
    """consol_template.json 의 detail 행 모두 추출."""
    with open(TEMPLATE_PATH, encoding='utf-8') as f:
        tpl = json.load(f)
    rows = []
    for r in tpl['rows']:
        if r.get('kind') != 'detail':
            continue
        code = str(r.get('code') or '').strip()
        if not code:
            continue
        rows.append({
            'code': code,
            'name': r.get('name', ''),
            'section': r.get('section', ''),     # BS or IS
            'sign': r.get('sign', 'D'),          # D or C
        })
    return rows


def build_pl_to_cf_index(cf_lines):
    """기존 draft 의 lookup_codes를 역인덱스화 → {coa_code: cf_code (가산/차감), sign}."""
    with open(DRAFT_PATH, encoding='utf-8') as f:
        draft = json.load(f)

    coa_to_cf = {}  # {coa_code: {cf_code, sign_mul}}
    for sec in draft['sections']:
        sign_rule = sec.get('sign_rule', '+')
        # WC asset = -1, WC liab = -1, 차감 = -1, 가산 = +1, 유입/유출 = -1 (모두 자연 부호 역방향)
        if sign_rule == '+':
            default_mul = '+'
        elif sign_rule == '-':
            default_mul = '-'
        else:  # WC
            default_mul = '-'
        for row in sec.get('rows', []):
            cf_code = row['cf_code']
            for code in (row.get('lookup_codes') or []):
                if not code:
                    continue
                # 첫번째 매핑만 사용 (중복 시 첫것 우선)
                coa_to_cf.setdefault(str(code), {
                    'cf_code': cf_code,
                    'sign': default_mul,
                })
    # extra_pl_mappings 도 흡수 (당기순이익에 들어가는 매출/매출원가/판관비)
    for e in draft.get('extra_pl_mappings', []):
        code = str(e.get('pl_code') or '')
        if code and code not in coa_to_cf:
            coa_to_cf[code] = {'cf_code': '_NI_', 'sign': '+'}
    return coa_to_cf


def default_for_coa(coa, draft_idx):
    """COA 한 행에 대한 기본 매핑 — draft 먼저, 없으면 prefix 휴리스틱.

    정책: BS 코드는 연결조정 분개에서 CF에 반영되면 안 됨 → 항상 _NONE_.
          (내부거래용 매핑은 build()에서 별도 처리)
    """
    code = coa['code']
    sign_kind = coa['sign']     # D or C
    section = coa['section']    # BS or IS

    # 정책 적용 — BS 연결조정은 무조건 _NONE_
    if section == 'BS':
        return '_NONE_', '+'

    # 1) draft 매핑이 있으면 우선 (IS만 해당)
    if code in draft_idx:
        d = draft_idx[code]
        return d['cf_code'], d['sign']

    # 2) prefix 휴리스틱
    p1 = code[:1] if code else ''
    p2 = code[:2] if len(code) >= 2 else ''

    # 자본 (3xxxxxx) → _NONE_ (재무 직접조정 외 무영향)
    if p1 == '3':
        return '_NONE_', '+'
    # 현금/자본금 등 — CF 직접 영향 없음
    if code in ('1110101',):   # 현금및현금등가물
        return '_NONE_', '+'
    # 법인세 (4800001)
    if code == '4800001':
        return '_NI_', '+'
    # 매출/매출원가/판관비/제조경비 → NI에 이미 반영
    if p2 in ('41', '42', '43', '52', '53'):
        return '_NI_', '+'
    # 영업외수익 (44xxxxx) → 차감
    if p2 == '44':
        return code, '-'  # 자기 코드 그대로 (있는 경우)
    # 영업외비용 (45xxxxx) / 손실 → 가산
    if p2 == '45':
        return code, '+'
    # BS 자산 (1xxxxxx) → CF1xxx (WC)
    if section == 'BS' and p1 == '1':
        cf_guess = 'CF' + code
        return cf_guess, '-'
    # BS 부채 (2xxxxxx) → CF2xxx
    if section == 'BS' and p1 == '2':
        cf_guess = 'CF' + code
        return cf_guess, '-'
    # 그 외 (당기순이익 4900001 등) → NI
    if code.startswith('49'):
        return '_NI_', '+'
    return '_NONE_', '+'


def build():
    cf_lines = load_cf_lines()
    coa_list = load_coa_list()
    draft_idx = build_pl_to_cf_index(cf_lines)
    valid_cf_codes = {c['cf_code'] for c in cf_lines}

    coa_out = []
    unmatched_guesses = 0
    for coa in coa_list:
        adj_cf, adj_sign = default_for_coa(coa, draft_idx)
        # adj_cf 가 cf_lines에 없으면 _NONE_
        if adj_cf not in valid_cf_codes:
            unmatched_guesses += 1
            adj_cf = '_NONE_'
            adj_sign = '+'
        # 내부거래 기본값:
        #  - PL(43xxx 판관비 등): _NI_
        #  - BS WC 계정: 동일 CF 라인 사용
        #  - 영업외 PL: _NI_ (내부거래는 보통 매출-매출원가 elim)
        if coa['section'] == 'BS':
            inter_cf, inter_sign = adj_cf, adj_sign
        elif coa['code'][:2] in ('41', '42'):
            inter_cf, inter_sign = '_NI_', '+'   # 매출/매출원가 상계 → NI 흡수
        else:
            inter_cf, inter_sign = '_NI_', '+'

        coa_out.append({
            'code': coa['code'],
            'name': coa['name'],
            'section': coa['section'],
            'sign': coa['sign'],
            'adj_cf_code':   adj_cf,
            'adj_sign':      adj_sign,
            'inter_cf_code': inter_cf,
            'inter_sign':    inter_sign,
        })

    result = {
        'version': '2026.1Q-v2-draft',
        'description':
            'COA 기반 매핑 — 각 BS/PL 코드를 연결조정/내부거래용 CF 라인에 각각 매핑.',
        'policy': {
            'bs_adj_locked': True,
            'bs_adj_locked_note':
                '연결조정 분개에서 BS 항목은 CF에 반영되지 않음 (PL 항목만 반영). '
                'BS 코드의 adj_cf_code는 항상 _NONE_, adj_sign 무시.',
        },
        'cf_lines': cf_lines,
        'coa': coa_out,
    }
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f'생성: {OUT_PATH}')
    print(f'CF 라인 수: {len(cf_lines)} (특수 _NI_/_NONE_ 포함)')
    print(f'COA 행 수: {len(coa_out)}')
    print(f'  - BS: {sum(1 for c in coa_out if c["section"] == "BS")}')
    print(f'  - IS: {sum(1 for c in coa_out if c["section"] == "IS")}')
    print(f'  unmatched (CF guess invalid → _NONE_): {unmatched_guesses}')

    # 매핑 분포 통계
    from collections import Counter
    adj_dist = Counter(c['adj_cf_code'] for c in coa_out)
    print(f'\n=== 연결조정 매핑 상위 ===')
    print(f'  _NI_   : {adj_dist["_NI_"]}건')
    print(f'  _NONE_ : {adj_dist["_NONE_"]}건')
    print(f'  실제 CF라인: {len(coa_out) - adj_dist["_NI_"] - adj_dist["_NONE_"]}건')


if __name__ == '__main__':
    build()
