"""
WCE (Worldwide Consolidation Equity) 시트 132행 이후 본사 입력용 스키마.

각 테이블은 columns(계정코드별 열) × rows(증감항목 행) 구조.
- columns: 자본 계정코드와 영문명
- rows   : 기초/유상증자/.../환산효과 등 증감항목 (key = 한글 라벨)
- summary: 테이블 5에만 있는 부가 요약섹션 (단일 열)
"""

WCE_TABLES = [
    {
        'id': 1,
        'title_ko': '자본금',
        'title_en': 'Paid-In Capital',
        'columns': [
            {'code': '3100101', 'name': 'Common Stock'},
            {'code': '3100102', 'name': 'Preferred Stock'},
        ],
        'rows': [
            {'key': '기초금액',   'name_en': 'Beginning'},
            {'key': '유상증자',   'name_en': 'Capital Increase'},
            {'key': '현물출자',   'name_en': 'Investment in Kind'},
            {'key': '감자',       'name_en': 'Capital Reduction'},
            {'key': '주식배당',   'name_en': 'Stock Dividends'},
            {'key': '출자전환',   'name_en': 'Debt for Equity Swap'},
            {'key': '기타',       'name_en': 'Others'},
            {'key': '환산효과',   'name_en': 'FX Effect'},
        ],
    },
    {
        'id': 2,
        'title_ko': '자본잉여금',
        'title_en': 'Capital Surplus',
        'columns': [
            {'code': '3200101', 'name': 'Paid-In Capital In Excess'},
            {'code': '3200199', 'name': 'Other Capital Surplus'},
        ],
        'rows': [
            {'key': '기초금액',              'name_en': 'Beginning'},
            {'key': '유상증자',              'name_en': 'Capital Increase'},
            {'key': '현물출자',              'name_en': 'Investment in Kind'},
            {'key': '감자',                  'name_en': 'Capital Reduction'},
            {'key': '출자전환',              'name_en': 'Debt for Equity Swap'},
            {'key': '자본전입',              'name_en': 'Transfer to R/E'},
            {'key': '연결실체간자본거래',    'name_en': 'Intragroup Equity Tx'},
            {'key': '자기주식의 처분',       'name_en': 'Disposal of Treasury'},
            {'key': '신종자본증권의 발행',   'name_en': 'Issuance of Hybrid Capital'},
            {'key': '지분법자본잉여금',      'name_en': 'Capital Surplus from Equity Method'},
            {'key': '기타',                  'name_en': 'Others'},
            {'key': '환산효과',              'name_en': 'FX Effect'},
        ],
    },
    {
        'id': 3,
        'title_ko': '자본조정',
        'title_en': 'Capital Adjustments',
        'columns': [
            {'code': '3300101', 'name': 'Discount on Share Issuance'},
            {'code': '3300102', 'name': 'Capital Adjustment Of Equity'},
            {'code': '3300103', 'name': 'Treasury Shares'},
            {'code': '3300199', 'name': 'Other Capital Adjustment'},
        ],
        'rows': [
            {'key': '기초금액',          'name_en': 'Beginning'},
            {'key': '유상증자',          'name_en': 'Capital Increase'},
            {'key': '현물출자',          'name_en': 'Investment in Kind'},
            {'key': '감자',              'name_en': 'Capital Reduction'},
            {'key': '지분법자본조정',    'name_en': 'Equity Method Capital Adj.'},
            {'key': '자기주식취득',      'name_en': 'Purchase of Treasury'},
            {'key': '자기주식처분',      'name_en': 'Disposal of Treasury'},
            {'key': '출자전환',          'name_en': 'Debt for Equity Swap'},
            {'key': '연결실체간자본거래','name_en': 'Intragroup Equity Tx'},
            {'key': '기타',              'name_en': 'Others'},
            {'key': '환산효과',          'name_en': 'FX Effect'},
        ],
    },
    {
        'id': 4,
        'title_ko': '기타포괄손익누계액',
        'title_en': 'Accumulated Other Comprehensive Income',
        'columns': [
            {'code': '3400101', 'name': 'Valuation Gain/Loss on AFS'},
            {'code': '3400102', 'name': 'Gain on Capital Investment'},
            {'code': '3400103', 'name': 'Loss on Capital Investment'},
            {'code': '3400104', 'name': 'Gain/Loss on Translation'},
            {'code': '3400105', 'name': 'Gain/Loss on Valuation'},
            {'code': '3400106', 'name': 'Gain on Revaluation'},
            {'code': '3400199', 'name': 'Other Comprehensive Income'},
        ],
        'rows': [
            {'key': '기초금액',              'name_en': 'Beginning'},
            {'key': '매도가능증권평가',      'name_en': 'Valuation of AFS'},
            {'key': '자산재평가',            'name_en': 'Revaluation of Assets'},
            {'key': '파생상품평가',          'name_en': 'Valuation of Derivatives'},
            {'key': '해외사업환산손익',      'name_en': 'FX Translation Gain/Loss'},
            {'key': '지분법자본변동',        'name_en': 'Gain on Capital Inv. (Equity Method)'},
            {'key': '부의지분법자본변동',    'name_en': 'Loss on Capital Inv. (Equity Method)'},
            {'key': '기타',                  'name_en': 'Others'},
            {'key': '환산효과',              'name_en': 'FX Effect'},
        ],
    },
    {
        'id': 5,
        'title_ko': '이익잉여금',
        'title_en': 'Retained Earnings',
        'columns': [
            {'code': '3500101',    'name': 'Legal Reserve'},
            {'code': '3500102',    'name': 'Voluntary Reserves'},
            {'code': '3500103',    'name': 'Actuarial Gain/Loss'},
            {'code': '3500104',    'name': 'Unappropriated R/E'},
            {'code': '3500105',    'name': 'Current Net Income'},
        ],
        'rows': [
            {'key': '기초금액',          'name_en': 'Beginning'},
            {'key': '당기순이익',        'name_en': 'Profit for the Year'},
            {'key': '보험수리적손익',    'name_en': 'Remeasurements'},
            {'key': '지분법이익잉여금',  'name_en': 'R/E Adjustment of Equity'},
            {'key': '자본전입',          'name_en': 'Transfer to R/E'},
            {'key': '배당',              'name_en': 'Dividends Paid'},
            {'key': '기타',              'name_en': 'Others'},
            {'key': '환산효과',          'name_en': 'FX Effect'},
        ],
    },
    {
        'id': 6,
        'title_ko': '비지배지분',
        'title_en': 'Non-Controlling Interest',
        'columns': [
            {'code': 'FS32000000', 'name': 'Non-controlling Interest'},
        ],
        'rows': [
            {'key': '기초금액',              'name_en': 'Beginning'},
            {'key': '유상증자',              'name_en': 'Capital Increase'},
            {'key': '현물출자',              'name_en': 'Investment in Kind'},
            {'key': '감자',                  'name_en': 'Capital Reduction'},
            {'key': '주식배당',              'name_en': 'Stock Dividends'},
            {'key': '출자전환',              'name_en': 'Debt for Equity Swap'},
            {'key': '당기순이익',            'name_en': 'Profit for the Year'},
            {'key': '보험수리적손익',        'name_en': 'Remeasurements'},
            {'key': '연결실체간자본거래',    'name_en': 'Intragroup Equity Tx'},
            {'key': '매도가능증권평가',      'name_en': 'Valuation of AFS'},
            {'key': '자산재평가',            'name_en': 'Revaluation of Assets'},
            {'key': '파생상품평가',          'name_en': 'Valuation of Derivatives'},
            {'key': '해외사업환산손익',      'name_en': 'FX Translation Gain/Loss'},
            {'key': '지분법자본변동',        'name_en': 'Equity Method Capital Chg'},
            {'key': '부의지분법자본변동',    'name_en': 'Equity Method Loss'},
            {'key': '배당',                  'name_en': 'Dividends Paid'},
            {'key': '기타',                  'name_en': 'Others'},
        ],
    },
]


# ─── WCE 로컬(1~132행) 섹션 레이아웃 ────────────────────────────────────────
# 각 테이블의 로컬 섹션 위치 (행 번호, 열별 코드 매핑, 데이터 행 범위)
WCE_LOCAL_LAYOUT = {
    '1': {  # 자본금 R11~R21
        'cols': {3: '3100101', 4: '3100102'},
        'data_rows': list(range(14, 22)),
    },
    '2': {  # 자본잉여금 R27~R41
        'cols': {3: '3200101', 4: '3200199'},
        'data_rows': list(range(30, 42)),
    },
    '3': {  # 자본조정 R46~R59
        'cols': {3: '3300101', 4: '3300102', 5: '3300103', 6: '3300199'},
        'data_rows': list(range(49, 60)),
    },
    '4': {  # 기타포괄손익누계액 R64~R75
        'cols': {3: '3400101', 4: '3400102', 5: '3400103',
                 6: '3400104', 7: '3400105', 8: '3400106', 9: '3400199'},
        'data_rows': list(range(67, 76)),
    },
    '5': {  # 이익잉여금 R80~R91
        'cols': {3: '3500101', 4: '3500102', 5: '3500103', 6: '3500104', 7: '3500105'},
        'data_rows': list(range(83, 92)),
    },
    '6': {  # 비지배지분 R97~R119 — 로컬은 3600101 코드 사용
        'cols': {3: '3600101'},
        'data_rows': list(range(100, 120)),
    },
}

# 스키마 코드 → 로컬 섹션 코드 매핑 (테이블 6의 비지배지분)
SCHEMA_TO_LOCAL_CODE = {
    'FS32000000': '3600101',
}

# 스키마 row.key → 로컬 섹션 한글 라벨 매핑 (라벨이 다른 경우만)
SCHEMA_TO_LOCAL_LABEL = {
    # 테이블 2/3: 자본잉여금/자본조정 - 패키지 표기 흡수
    '연결실체간자본거래': '연결실체 내 자본거래',
    # 테이블 4: 기타포괄손익누계액 - 패키지 표기 흡수
    '자산재평가': '자산재평가차익',
    # 테이블 5: 이익잉여금 - 패키지 표기 흡수
    '자본전입': '적립금 적립',
    # 테이블 5: 이익잉여금 — 스키마와 로컬 라벨 동일하므로 매핑 불필요
    # 테이블 4/6: 지분법자본변동/부의지분법자본변동 — 스키마와 로컬 라벨 동일
}


# BS 자본 소계 코드 → 구성 자식 코드 (WCE 데이터로 재계산 시 사용)
WCE_EQUITY_GROUPS = {
    '3100000': ['3100101', '3100102'],
    '3200000': ['3200101', '3200199'],
    '3300000': ['3300101', '3300102', '3300103', '3300199'],
    '3400000': ['3400101', '3400102', '3400103', '3400104', '3400105', '3400106', '3400199'],
    '3500000': ['3500101', '3500102', '3500103', '3500104', '3500105'],
    # 3000000(자본총계)은 모든 leaf + FS32000000 합계로 별도 처리
}
# 자본총계(3000000)에 포함될 모든 leaf 코드
WCE_ALL_EQUITY_CODES = [code for t in WCE_TABLES for col in t['columns'] for code in [col['code']]]


def get_table(table_id):
    for t in WCE_TABLES:
        if t['id'] == table_id:
            return t
    return None


def to_local_code(schema_code):
    return SCHEMA_TO_LOCAL_CODE.get(schema_code, schema_code)


def to_local_label(row_key):
    return SCHEMA_TO_LOCAL_LABEL.get(row_key, row_key)


def empty_overrides():
    """빈 override 구조 생성 (모든 셀이 0)."""
    return {
        str(t['id']): {
            col['code']: {row['key']: 0 for row in t['rows']}
            for col in t['columns']
        }
        for t in WCE_TABLES
    }
