"""app.py 라우트에 권한 키 기반 데코레이터를 일괄 적용.

규칙:
- 명시된 view 함수명 → @admin_required를 @require_permission('<key>')로 교체
- @login_required 만 있던 라우트에는 @require_permission('<key>') 한 줄 추가
"""
from pathlib import Path

P = Path(r'C:\패키지프로그램\app.py')
text = P.read_text(encoding='utf-8')
orig = text

# (함수명, 새 권한키) — @admin_required 를 @require_permission('<key>')로 교체
ADMIN_TO_PERM = {
    'admin_users': 'users.manage',
    'admin_create_user': 'users.manage',
    'admin_delete_user': 'users.manage',
    'admin_reset_password': 'users.manage',
    'admin_toggle_admin': 'users.manage',
    'admin_get_user_companies': 'users.manage',
    'admin_set_user_companies': 'users.manage',
    'admin_all_companies': 'users.manage',
    'admin_wce_index': 'wce.manage',
    'admin_wce_edit': 'wce.manage',
    'admin_wce_save': 'wce.manage',
    'admin_wce_delete': 'wce.manage',
    'admin_wce_aggregate_page': 'wce.manage',
    'admin_wce_aggregate_data': 'wce.manage',
    'admin_fx_rates_data': 'fx.manage',
    'admin_fx_rates_save': 'fx.manage',
    'admin_fx_rates_reapply': 'fx.manage',
    'admin_fx_rates_pull': 'fx.manage',
    'lock_year': 'years.manage',
    'unlock_year': 'years.manage',
    'add_year': 'years.manage',
    'delete_year': 'years.manage',
    'delete_all_files': 'files.delete',
    'delete_file': 'files.delete',
    'consolidation_groups_create': 'consol.groups',
    'consolidation_groups_delete': 'consol.groups',
}

# (함수명, 권한키) — @login_required 만 있는 라우트에 @require_permission('<key>') 추가
LOGIN_TO_PERM_ADD = {
    'upload': 'files.upload',
    'reanalyze_all': 'files.reanalyze',
    'reanalyze_one': 'files.reanalyze',
    'run_aggregate': 'aggregate.run',
    'consolidation_prior_save': 'prior.edit',
    'consolidation_journal_delete': 'consol.journal',
    'consolidation_journal_upload': 'consol.journal',
    'consolidation_compute': 'consol.compute',
}

import re

# 패턴: @admin_required\n(어쩌면 다른 데코)... \ndef FUNC_NAME(
# 안전하게 처리하려면 함수명 매칭 후 그 위 데코레이터 영역에서 @admin_required를 찾아 치환.

def replace_admin_required_for(text, func_name, perm_key):
    """def FUNC_NAME 바로 위 데코레이터 블록의 @admin_required를 @require_permission('key')로 교체."""
    pattern = re.compile(
        r'((?:^@[\w\.\(\)\'\",\s]+\n)+)def ' + re.escape(func_name) + r'\b',
        re.MULTILINE,
    )
    def _sub(m):
        block = m.group(1)
        new_block = block.replace('@admin_required\n', f"@require_permission('{perm_key}')\n")
        return new_block + 'def ' + func_name
    new_text, n = pattern.subn(_sub, text)
    return new_text, n

def add_perm_decorator_for(text, func_name, perm_key):
    """def FUNC_NAME 바로 위 데코레이터 블록의 @login_required 다음 줄에 @require_permission('key')를 삽입.
    이미 추가된 경우(@require_permission 이 존재하면) 변경 없음."""
    pattern = re.compile(
        r'((?:^@[\w\.\(\)\'\",\s]+\n)+)def ' + re.escape(func_name) + r'\b',
        re.MULTILINE,
    )
    def _sub(m):
        block = m.group(1)
        if f"@require_permission('{perm_key}')" in block:
            return block + 'def ' + func_name
        # @login_required 다음 줄에 추가
        if '@login_required\n' in block:
            new_block = block.replace(
                '@login_required\n',
                f"@login_required\n@require_permission('{perm_key}')\n",
                1,
            )
        else:
            new_block = block + f"@require_permission('{perm_key}')\n"
        return new_block + 'def ' + func_name
    new_text, n = pattern.subn(_sub, text)
    return new_text, n


total_replaced = 0
total_added = 0
for fn, key in ADMIN_TO_PERM.items():
    text, n = replace_admin_required_for(text, fn, key)
    if n > 0:
        total_replaced += n
        print(f'  REPLACE {fn}: @admin_required → @require_permission({key!r}) ({n}건)')
    else:
        print(f'  MISS    {fn} (admin)')

for fn, key in LOGIN_TO_PERM_ADD.items():
    text, n = add_perm_decorator_for(text, fn, key)
    if n > 0:
        total_added += n
        print(f'  ADD     {fn}: @require_permission({key!r}) 추가 ({n}건)')
    else:
        print(f'  MISS    {fn} (login)')

print(f'\n총 변경: 교체 {total_replaced}건, 추가 {total_added}건')

if text != orig:
    P.write_text(text, encoding='utf-8')
    print('app.py 저장 완료')
else:
    print('변경 사항 없음 — 파일 미수정')
