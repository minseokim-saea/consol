"""연결감사보고서 DSD → 감사전 연결재무제표 DSD 형식 변환.

감사인이 발행한 연결감사보고서(DART 문서코드 00761)를 틀로 올려도, 회사가 제출하는
감사전 연결재무제표(00861) 형식으로 산출물을 만들기 위해 틀 자체를 먼저 변환한다.
(DART 편집기의 '연결감사보고서 변환' 으로 만든 감사전 문서와 같은 구조)

  - 문서 머리: 문서명·코드를 감사전 연결재무제표로, 감사인 관련 추출항목(AUDIT_CIK·SUPV_OPIN) 삭제
  - 본문 앞: 표지·목차·독립된 감사인의 감사보고서 → '보고서 종류: 연결감사보고서 변환' 한 칸으로
  - '(첨부)연결재무제표' 장(재무제표 4종 + 주석)은 그대로 두고 제목의 '(첨부)' 만 뺀다
  - 본문 뒤: 내부회계관리제도 검토의견·외부감사 실시내용 등 삭제 (연결부속명세서 목차 삽입 칸만 남김)
  - 목차 번호(ATOCID)를 감사전 문서 순서로 다시 매김, meta.xml 의 문서코드도 바꿈
"""
from __future__ import annotations

import re
import zipfile

AUDIT_CODE = '00761'          # 연결감사보고서
PREAUDIT_CODE = '00861'       # 감사전 연결재무제표
PREAUDIT_BSN = '00860'

_REP_KND = '''
<PART ACLASS="MANDATORY" APARTSOURCE="SOURCE">

<TABLE-GROUP ACLASS="REP_KND" ADELETETABLE="N">

<TABLE ACLASS="NORMAL" AFIXTABLE="Y" WIDTH="600" BORDER="0">

<COLGROUP WIDTH="600">
<COL WIDTH="600"></COL>
</COLGROUP>

<TBODY>

<TR ACOPY="N" ADELETE="N" HEIGHT="30">
<TU WIDTH="591" HEIGHT="23" VALIGN="MIDDLE" AUNIT="+REP_KND" AUNITVALUE="00761">연결감사보고서 변환</TU>
</TR>
</TBODY>
</TABLE>
</TABLE-GROUP>
</PART>

'''

_TAIL = '''

<INSERTION ABASISNUMBER="SE1" AFREQUENCY="0" ADUPLICATION="N">
<COMMENT>◆click◆ 『연결부속명세서 목차』 삽입</COMMENT>
<LIBRARYLIST>00761#*연결부속명세서_목차.dsl</LIBRARYLIST>
</INSERTION>
'''

_RE_DOCNAME = re.compile(r'<DOCUMENT-NAME\s+ACODE="(\d+)"\s*>[^<]*</DOCUMENT-NAME>')


def doc_code(text: str) -> str | None:
    m = _RE_DOCNAME.search(text)
    return m.group(1) if m else None


def _fs_section_span(text: str) -> tuple[int, int]:
    """'(첨부)연결재무제표' 장 SECTION-1 의 [시작, 끝) — 제목의 AASSOCNOTE 가 D-0-0-0-0 인 것."""
    m = re.search(r'<SECTION-1\b[^>]*>\s*<TITLE\b[^>]*AASSOCNOTE="D-0-0-0-0"', text)
    if not m:
        raise ValueError('연결감사보고서에서 (첨부)연결재무제표 장을 찾지 못했습니다.')
    start = m.start()
    depth, pos = 0, start
    for t in re.finditer(r'<(/?)SECTION-1\b[^>]*>', text[start:]):
        depth += -1 if t.group(1) else 1
        if depth == 0:
            pos = start + t.end()
            break
    else:
        raise ValueError('(첨부)연결재무제표 장의 끝을 찾지 못했습니다.')
    return start, pos


def convert_contents(text: str) -> str:
    """연결감사보고서 contents.xml 본문 → 감사전 연결재무제표 본문."""
    if doc_code(text) != AUDIT_CODE:
        raise ValueError('연결감사보고서(00761) 문서가 아닙니다.')
    t = _RE_DOCNAME.sub(f'<DOCUMENT-NAME ACODE="{PREAUDIT_CODE}">감사전 연결재무제표</DOCUMENT-NAME>',
                        text, count=1)
    # 감사인 관련 요약 추출항목 삭제 (감사전 문서에는 없음)
    t = re.sub(r'[ \t]*<EXTRACTION\s+ACODE="(?:AUDIT_CIK|SUPV_OPIN)"[^>]*>[^<]*</EXTRACTION>\r?\n?', '', t)

    body = re.search(r'<BODY\b[^>]*>', t)
    body_end = t.rfind('</BODY>')
    if not body or body_end < 0:
        raise ValueError('DSD 본문(BODY)을 찾지 못했습니다.')
    a, b = _fs_section_span(t)
    fs = t[a:b]
    # 제목 '(첨부)' 제거, 영문 제목도 맞춤
    fs = re.sub(r'(<TITLE\b[^>]*AASSOCNOTE="D-0-0-0-0"[^>]*>)\s*\(첨부\)\s*', r'\1', fs, count=1)
    fs = fs.replace('ENG="(Attachment) Consolidated Financial Statements"', '', 1)
    fs = re.sub(r'(<TITLE\b[^>]*?)\s+>', r'\1>', fs)

    out = t[:body.end()] + '\n' + _REP_KND + fs + _TAIL + t[body_end:]
    return _renumber_toc(out)


def _renumber_toc(t: str) -> str:
    """ATOCID: 연결재무제표 장 1, 주석 2, 나머지(재무제표 4종)는 문서 순서대로 3~. BODY 는 최댓값."""
    titles = list(re.finditer(r'<(?:TITLE|COVER-TITLE)\b[^>]*\bATOCID="\d+"[^>]*>', t))
    ids, nxt = {}, 3
    for m in titles:
        tag = m.group(0)
        if 'AASSOCNOTE="D-0-0-0-0"' in tag:
            ids[m.start()] = 1
        elif 'AASSOCNOTE="D-0-1-0-0"' in tag:
            ids[m.start()] = 2
        else:
            ids[m.start()] = nxt
            nxt += 1
    parts, last = [], 0
    for m in titles:
        parts.append(t[last:m.start()])
        parts.append(re.sub(r'\bATOCID="\d+"', f'ATOCID="{ids[m.start()]}"', m.group(0), count=1))
        last = m.end()
    parts.append(t[last:])
    t = ''.join(parts)
    top = max(ids.values()) if ids else 0
    return re.sub(r'(<BODY\b[^>]*\bATOCID=")\d+(")', rf'\g<1>{top}\2', t, count=1)


def convert_meta(meta: str) -> str:
    m = re.sub(r'bsn-id="\d+"', f'bsn-id="{PREAUDIT_BSN}"', meta, count=1)
    m = re.sub(r'rpt-id="\d+"', f'rpt-id="{PREAUDIT_CODE}"', m, count=1)
    m = re.sub(r'doc-id="\d+"', f'doc-id="{PREAUDIT_CODE}"', m, count=1)
    m = re.sub(r'<FILE-INFO\b[^>]*/>', '<FILE-INFO date="" size="" rnnumber="" rncheck="" rnimgfile=""/>', m, count=1)
    m = re.sub(r'\s*<EXTRACT\b[^>]*>.*?</EXTRACT>', '', m, flags=re.S)
    return m


def convert_file_if_audit(path: str) -> bool:
    """path 의 DSD 가 연결감사보고서면 감사전 연결재무제표 형식으로 덮어쓴다 (작업 사본에만 사용).
    변환했으면 True, 원래 감사전 등 다른 문서면 그대로 두고 False."""
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        members = {i.filename: z.read(i.filename) for i in infos}
    raw = members.get('contents.xml')
    if raw is None:
        return False
    bom = raw.startswith(b'\xef\xbb\xbf')
    text = raw.decode('utf-8-sig')
    if doc_code(text) != AUDIT_CODE:
        return False
    new = convert_contents(text).encode('utf-8')
    members['contents.xml'] = (b'\xef\xbb\xbf' if bom else b'') + new
    if 'meta.xml' in members:
        mraw = members['meta.xml']
        mbom = mraw.startswith(b'\xef\xbb\xbf')
        mtxt = convert_meta(mraw.decode('utf-8-sig')).encode('utf-8')
        members['meta.xml'] = (b'\xef\xbb\xbf' if mbom else b'') + mtxt
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i.filename, members[i.filename])
    return True
