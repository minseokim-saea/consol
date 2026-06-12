"""
배포 패키지 무결성 토큰 (HMAC + docProps/custom.xml).

배포 파일 생성 시:
  · payload = {company, year, generated_at, nonce, ...}
  · signature = HMAC-SHA256(secret_key, payload_json)
  · 두 값을 .xlsm 의 docProps/custom.xml 에 'PackageToken' / 'PackageSig' 로 삽입

업로드 시:
  · 같은 위치에서 토큰 읽고 HMAC 검증
  · company / year 매칭으로 옛 분기·타사 파일 재활용 차단
  · 관리자 권한 사용자는 우회 가능 (운영상 옛 파일 재업로드용)

토큰은 docProps 영역이라 일반 사용자가 잘 안 보고, secret_key 없이는 위조 불가.
"""

from __future__ import annotations
import hmac
import hashlib
import json
import re
import secrets as _secrets
import shutil
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SECRET_KEY_FILE = Path(__file__).resolve().parent / 'secret_key.bin'

_NS_CP  = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
_NS_VT  = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
_FMTID  = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'
_TYPE_CUSTOM = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'
_REL_CUSTOM  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'

_PROP_PAYLOAD = 'PackageToken'
_PROP_SIG     = 'PackageSig'


# ─── secret_key 로드 ────────────────────────────────────────────

def _load_secret(secret: bytes | None = None) -> bytes:
    if secret is not None:
        return secret
    if not SECRET_KEY_FILE.exists():
        # 없으면 app.py 와 동일하게 생성
        SECRET_KEY_FILE.write_bytes(_secrets.token_bytes(64))
    return SECRET_KEY_FILE.read_bytes()


def _hmac_sign(secret: bytes, payload_str: str) -> str:
    return hmac.new(secret, payload_str.encode('utf-8'), hashlib.sha256).hexdigest()


# ─── 토큰 생성 ──────────────────────────────────────────────────

def generate_token(company: str, year: str,
                   template_version: str = '',
                   extra: dict | None = None,
                   secret: bytes | None = None) -> tuple[str, str]:
    """payload JSON 문자열과 signature hex 반환."""
    key = _load_secret(secret)
    payload = {
        'company': company,
        'year':    year,
        'template_version': template_version,
        'generated_at': time.strftime('%Y-%m-%dT%H:%M:%S'),
        'nonce': _secrets.token_hex(8),
    }
    if extra:
        payload.update(extra)
    payload_str = json.dumps(payload, sort_keys=True, ensure_ascii=False, separators=(',', ':'))
    return payload_str, _hmac_sign(key, payload_str)


# ─── docProps/custom.xml 빌드 ───────────────────────────────────

def _xml_escape(s: str) -> str:
    return (str(s)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


def _build_custom_xml(props: dict[str, str]) -> bytes:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<Properties xmlns="{_NS_CP}" xmlns:vt="{_NS_VT}">',
    ]
    for pid, (name, value) in enumerate(props.items(), start=2):
        parts.append(
            f'<property fmtid="{_FMTID}" pid="{pid}" name="{_xml_escape(name)}">'
            f'<vt:lpwstr>{_xml_escape(value)}</vt:lpwstr></property>'
        )
    parts.append('</Properties>')
    return ''.join(parts).encode('utf-8')


def _patch_content_types(xml_bytes: bytes) -> bytes:
    text = xml_bytes.decode('utf-8')
    if '/docProps/custom.xml' in text:
        return xml_bytes
    override = (
        f'<Override PartName="/docProps/custom.xml" '
        f'ContentType="{_TYPE_CUSTOM}"/>'
    )
    return text.replace('</Types>', override + '</Types>').encode('utf-8')


def _patch_root_rels(xml_bytes: bytes) -> bytes:
    text = xml_bytes.decode('utf-8')
    if 'docProps/custom.xml' in text:
        return xml_bytes
    used = set(re.findall(r'Id="(rId\d+)"', text))
    n = 1
    while f'rId{n}' in used:
        n += 1
    rel = (
        f'<Relationship Id="rId{n}" Type="{_REL_CUSTOM}" '
        f'Target="docProps/custom.xml"/>'
    )
    return text.replace('</Relationships>', rel + '</Relationships>').encode('utf-8')


# ─── 토큰 임베드 ────────────────────────────────────────────────

def embed_token(file_path, payload_str: str, signature: str) -> None:
    """xlsm/xlsx 에 PackageToken / PackageSig 를 주입.
    기존 custom.xml 이 있으면 두 키만 교체/추가, 다른 속성은 보존."""
    fp = Path(file_path)
    tmp = fp.with_suffix(fp.suffix + '.tok.tmp')

    new_props = {
        _PROP_PAYLOAD: payload_str,
        _PROP_SIG:     signature,
    }

    with zipfile.ZipFile(fp, 'r') as src:
        names = src.namelist()
        has_custom = 'docProps/custom.xml' in names

        # 기존 custom.xml 의 다른 속성 보존
        if has_custom:
            try:
                existing = ET.fromstring(src.read('docProps/custom.xml'))
                for prop in existing.findall(f'{{{_NS_CP}}}property'):
                    name = prop.get('name')
                    if not name or name in new_props:
                        continue
                    v = prop.find(f'{{{_NS_VT}}}lpwstr')
                    if v is not None and v.text is not None:
                        new_props[name] = v.text
            except Exception:
                pass

        custom_xml = _build_custom_xml(new_props)

        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as dst:
            for n in names:
                data = src.read(n)
                if n == 'docProps/custom.xml':
                    continue   # 새 버전으로 교체
                if n == '[Content_Types].xml' and not has_custom:
                    data = _patch_content_types(data)
                if n == '_rels/.rels' and not has_custom:
                    data = _patch_root_rels(data)
                dst.writestr(n, data)
            dst.writestr('docProps/custom.xml', custom_xml)

    shutil.move(str(tmp), str(fp))


# ─── 토큰 읽기 / 검증 ───────────────────────────────────────────

def read_token(file_path) -> dict | None:
    """파일의 custom properties 전체를 dict 로 반환. 없으면 None."""
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            if 'docProps/custom.xml' not in zf.namelist():
                return None
            xml_bytes = zf.read('docProps/custom.xml')
    except Exception:
        return None
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return None
    out = {}
    for prop in root.findall(f'{{{_NS_CP}}}property'):
        name = prop.get('name')
        if not name:
            continue
        v = prop.find(f'{{{_NS_VT}}}lpwstr')
        if v is not None and v.text is not None:
            out[name] = v.text
    return out


def verify_token(file_path, secret: bytes | None = None) -> dict:
    """파일 검증.

    반환:
      {
        'ok': bool,                # 시그 유효 여부
        'reason': str | None,      # 실패 사유 코드: 'token_missing'/'sig_invalid'/'payload_corrupt'
        'payload': dict | None,    # 유효한 경우 파싱된 payload
      }
    """
    key = _load_secret(secret)
    props = read_token(file_path) or {}
    payload_str = props.get(_PROP_PAYLOAD)
    sig = props.get(_PROP_SIG)
    if not payload_str or not sig:
        return {'ok': False, 'reason': 'token_missing', 'payload': None}

    expected = _hmac_sign(key, payload_str)
    if not hmac.compare_digest(expected, sig):
        return {'ok': False, 'reason': 'sig_invalid', 'payload': None}

    try:
        payload = json.loads(payload_str)
    except Exception:
        return {'ok': False, 'reason': 'payload_corrupt', 'payload': None}

    return {'ok': True, 'reason': None, 'payload': payload}
