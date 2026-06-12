"""구글 OTP(Google Authenticator) 호환 TOTP 유틸 — 표준 라이브러리만 사용.

RFC 6238(TOTP) / RFC 4226(HOTP) 기준. 외부 패키지(pyotp 등) 불필요.
Google Authenticator 기본값(SHA1, 6자리, 30초)과 호환된다.
"""

import base64
import hashlib
import hmac
import secrets
import struct
import time
import urllib.parse

PERIOD = 30
DIGITS = 6
ISSUER = '연결재무보고시스템'


def generate_secret(num_bytes: int = 20) -> str:
    """무작위 base32 시크릿 생성(패딩 '=' 제거). 기본 20바이트(=160bit)."""
    return base64.b32encode(secrets.token_bytes(num_bytes)).decode('ascii').rstrip('=')


def _hotp(secret_b32: str, counter: int, digits: int = DIGITS) -> str:
    """RFC 4226 HOTP. secret_b32 는 base32 문자열(패딩 유무 무관)."""
    s = (secret_b32 or '').strip().replace(' ', '').upper()
    s += '=' * (-len(s) % 8)                    # base32 패딩 복원
    key = base64.b32decode(s)
    msg = struct.pack('>Q', counter)
    digest = hmac.new(key, msg, hashlib.sha1).digest()
    offset = digest[-1] & 0x0F
    code_int = struct.unpack('>I', digest[offset:offset + 4])[0] & 0x7FFFFFFF
    return str(code_int % (10 ** digits)).zfill(digits)


def totp_now(secret_b32: str, at: float = None, period: int = PERIOD, digits: int = DIGITS) -> str:
    """현재(또는 지정 시각)의 TOTP 코드."""
    if at is None:
        at = time.time()
    return _hotp(secret_b32, int(at // period), digits)


def verify(secret_b32: str, code: str, at: float = None,
           period: int = PERIOD, digits: int = DIGITS, window: int = 1) -> bool:
    """코드 검증. window=1 → 직전/직후 30초 스텝까지 허용(시계 오차 보정)."""
    if not secret_b32 or code is None:
        return False
    code = str(code).strip().replace(' ', '').replace('-', '')
    if not code.isdigit() or len(code) != digits:
        return False
    if at is None:
        at = time.time()
    counter = int(at // period)
    for w in range(-window, window + 1):
        if counter + w < 0:
            continue
        if hmac.compare_digest(_hotp(secret_b32, counter + w, digits), code):
            return True
    return False


def provisioning_uri(secret_b32: str, account_name: str, issuer: str = ISSUER,
                     period: int = PERIOD, digits: int = DIGITS) -> str:
    """otpauth:// URI — QR로 변환해 인증 앱에 등록하는 데 사용."""
    label = urllib.parse.quote(f'{issuer}:{account_name}')
    params = urllib.parse.urlencode({
        'secret': secret_b32,
        'issuer': issuer,
        'algorithm': 'SHA1',
        'digits': digits,
        'period': period,
    })
    return f'otpauth://totp/{label}?{params}'
