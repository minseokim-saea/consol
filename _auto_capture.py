"""
연결재무보고 시스템 자동 화면 캡쳐
- Flask 서버를 백그라운드로 띄우고
- Playwright(Chromium)로 주요 화면을 캡쳐하여 ppt_screenshots/ 에 저장
- 사용법:
    $env:ADMIN_PASSWORD = "본인_관리자_비밀번호"
    $env:ADMIN_USER     = "admin"   # (선택, 기본 admin)
    python _auto_capture.py
"""
import os, sys, time, subprocess, signal, socket, getpass
from pathlib import Path

BASE = Path(__file__).resolve().parent
SHOTS = BASE / "ppt_screenshots"
SHOTS.mkdir(exist_ok=True)
HOST = "127.0.0.1"
PORT = 5000
BASE_URL = f"http://{HOST}:{PORT}"

USER = os.environ.get("ADMIN_USER", "admin")
PWD  = os.environ.get("ADMIN_PASSWORD")
if not PWD:
    try:
        PWD = getpass.getpass(f"[?] '{USER}' 계정의 비밀번호를 입력하세요: ")
    except Exception:
        print("[ERR] 비밀번호가 필요합니다. 환경변수 ADMIN_PASSWORD 또는 입력 프롬프트로 제공하세요.")
        sys.exit(1)

def port_in_use(host, port, timeout=0.5):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(timeout)
    try:
        s.connect((host, port)); s.close(); return True
    except Exception:
        return False

# 1) Flask 서버 기동 (이미 떠 있으면 재사용)
server = None
if port_in_use(HOST, PORT):
    print(f"[i] 이미 {BASE_URL} 에 서버가 실행 중. 그대로 사용합니다.")
else:
    print(f"[i] Flask 서버 기동 중 ({BASE_URL}) ...")
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    server = subprocess.Popen(
        [sys.executable, "app.py"],
        cwd=str(BASE),
        env=env,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0,
    )
    # wait until up
    for _ in range(60):
        if port_in_use(HOST, PORT):
            break
        time.sleep(0.5)
    else:
        print("[ERR] 서버 기동 실패 (60초 초과)")
        server.terminate()
        sys.exit(1)
    print("[OK] 서버 기동 완료")

# 2) Playwright 캡쳐
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# (slide_num, filename_prefix, label, url, requires_login, post_login_action)
TARGETS = [
    # 로그인 화면 — 미로그인 상태로 캡쳐 (가장 먼저)
    (9,  "09_login",         "로그인 화면",          "/login",                     False, None),
    # 로그인 이후
    (8,  "08_main",          "메인 대시보드",        "/",                          True,  None),
    (10, "10_admin_users",   "사용자 관리",          "/admin/users",               True,  None),
    (10, "10_admin_permission", "권한 그룹 관리",     "/admin/permission-groups",   True,  None),
    (11, "11_upload",        "패키지 업로드",        "/",                          True,  "scroll_upload"),
    (12, "12_aggregate",     "자동 합산 결과",       "/",                          True,  "scroll_aggregate"),
    (13, "13_fx",            "환율 관리",            "/",                          True,  "scroll_fx"),
    (14, "14_wce_input",     "WCE 입력",             "/admin/wce",                 True,  None),
    (15, "15_wce_aggregate", "WCE 통합 집계",        "/admin/wce/aggregate",       True,  None),
    (16, "16_consolidation", "연결 그룹 관리",       "/consolidation",             True,  None),
    (17, "17_journal",       "연결조정 분개",        "/consolidation",             True,  "scroll_journal"),
    (18, "18_dashboard",     "연결 KPI 대시보드",    "/consolidation/dashboard",   True,  None),
]

shot_results = []
with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=True)
    ctx = browser.new_context(viewport={"width": 1600, "height": 1000}, locale="ko-KR")
    page = ctx.new_page()

    # 먼저 로그인 페이지 캡쳐
    print(f"[*] {TARGETS[0][3]} (비로그인 캡쳐)")
    try:
        page.goto(BASE_URL + "/login", wait_until="networkidle", timeout=15000)
        page.wait_for_timeout(800)
        path = SHOTS / f"{TARGETS[0][1]}.png"
        page.screenshot(path=str(path), full_page=True)
        shot_results.append((TARGETS[0][1], "OK", str(path)))
        print(f"    -> {path.name}")
    except Exception as e:
        shot_results.append((TARGETS[0][1], f"FAIL: {e}", ""))

    # 로그인
    print(f"[*] 로그인 시도 ({USER})")
    try:
        page.goto(BASE_URL + "/login", wait_until="networkidle", timeout=15000)
        # 일반적인 폼 셀렉터들을 시도
        filled = False
        for u_sel, p_sel in [
            ("input[name='username']",  "input[name='password']"),
            ("input[id='username']",    "input[id='password']"),
            ("input[type='text']",      "input[type='password']"),
        ]:
            if page.locator(u_sel).count() and page.locator(p_sel).count():
                page.fill(u_sel, USER)
                page.fill(p_sel, PWD)
                filled = True
                break
        if not filled:
            raise RuntimeError("로그인 폼 input을 찾지 못했습니다.")
        # submit
        if page.locator("button[type='submit']").count():
            page.click("button[type='submit']")
        elif page.locator("input[type='submit']").count():
            page.click("input[type='submit']")
        else:
            page.keyboard.press("Enter")
        page.wait_for_load_state("networkidle", timeout=15000)
        # 비번 변경 강제 페이지로 이동될 수 있음 - 그 경우 패스
        if "/login" in page.url:
            print(f"[ERR] 로그인 실패 (여전히 /login). 비밀번호 또는 계정명을 확인하세요.")
            print(f"      현재 url: {page.url}")
            sys.exit(2)
        print(f"    -> 로그인 성공, 현재 url: {page.url}")
    except Exception as e:
        print(f"[ERR] 로그인 실패: {e}")
        sys.exit(2)

    # 비밀번호 강제 변경 화면일 수 있음 — change-password 화면도 캡쳐 후 메인으로
    if "/change-password" in page.url or "change_password" in page.url.lower():
        try:
            path = SHOTS / "09_change_password.png"
            page.screenshot(path=str(path), full_page=True)
            print(f"    -> {path.name} (비밀번호 변경 화면)")
        except Exception:
            pass

    # 나머지 캡쳐
    for slide_num, prefix, label, url, needs_login, action in TARGETS[1:]:
        full_url = BASE_URL + url
        print(f"[*] {label}  ({full_url})")
        try:
            page.goto(full_url, wait_until="networkidle", timeout=20000)
            page.wait_for_timeout(1500)  # 차트/AJAX 렌더 대기
            # 액션
            if action == "scroll_upload":
                # 업로드 섹션이 있으면 그쪽으로 스크롤
                for sel in ['#upload-section', '#uploadSection', 'form[enctype="multipart/form-data"]', 'input[type="file"]']:
                    if page.locator(sel).count():
                        try: page.locator(sel).first.scroll_into_view_if_needed()
                        except: pass
                        break
                page.wait_for_timeout(500)
            elif action == "scroll_aggregate":
                for sel in ['#aggregate', '#result', '#aggregateSection', 'button:has-text("합산")', 'a:has-text("다운로드")']:
                    if page.locator(sel).count():
                        try: page.locator(sel).first.scroll_into_view_if_needed()
                        except: pass
                        break
                page.wait_for_timeout(500)
            elif action == "scroll_fx":
                for sel in ['#fx', '#fx-rates', 'text=환율', 'text=FX']:
                    if page.locator(sel).count():
                        try: page.locator(sel).first.scroll_into_view_if_needed()
                        except: pass
                        break
                page.wait_for_timeout(500)
            elif action == "scroll_journal":
                for sel in ['#journal', 'text=분개', 'text=연결조정']:
                    if page.locator(sel).count():
                        try: page.locator(sel).first.scroll_into_view_if_needed()
                        except: pass
                        break
                page.wait_for_timeout(500)
            path = SHOTS / f"{prefix}.png"
            page.screenshot(path=str(path), full_page=True)
            shot_results.append((prefix, "OK", str(path)))
            print(f"    -> {path.name}")
        except PWTimeout as e:
            shot_results.append((prefix, f"TIMEOUT", ""))
            print(f"    [!] 타임아웃: {e}")
        except Exception as e:
            shot_results.append((prefix, f"FAIL: {e}", ""))
            print(f"    [!] 실패: {e}")

    browser.close()

# 결과 요약
print("\n" + "="*60)
print("캡쳐 결과 요약")
print("="*60)
ok = sum(1 for _, s, _ in shot_results if s == "OK")
print(f"성공: {ok} / {len(shot_results)}")
for prefix, status, path in shot_results:
    mark = "[OK]" if status == "OK" else "[--]"
    print(f"  {mark} {prefix:<28} {status}")

# 서버 종료
if server is not None:
    print("\n[i] Flask 서버 종료 중...")
    try:
        server.terminate()
        server.wait(timeout=5)
    except Exception:
        try: server.kill()
        except Exception: pass

# PPT 재빌드
print("\n[i] PPT 재빌드 중...")
rc = subprocess.run([sys.executable, str(BASE / "_build_ppt.py")], cwd=str(BASE)).returncode
if rc == 0:
    print("[OK] PPT 재빌드 완료")
else:
    print("[ERR] PPT 재빌드 실패")
    sys.exit(rc)
