"""파일명 패턴 + 실제 build_distribution_package 결과 검증"""
import sys, json
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))

import distribute_builder as db

# 1) 패턴 점검 (in-memory)
import re
for y, q in [(2026, 1), (2026, 2), (2026, 4), (2025, 3), (2030, 2)]:
    yy = f"{int(y) % 100:02d}"
    safe_co = re.sub(r'[\\/:*?"<>|]', '_', "글로벌세아")
    print(f"  Y={y} Q={q}  →  {safe_co}_{yy}Q{q}.xlsm")

# 2) 실제 build_distribution_package 호출 시 파일명 확인
print("\n=== build_distribution_package 실제 출력 ===")
state = json.loads(Path("uploads/_state.json").read_text(encoding="utf-8"))
tpl = db.get_template_path("2026")
out_dir = Path("results/distribute/_test_filename")
out_dir.mkdir(parents=True, exist_ok=True)
for p in out_dir.glob("*.xlsm"):
    p.unlink()
r = db.build_distribution_package(
    template_path=tpl, output_dir=out_dir, company="글로벌세아",
    target_year=2026, target_quarter=1,
    uploaded_files=state, file_password=None, sheet_protect_password=None,
)
print(f"output_path: {r['output_path']}")
print(f"basename:    {Path(r['output_path']).name}")
expected = "글로벌세아_26Q1.xlsm"
if Path(r['output_path']).name == expected:
    print(f"[PASS] 파일명 일치 — {expected}")
else:
    print(f"[FAIL] 기대: {expected}, 실제: {Path(r['output_path']).name}")
