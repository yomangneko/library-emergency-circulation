"""v2.0.3 고침 게이트 ⓐ~ⓒ (지시 v1 2026-11-07 §1-4). 대조군은 git 태그 v2.0.2 그대로. 화면 없이 돈다.

    python -I -X utf8 tests/fix2_gates.py

  ⓐ 판정 규칙 전 꼴(사서 결정 「나」): 손으로 넣은 기록 줄의 처리시간 「9:30」 글자 · 엑셀 시각 값 · 날짜 값 · 날짜·시각 값 ·
     「오전 9시」 · 빈칸 → 빈칸만 세지 않음. 구분이 대출/반납이 아닌 줄 · 도서번호 빈 줄은 세지 않음. 합계는 시험이 셈.
  ⓑ 수식 전 자리(㉹): 끼운 자리 아래(합계 칸 · 합계 아래 메모 칸) · 위(기록 줄 칸 · 머리 칸) · 범위(위쪽만 · 걸침 · 아래쪽만)
     → 다음 기록 뒤 수식이 가리키는 값 = 엑셀에서 같은 자리에 줄을 끼웠을 때의 기대값(시험이 셈).
  ⓒ 고칠 수 없는 수식이 있는 파일 → 저장하지 않음 · 파일 그대로 · 일지에 남음 · 빨간 문구.
"""

from __future__ import annotations

import io
import json
import os
import shutil
import subprocess
import sys
import tarfile
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIX_PROBE = os.path.join(HERE, "fix_probe.py")
CONTROL_TAG = "v2.0.2"
CONTROL_COMMIT_PREFIX = "0e71ac6"  # 지시서 §대상


def control_tree(work):
    commit = subprocess.run(["git", "-C", ROOT, "rev-parse", f"{CONTROL_TAG}^{{commit}}"], capture_output=True,
                            text=True, check=True).stdout.strip()
    if not commit.startswith(CONTROL_COMMIT_PREFIX):
        raise RuntimeError(f"대조군 커밋이 다름: {commit}")
    data = subprocess.run(["git", "-C", ROOT, "archive", "--format=tar", CONTROL_TAG, "library_emergency"],
                          capture_output=True, check=True).stdout
    dest = os.path.join(work, "control_v2_0_2")
    with tarfile.open(fileobj=io.BytesIO(data)) as tf:
        tf.extractall(dest, filter="data")
    return dest, commit


def run(pkg, mode):
    d = tempfile.mkdtemp(prefix=f"le_fix2_{mode}_", dir=os.getcwd())
    try:
        cmd = [sys.executable, "-I", "-X", "utf8", FIX_PROBE, "--pkg", pkg, "--folder", d, "--mode", mode]
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=900)
        line = [ln for ln in p.stdout.splitlines() if ln.startswith("RESULT ")]
        if not line:
            raise RuntimeError(f"탐침 실패 {mode} rc={p.returncode} {p.stderr[-800:]}")
        return json.loads(line[-1][7:])
    finally:
        shutil.rmtree(d, ignore_errors=True)


def table_gate(results, title, mode, ctl, describe):
    t0 = time.monotonic()
    o, f = run(ctl, mode), run(ROOT, mode)
    print(title)
    for name in f:
        print(f"    {name:<16} | v2.0.2: {'초록' if o[name]['ok'] else '붉음'} ({describe(o[name])}) | 고친 판: "
              f"{'초록' if f[name]['ok'] else '붉음'} ({describe(f[name])})")
    red_before = [k for k in f if not o[k]["ok"]]
    passed = bool(red_before) and all(v["ok"] for v in f.values())
    secs = time.monotonic() - t0
    return o, f, red_before, passed, secs


def main():
    work = tempfile.mkdtemp(prefix="le_fix2_ctl_", dir=os.getcwd())
    results = []
    try:
        ctl, commit = control_tree(work)
        print(f"python {sys.version.split()[0]} · {sys.platform} · 대조군 {CONTROL_TAG} = {commit}\n")

        def d_rule(x):
            if "exception" in x:
                return x["exception"]
            return (f"합계 {x['totals'].get('전체 건수')} / 기대 {x['want']['전체 건수']} · 화면 {x['view_rows']}건"
                    + (f" · 처리시간 표시 「{x['hand_time_shown']}」" if x.get("hand_time_shown") else ""))

        o, f, red, passed, secs = table_gate(
            results, "[ⓐ 판정 규칙 전 꼴] 3건 기록 → 기록 사이에 손으로 한 줄 → 다음 기록 (합계는 시험이 셈)", "rule", ctl, d_rule)
        print(f"  {'PASS' if passed else 'FAIL'}  ⓐ 판정 규칙 — 대조군 붉은 꼴 {len(red)}개 {red} → 고친 판 "
              f"{sum(v['ok'] for v in f.values())}/{len(f)} 초록 · {secs:.1f}초")
        results.append({"name": "ⓐ 판정 규칙 전 꼴", "pass": passed, "secs": round(secs, 1), "control": o, "fixed": f})

        def d_formula(x):
            if "exception" in x:
                return x["exception"]
            return f"{x['formula_after']} → {x['value']!r} / 기대 {x['want']!r}"

        o, f, red, passed, secs = table_gate(
            results, "\n[ⓑ 수식 전 자리] 기록 5~7행 · 합계 9~11행 · 다음 기록은 8행에 끼워짐", "formula", ctl, d_formula)
        print(f"  {'PASS' if passed else 'FAIL'}  ⓑ 수식 — 대조군 붉은 자리 {len(red)}곳 {red} → 고친 판 "
              f"{sum(v['ok'] for v in f.values())}/{len(f)} 초록 · {secs:.1f}초")
        results.append({"name": "ⓑ 수식 전 자리", "pass": passed, "secs": round(secs, 1), "control": o, "fixed": f})

        t0 = time.monotonic()
        o, f = run(ctl, "badformula"), run(ROOT, "badformula")
        good = (not f["saved"]) and f["file_unchanged"] and f["in_journal"] and f["status"].startswith("일지에만 기록됨") \
            and "저장 완료" not in f["status"]
        ctl_red = o["saved"]
        secs = time.monotonic() - t0
        print("\n[ⓒ 고칠 수 없는 수식(표 참조 =SUM(표1[금액]))이 있는 파일 → 다음 기록]")
        for who, x in (("v2.0.2", o), ("고친 판", f)):
            print(f"    {who}: 저장 {'됨' if x['saved'] else '안 됨'} · 파일 그대로={x['file_unchanged']} · 일지에 남음="
                  f"{x['in_journal']} · 상태 「{x['status']}」")
        passed = good and ctl_red
        print(f"  {'PASS' if passed else 'FAIL'}  ⓒ 고칠 수 없는 수식 — 대조군 {'붉음(그대로 저장)' if ctl_red else '초록'} → "
              f"고친 판 {'초록' if good else '붉음'} · {secs:.1f}초")
        results.append({"name": "ⓒ 고칠 수 없는 수식", "pass": passed, "secs": round(secs, 1), "control": o, "fixed": f})
    finally:
        shutil.rmtree(work, ignore_errors=True)
    n_pass = sum(r["pass"] for r in results)
    print(f"\n고침2 게이트 합계: {n_pass}/3 통과")
    with open(os.path.join(os.getcwd(), "fix2_result.json"), "w", encoding="utf-8") as fh:
        json.dump({"python": sys.version, "platform": sys.platform, "results": results}, fh, ensure_ascii=False,
                  indent=1)
    return 0 if n_pass == 3 and len(results) == 3 else 1


if __name__ == "__main__":
    sys.exit(main())
