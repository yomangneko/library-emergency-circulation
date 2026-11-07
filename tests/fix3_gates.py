"""v2.0.4 고침 게이트 ⓐ~ⓒ + 1-1 부품 표 (고침 지시 FIX3 2026-11-07 §1-4). 대조군은 git 태그 v2.0.3 그대로. 화면 없이 돈다.

    python -I -X utf8 tests/fix3_gates.py

  표  1-1: openpyxl 3.1.5 자체 실측 — 부품마다 「읽고 저장하면 남는가 · 줄을 끼우면 옮기는가」(판과 상관없음)
  ⓐ 옮기는 부품: 부품마다 · 자리마다(합계 줄 · 기록 줄 · 머리 칸 · 걸침) 3건 기록 → 부품 → 다음 기록
     → 가리키는 자리 = 시험이 직접 셈(8행에 1줄 끼움). 고친 판 전부 초록 · 대조군 붉은 자리 하나 이상
  ⓑ 쓰지 않는 부품: 파일 지문 그대로 · 일지에 남음 · 빨간 문구(까닭에 부품 이름). 고친 판 전부 초록 · 대조군 전부 붉음
  ⓒ 실제 프로그램(Numbers)으로 만든 시험 자료(tests/fixtures): 고친 판 전부 초록 · 대조군 붉음 하나 이상
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
PROBE = os.path.join(HERE, "fix3_probe.py")
CONTROL_TAG = "v2.0.3"
CONTROL_COMMIT_PREFIX = "7c033e9"  # 지시서 §대상


def control_tree(work):
    commit = subprocess.run(["git", "-C", ROOT, "rev-parse", f"{CONTROL_TAG}^{{commit}}"], capture_output=True,
                            text=True, check=True).stdout.strip()
    if not commit.startswith(CONTROL_COMMIT_PREFIX):
        raise RuntimeError(f"대조군 커밋이 다름: {commit}")
    data = subprocess.run(["git", "-C", ROOT, "archive", "--format=tar", CONTROL_TAG, "library_emergency"],
                          capture_output=True, check=True).stdout
    dest = os.path.join(work, "control_v2_0_3")
    with tarfile.open(fileobj=io.BytesIO(data)) as tf:
        tf.extractall(dest, filter="data")
    return dest, commit


def run(pkg, mode):
    d = tempfile.mkdtemp(prefix=f"le_fix3_{mode}_", dir=os.getcwd())
    try:
        cmd = [sys.executable, "-I", "-X", "utf8", PROBE, "--pkg", pkg, "--folder", d, "--mode", mode]
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=1800)
        line = [ln for ln in p.stdout.splitlines() if ln.startswith("RESULT ")]
        if not line:
            raise RuntimeError(f"탐침 실패 {mode} rc={p.returncode} {p.stderr[-800:]}")
        return json.loads(line[-1][7:])
    finally:
        shutil.rmtree(d, ignore_errors=True)


def g(ok):
    return "초록" if ok else "붉음"


def main():
    work = tempfile.mkdtemp(prefix="le_fix3_ctl_", dir=os.getcwd())
    results = []
    try:
        ctl, commit = control_tree(work)
        import openpyxl

        print(f"python {sys.version.split()[0]} · {sys.platform} · openpyxl {openpyxl.__version__} · "
              f"대조군 {CONTROL_TAG} = {commit}\n")

        t0 = time.monotonic()
        inv = run(ROOT, "inventory")
        print("[1-1 부품 표 — openpyxl 자체 실측 · 합계 줄에 붙이고 8행에 1줄 끼움]")
        print(f"    {'부품':<24} | 읽고 저장하면 남는가 | 줄을 끼우면 옮기는가")
        for r in inv:
            kept = "남음" if r["kept"] else "지워짐"
            if r.get("kept_rich_text") is not None and r["kept_rich_text"] != r["kept"]:
                kept += f" (꾸밈 읽기 rich_text 로는 {'남음' if r['kept_rich_text'] else '지워짐'})"
            moved = "—" if r["moved"] is None else ("옮김" if r["moved"] else "안 옮김")
            print(f"    {r['part']:<24} | {kept} | {moved}")
        results.append({"name": "1-1 부품 표", "pass": True, "secs": round(time.monotonic() - t0, 1),
                        "rows": inv})

        t0 = time.monotonic()
        o, f = run(ctl, "move"), run(ROOT, "move")
        print("\n[ⓐ 옮기는 부품] 3건 기록 → 부품 → 다음 기록(8행에 1줄) · 기대값은 시험이 직접 셈")
        for k in f:
            print(f"    {k:<28} | v2.0.3: {g(o[k]['ok'])} {o[k].get('got', o[k].get('exception'))} | "
                  f"고친 판: {g(f[k]['ok'])} {f[k].get('got', f[k].get('exception'))} | 기대 {f[k].get('want')}")
        red = [k for k in f if not o[k]["ok"]]
        passed = bool(red) and all(v["ok"] for v in f.values())
        secs = time.monotonic() - t0
        print(f"  {'PASS' if passed else 'FAIL'}  ⓐ 옮기는 부품 — 대조군 붉은 자리 {len(red)}/{len(f)} → 고친 판 "
              f"{sum(v['ok'] for v in f.values())}/{len(f)} 초록 · {secs:.1f}초")
        results.append({"name": "ⓐ 옮기는 부품", "pass": passed, "secs": round(secs, 1), "control": o, "fixed": f})

        t0 = time.monotonic()
        o, f = run(ctl, "refuse"), run(ROOT, "refuse")
        print("\n[ⓑ 쓰지 않는 부품] 3건 기록 → 부품 → 다음 기록 · 파일 그대로 · 일지에 남음 · 빨간 문구")
        for k in f:
            print(f"    {k:<24} | v2.0.3: {g(o[k]['ok'])} (저장 {'됨' if o[k].get('saved') else '안 됨'}) | "
                  f"고친 판: {g(f[k]['ok'])} (저장 {'됨' if f[k].get('saved') else '안 됨'} · 파일 그대로="
                  f"{f[k].get('file_unchanged')} · 일지={f[k].get('in_journal')}) 「{f[k].get('status', f[k].get('exception'))}」")
        passed = all(not v["ok"] for v in o.values()) and all(v["ok"] for v in f.values())
        secs = time.monotonic() - t0
        print(f"  {'PASS' if passed else 'FAIL'}  ⓑ 쓰지 않는 부품 — 대조군 붉음 {sum(not v['ok'] for v in o.values())}"
              f"/{len(o)} → 고친 판 {sum(v['ok'] for v in f.values())}/{len(f)} 초록 · {secs:.1f}초")
        results.append({"name": "ⓑ 쓰지 않는 부품", "pass": passed, "secs": round(secs, 1), "control": o, "fixed": f})

        t0 = time.monotonic()
        o, f = run(ctl, "real"), run(ROOT, "real")
        print("\n[ⓒ 실제 프로그램(Numbers)이 내보낸 시험 자료]")
        for k in f:
            print(f"    {k} (기대: {f[k].get('expect')}) | v2.0.3: {g(o[k]['ok'])} (저장 "
                  f"{'됨' if o[k].get('saved') else '안 됨'}) | 고친 판: {g(f[k]['ok'])} 「{f[k].get('status', f[k].get('exception'))}」")
            if "got" in f[k]:
                print(f"        기대 {f[k]['want']}\n        고친 판 {f[k]['got']}\n        v2.0.3 {o[k].get('got')}")
        passed = bool(f) and all(v["ok"] for v in f.values()) and any(not v["ok"] for v in o.values())
        secs = time.monotonic() - t0
        print(f"  {'PASS' if passed else 'FAIL'}  ⓒ 실제 프로그램 시험 자료 — 대조군 붉음 {sum(not v['ok'] for v in o.values())}"
              f"/{len(o)} → 고친 판 {sum(v['ok'] for v in f.values())}/{len(f)} 초록 · {secs:.1f}초")
        results.append({"name": "ⓒ 실제 프로그램 시험 자료", "pass": passed, "secs": round(secs, 1), "control": o,
                        "fixed": f})
    finally:
        shutil.rmtree(work, ignore_errors=True)
    gates = [r for r in results if r["name"] != "1-1 부품 표"]
    n_pass = sum(r["pass"] for r in gates)
    print(f"\n고침3 게이트 합계: {n_pass}/3 통과")
    with open(os.path.join(os.getcwd(), "fix3_result.json"), "w", encoding="utf-8") as fh:
        json.dump({"python": sys.version, "platform": sys.platform, "results": results}, fh, ensure_ascii=False,
                  indent=1, default=str)
    return 0 if n_pass == 3 and len(gates) == 3 else 1


if __name__ == "__main__":
    sys.exit(main())
