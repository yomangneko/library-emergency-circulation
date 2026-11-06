"""v2.0.2 고침 게이트 ⓐ~ⓓ (지시 v1 2026-11-06 §1-5). 대조군은 git 태그 v2.0.1 그대로.

    python -I -X utf8 tests/fix_gates.py

  ⓐ 닫기(㉷): 그날 1,000 · 5,000 · 10,000건 위에서 5건 연달아 찍고 곧바로 닫기 → 엑셀 = 일지(빠짐 0).
     엑셀을 잠근 채 닫기 → 알림이 뜨는가 · 「그냥 닫기」 뒤 다시 켜면 채워지는가 · 「다시 시도」로 채우고 닫히는가.
  ⓑ 끊긴 줄 전 자리(㉸): 한 줄의 첫 바이트부터 마지막 바이트까지 모든 자리에서 끊은 꼴 → 다시 켜기 → 다음 스캔.
     + v2.0.1 이 만든 일지를 새 판이 같은 줄 수로 읽는가 + v2.0.1 이 쓰다 끊긴 일지를 새 판이 이어받는 꼴.
  ⓒ 메모 전 자리(②나): 합계 아래 · 기록과 합계 사이 빈 줄 · 합계 줄 옆 칸 · 기록 줄 오른쪽 칸 · 기록 줄 사이 · 머리 위.
  ⓓ 빨간 문구 끝 「엑셀을 닫으면 곧 채워집니다」(①가).
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
SPEED_PROBE = os.path.join(HERE, "speed_probe.py")
FIX_PROBE = os.path.join(HERE, "fix_probe.py")
CONTROL_TAG = "v2.0.1"
CONTROL_COMMIT_PREFIX = "12a2cf0"  # 지시서 §대상
WITH_GUI = "--no-gui" not in sys.argv  # 맥에서 창 없는 부분만 돌릴 때 --no-gui (화면 시험은 윈도우 빌드 기계에서)

sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)
from speed_probe import DAY, jname, journal_r, prefill, read_sheet, xname  # noqa: E402
from library_emergency import core as fixed_core  # noqa: E402  미리 채우기(시험 준비) · ⓓ 문구에만 쓴다
from library_emergency.messages import status_for  # noqa: E402


def control_tree(work):
    commit = subprocess.run(["git", "-C", ROOT, "rev-parse", f"{CONTROL_TAG}^{{commit}}"], capture_output=True,
                            text=True, check=True).stdout.strip()
    if not commit.startswith(CONTROL_COMMIT_PREFIX):
        raise RuntimeError(f"대조군 커밋이 다름: {commit}")
    data = subprocess.run(["git", "-C", ROOT, "archive", "--format=tar", CONTROL_TAG, "library_emergency"],
                          capture_output=True, check=True).stdout
    dest = os.path.join(work, "control_v2_0_1")
    with tarfile.open(fileobj=io.BytesIO(data)) as tf:
        tf.extractall(dest, filter="data")
    return dest, commit


def run(script, pkg, folder, mode, *extra, timeout=1800):
    cmd = [sys.executable, "-I", "-X", "utf8", script, "--pkg", pkg, "--folder", folder, "--mode", mode, *extra]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=timeout)
    line = [ln for ln in p.stdout.splitlines() if ln.startswith("RESULT ")]
    if not line:
        raise RuntimeError(f"탐침 실패 {mode} rc={p.returncode} {p.stderr[-800:]}")
    return json.loads(line[-1][7:])


def newdir(tag):
    return tempfile.mkdtemp(prefix=f"le_fix_{tag}_", dir=os.getcwd())


def counts(folder):
    xp = os.path.join(folder, xname(DAY))
    xl = len(read_sheet(xp)[0]) if os.path.exists(xp) else 0
    return xl, journal_r(os.path.join(folder, jname(DAY)))


# ------------------------------------------------------------------ ⓐ


def close_after(pkg, n):
    d = newdir("close")
    try:
        prefill(fixed_core, d, n)
        out = run(SPEED_PROBE, pkg, d, "closeafter", "--n", str(n))
        time.sleep(0.5)
        xl, jn = counts(d)
        out.update({"excel_rows": xl, "journal_rows": jn, "missing": jn - xl})
        return out
    finally:
        shutil.rmtree(d, ignore_errors=True)


def lock_close(pkg, answer):
    d = newdir("lockclose")
    try:
        prefill(fixed_core, d, 1000)
        out = run(SPEED_PROBE, pkg, d, "lockclose", "--n", "1000", "--answer", answer)
        time.sleep(0.5)
        xl, jn = counts(d)
        out.update({"excel_rows_after_close": xl, "journal_rows": jn})
        r = run(SPEED_PROBE, pkg, d, "restart")
        out.update({"excel_rows_after_restart": r["excel_rows"], "restart_journal_rows": r["journal_rows"]})
        return out
    finally:
        shutil.rmtree(d, ignore_errors=True)


def gate_a(ctl, results):
    t0 = time.monotonic()
    print("[ⓐ 닫기] 그날 n건 위에서 5건(0.15초 간격) → 곧바로 닫기 → 엑셀에 빠진 줄")
    print("    n건 | v2.0.1(대조군) | 고친 판 | 고친 판 닫히기까지")
    rows = []
    ctl_red = False
    fix_ok = True
    for n in (1000, 5000, 10000):
        o = close_after(ctl, n)
        f = close_after(ROOT, n)
        rows.append({"n": n, "control": o, "fixed": f})
        ctl_red |= o["missing"] != 0
        good = f["missing"] == 0 and not f["dialogs"]
        fix_ok &= good
        print(f"  {n:>6} | 빠짐 {o['missing']}{' 붉음' if o['missing'] else ' 초록'} | 빠짐 {f['missing']}"
              f"{' 초록' if good else ' 붉음'} | {f['close_secs']}초")
    print("  엑셀이 잠긴 채 5건 → 닫기")
    lc = {}
    for who, pkg in (("대조군", ctl), ("고친 판", ROOT)):
        for answer in ("close", "retry"):
            lc[(who, answer)] = r = lock_close(pkg, answer)
            print(f"    {who} · 답 「{'그냥 닫기' if answer == 'close' else '다시 시도'}」: 알림 {len(r['dialogs'])}번"
                  f"{' (엑셀에 없는 건수 ' + str(r['dialogs'][0]['missing']) + ')' if r['dialogs'] else ''}"
                  f" · 닫힌 직후 엑셀 {r['excel_rows_after_close']}줄 / 일지 {r['journal_rows']}줄"
                  f" · 다시 켠 뒤 엑셀 {r['excel_rows_after_restart']}줄")
    fc, fr = lc[("고친 판", "close")], lc[("고친 판", "retry")]
    oc = lc[("대조군", "close")]
    lock_fix_ok = (len(fc["dialogs"]) == 1 and fc["dialogs"][0]["missing"] == 5 and
                   fc["excel_rows_after_restart"] == fc["restart_journal_rows"] == 1005 and
                   len(fr["dialogs"]) == 1 and fr["excel_rows_after_close"] == fr["journal_rows"] == 1005)
    lock_ctl_red = not oc["dialogs"] and oc["excel_rows_after_close"] < oc["journal_rows"]
    passed = ctl_red and fix_ok and lock_fix_ok and lock_ctl_red
    secs = time.monotonic() - t0
    print(f"  {'PASS' if passed else 'FAIL'}  ⓐ 닫기 — 대조군 {'붉음' if ctl_red and lock_ctl_red else '초록(잣대 불인정)'}"
          f" → 고친 판 {'초록' if fix_ok and lock_fix_ok else '붉음'} · {secs:.0f}초")
    results.append({"name": "ⓐ 닫기", "pass": passed, "secs": round(secs, 1), "close_after": rows,
                    "lock_close": {f"{k[0]}/{k[1]}": v for k, v in lc.items()}})


# ------------------------------------------------------------------ ⓑ


def trunc(writer, reader):
    d = newdir("trunc")
    try:
        w = run(FIX_PROBE, writer, d, "trunc_write")
        r = run(FIX_PROBE, reader, d, "trunc_read")
        r["line_bytes"] = w["line_bytes"]
        return r
    finally:
        shutil.rmtree(d, ignore_errors=True)


def gate_b(ctl, results):
    t0 = time.monotonic()
    print("\n[ⓑ 끊긴 줄 전 자리] 셋째 줄을 첫 바이트부터 마지막 바이트까지 자리마다 끊음 → 다시 켬 → 다음 스캔")
    o = trunc(ctl, ctl)
    f = trunc(ROOT, ROOT)
    up = trunc(ctl, ROOT)  # v2.0.1 이 쓰다 끊긴 일지를 새 판이 이어받음
    for label, r in (("v2.0.1 → v2.0.1(대조군)", o), ("고친 판 → 고친 판", f), ("v2.0.1 일지 → 고친 판(판 올림)", up)):
        ex = "; ".join(json.dumps(x, ensure_ascii=False) for x in r["bad_examples"][:2])
        print(f"    {label}: 줄 {r['line_bytes']}바이트 · 끊은 자리 {r['cases']}곳 중 틀림 {r['bad']}"
              + (f" — 예: {ex}" if r["bad"] else ""))
    d = newdir("journal")
    try:
        jw = run(FIX_PROBE, ctl, d, "journal_write")
        jr = run(FIX_PROBE, ROOT, d, "journal_read")
    finally:
        shutil.rmtree(d, ignore_errors=True)
    same = jw["rows"] == jr["rows"] and jw["saved_count"] == jr["saved_count"] and jw["keys"] == jr["keys"]
    print(f"    v2.0.1 이 만든 일지(저장 실패 구간 섞음): v2.0.1 읽기 {jw['rows']}줄 · 옮김 표시 {jw['saved_count']} / "
          f"고친 판 읽기 {jr['rows']}줄 · 옮김 표시 {jr['saved_count']} · 줄 내용까지 같음={jw['keys'] == jr['keys']}")
    passed = o["bad"] > 0 and f["bad"] == 0 and up["bad"] == 0 and f["cases"] == f["line_bytes"] and same
    secs = time.monotonic() - t0
    print(f"  {'PASS' if passed else 'FAIL'}  ⓑ 끊긴 줄 — 대조군 {'붉음' if o['bad'] else '초록(잣대 불인정)'} → "
          f"고친 판 {'초록' if f['bad'] == 0 and up['bad'] == 0 else '붉음'} · 옛 일지 같게 읽음={same} · {secs:.0f}초")
    results.append({"name": "ⓑ 끊긴 줄 전 자리", "pass": passed, "secs": round(secs, 1), "control": o, "fixed": f,
                    "upgrade": up, "journal_compat": {"v2.0.1": {k: jw[k] for k in ("rows", "saved_count")},
                                                      "fixed": {k: jr[k] for k in ("rows", "saved_count")},
                                                      "same_keys": jw["keys"] == jr["keys"]}})


# ------------------------------------------------------------------ ⓒ

PLACE_NAMES = {"below_totals": "합계 아래", "gap_before_totals": "기록과 합계 사이 빈 줄", "totals_side": "합계 줄 옆 칸",
               "record_right": "기록 줄 오른쪽 칸", "between_records": "기록 줄 사이 메모 줄", "header": "머리 위(3행)"}


def gate_c(ctl, results):
    t0 = time.monotonic()
    print("\n[ⓒ 메모 전 자리] 3건 기록 → 사람이 메모 → 다음 기록. 메모 그대로 · 합계 = 시험이 센 4건 · 엑셀 저장됨")
    d1, d2 = newdir("memo_ctl"), newdir("memo_fix")
    try:
        o = run(FIX_PROBE, ctl, d1, "memo")
        f = run(FIX_PROBE, ROOT, d2, "memo")
    finally:
        shutil.rmtree(d1, ignore_errors=True)
        shutil.rmtree(d2, ignore_errors=True)

    def desc(x):
        if "exception" in x:
            return x["exception"]
        return (f"메모 {'그대로' if x['memo_kept'] else '사라짐'} · 합계 {x['totals'].get('전체 건수')} · "
                f"저장 {'됨' if x['saved'] else '안 됨'}")

    for k, name in PLACE_NAMES.items():
        print(f"    {name:<14} | v2.0.1: {'초록' if o[k]['ok'] else '붉음'} ({desc(o[k])}) | 고친 판: "
              f"{'초록' if f[k]['ok'] else '붉음'} ({desc(f[k])})")
    red_before = [k for k in PLACE_NAMES if not o[k]["ok"]]
    passed = bool(red_before) and all(f[k]["ok"] for k in PLACE_NAMES)
    secs = time.monotonic() - t0
    print(f"  {'PASS' if passed else 'FAIL'}  ⓒ 메모 — 대조군 붉은 자리 {len(red_before)}곳 → 고친 판 "
          f"{sum(f[k]['ok'] for k in PLACE_NAMES)}/{len(PLACE_NAMES)}곳 초록 · {secs:.0f}초")
    results.append({"name": "ⓒ 메모 전 자리", "pass": passed, "secs": round(secs, 1), "control": o, "fixed": f})


# ------------------------------------------------------------------ ⓓ


def gate_d(results):
    row = fixed_core.Row("2026-11-06 09:00:00", "대출", "M1", "B1")
    text = status_for(fixed_core.RecordResult(True, row=row, excel_saved=False, pending=3,
                                              error="엑셀 파일이 열려 있어 저장하지 못했습니다"))[0]
    want_end = "엑셀 미반영 3건은 엑셀을 닫으면 곧 채워집니다"
    passed = text.endswith(want_end) and "다음 기록 때" not in text
    print(f"\n[ⓓ 빨간 문구]\n  {'PASS' if passed else 'FAIL'}  「{text}」")
    results.append({"name": "ⓓ 빨간 문구", "pass": passed, "text": text})


def main():
    work = newdir("ctl")
    results = []
    try:
        ctl, commit = control_tree(work)
        print(f"python {sys.version.split()[0]} · {sys.platform} · 대조군 {CONTROL_TAG} = {commit}\n")
        if WITH_GUI:
            gate_a(ctl, results)
        else:
            print("[ⓐ 닫기] 화면 시험 — 이 기계에서는 건너뜀(--no-gui)")
        gate_b(ctl, results)
        gate_c(ctl, results)
        gate_d(results)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    want = 4 if WITH_GUI else 3
    n_pass = sum(r["pass"] for r in results)
    print(f"\n고침 게이트 합계: {n_pass}/{want} 통과")
    with open(os.path.join(os.getcwd(), "fix_result.json"), "w", encoding="utf-8") as fh:
        json.dump({"python": sys.version, "platform": sys.platform, "results": results}, fh, ensure_ascii=False,
                  indent=1)
    return 0 if n_pass == want and len(results) == want else 1


if __name__ == "__main__":
    sys.exit(main())
