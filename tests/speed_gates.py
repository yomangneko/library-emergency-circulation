"""속도 고침 게이트 ⓐ~ⓔ (지시 v1 2026-11-06 §2-7).

    python -I -X utf8 tests/speed_gates.py

  ⓐ 속도: 그날 1 · 1,000 · 5,000 · 10,000건째 Enter → 다음 입력까지(화면을 띄운 채). 기준 0.2초 이하.
     대조군은 흉내 사본이 아니라 v2.0.0 코드(git 태그에서 그대로 꺼냄) — 대조군이 붉고 고친 판이 초록일 때만 통과.
  ⓑ 10,000건 대조: 엑셀 줄 수 = 일지 줄 수 = 화면 건수 = 합계(기대값은 탐침이 직접 셈)
  ⓒ 바깥 수정 보존: 사람이 지운 줄이 돌아오지 않음 · 합계 아래 메모 보존
  ⓓ 잠금 중 연속 스캔 20건: 화면 멈춤 없음 · 잠금을 푼 뒤 스캔 없이 20건 채워짐(윈도우는 실제 파일 잠금)
  ⓔ 저장 중 프로세스 종료: 엑셀 파일 온전 · 다시 켠 뒤 일지와 같아짐
"""

from __future__ import annotations

import glob
import json
import os
import shutil
import subprocess
import sys
import tarfile
import io
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PROBE = os.path.join(HERE, "speed_probe.py")
THRESHOLD = 0.2
POINTS = (1, 1000, 5000, 10000)
CONTROL_TAG = "v2.0.0"
CONTROL_COMMIT_PREFIX = "d3bbb9f"  # 지시서 §대상

sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)
from speed_probe import jname, read_sheet, xname, DAY, journal_r, prefill  # noqa: E402
from library_emergency import core as fixed_core  # noqa: E402  미리 채우기(시험 준비)에만 쓴다


def control_tree(work):
    """v2.0.0 태그의 library_emergency 를 git 에서 그대로 꺼낸다(손대지 않음)."""
    commit = subprocess.run(["git", "-C", ROOT, "rev-parse", f"{CONTROL_TAG}^{{commit}}"], capture_output=True,
                            text=True, check=True).stdout.strip()
    if not commit.startswith(CONTROL_COMMIT_PREFIX):
        raise RuntimeError(f"대조군 커밋이 다름: {commit}")
    data = subprocess.run(["git", "-C", ROOT, "archive", "--format=tar", CONTROL_TAG, "library_emergency"],
                          capture_output=True, check=True).stdout
    dest = os.path.join(work, "control_v2_0_0")
    with tarfile.open(fileobj=io.BytesIO(data)) as tf:
        tf.extractall(dest, filter="data")
    return dest, commit


def probe(pkg, mode, n=0, timeout=3600, pre=0, extra=()):
    folder = tempfile.mkdtemp(prefix=f"le_speed_{mode}_", dir=os.getcwd())
    try:
        prefill(fixed_core, folder, pre)
        cmd = [sys.executable, "-I", "-X", "utf8", PROBE, "--pkg", pkg, "--folder", folder, "--mode", mode,
               "--n", str(n), *extra]
        t0 = time.monotonic()
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=timeout)
        secs = time.monotonic() - t0
        line = [ln for ln in p.stdout.splitlines() if ln.startswith("RESULT ")]
        if p.returncode != 0 or not line:
            raise RuntimeError(f"탐침 실패 rc={p.returncode} {p.stderr[-800:]}")
        out = json.loads(line[-1][7:])
        out["probe_secs"] = round(secs, 1)
        return out
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def kill_during_save(pkg, n):
    """ⓔ: n-1 줄이 있는 날에 한 건을 찍고, 일꾼의 임시 저장 파일이 보이는 순간 프로세스를 죽인다."""
    folder = tempfile.mkdtemp(prefix="le_speed_kill_", dir=os.getcwd())
    try:
        prefill(fixed_core, folder, n - 1)
        cmd = [sys.executable, "-I", "-X", "utf8", PROBE, "--pkg", pkg, "--folder", folder, "--mode", "killme",
               "--n", str(n)]
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8")
        seen = False
        t0 = time.monotonic()
        try:
            # 미리 채우기(그 저장도 임시 파일을 쓴다)가 끝나고 스캔했다는 신호를 받은 뒤부터 본다
            first = p.stdout.readline()
            if '"scanned": true' not in first:
                raise RuntimeError(f"탐침 시작 실패: {first!r}")
            t0 = time.monotonic()
            while time.monotonic() - t0 < 300:
                if glob.glob(os.path.join(folder, ".저장중_*")):
                    seen = True
                    break
                if p.poll() is not None:
                    break
                time.sleep(0.002)
        finally:
            p.kill()
            p.wait(30)
        killed_at = round(time.monotonic() - t0, 2)
        time.sleep(0.5)
        xp = os.path.join(folder, xname(DAY))
        try:
            rows_after_kill = len(read_sheet(xp)[0])
            intact = True
        except Exception as exc:  # 깨졌으면 붉음
            rows_after_kill, intact = f"{exc.__class__.__name__}", False
        jn = journal_r(os.path.join(folder, jname(DAY)))
        temp_left = len(glob.glob(os.path.join(folder, ".저장중_*")))
        cmd2 = [sys.executable, "-I", "-X", "utf8", PROBE, "--pkg", pkg, "--folder", folder, "--mode", "restart"]
        r = subprocess.run(cmd2, capture_output=True, text=True, encoding="utf-8", timeout=900)
        line = [ln for ln in r.stdout.splitlines() if ln.startswith("RESULT ")]
        restart = json.loads(line[-1][7:]) if line else {"error": r.stderr[-800:]}
        return {"tmp_seen_at_kill": seen, "killed_after_secs": killed_at, "excel_intact_after_kill": intact,
                "excel_rows_after_kill": rows_after_kill, "journal_rows": jn, "temp_files_after_kill": temp_left,
                "restart": restart}
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def fmt(x):
    return f"{x:.3f}초"


def main():
    work = tempfile.mkdtemp(prefix="le_speed_ctl_", dir=os.getcwd())
    results = []
    fixed_pkg = ROOT
    try:
        ctl_pkg, ctl_commit = control_tree(work)
        print(f"python {sys.version.split()[0]} · {sys.platform} · 대조군 {CONTROL_TAG} = {ctl_commit}")
        print(f"기준: Enter → 다음 입력까지 {THRESHOLD}초 이하\n")

        # ⓑ 먼저(ⓐ 의 연속 실행 값도 여기서 나온다)
        t0 = time.monotonic()
        cont = probe(fixed_pkg, "continuous", 10000)
        b_secs = time.monotonic() - t0

        # ⓐ
        t0 = time.monotonic()
        table = []
        for n in POINTS:
            # 같은 상태(그날 n-1 줄)에서 n번째: 대조군은 그 한 건, 고친 판은 n-1 번째를 찍어 일꾼을 돌려 놓고 5건
            old = probe(ctl_pkg, "point", n, pre=n - 1, extra=("--samples", "1"))
            new = probe(fixed_pkg, "point", n, pre=max(0, n - 2), extra=("--samples", "5", "--warm"))
            table.append((n, old, new))
        a_secs = time.monotonic() - t0 + b_secs
        print("[ⓐ 속도] 그날 n번째 Enter → 다음 입력까지 (대조군은 n번째 한 건 · 고친 판은 n번째부터 5건 중 가장 느린 값)")
        print("   n번째 | v2.0.0(대조군) | 고친 판(같은 조건) | 고친 판(10,000건 연속 중) | 고친 판 일꾼 바쁨")
        old_over = False
        new_ok = True
        for n, old, new in table:
            o = max(s["secs"] for s in old["samples"])
            w = max(s["secs"] for s in new["samples"])
            c = cont["points"][str(n)]
            busy = sum(s["worker_busy"] for s in new["samples"])
            old_over |= o > THRESHOLD
            ready = all(s["ready"] for s in new["samples"]) and c["ready"]
            good = w <= THRESHOLD and c["secs"] <= THRESHOLD and ready
            new_ok &= good
            print(f"   {n:>6} | {fmt(o):>9}{' 붉음' if o > THRESHOLD else ' 초록'} | {fmt(w):>9}"
                  f"{' 초록' if w <= THRESHOLD else ' 붉음'} | {fmt(c['secs']):>9}{' 초록' if c['secs'] <= THRESHOLD else ' 붉음'}"
                  f" | {busy}/5")
        a_pass = old_over and new_ok
        print(f"   10,000건 연속: 가운데 {fmt(cont['median'])} · 상위1% {fmt(cont['p99'])} · 최대 {fmt(cont['max'])}"
              f" · 기준 넘은 스캔 {cont['over_threshold']}건")
        print(f"  {'PASS' if a_pass else 'FAIL'}  ⓐ 속도 — 대조군 {'붉음' if old_over else '초록(잣대 불인정)'} → "
              f"고친 판 {'초록' if new_ok else '붉음'} · {a_secs:.0f}초")
        results.append({"name": "ⓐ 속도", "pass": a_pass, "secs": round(a_secs, 1),
                        "points": [{"n": n, "control": old, "fixed": new} for n, old, new in table],
                        "continuous_points": cont["points"]})

        # ⓑ
        want = cont["want"]
        b_pass = (cont["excel_rows"] == cont["journal_rows"] == cont["list_rows"] == 10000 and
                  cont["totals"] == want and cont["drained"] and cont["pink_rows"] == 0)
        print(f"\n[ⓑ 10,000건 대조]\n  {'PASS' if b_pass else 'FAIL'}  엑셀 {cont['excel_rows']}줄 · 일지 "
              f"{cont['journal_rows']}줄 · 화면 {cont['list_rows']}건 · 합계 {cont['totals']} (기대 {want}) · "
              f"분홍 줄 {cont['pink_rows']} · 찍는 데 {cont['scan_secs']}초 · 마지막 스캔 뒤 엑셀이 따라잡기까지 "
              f"{cont['drain_secs']}초 · {b_secs:.0f}초")
        results.append({"name": "ⓑ 10,000건 대조", "pass": b_pass, "secs": round(b_secs, 1), "detail": cont})

        # ⓒ
        t0 = time.monotonic()
        ext = probe(fixed_pkg, "external")
        c_secs = time.monotonic() - t0
        want_books = ["EB1", "EB3", "EB4", "EB5", "EB6"]
        c_pass = (ext["deleted"] == "EB2" and ext["after_delete_books"] == want_books and
                  ext["after_delete_totals"].get("전체 건수") == 5 and ext["memo"] == ["메모: 3번 책 표지 파손"] and
                  ext["after_memo_books"] == want_books and "양식" in ext["status_after_memo"] and
                  "저장 완료" not in ext["status_after_memo"] and ext["journal_rows"] == 7)
        print(f"\n[ⓒ 바깥 수정 보존]\n  {'PASS' if c_pass else 'FAIL'}  지운 줄 {ext['deleted']} → 다음 기록 뒤 엑셀 "
              f"{ext['after_delete_books']} · 합계 {ext['after_delete_totals'].get('전체 건수')} · 메모 뒤 기록 → 메모 "
              f"{ext['memo']} · 엑셀 {ext['after_memo_books']} · 상태 「{ext['status_after_memo'][:40]}…」 · "
              f"일지 {ext['journal_rows']}줄 · {c_secs:.0f}초")
        results.append({"name": "ⓒ 바깥 수정 보존", "pass": c_pass, "secs": round(c_secs, 1), "detail": ext})

        # ⓓ
        t0 = time.monotonic()
        lk = probe(fixed_pkg, "lock", 1000, pre=1000)
        d_secs = time.monotonic() - t0
        d_pass = (lk["max"] <= THRESHOLD and lk["red_shown"] and lk["filled_without_scan"] and
                  lk["excel_rows"] == lk["want_rows"] and lk["status_after"].startswith("저장 완료") and
                  (not lk["real_lock"] or lk["excel_rows_while_locked"] == lk["prefill"]))
        print(f"\n[ⓓ 잠금 중 연속 스캔 20건] ({'실제 파일 잠금 · 읽기만 공유' if lk['real_lock'] else '저장 실패 흉내'})")
        print(f"  {'PASS' if d_pass else 'FAIL'}  잠긴 동안 최대 {fmt(lk['max'])} · 빨간 문구 {lk['red_shown']} · "
              f"잠긴 동안 엑셀 {lk['excel_rows_while_locked']}줄 · 분홍 {lk['pink_while_locked']}줄 → 잠금 푼 뒤 스캔 없이 "
              f"{lk['fill_secs']}초 만에 엑셀 {lk['excel_rows']}줄(기대 {lk['want_rows']}) · 상태 「{lk['status_after'][:20]}」"
              f" · {d_secs:.0f}초")
        results.append({"name": "ⓓ 잠금 중 연속 스캔", "pass": d_pass, "secs": round(d_secs, 1), "detail": lk})

        # ⓔ
        t0 = time.monotonic()
        k = kill_during_save(fixed_pkg, 5000)
        e_secs = time.monotonic() - t0
        rs = k["restart"]
        e_pass = (k["tmp_seen_at_kill"] and k["excel_intact_after_kill"] and k["journal_rows"] == 5000 and
                  rs.get("excel_rows") == rs.get("journal_rows") == 5000 and
                  rs.get("totals", {}).get("전체 건수") == 5000)
        print(f"\n[ⓔ 저장 중 프로세스 종료]\n  {'PASS' if e_pass else 'FAIL'}  임시 저장 파일이 보이는 순간 종료="
              f"{k['tmp_seen_at_kill']} · 종료 직후 엑셀 온전={k['excel_intact_after_kill']}({k['excel_rows_after_kill']}줄)"
              f" · 일지 {k['journal_rows']}줄 · 남은 임시 파일 {k['temp_files_after_kill']} → 다시 켠 뒤 엑셀 "
              f"{rs.get('excel_rows')}줄 = 일지 {rs.get('journal_rows')}줄 · 합계 {rs.get('totals', {}).get('전체 건수')}"
              f" · {e_secs:.0f}초")
        results.append({"name": "ⓔ 저장 중 프로세스 종료", "pass": e_pass, "secs": round(e_secs, 1), "detail": k})
    finally:
        shutil.rmtree(work, ignore_errors=True)

    n_pass = sum(r["pass"] for r in results)
    print(f"\n속도 게이트 합계: {n_pass}/5 통과")
    with open(os.path.join(os.getcwd(), "speed_result.json"), "w", encoding="utf-8") as fh:
        json.dump({"python": sys.version, "platform": sys.platform, "results": results}, fh, ensure_ascii=False,
                  indent=1)
    return 0 if n_pass == 5 and len(results) == 5 else 1


if __name__ == "__main__":
    sys.exit(main())
