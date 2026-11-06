"""속도 게이트의 탐침 — 판 하나(--pkg 폴더의 library_emergency)를 화면을 띄운 채 돌리고 JSON 한 줄을 낸다.

    python -I -X utf8 tests/speed_probe.py --pkg <판 폴더> --folder <기록 폴더> --mode <point|continuous|external|lock|killme> [--n N]

판마다 따로 프로세스로 돈다(v2.0.0 과 고친 판이 같은 이름의 꾸러미라 한 프로세스에 둘을 못 올린다).
시간은 「책 칸에서 Enter → 화면이 다음 입력을 받을 수 있을 때(그 Enter 처리와 다시 그리기가 끝남)」로 잰다.
기대값(줄 수 · 합계)은 이 탐침이 직접 세고, 엑셀은 openpyxl 로 따로 읽는다(판의 함수를 쓰지 않는다).
"""

from __future__ import annotations

import argparse
import datetime as dt
import glob
import json
import os
import sys
import time

THRESHOLD = 0.2


class FakeClock:
    """한 날 안에서만 흐르는 시계(빌드 기계가 자정 무렵에 돌아도 날짜가 바뀌지 않게)."""

    def __init__(self):
        self.now = dt.datetime(2026, 11, 6, 9, 0, 0)

    def __call__(self):
        self.now += dt.timedelta(milliseconds=1)
        return self.now


def xname(day):
    return f"도서기록_{day:%Y%m%d}.xlsx"


def jname(day):
    return f"도서기록_{day:%Y%m%d}_일지.txt"


T_SHEET = "대출반납기록"
T_LABELS = ("전체 건수", "대출 건수", "반납 건수")
DAY = dt.date(2026, 11, 6)


def read_sheet(path):
    import openpyxl

    wb = openpyxl.load_workbook(path)
    ws = wb[T_SHEET]
    rows = []
    r = 5
    while r <= ws.max_row:
        vals = [ws.cell(r, c).value for c in range(1, 5)]
        if all(v in (None, "") for v in vals) or any(isinstance(v, str) and v.strip() in T_LABELS for v in vals):
            break
        rows.append(tuple("" if v is None else str(v) for v in vals))
        r += 1
    totals = {}
    memo = []
    for rr in range(5, ws.max_row + 1):
        for cc in range(1, ws.max_column + 1):
            v = ws.cell(rr, cc).value
            if isinstance(v, str) and v.strip() in T_LABELS:
                totals[v.strip()] = ws.cell(rr, cc + 1).value
            elif isinstance(v, str) and v.startswith("메모"):
                memo.append(v)
    return rows, totals, memo


def journal_r(path):
    if not os.path.exists(path):
        return 0
    with open(path, encoding="utf-8") as fh:
        return sum(1 for ln in fh.read().split("\n") if ln.startswith("R\t"))


def kind_of(i):
    return "대출" if i % 3 else "반납"


def prefill(core, folder, n):
    """그날 기록 n줄이 이미 일지와 엑셀에 있는 상태를 만든다. 부모(게이트)가 고친 판의 양식 함수로 부른다 —
    두 판의 엑셀 양식은 같고, v2.0.0 의 느린 자리로 미리 채우면 1만 줄에 몇 분이 걸린다. 두 판 모두 같은 파일에서 시작한다."""
    if n <= 0:
        return
    rows = [core.Row(f"2026-11-06 08:{(i // 60) % 60:02d}:{i % 60:02d}", kind_of(i), f"P{i % 41}", f"PB{i:06d}")
            for i in range(n)]
    with open(os.path.join(folder, jname(DAY)), "w", encoding="utf-8", newline="") as fh:
        fh.write(core.JOURNAL_HEADER)
        for r in rows:
            fh.write("\t".join(("R",) + r.key()) + "\n")
        fh.write(f"S\t{n}\t2026-11-06 08:59:59\n")
    wb = core.new_workbook(DAY, dt.datetime(2026, 11, 6, 8, 0, 0))
    core.write_rows_and_totals(wb[core.SHEET_TITLE], core.HEADER_ROW, rows)
    core.save_atomic(wb, os.path.join(folder, xname(DAY)))


class Driver:
    def __init__(self, pkg, folder):
        sys.path.insert(0, pkg)
        import tkinter as tk

        from library_emergency import app as appmod
        from library_emergency import core

        self.tk, self.core, self.appmod = tk, core, appmod
        self.folder = folder
        self.clock = FakeClock()
        self.fixed = hasattr(core, "ExcelWorker")
        # 탐침에서 모달 창이 뜨면 화면이 멈춘 채 기다린다 — 창은 띄우지 않고 기록만 남긴다(시험 쪽 대체)
        self.dialogs = []
        for name in ("showwarning", "showerror", "showinfo"):
            setattr(appmod.messagebox, name, lambda *a, _n=name, **k: self.dialogs.append((_n, a[:2])))

    def start(self):
        if hasattr(self.appmod, "tune_threads"):
            self.appmod.tune_threads()
        self.store = self.core.RecordStore(self.folder, clock=self.clock)
        self.root = self.tk.Tk()
        self.app = self.appmod.EmergencyLibraryApp(self.root, self.store)
        self.root.geometry("980x720+20+20")
        # 창이 탐침 밖에서 닫히면(사람이 닫음 등) 누가 닫았는지 남긴다
        real_destroy = self.root.destroy

        def destroy_logged():
            import traceback
            sys.stderr.write(f"[탐침] 창 닫힘 · 스캔 {self.n_scans}건 뒤\n" + "".join(traceback.format_stack(limit=6)))
            real_destroy()

        self.root.destroy = destroy_logged
        self.root.update()
        self.n_scans = 0

    def worker_busy(self):
        w = getattr(self.app, "worker", None)
        return bool(w and not w.idle())

    def scan(self, member, book):
        """회원번호 Enter → 책 바코드 → (시간 재기) Enter → 화면이 다음 입력을 받을 수 있을 때까지."""
        app, root = self.app, self.root
        # 키 이벤트는 초점이 있는 칸으로 간다 — 스캐너가 찍는 칸에 초점을 먼저 둔다
        app.member_entry.focus_force()
        root.update_idletasks()
        app.member_var.set(member)
        app.member_entry.event_generate("<Return>")
        app.book_entry.focus_force()
        root.update_idletasks()
        app.book_var.set(book)
        busy = self.worker_busy()
        t0 = time.perf_counter()
        app.book_entry.event_generate("<Return>")
        root.update()
        secs = time.perf_counter() - t0
        self.n_scans += 1
        ready = app.book_var.get() == "" and root.focus_get() in (app.book_entry, app.member_entry)
        return secs, busy, ready and not self.dialogs

    def settle(self, until, timeout):
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            self.root.update()
            if until():
                return True
            time.sleep(0.02)
        self.root.update()
        return until()

    def drained(self):
        w = getattr(self.app, "worker", None)
        if w is None:
            return True
        return w.idle() and self.app.results.empty() and not self.pinks()

    def pinks(self):
        t = self.app.tree
        return [i for i in t.get_children() if "pending" in t.item(i, "tags")]

    def close(self):
        if hasattr(self.app, "close"):
            self.app.close()
        else:
            self.root.destroy()


def mode_point(args):
    """그날 n번째 스캔의 시간. --warm 이면 n-1 번째를 먼저 찍고(재지 않음, 일꾼이 엑셀 정리를 시작함)
    n번째부터 --samples 건을 쉬지 않고 이어 잰다. 기록 폴더는 부모가 미리 채워 둔다(판의 함수를 쓰지 않게)."""
    d = Driver(args.pkg, args.folder)
    n = args.n
    t0 = time.monotonic()
    d.start()
    startup = time.monotonic() - t0
    if args.warm and n >= 2:
        d.scan("S1", f"SB{n - 1:06d}")
    samples = []
    for k in range(args.samples):
        secs, busy, ready = d.scan("S1", f"SB{n + k:06d}")
        samples.append({"nth": n + k, "secs": round(secs, 4), "worker_busy": busy, "ready": ready})
    drain_secs = None
    if d.fixed:
        t1 = time.monotonic()
        d.settle(d.drained, 600)
        drain_secs = round(time.monotonic() - t1, 2)
    rows, totals, _ = read_sheet(os.path.join(args.folder, xname(DAY)))
    jn = journal_r(os.path.join(args.folder, jname(DAY)))
    out = {"mode": "point", "n": n, "fixed": d.fixed, "samples": samples, "startup_secs": round(startup, 2),
           "drain_secs": drain_secs, "excel_rows": len(rows), "journal_rows": jn, "list_rows": d.app.list_count(),
           "want_rows": n + args.samples - 1, "totals": totals}
    d.close()
    return out


def mode_continuous(args):
    """빈 날에서 n건을 화면으로 쉬지 않고 찍는다. 매 스캔 시간 + 끝난 뒤 네 곳 대조."""
    d = Driver(args.pkg, args.folder)
    d.start()
    n = args.n
    lat = []
    loans = 0
    t0 = time.monotonic()
    marks = {1, 1000, 5000, 10000, n}
    points = {}
    for i in range(1, n + 1):
        k = kind_of(i)
        if (k == "대출") != (d.app.mode == "대출"):
            d.root.event_generate("<F1>" if k == "대출" else "<F2>")
            d.root.update()
        loans += k == "대출"
        secs, busy, ready = d.scan(f"M{i % 37}", f"B{i:06d}")
        lat.append(secs)
        if i in marks:
            points[str(i)] = {"secs": round(secs, 4), "worker_busy": busy, "ready": ready}
    scan_secs = time.monotonic() - t0
    t1 = time.monotonic()
    ok_drain = d.settle(d.drained, 900)
    drain = time.monotonic() - t1
    rows, totals, _ = read_sheet(os.path.join(args.folder, xname(DAY)))
    jn = journal_r(os.path.join(args.folder, jname(DAY)))
    s = sorted(lat)
    out = {"mode": "continuous", "n": n, "points": points, "max": round(s[-1], 4),
           "p99": round(s[int(len(s) * 0.99) - 1], 4), "median": round(s[len(s) // 2], 4),
           "over_threshold": sum(1 for x in lat if x > THRESHOLD), "scan_secs": round(scan_secs, 1),
           "drain_secs": round(drain, 1), "drained": ok_drain, "excel_rows": len(rows), "journal_rows": jn,
           "list_rows": d.app.list_count(), "pink_rows": len(d.pinks()), "totals": totals,
           "want": {T_LABELS[0]: n, T_LABELS[1]: loans, T_LABELS[2]: n - loans}}
    d.close()
    return out


def mode_external(args):
    """ⓒ 기록 중간에 사람이 엑셀에서 줄을 지움 → 다음 기록 · 합계 아래 메모 → 다음 기록."""
    import openpyxl

    d = Driver(args.pkg, args.folder)
    d.start()
    p = os.path.join(args.folder, xname(DAY))
    for i in range(1, 6):
        d.scan("E1", f"EB{i}")
    d.settle(d.drained, 120)
    wb = openpyxl.load_workbook(p)
    ws = wb[T_SHEET]
    deleted = ws.cell(6, 4).value  # 둘째 기록 줄(사람이 겹침 스캔이라 보고 지움)
    ws.delete_rows(6, 1)
    wb.save(p)
    d.scan("E1", "EB6")
    d.settle(d.drained, 120)
    rows1, totals1, _ = read_sheet(p)
    # 합계 아래 메모
    wb = openpyxl.load_workbook(p)
    ws = wb[T_SHEET]
    memo_row = ws.max_row + 3
    ws.cell(memo_row, 1, "메모: 3번 책 표지 파손")
    wb.save(p)
    d.scan("E1", "EB7")
    d.settle(lambda: d.app.results.empty() and "양식" in d.app.status.cget("text"), 60)
    status = d.app.status.cget("text")
    rows2, totals2, memo2 = read_sheet(p)
    jn = journal_r(os.path.join(args.folder, jname(DAY)))
    out = {"mode": "external", "deleted": deleted, "after_delete_books": [r[3] for r in rows1],
           "after_delete_totals": totals1, "after_memo_books": [r[3] for r in rows2], "memo": memo2,
           "status_after_memo": status, "journal_rows": jn, "pink_rows": len(d.pinks())}
    d.close()
    return out


def _win_lock(path, share):
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                                wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    h = k32.CreateFileW(path, 0x80000000 | 0x40000000, share, None, 3, 0x80, None)
    if h in (None, wintypes.HANDLE(-1).value):
        raise OSError(ctypes.get_last_error(), "잠금 실패")
    return lambda: k32.CloseHandle(h)


def _fault_lock():
    """윈도우가 아닌 기계: 엑셀 저장을 PermissionError 로 실패시키는 흉내."""
    from openpyxl.workbook.workbook import Workbook

    orig = Workbook.save

    def bad(self, filename):
        raise PermissionError(13, "열려 있음(흉내)", filename)

    Workbook.save = bad

    def release():
        Workbook.save = orig

    return release


def mode_lock(args):
    """ⓓ 엑셀이 열려 잠긴 채 20건 → 화면 멈춤 없음 · 잠금을 푼 뒤 스캔 없이 20건이 엑셀에 채워짐."""
    d = Driver(args.pkg, args.folder)
    d.start()
    p = os.path.join(args.folder, xname(DAY))
    real = sys.platform.startswith("win")
    release = _win_lock(p, 1) if real else _fault_lock()  # 엑셀식 잠금: 읽기만 공유
    lat = []
    try:
        for i in range(1, 21):
            secs, busy, ready = d.scan("K1", f"KB{i:02d}")
            lat.append(round(secs, 4))
        red = d.settle(lambda: "일지에만 기록됨" in d.app.status.cget("text"), 30)
        status_locked = d.app.status.cget("text")
        rows_locked = len(read_sheet(p)[0]) if real else None
        pinks_locked = len(d.pinks())
    finally:
        release()
    t1 = time.monotonic()
    scans_before = d.n_scans
    filled = d.settle(lambda: d.drained() and d.app.status.cget("text").startswith("저장 완료"), 120)
    fill_secs = time.monotonic() - t1
    rows, totals, _ = read_sheet(p)
    out = {"mode": "lock", "real_lock": real, "prefill": args.n, "latencies": lat, "max": max(lat),
           "red_shown": red, "status_locked": status_locked, "excel_rows_while_locked": rows_locked,
           "pink_while_locked": pinks_locked, "filled_without_scan": filled and d.n_scans == scans_before,
           "fill_secs": round(fill_secs, 2), "excel_rows": len(rows), "want_rows": args.n + 20,
           "status_after": d.app.status.cget("text"), "totals": totals}
    d.close()
    return out


def mode_killme(args):
    """ⓔ 엑셀 정리 중에 죽임을 당할 프로세스: (부모가 n-1 줄을 채워 둔 날에) 한 건 찍어 일꾼이 저장하게 한 뒤 기다린다."""
    d = Driver(args.pkg, args.folder)
    d.start()
    d.scan("Z1", "ZB-LAST")
    print(json.dumps({"mode": "killme", "scanned": True}), flush=True)
    d.settle(lambda: False, 600)  # 부모가 죽일 때까지 화면을 돌린다(일꾼은 계속 다시 저장)
    return {}


def mode_restart(args):
    """ⓔ 다시 켬: 실제 시작 순서(main)대로 최근 일지에서 엑셀을 채운다(화면 없이)."""
    sys.path.insert(0, args.pkg)
    from library_emergency import core

    store = core.RecordStore(args.folder, clock=FakeClock())
    rec = store.recover_recent()
    ok, err = store.ensure_today_file()
    rows, totals, _ = read_sheet(os.path.join(args.folder, xname(DAY)))
    return {"mode": "restart", "recovered": [(str(a), b, c) for a, b, c in rec], "ok": ok, "err": err,
            "excel_rows": len(rows), "journal_rows": journal_r(os.path.join(args.folder, jname(DAY))),
            "totals": totals, "temp_files": len(glob.glob(os.path.join(args.folder, ".저장중_*")))}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pkg", required=True)
    ap.add_argument("--folder", required=True)
    ap.add_argument("--mode", required=True)
    ap.add_argument("--n", type=int, default=0)
    ap.add_argument("--samples", type=int, default=5)
    ap.add_argument("--warm", action="store_true")
    args = ap.parse_args()
    fn = {"point": mode_point, "continuous": mode_continuous, "external": mode_external, "lock": mode_lock,
          "killme": mode_killme, "restart": mode_restart}[args.mode]
    out = fn(args)
    print("RESULT " + json.dumps(out, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
