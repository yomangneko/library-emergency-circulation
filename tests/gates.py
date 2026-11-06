"""게이트 — 화면 없이 기록 처리를 잰다.

    python -I tests/gates.py            # 전 게이트, 결과표 + gates_result.json
    python -I tests/gates.py --windows  # 윈도우 빌드 기계: 실제 파일 잠금 게이트까지 (윈도우가 아니면 실패)

규율:
  - 기대값은 프로그램의 함수 · 상수를 쓰지 않고 시험이 직접 계산한다(엑셀은 시험이 따로 읽는다).
  - 버그 부류마다 「고치기 전 꼴을 흉내 낸 사본」(tests/legacy_mimic.py)에서 붉고, 고친 판에서 초록이어야 통과.
"""

from __future__ import annotations

import builtins
import contextlib
import datetime as dt
import importlib
import json
import os
import re
import shutil
import socket
import stat
import sys
import tempfile
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

import openpyxl  # noqa: E402

from library_emergency import core  # noqa: E402
from library_emergency.messages import status_for  # noqa: E402
import legacy_mimic  # noqa: E402

# 시험이 따로 적어 두는 기대 문면(프로그램 상수를 가져다 쓰지 않는다)
T_SHEET = "대출반납기록"
T_LABELS = ("전체 건수", "대출 건수", "반납 건수")
T_LOAN, T_RET = "대출", "반납"


# ------------------------------------------------------------------ 도구


class FakeClock:
    def __init__(self, start: dt.datetime):
        self.now = start

    def __call__(self):
        return self.now

    def tick(self, seconds=1):
        self.now += dt.timedelta(seconds=seconds)


def xname(day):
    return f"도서기록_{day:%Y%m%d}.xlsx"


def jname(day):
    return f"도서기록_{day:%Y%m%d}_일지.txt"


def read_sheet(path):
    """시험 쪽 독립 판독기: (기록 줄들, 합계 dict, 머리글, 제목)."""
    wb = openpyxl.load_workbook(path)
    ws = wb[T_SHEET]
    rows = []
    r = 5
    while r <= ws.max_row:
        vals = [ws.cell(r, c).value for c in range(1, 5)]
        if all(v in (None, "") for v in vals):
            break
        if any(isinstance(v, str) and v.strip() in T_LABELS for v in vals):
            break
        rows.append(tuple("" if v is None else str(v) for v in vals))
        r += 1
    totals = {}
    for rr in range(5, ws.max_row + 1):
        for cc in range(1, ws.max_column + 1):
            v = ws.cell(rr, cc).value
            if isinstance(v, str) and v.strip() in T_LABELS:
                totals[v.strip()] = eval_cell(ws, ws.cell(rr, cc + 1).value)
    header = tuple(ws.cell(4, c).value for c in range(1, 5))
    return rows, totals, header, ws["A1"].value


def eval_cell(ws, v):
    """합계 칸이 수식이면 시험이 직접 셈한다(COUNTA · COUNTIF 만)."""
    if not (isinstance(v, str) and v.startswith("=")):
        return v
    m = re.fullmatch(r"=COUNTA\(([A-Z])(\d+):([A-Z])(\d+)\)", v)
    if m:
        col, a, b = m.group(1), int(m.group(2)), int(m.group(4))
        return sum(1 for r in range(a, b + 1) if ws[f"{col}{r}"].value not in (None, ""))
    m = re.fullmatch(r'=COUNTIF\(([A-Z])(\d+):([A-Z])(\d+),"(.+)"\)', v)
    if m:
        col, a, b, word = m.group(1), int(m.group(2)), int(m.group(4)), m.group(5)
        return sum(1 for r in range(a, b + 1) if ws[f"{col}{r}"].value == word)
    return v


def journal_r_lines(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as fh:
        return [ln for ln in fh.read().split("\n") if ln.startswith("R\t")]


@contextlib.contextmanager
def excel_save_fault(kind="permission"):
    """엑셀 저장을 실패시키는 흉내(모든 운영체제). 잠금 자체는 --windows 게이트가 실제로 잰다."""
    from openpyxl.workbook.workbook import Workbook

    orig = Workbook.save

    def bad(self, filename):
        if kind == "partial":
            with open(filename, "wb") as fh:
                fh.write(b"PK\x03\x04 partial write")
            raise OSError("disk full (흉내)")
        raise PermissionError(13, "열려 있음(흉내)", filename)

    Workbook.save = bad
    try:
        yield
    finally:
        Workbook.save = orig


def new_store(folder, clock):
    return core.RecordStore(folder, clock=clock)


def old_store(folder, clock):
    return legacy_mimic.LegacyStore(folder, clock)


START = dt.datetime(2026, 11, 6, 9, 0, 0)


# ------------------------------------------------------------------ 버그 부류 검사 (사본 붉음 → 고친 판 초록)
# 각 검사는 (통과 여부, 근거 한 줄)을 돌려준다. 예외는 붉음.


def cls_a_loss(make, d):
    """㉮ 기록 유실: 엑셀 저장 실패 중 스캔이 사라지지 않고, 실패를 「저장 완료」로 알리지 않는다."""
    clk = FakeClock(START)
    s = make(d, clk)
    s.record(T_LOAN, "M1", "B1")
    clk.tick(30)
    with excel_save_fault():
        rb = s.record(T_LOAN, "M2", "B2")
    claimed_done = False
    if hasattr(rb, "row") and rb.recorded and rb.row is not None:
        claimed_done = status_for(rb)[0].startswith("저장 완료")
    elif rb.recorded and getattr(rb, "excel_saved", False):
        claimed_done = True
    clk.tick(30)
    s.record(T_RET, "M3", "B3")
    rows, _, _, _ = read_sheet(os.path.join(d, xname(START.date())))
    books = [r[3] for r in rows]
    ok = books == ["B1", "B2", "B3"] and not claimed_done
    return ok, f"엑셀 도서번호 {books} · 실패를 저장 완료로 알림={claimed_done}"


def cls_b_rollover(make, d):
    """㉯ 날짜 넘김: 켜 둔 채 자정을 넘기면 그날 파일로 바뀐다."""
    clk = FakeClock(dt.datetime(2026, 11, 6, 23, 59, 58))
    s = make(d, clk)
    s.record(T_LOAN, "M1", "B1")
    clk.now = dt.datetime(2026, 11, 7, 0, 0, 3)
    s.record(T_LOAN, "M2", "B2")
    p1 = os.path.join(d, xname(dt.date(2026, 11, 6)))
    p2 = os.path.join(d, xname(dt.date(2026, 11, 7)))
    b1 = [r[3] for r in read_sheet(p1)[0]] if os.path.exists(p1) else None
    b2 = [r[3] for r in read_sheet(p2)[0]] if os.path.exists(p2) else None
    return (b1 == ["B1"] and b2 == ["B2"]), f"11-06 파일 {b1} · 11-07 파일 {b2}"


def cls_c_double(make, d):
    """㉰ 겹침 스캔: 기록은 하되 「방금 같은 기록이 있음」을 알린다(조용히 버리지 않음)."""
    clk = FakeClock(START)
    s = make(d, clk)
    r1 = s.record(T_LOAN, "M1", "B1")
    clk.tick(1)
    r2 = s.record(T_LOAN, "M1", "B1")
    clk.tick(1)
    r3 = s.record(T_LOAN, "M1", "B2")  # 다른 책은 알리지 않는다
    clk.tick(60)
    r4 = s.record(T_LOAN, "M1", "B2")  # 오래 지나면 알리지 않는다
    rows = read_sheet(os.path.join(d, xname(START.date())))[0]
    flags = [getattr(r, "duplicate_seconds", None) for r in (r1, r2, r3, r4)]
    text = status_for(r2)[0] if hasattr(r2, "row") and r2.row is not None else ""
    ok = len(rows) == 4 and flags[0] is None and flags[1] is not None and flags[2] is None and flags[3] is None \
        and "방금" in text
    return ok, f"엑셀 {len(rows)}줄 · 알림 {flags}"


def cls_d_folder(check, d):
    """㉱ 저장 폴더: 못 만들거나 못 쓰는 자리 · 임시 폴더를 시작할 때 알아챈다."""
    out = []
    # ⓐ 같은 이름의 파일이 막고 있다
    blocker = os.path.join(d, "exports")
    with open(blocker, "w") as fh:
        fh.write("x")
    try:
        c = check(blocker)
        out.append(("파일이 막음", not c.ok))
    except Exception:
        out.append(("파일이 막음", False))
    # ⓑ 읽기 전용 폴더(윈도우는 폴더 읽기 전용 속성이 쓰기를 막지 않으므로 POSIX 에서만)
    if os.name == "posix" and os.geteuid() != 0:
        ro = os.path.join(d, "ro")
        os.makedirs(ro)
        os.chmod(ro, stat.S_IRUSR | stat.S_IXUSR)
        try:
            c = check(os.path.join(ro, "exports"))
            out.append(("읽기 전용", not c.ok))
        except Exception:
            out.append(("읽기 전용", False))
        finally:
            os.chmod(ro, stat.S_IRWXU)
    # ⓒ 임시 폴더 = 압축 파일 안에서 바로 실행한 자리
    tmpd = tempfile.mkdtemp(prefix="Temp1_LibraryEmergency.zip_")
    try:
        c = check(os.path.join(tmpd, "exports"))
        out.append(("임시 폴더", bool(c.ok and c.warning)))
    except Exception:
        out.append(("임시 폴더", False))
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)
    # ⓓ 정상 자리는 경고 없이 통과(긍정 대조) — 임시 폴더 밖이어야 하므로 실행 자리(작업 폴더) 아래
    gooddir = tempfile.mkdtemp(prefix="le_good_", dir=os.getcwd())
    try:
        c = check(os.path.join(gooddir, "exports"))
        out.append(("정상 자리", bool(c.ok and not c.warning)))
    except Exception:
        out.append(("정상 자리", False))
    finally:
        shutil.rmtree(gooddir, ignore_errors=True)
    return all(v for _, v in out), " · ".join(f"{k}={'맞음' if v else '틀림'}" for k, v in out)


def cls_e_input(make, d):
    """㉲ 입력값: 앞뒤 공백 · 빈 값 · 바코드 기계 꼬리 글자 · 보이지 않는 글자."""
    clk = FakeClock(START)
    s = make(d, clk)
    cases = [
        ((" M1 ", "B1\r\n"), ("M1", "B1")),
        (("​M2", "B2\t"), ("M2", "B2")),
        (("﻿M3", "　B-03 "), ("M3", "B-03")),
        (("M4\r", "B4\n"), ("M4", "B4")),
    ]
    for (m, b), _ in cases:
        s.record(T_LOAN, m, b)
        clk.tick(20)
    blank_member = s.record(T_LOAN, "   ", "B9")
    blank_book = s.record(T_LOAN, "M9", "\r\n")
    zw_only = s.record(T_LOAN, "​", "B9")
    rows = read_sheet(os.path.join(d, xname(START.date())))[0]
    got = [(r[2], r[3]) for r in rows]
    want = [w for _, w in cases]
    misses = [getattr(x, "missing", "") for x in (blank_member, blank_book, zw_only)]
    ok = got == want and misses == ["member", "book", "member"]
    return ok, f"엑셀 {got!r} · 빈 값 처리 {misses}"


def cls_f_totals(make, d):
    """㉳ 합계 줄: 줄을 더할 때마다 전체 · 대출 · 반납 건수가 맞는다."""
    clk = FakeClock(START)
    s = make(d, clk)
    kinds = [T_LOAN, T_LOAN, T_RET, T_LOAN, T_RET, T_RET, T_LOAN]
    bad = []
    for i, k in enumerate(kinds):
        s.record(k, f"M{i}", f"B{i}")
        clk.tick(20)
        rows, totals, _, _ = read_sheet(os.path.join(d, xname(START.date())))
        want = {T_LABELS[0]: len(rows), T_LABELS[1]: sum(r[1] == T_LOAN for r in rows),
                T_LABELS[2]: sum(r[1] == T_RET for r in rows)}
        if len(rows) != i + 1 or totals != want:
            bad.append((i + 1, totals))
    return not bad, ("매 줄 합계 일치" if not bad else f"어긋남(줄 수, 합계) {bad[:3]}")


def cls_g_corrupt(make, d):
    """㉴-1 깨진 엑셀: 덮어쓰지 않고, 스캔은 일지에 남는다."""
    clk = FakeClock(START)
    s = make(d, clk)
    s.record(T_LOAN, "M1", "B1")
    p = os.path.join(d, xname(START.date()))
    with open(p, "wb") as fh:
        fh.write(b"not an xlsx")
    clk.tick(30)
    r = s.record(T_LOAN, "M2", "B2")
    with open(p, "rb") as fh:
        untouched = fh.read() == b"not an xlsx"
    in_journal = any(ln.endswith("\tM2\tB2") for ln in journal_r_lines(os.path.join(d, jname(START.date()))))
    ok = untouched and in_journal and r.recorded and not r.excel_saved
    return ok, f"원 파일 그대로={untouched} · 일지에 남음={in_journal}"


def cls_h_partial(make, d):
    """㉴-2 저장 중 끊김: 반쯤 쓴 파일이 원래 엑셀을 망가뜨리지 않고, 다음 저장 때 채워진다."""
    clk = FakeClock(START)
    s = make(d, clk)
    s.record(T_LOAN, "M1", "B1")
    clk.tick(30)
    with excel_save_fault("partial"):
        s.record(T_LOAN, "M2", "B2")
    p = os.path.join(d, xname(START.date()))
    try:
        mid = [r[3] for r in read_sheet(p)[0]]
    except Exception:
        mid = None
    clk.tick(30)
    s.record(T_LOAN, "M3", "B3")
    try:
        end = [r[3] for r in read_sheet(p)[0]]
    except Exception:
        end = None
    leftovers = [f for f in os.listdir(d) if f.startswith(".저장중_")]
    ok = mid == ["B1"] and end == ["B1", "B2", "B3"] and not leftovers
    return ok, f"끊긴 직후 {mid} · 다음 저장 뒤 {end} · 임시 파일 {len(leftovers)}"


BUG_CLASSES = [
    ("㉮ 기록 유실", cls_a_loss, "store"),
    ("㉯ 날짜 넘김", cls_b_rollover, "store"),
    ("㉰ 겹침 스캔", cls_c_double, "store"),
    ("㉱ 저장 폴더", cls_d_folder, "folder"),
    ("㉲ 입력값", cls_e_input, "store"),
    ("㉳ 합계 줄", cls_f_totals, "store"),
    ("㉴-1 깨진 엑셀 덮어쓰기", cls_g_corrupt, "store"),
    ("㉴-2 저장 중 끊김", cls_h_partial, "store"),
]


# ------------------------------------------------------------------ 고친 판만 재는 게이트


def gate_1000(d):
    """대출 · 반납 1,000건: 엑셀 줄 수 = 일지 줄 수 = 화면 건수 = 합계 줄."""
    clk = FakeClock(START)
    s = core.RecordStore(d, clock=clk)
    t0 = time.time()
    loans = 0
    for i in range(1000):
        k = T_LOAN if i % 3 else T_RET
        loans += k == T_LOAN
        s.record(k, f"M{i % 37}", f"B{i:05d}")
        clk.tick(2)
    secs = time.time() - t0
    day = START.date()
    rows, totals, _, _ = read_sheet(os.path.join(d, xname(day)))
    jn = len(journal_r_lines(os.path.join(d, jname(day))))
    view = len(s.day_view(day).rows)
    want = {T_LABELS[0]: 1000, T_LABELS[1]: loans, T_LABELS[2]: 1000 - loans}
    ok = len(rows) == jn == view == 1000 and totals == want
    return ok, f"엑셀 {len(rows)}줄 · 일지 {jn}줄 · 화면 {view}건 · 합계 {totals} · {secs:.1f}초"


def gate_rollover_view(d):
    """날짜 넘김 뒤 화면 목록도 그날 것(자정 이후 기록만)."""
    clk = FakeClock(dt.datetime(2026, 11, 6, 23, 59, 50))
    s = core.RecordStore(d, clock=clk)
    s.record(T_LOAN, "M1", "B1")
    clk.now = dt.datetime(2026, 11, 7, 0, 0, 10)
    before = len(s.day_view().rows)
    s.record(T_LOAN, "M2", "B2")
    after = [r.book for r, _ in s.day_view().rows]
    return before == 0 and after == ["B2"], f"자정 직후 목록 {before}건 · 기록 뒤 {after}"


def gate_pending_view_and_recovery(d):
    """저장 실패 줄이 화면에 「엑셀 대기」로 보이고, 다시 켰을 때 일지에서 엑셀로 되살아난다."""
    clk = FakeClock(START)
    s = core.RecordStore(d, clock=clk)
    s.record(T_LOAN, "M1", "B1")
    clk.tick(30)
    with excel_save_fault():
        r = s.record(T_LOAN, "M2", "B2")
        flags = [(row.book, p) for row, p in s.day_view().rows]
    text = status_for(r)[0]
    # 프로그램을 껐다 켬
    s2 = core.RecordStore(d, clock=clk)
    rec = s2.recover_recent()
    rows = [x[3] for x in read_sheet(os.path.join(d, xname(START.date())))[0]]
    ok = flags == [("B1", False), ("B2", True)] and "일지에만 기록됨" in text and rows == ["B1", "B2"] \
        and rec == [(START.date(), True, 0)]
    return ok, f"화면 {flags} · 문구 앞머리 「{text[:8]}」 · 다시 켠 뒤 엑셀 {rows}"


def gate_deleted_row_respected(d):
    """사람이 엑셀에서 지운 줄(겹침 스캔 정리)을 일지가 다시 살려 내지 않는다."""
    clk = FakeClock(START)
    s = core.RecordStore(d, clock=clk)
    s.record(T_LOAN, "M1", "B1")
    clk.tick(1)
    s.record(T_LOAN, "M1", "B1")
    p = os.path.join(d, xname(START.date()))
    wb = openpyxl.load_workbook(p)
    wb[T_SHEET].delete_rows(6, 1)
    wb.save(p)
    clk.tick(30)
    s.record(T_LOAN, "M2", "B2")
    rows, totals, _, _ = read_sheet(p)
    ok = [r[3] for r in rows] == ["B1", "B2"] and totals[T_LABELS[0]] == 2
    return ok, f"엑셀 {[r[3] for r in rows]} · 전체 {totals.get(T_LABELS[0])}"


def gate_truncated_journal(d):
    """일지 마지막 줄이 끊겨 있어도(전원 차단) 다음 줄이 그 뒤에 붙지 않고 바르게 읽힌다."""
    clk = FakeClock(START)
    s = core.RecordStore(d, clock=clk)
    s.record(T_LOAN, "M1", "B1")
    jp = os.path.join(d, jname(START.date()))
    with open(jp, "ab") as fh:
        fh.write("R\t2026-11-06 09:00:05\t대출\tM9".encode("utf-8"))  # 끊긴 줄
    clk.tick(30)
    s.record(T_LOAN, "M2", "B2")
    rows = [r[3] for r in read_sheet(os.path.join(d, xname(START.date())))[0]]
    lines = journal_r_lines(jp)
    ok = rows == ["B1", "B2"] and lines[-1].endswith("\tM2\tB2")
    return ok, f"엑셀 {rows} · 일지 마지막 줄 끝 「{lines[-1][-6:]}」"


def gate_unexpected_content(d):
    """합계 아래에 사람이 적은 메모가 있으면 지우지 않고 손대지 않는다(일지에는 남음)."""
    clk = FakeClock(START)
    s = core.RecordStore(d, clock=clk)
    s.record(T_LOAN, "M1", "B1")
    p = os.path.join(d, xname(START.date()))
    wb = openpyxl.load_workbook(p)
    wb[T_SHEET]["A20"] = "메모: 3번 책 파손"
    wb.save(p)
    clk.tick(30)
    r = s.record(T_LOAN, "M2", "B2")
    wb = openpyxl.load_workbook(p)
    memo = wb[T_SHEET]["A20"].value
    ok = memo == "메모: 3번 책 파손" and r.recorded and not r.excel_saved and "양식" in r.error
    return ok, f"메모 보존={memo is not None} · 저장={r.excel_saved}"


def gate_return_needs_member(d):
    """사서 칸 ①가: 반납도 회원번호를 받는다."""
    clk = FakeClock(START)
    s = core.RecordStore(d, clock=clk)
    r = s.record(T_RET, "", "B1")
    r2 = s.record(T_RET, "M1", "B1")
    return r.missing == "member" and r2.recorded, f"회원번호 없는 반납 missing={r.missing!r}"


def _legacy_file(path, variant):
    """부록 A 대로 만든 기존 판 양식 사본. 합계 자리는 채록에 없어 두 꼴로 잰다."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = T_SHEET
    ws.merge_cells("A1:D1")
    ws["A1"] = "도서기록 2026-11-06"
    ws.merge_cells("A2:D2")
    ws["A2"] = "생성시각: 2026-11-06 08:30:00"
    for i, h in enumerate(("처리시간", "처리구분", "회원번호", "도서번호"), 1):  # 둘째 머리글은 미확인 → 다른 이름
        ws.cell(4, i, h)
    data = [("2026-11-06 08:31:00", "대출", "L1", "X1"), ("2026-11-06 08:32:00", "반납", "L2", "X2"),
            ("2026-11-06 08:33:00", "대출", "L3", "X3")]
    for r, row in enumerate(data, 5):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
    if variant == "A":  # 빈 줄 하나 뒤 라벨 A칸 · 값 B칸(수식)
        ws["A9"], ws["B9"] = "전체 건수", "=COUNTA(D5:D7)"
        ws["A10"], ws["B10"] = "대출 건수", '=COUNTIF(B5:B7,"대출")'
        ws["A11"], ws["B11"] = "반납 건수", '=COUNTIF(B5:B7,"반납")'
    else:  # 바로 아래 · 라벨 C칸 · 값 D칸 숫자 · A:B 합친 칸
        ws.merge_cells("A8:B8")
        ws["C8"], ws["D8"] = "전체 건수", 3
        ws["C9"], ws["D9"] = "대출 건수", 2
        ws["C10"], ws["D10"] = "반납 건수", 1
    wb.save(path)
    return data


def gate_legacy_append(d):
    """기존 판이 만든 그날 파일이 있으면 그 파일에 이어 적는다(머리글 · 제목 · 기존 줄 보존)."""
    out = []
    for variant in ("A", "B"):
        sub = os.path.join(d, variant)
        os.makedirs(sub)
        data = _legacy_file(os.path.join(sub, xname(START.date())), variant)
        clk = FakeClock(START)
        s = core.RecordStore(sub, clock=clk)
        s.record(T_LOAN, "M1", "B1")
        clk.tick(30)
        s.record(T_RET, "M2", "B2")
        rows, totals, header, title = read_sheet(os.path.join(sub, xname(START.date())))
        want_rows = [tuple(x) for x in data] + [("2026-11-06 09:00:00", "대출", "M1", "B1"),
                                                 ("2026-11-06 09:00:30", "반납", "M2", "B2")]
        good = rows == want_rows and header[1] == "처리구분" and title == "도서기록 2026-11-06" \
            and totals == {T_LABELS[0]: 5, T_LABELS[1]: 3, T_LABELS[2]: 2}
        out.append((variant, good, len(rows), totals))
    return all(g for _, g, _, _ in out), " · ".join(f"꼴{v}: {n}줄 {t}" for v, _, n, t in out)


def gate_offline(d):
    """인터넷 차단 상태에서 시작 · 기록 — 네트워크를 건드리면 붉음. 소스에 네트워크 모듈 import 0."""
    touched = []

    def deny(*a, **k):
        touched.append(a[:1])
        raise OSError("인터넷 차단(게이트)")

    saved = (socket.socket.connect, socket.create_connection, socket.getaddrinfo)
    socket.socket.connect = deny
    socket.create_connection = deny
    socket.getaddrinfo = deny
    try:
        for mod in ("library_emergency.core", "library_emergency.messages"):
            importlib.reload(importlib.import_module(mod))
        from library_emergency import core as c2

        clk = FakeClock(START)
        s = c2.RecordStore(d, clock=clk)
        s.ensure_today_file()
        r = s.record(T_LOAN, "M1", "B1")
    finally:
        socket.socket.connect, socket.create_connection, socket.getaddrinfo = saved
    banned = re.compile(r"^\s*(import|from)\s+(socket|urllib|http|requests|ftplib|smtplib|ssl|asyncio|webbrowser)\b",
                        re.M)
    hits = []
    for base, _, files in os.walk(os.path.join(ROOT, "library_emergency")):
        for f in files:
            if f.endswith(".py"):
                with open(os.path.join(base, f), encoding="utf-8") as fh:
                    for m in banned.finditer(fh.read()):
                        hits.append((f, m.group(0).strip()))
    ok = r.recorded and r.excel_saved and not touched and not hits
    return ok, f"기록={r.recorded} · 네트워크 시도 {len(touched)} · 금지 import {hits}"


def gate_status_text(d):
    """문구: 엑셀 실패는 「저장 완료」가 아니고, 일지 실패는 「기록 실패」."""
    row = core.Row("2026-11-06 09:00:00", T_LOAN, "M1", "B1")
    a = status_for(core.RecordResult(True, row=row, excel_saved=True))[0]
    b = status_for(core.RecordResult(True, row=row, excel_saved=False, pending=2, error="x"))[0]
    c = status_for(core.RecordResult(False, row=row, error="y"))[0]
    ok = a.startswith("저장 완료: 회원 M1 / 도서 B1") and "저장 완료" not in b and "2건" in b and c.startswith("기록 실패")
    return ok, f"성공 「{a[:14]}」 · 엑셀 실패 「{b[:10]}」 · 일지 실패 「{c[:5]}」"


def gate_journal_fail(d):
    """일지조차 못 쓰면 「기록됨」이라 하지 않는다(입력을 지우지 않게 recorded=False)."""
    clk = FakeClock(START)
    s = core.RecordStore(os.path.join(d, "x"), clock=clk)
    real_open = builtins.open

    def deny(path, *a, **k):
        if str(path).endswith("_일지.txt"):
            raise PermissionError(13, "막힘(게이트)")
        return real_open(path, *a, **k)

    builtins.open = deny
    try:
        r = s.record(T_LOAN, "M1", "B1")
    finally:
        builtins.open = real_open
    return (not r.recorded and bool(r.error)), f"recorded={r.recorded} · error={r.error!r}"


# ------------------------------------------------------------------ 윈도우 실제 잠금


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
    return k32, h


def win_lock_check(make, d, share):
    clk = FakeClock(START)
    s = make(d, clk)
    s.record(T_LOAN, "M1", "B1")
    p = os.path.join(d, xname(START.date()))
    k32, h = _win_lock(p, share)
    try:
        clk.tick(30)
        r = s.record(T_LOAN, "M2", "B2")
    finally:
        k32.CloseHandle(h)
    clk.tick(30)
    s.record(T_LOAN, "M3", "B3")
    books = [x[3] for x in read_sheet(p)[0]]
    saved_claim = bool(getattr(r, "excel_saved", False))
    return books == ["B1", "B2", "B3"] and not saved_claim, f"엑셀 {books} · 잠긴 동안 저장 완료 주장={saved_claim}"


# ------------------------------------------------------------------ 실행


def run_one(fn, *args):
    d = tempfile.mkdtemp(prefix="le_gate_")
    try:
        return fn(*args, d)
    except Exception as exc:  # 예외는 붉음
        return False, f"예외 {exc.__class__.__name__}: {exc} | {traceback.format_exc(limit=2).splitlines()[-1]}"
    finally:
        shutil.rmtree(d, ignore_errors=True)


def main(argv):
    windows = "--windows" in argv
    results = []
    print(f"python {sys.version.split()[0]} · {sys.platform} · openpyxl {openpyxl.__version__}")
    print("\n[버그 부류] 흉내 사본 → 고친 판 (사본 붉음 · 고친 판 초록일 때만 통과)")
    for name, fn, kind in BUG_CLASSES:
        if kind == "folder":
            old = run_one(fn, legacy_mimic.legacy_check_export_dir)
            new = run_one(fn, core.check_export_dir)
        else:
            old = run_one(fn, old_store)
            new = run_one(fn, new_store)
        passed = (not old[0]) and new[0]
        results.append({"group": "bug-class", "name": name, "pass": passed, "mimic": old, "fixed": new})
        print(f"  {'PASS' if passed else 'FAIL'}  {name}\n        사본: {'초록' if old[0] else '붉음'} — {old[1]}"
              f"\n        고친 판: {'초록' if new[0] else '붉음'} — {new[1]}")

    print("\n[고친 판 게이트]")
    gates = [
        ("1,000건 대조", gate_1000),
        ("날짜 넘김 뒤 화면 목록", gate_rollover_view),
        ("엑셀 대기 표시 · 다시 켜서 되살림", gate_pending_view_and_recovery),
        ("지운 줄을 되살리지 않음", gate_deleted_row_respected),
        ("끊긴 일지 줄", gate_truncated_journal),
        ("합계 아래 메모 보존", gate_unexpected_content),
        ("사서 칸 ①가 반납도 회원번호", gate_return_needs_member),
        ("기존 판 양식 파일에 이어 적기", gate_legacy_append),
        ("인터넷 차단 상태 시작", gate_offline),
        ("상태 문구", gate_status_text),
        ("일지 실패는 기록 실패", gate_journal_fail),
    ]
    for name, fn in gates:
        ok, detail = run_one(fn)
        results.append({"group": "gate", "name": name, "pass": ok, "detail": detail})
        print(f"  {'PASS' if ok else 'FAIL'}  {name} — {detail}")

    if windows:
        print("\n[윈도우 실제 파일 잠금] 흉내 사본 → 고친 판")
        if not sys.platform.startswith("win"):
            results.append({"group": "windows", "name": "윈도우 아님", "pass": False})
            print("  FAIL  윈도우가 아닌 기계에서 --windows")
        else:
            for label, share in (("완전 잠금(공유 0)", 0), ("엑셀식 잠금(읽기만 공유)", 1)):
                old = run_one(win_lock_check, old_store, share)
                new = run_one(win_lock_check, new_store, share)
                passed = (not old[0]) and new[0]
                results.append({"group": "windows", "name": label, "pass": passed, "mimic": old, "fixed": new})
                print(f"  {'PASS' if passed else 'FAIL'}  {label}\n        사본: {'초록' if old[0] else '붉음'} — "
                      f"{old[1]}\n        고친 판: {'초록' if new[0] else '붉음'} — {new[1]}")

    n_pass = sum(r["pass"] for r in results)
    print(f"\n합계: {n_pass}/{len(results)} 통과")
    with open(os.path.join(os.getcwd(), "gates_result.json"), "w", encoding="utf-8") as fh:
        json.dump({"python": sys.version, "platform": sys.platform, "results": results}, fh, ensure_ascii=False,
                  indent=1)
    return 0 if n_pass == len(results) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
