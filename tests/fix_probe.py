"""v2.0.2 게이트 ⓑ · ⓒ 의 탐침 — 화면 없이 판 하나(--pkg 의 library_emergency)의 기록 처리를 돌리고 JSON 한 줄을 낸다.

    python -I -X utf8 tests/fix_probe.py --pkg <판 폴더> --folder <폴더> --mode <trunc_write|trunc_read|journal_write|journal_read|memo>

판마다 따로 프로세스로 돈다(v2.0.1 과 고친 판이 같은 이름의 꾸러미라서).
엑셀은 시험이 openpyxl 로 직접 훑어 읽는다 — 판의 「기록 줄」 판정을 쓰지 않는다.
"""

from __future__ import annotations

import argparse
import contextlib
import datetime as dt
import json
import os
import shutil
import sys

DAY = dt.date(2026, 11, 6)
T_SHEET = "대출반납기록"
T_LABELS = ("전체 건수", "대출 건수", "반납 건수")

# 끊긴 줄 시험에 쓰는 진짜 기록 넷(셋째가 쓰다 끊기는 줄)
GENUINE = [("대출", "M1", "EM0000011111"), ("반납", "M1", "EM0000022222"), ("대출", "회원M2", "EM0000067890"),
           ("반납", "M3", "EM0000044444")]


class Clock:
    def __init__(self):
        self.now = dt.datetime(2026, 11, 6, 9, 0, 0)

    def __call__(self):
        return self.now

    def tick(self, s=20):
        self.now += dt.timedelta(seconds=s)


def xname(day):
    return f"도서기록_{day:%Y%m%d}.xlsx"


def jname(day):
    return f"도서기록_{day:%Y%m%d}_일지.txt"


def sheet_cells(path):
    """시트의 모든 값 있는 칸 {(행, 열): 값} — 시험 쪽 독립 판독."""
    import openpyxl

    ws = openpyxl.load_workbook(path)[T_SHEET]
    out = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value not in (None, ""):
                out[(c.row, c.column)] = c.value
    return out


def books_in(path, known):
    """D열에서 시험이 아는 책 번호를 찾아 센다 + 아는 번호가 아닌 「번호처럼 보이는」 값(EM…)도 모은다."""
    cells = sheet_cells(path)
    found = [v for (r, c), v in sorted(cells.items()) if c == 4 and r >= 5 and isinstance(v, str)]
    hits = [v for v in found if v in known]
    strays = [v for v in found if v not in known and v != "도서번호" and not v.startswith("메모")]
    return hits, strays, cells


def totals_of(cells):
    out = {}
    for (r, c), v in cells.items():
        if isinstance(v, str) and v.strip() in T_LABELS:
            out[v.strip()] = cells.get((r, c + 1))
    return out


@contextlib.contextmanager
def save_fault():
    from openpyxl.workbook.workbook import Workbook

    orig = Workbook.save

    def bad(self, filename):
        raise PermissionError(13, "열려 있음(흉내)", filename)

    Workbook.save = bad
    try:
        yield
    finally:
        Workbook.save = orig


def load(pkg):
    sys.path.insert(0, pkg)
    from library_emergency import core

    return core


# ------------------------------------------------------------------ ⓑ 끊긴 줄


def trunc_write(core, root):
    """진짜 기록 둘을 적고, 셋째 줄(일지에만 적힘)을 첫 바이트부터 마지막 바이트까지 자리마다 끊은 폴더를 만든다."""
    base = os.path.join(root, "base")
    os.makedirs(base)
    clk = Clock()
    s = core.RecordStore(base, clock=clk)
    for k, m, b in GENUINE[:2]:
        s.record(k, m, b)
        clk.tick()
    jp = os.path.join(base, jname(DAY))
    size0 = os.path.getsize(jp)
    kind, m, b = GENUINE[2]
    r = s.record(kind, m, b, flush=False)  # 일지에만 — 엑셀 저장 전에 전원이 나간 꼴
    with open(jp, "rb") as fh:
        full = fh.read()
    line = full[size0:]
    assert line.endswith(b"\n")
    line = line[:-1]
    cases = []
    for k in range(1, len(line) + 1):
        d = os.path.join(root, f"cut_{k:03d}")
        shutil.copytree(base, d)
        with open(os.path.join(d, jname(DAY)), "r+b") as fh:
            fh.truncate(size0 + k)
        cases.append(k)
    shutil.rmtree(base)
    return {"line_bytes": len(line), "line": line.decode("utf-8"), "cases": len(cases), "recorded": r.recorded}


def trunc_read(core, root):
    """폴더마다: 다시 켬(최근 일지 되살리기) → 다음 스캔 → 엑셀에 없는 번호가 생겼는가 · 진짜 기록이 다 있는가."""
    known = {b for _, _, b in GENUINE}
    bad = []
    n = 0
    for name in sorted(os.listdir(root)):
        if not name.startswith("cut_"):
            continue
        n += 1
        d = os.path.join(root, name)
        clk = Clock()
        clk.tick(600)
        try:
            s = core.RecordStore(d, clock=clk)
            s.recover_recent()
            k, m, b = GENUINE[3]
            r = s.record(k, m, b)
            hits, strays, cells = books_in(os.path.join(d, xname(DAY)), known)
            need = ["EM0000011111", "EM0000022222", "EM0000044444"]
            ok = all(hits.count(x) == 1 for x in need) and hits.count("EM0000067890") <= 1 and not strays \
                and r.recorded and r.excel_saved
            # 셋째 줄이 엑셀에 있다면 회원까지 온전해야 한다
            for (rr, c), v in cells.items():
                if c == 4 and v == "EM0000067890" and cells.get((rr, 3)) != "회원M2":
                    ok = False
            if not ok:
                bad.append({"cut": name, "books": hits, "strays": strays, "saved": r.excel_saved,
                            "error": r.error})
        except Exception as exc:
            bad.append({"cut": name, "exception": f"{exc.__class__.__name__}: {exc}"})
    return {"cases": n, "bad": len(bad), "bad_examples": bad[:6]}


# ------------------------------------------------------------------ ⓑ 옛 일지 읽기


def journal_write(core, root):
    """이 판으로 일지를 만든다: 40건, 중간에 엑셀 저장 실패 구간(일지에만 남음)과 되살림을 섞는다."""
    clk = Clock()
    s = core.RecordStore(root, clock=clk)
    for i in range(40):
        kind = "대출" if i % 3 else "반납"
        if 10 <= i < 15 or 30 <= i < 33:
            with save_fault():
                s.record(kind, f"회원{i}", f"BK{i:04d}")
        else:
            s.record(kind, f"회원{i}", f"BK{i:04d}")
        clk.tick()
    with save_fault():
        s.record("대출", "회원끝", "BK9999")  # 마지막은 엑셀에 못 옮긴 채
    st = core.read_journal(os.path.join(root, jname(DAY)))
    return {"rows": len(st.rows), "saved_count": st.saved_count, "keys": [list(r.key()) for r in st.rows]}


def journal_read(core, root):
    st = core.read_journal(os.path.join(root, jname(DAY)))
    return {"rows": len(st.rows), "saved_count": st.saved_count, "keys": [list(r.key()) for r in st.rows]}


# ------------------------------------------------------------------ ⓒ 메모 전 자리

MEMO = "메모: 3번 책 표지 파손"


def place_memo(path, where):
    import openpyxl

    wb = openpyxl.load_workbook(path)
    ws = wb[T_SHEET]
    label_rows = [r for r in range(5, ws.max_row + 1) for c in range(1, ws.max_column + 1)
                  if ws.cell(r, c).value == "전체 건수"]
    if where == "below_totals":
        ws.cell(ws.max_row + 2, 1, MEMO)
    elif where == "gap_before_totals":
        ws.cell(label_rows[0] - 1, 1, MEMO)  # 기록과 합계 사이 빈 줄
    elif where == "totals_side":
        ws.cell(label_rows[0], 3, MEMO)  # 합계 줄 옆 칸(라벨 A · 숫자 B 다음)
    elif where == "record_right":
        ws.cell(5, 5, MEMO)  # 기록 줄 오른쪽(E열)
    elif where == "between_records":
        ws.insert_rows(6, 1)
        ws.cell(6, 1, MEMO)  # 기록 줄 사이에 끼운 메모 줄
    elif where == "header":
        ws.cell(3, 1, MEMO)  # 머리 위(1~3행)
    wb.save(path)


MEMO_PLACES = ("below_totals", "gap_before_totals", "totals_side", "record_right", "between_records", "header")


def memo(core, root):
    out = {}
    for where in MEMO_PLACES:
        d = os.path.join(root, where)
        os.makedirs(d)
        clk = Clock()
        s = core.RecordStore(d, clock=clk)
        scans = [("대출", "M1", "MB1"), ("반납", "M2", "MB2"), ("대출", "M3", "MB3"), ("대출", "M4", "MB4")]
        try:
            for k, m, b in scans[:3]:
                s.record(k, m, b)
                clk.tick()
            p = os.path.join(d, xname(DAY))
            place_memo(p, where)
            k, m, b = scans[3]
            r = s.record(k, m, b)
            hits, strays, cells = books_in(p, {b for _, _, b in scans})
            memos = [v for v in cells.values() if v == MEMO]
            tot = totals_of(cells)
            want = {"전체 건수": 4, "대출 건수": 3, "반납 건수": 1}
            ok = len(memos) == 1 and sorted(hits) == ["MB1", "MB2", "MB3", "MB4"] and tot == want \
                and bool(getattr(r, "excel_saved", False))
            out[where] = {"ok": ok, "memo_kept": len(memos) == 1, "books": hits, "totals": tot,
                          "saved": bool(getattr(r, "excel_saved", False)), "error": getattr(r, "error", "")}
        except Exception as exc:
            out[where] = {"ok": False, "exception": f"{exc.__class__.__name__}: {exc}"}
    return out


# ------------------------------------------------------------------ v2.0.3 ⓐ 판정 규칙 전 꼴

HAND_CASES = [
    # (이름, 처리시간 값, 구분, 도서번호, 세어야 하는가)
    ("글자 9:30", "9:30", "대출", "HAND1", True),
    ("엑셀 시각 값", dt.time(9, 30), "대출", "HAND1", True),
    ("날짜 값", dt.date(2026, 11, 6), "대출", "HAND1", True),
    ("날짜·시각 값", dt.datetime(2026, 11, 6, 9, 30), "대출", "HAND1", True),
    ("글자 오전 9시", "오전 9시", "대출", "HAND1", True),
    ("처리시간 빈칸", None, "대출", "HAND1", False),
    ("구분이 대출처리", "9:31", "대출처리", "HAND1", False),
    ("도서번호 빈칸", "9:32", "대출", None, False),
]


def rule(core, root):
    """3건 기록 → 직원이 기록 사이(7행)에 손으로 한 줄 → 다음 기록. 합계 = 시험이 센 수 · 화면 목록 건수도 같음."""
    import openpyxl

    out = {}
    for name, a, kind, book, counts in HAND_CASES:
        d = os.path.join(root, f"h{len(out)}")
        os.makedirs(d)
        clk = Clock()
        try:
            s = core.RecordStore(d, clock=clk)
            scans = [("대출", "M1", "MB1"), ("반납", "M2", "MB2"), ("대출", "M3", "MB3"), ("대출", "M4", "MB4")]
            for k, m, b in scans[:3]:
                s.record(k, m, b)
                clk.tick()
            p = os.path.join(d, xname(DAY))
            wb = openpyxl.load_workbook(p)
            ws = wb[T_SHEET]
            ws.insert_rows(7, 1)
            for col, v in enumerate((a, kind, "M7", book), start=1):
                ws.cell(7, col, v)
            if isinstance(a, dt.time):
                ws.cell(7, 1).number_format = "h:mm"
            wb.save(p)
            k, m, b = scans[3]
            r = s.record(k, m, b)
            cells = sheet_cells(p)
            tot = totals_of(cells)
            n = 4 + (1 if counts else 0)
            loans = 3 + (1 if counts and kind == "대출" else 0)
            want = {"전체 건수": n, "대출 건수": loans, "반납 건수": n - loans}
            view = s.day_view(DAY)
            view_times = [row.time for row, _ in view.rows]
            hand_kept = any(v == "M7" for (rr, c), v in cells.items() if c == 3)
            ok = tot == want and len(view.rows) == n and hand_kept and bool(r.excel_saved) \
                and all(isinstance(t, str) and t for t in view_times)
            out[name] = {"ok": ok, "counts_expected": counts, "totals": tot, "want": want, "view_rows": len(view.rows),
                         "hand_time_shown": next((row.time for row, _ in view.rows if row.member == "M7"), None),
                         "saved": bool(r.excel_saved), "error": r.error}
        except Exception as exc:
            out[name] = {"ok": False, "exception": f"{exc.__class__.__name__}: {exc}"}
    return out


# ------------------------------------------------------------------ v2.0.3 ⓑ 수식 전 자리

# 기록 5 · 6 · 7행, 빈 줄 8행, 합계 9(전체) · 10(대출) · 11(반납)행, 다음 기록은 8행에 끼워진다.
FORMULA_CASES = [
    # (이름, 수식 칸, 수식, 메모 (칸, 값) 또는 None, 기대값을 내는 법)
    ("아래 · 합계 칸", "E5", "=B9", None, "total"),
    ("아래 · 합계 아래 메모 칸", "E6", "=A13", ("A13", "메모X"), "메모X"),
    ("위 · 기록 줄 칸", "A13", "=D5", None, "MB1"),
    ("위 · 머리 칸", "A13", "=A1", None, "title"),
    ("범위 · 위쪽만", "A13", "=COUNTA(D5:D6)", None, 2),
    ("범위 · 끼운 자리를 걸침", "A13", "=COUNTA(D5:D8)", None, 4),
    ("범위 · 아래쪽만", "A13", "=SUM(B9:B11)", None, "sum_totals"),
]


def eval_formula(ws, f):
    """시험 쪽 작은 계산기: =칸 · =COUNTA(범위) · =SUM(범위). 시트 이름 · $ 는 떼고 읽는다."""
    import re

    body = f[1:].replace("$", "")
    body = re.sub(r"^'?[^'!]*'?!", "", body)

    def cells(rng):
        a, b = rng.split(":")
        ma, mb = re.fullmatch(r"([A-Z]+)(\d+)", a), re.fullmatch(r"([A-Z]+)(\d+)", b)
        col = ma.group(1)
        return [ws[f"{col}{r}"].value for r in range(int(ma.group(2)), int(mb.group(2)) + 1)]

    m = re.fullmatch(r"COUNTA\(([A-Z]+\d+:[A-Z]+\d+)\)", body)
    if m:
        return sum(1 for v in cells(m.group(1)) if v not in (None, ""))
    m = re.fullmatch(r"SUM\(([A-Z]+\d+:[A-Z]+\d+)\)", body)
    if m:
        return sum(v for v in cells(m.group(1)) if isinstance(v, (int, float)))
    m = re.fullmatch(r"[A-Z]+\d+", body)
    if m:
        return ws[body].value
    raise ValueError(f)


def formula(core, root):
    import openpyxl

    out = {}
    for name, at, f, memo_cell, how in FORMULA_CASES:
        d = os.path.join(root, f"f{len(out)}")
        os.makedirs(d)
        clk = Clock()
        try:
            s = core.RecordStore(d, clock=clk)
            scans = [("대출", "M1", "MB1"), ("반납", "M2", "MB2"), ("대출", "M3", "MB3"), ("대출", "M4", "MB4")]
            for k, m, b in scans[:3]:
                s.record(k, m, b)
                clk.tick()
            p = os.path.join(d, xname(DAY))
            wb = openpyxl.load_workbook(p)
            ws = wb[T_SHEET]
            title = ws["A1"].value
            assert ws["A9"].value == "전체 건수", "합계 자리가 예상과 다름"
            if memo_cell:
                ws[memo_cell[0]] = memo_cell[1]
            ws[at] = f
            wb.save(p)
            k, m, b = scans[3]
            r = s.record(k, m, b)
            ws = openpyxl.load_workbook(p)[T_SHEET]
            found = [(c.coordinate, c.value) for row in ws.iter_rows() for c in row
                     if isinstance(c.value, str) and c.value.startswith("=")]
            (coord, newf), = found
            got = eval_formula(ws, newf)
            tot = totals_of(sheet_cells(p))
            want = {"total": tot.get("전체 건수"), "title": title,
                    "sum_totals": sum(v for v in tot.values() if isinstance(v, int))}.get(how, how)
            # 합계 자체도 시험이 센 수(4 · 3 · 1)와 맞아야 기대값으로 쓴다
            totals_ok = tot == {"전체 건수": 4, "대출 건수": 3, "반납 건수": 1}
            ok = got == want and totals_ok and bool(r.excel_saved)
            out[name] = {"ok": ok, "formula_after": f"{coord} {newf}", "value": got, "want": want,
                         "saved": bool(r.excel_saved)}
        except Exception as exc:
            out[name] = {"ok": False, "exception": f"{exc.__class__.__name__}: {exc}"}
    return out


# ------------------------------------------------------------------ v2.0.3 ⓒ 고칠 수 없는 수식


def badformula(core, root):
    import hashlib

    import openpyxl
    from library_emergency.messages import status_for

    clk = Clock()
    s = core.RecordStore(root, clock=clk)
    for k, m, b in [("대출", "M1", "MB1"), ("반납", "M2", "MB2")]:
        s.record(k, m, b)
        clk.tick()
    p = os.path.join(root, xname(DAY))
    wb = openpyxl.load_workbook(p)
    wb[T_SHEET]["E5"] = "=SUM(표1[금액])"  # 표(구조적) 참조 — 칸 이름이 아니라 고칠 수 없다
    wb.save(p)
    with open(p, "rb") as fh:
        before = hashlib.sha256(fh.read()).hexdigest()
    r = s.record("대출", "M3", "MB3")
    with open(p, "rb") as fh:
        after = hashlib.sha256(fh.read()).hexdigest()
    text = status_for(r)[0]
    in_journal = any(row.book == "MB3" for row in core.read_journal(os.path.join(root, jname(DAY))).rows)
    return {"saved": bool(r.excel_saved), "file_unchanged": before == after, "in_journal": in_journal,
            "status": text}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pkg", required=True)
    ap.add_argument("--folder", required=True)
    ap.add_argument("--mode", required=True)
    a = ap.parse_args()
    core = load(a.pkg)
    fn = {"trunc_write": trunc_write, "trunc_read": trunc_read, "journal_write": journal_write,
          "journal_read": journal_read, "memo": memo, "rule": rule, "formula": formula,
          "badformula": badformula}[a.mode]
    print("RESULT " + json.dumps(fn(core, a.folder), ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
