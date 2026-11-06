"""기록 처리 — 화면 없이 돌고 시험할 수 있는 부분.

한 번의 스캔이 지나는 길:
  1) 입력값 정리(앞뒤 공백 · 바코드 기계 꼬리 글자 · 보이지 않는 글자)
  2) 그날 일지 파일(덧붙이기 전용)에 한 줄을 먼저 쓴다 — 여기까지가 「기록됨」
  3) 그날 엑셀 파일에 일지에서 아직 안 옮긴 줄을 모두 옮겨 적는다
     실패하면(엑셀이 열려 있음 · 저장 실패) 줄은 일지에 남고, 다음 저장 때 다시 채운다

인터넷 · 외부 프로그램을 쓰지 않는다.
"""

from __future__ import annotations

import datetime as _dt
import os
import tempfile
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from typing import Callable, Iterable, List, Optional, Sequence, Tuple

KIND_LOAN = "대출"
KIND_RETURN = "반납"
KINDS = (KIND_LOAN, KIND_RETURN)

SHEET_TITLE = "대출반납기록"
HEADERS = ("처리시간", "구분", "회원번호", "도서번호")
TOTAL_LABELS = ("전체 건수", "대출 건수", "반납 건수")
HEADER_ROW = 4
FIRST_DATA_ROW = 5
TIME_FORMAT = "%Y-%m-%d %H:%M:%S"

#: 같은 회원 · 같은 책 · 같은 구분이 이 시간(초) 안에 다시 오면 「방금 같은 기록이 있음」을 알린다.
DUPLICATE_WINDOW_SECONDS = 10

JOURNAL_HEADER = "# 도서기록 일지 — 덧붙이기 전용. R=기록 줄, S=엑셀에 옮겨 적은 줄 수. 지우거나 고치지 마세요.\n"

Clock = Callable[[], _dt.datetime]


def excel_name(day: _dt.date) -> str:
    return f"도서기록_{day:%Y%m%d}.xlsx"


def journal_name(day: _dt.date) -> str:
    return f"도서기록_{day:%Y%m%d}_일지.txt"


# ---------------------------------------------------------------- 입력값


_INVISIBLE = {"​", "‌", "‍", "⁠", "﻿"}


def normalize_input(value: Optional[str]) -> str:
    """앞뒤 공백 · 줄바꿈 · 탭 · 보이지 않는 글자 · 제어 글자를 걷어낸다.

    보이는 글자는 하나도 바꾸지 않는다(대소문자 · 하이픈 그대로).
    """
    if value is None:
        return ""
    kept = []
    for ch in str(value):
        if ch in _INVISIBLE:
            continue
        cat = unicodedata.category(ch)
        if cat in ("Cc", "Cf"):
            # 제어 글자(\r \n \t 등)는 칸 안의 구분자가 될 수 없으므로 공백으로 바꿔 둔다
            kept.append(" ")
            continue
        kept.append(ch)
    return "".join(kept).strip()


# ---------------------------------------------------------------- 저장 폴더


@dataclass
class FolderCheck:
    ok: bool
    path: str
    reason: str = ""  # ok=False 일 때 화면에 보일 까닭
    warning: str = ""  # ok=True 여도 알려야 할 것(임시 폴더 등)


def check_export_dir(path: str) -> FolderCheck:
    """폴더를 만들고 실제로 써 보고 지워 본다. 임시 폴더(압축 파일 안에서 바로 실행)면 경고."""
    path = os.path.abspath(path)
    try:
        os.makedirs(path, exist_ok=True)
    except OSError as exc:
        return FolderCheck(False, path, f"저장 폴더를 만들 수 없습니다 ({exc.__class__.__name__})")
    if not os.path.isdir(path):
        return FolderCheck(False, path, "저장 폴더 자리에 같은 이름의 파일이 있습니다")
    probe = os.path.join(path, ".쓰기확인.tmp")
    try:
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("ok")
            fh.flush()
            os.fsync(fh.fileno())
        os.remove(probe)
    except OSError as exc:
        return FolderCheck(False, path, f"저장 폴더에 쓸 수 없습니다 ({exc.__class__.__name__})")
    warning = ""
    if is_temporary_location(path):
        warning = (
            "프로그램이 임시 폴더에서 실행되고 있습니다(압축 파일 안에서 바로 실행한 경우 등). "
            "여기 쌓인 기록은 나중에 지워질 수 있습니다."
        )
    return FolderCheck(True, path, "", warning)


def is_temporary_location(path: str) -> bool:
    try:
        tmp = os.path.normcase(os.path.realpath(tempfile.gettempdir()))
        here = os.path.normcase(os.path.realpath(path))
    except OSError:
        return False
    return here == tmp or here.startswith(tmp + os.sep)


# ---------------------------------------------------------------- 일지


@dataclass(frozen=True)
class Row:
    time: str
    kind: str
    member: str
    book: str

    def key(self) -> Tuple[str, str, str, str]:
        return (self.time, self.kind, self.member, self.book)


@dataclass
class JournalState:
    rows: List[Row] = field(default_factory=list)
    saved_count: int = 0  # 엑셀에 옮겨 적었다고 확인된 앞쪽 줄 수


def read_journal(path: str) -> JournalState:
    state = JournalState()
    if not os.path.exists(path):
        return state
    with open(path, "r", encoding="utf-8", newline="") as fh:
        for raw in fh:
            if not raw.endswith("\n"):
                continue  # 쓰다 끊긴 마지막 줄은 믿지 않는다
            line = raw.rstrip("\r\n")
            if not line or line.startswith("#"):
                continue
            parts = line.split("\t")
            if parts[0] == "R" and len(parts) == 5 and parts[2] in KINDS:
                state.rows.append(Row(parts[1], parts[2], parts[3], parts[4]))
            elif parts[0] == "S" and len(parts) >= 2 and parts[1].isdigit():
                state.saved_count = min(int(parts[1]), len(state.rows))
    return state


def _append_line(path: str, line: str) -> None:
    new_file = not os.path.exists(path)
    with open(path, "a+b") as fh:
        prefix = b""
        if new_file:
            prefix = JOURNAL_HEADER.encode("utf-8")
        else:
            fh.seek(0, os.SEEK_END)
            if fh.tell() > 0:
                fh.seek(-1, os.SEEK_END)
                if fh.read(1) != b"\n":
                    prefix = b"\n"  # 끊긴 줄 뒤에 붙지 않게
        fh.seek(0, os.SEEK_END)
        fh.write(prefix + line.encode("utf-8") + b"\n")
        fh.flush()
        os.fsync(fh.fileno())


# ---------------------------------------------------------------- 엑셀


class ExcelFormatError(Exception):
    """엑셀 파일이 예상한 양식이 아니다 — 덮어쓰지 않고 멈춘다."""


def _load_openpyxl():
    import openpyxl  # noqa: WPS433 — 묶음 판에는 들어 있다

    return openpyxl


def _row_has_total_label(ws, r: int) -> bool:
    for c in range(1, max(ws.max_column, 4) + 1):
        v = ws.cell(r, c).value
        if isinstance(v, str) and v.strip() in TOTAL_LABELS:
            return True
    return False


def _row_empty(ws, r: int) -> bool:
    return all(ws.cell(r, c).value in (None, "") for c in range(1, max(ws.max_column, 4) + 1))


def _scan_layout(ws) -> Tuple[List[Row], int]:
    """(기록 줄들, 마지막 기록 줄 번호)를 돌려준다. 양식이 다르면 ExcelFormatError."""
    rows: List[Row] = []
    last = HEADER_ROW
    r = FIRST_DATA_ROW
    max_row = ws.max_row
    while r <= max_row:
        if _row_has_total_label(ws, r) or _row_empty(ws, r):
            break
        vals = [ws.cell(r, c).value for c in range(1, 5)]
        t, k, m, b = ("" if v is None else str(v) for v in vals)
        if isinstance(vals[0], _dt.datetime):
            t = vals[0].strftime(TIME_FORMAT)
        rows.append(Row(t, k.strip(), m, b))
        last = r
        r += 1
    # 기록 줄 아래에는 빈 줄과 합계 줄만 있어야 한다
    for rr in range(last + 1, max_row + 1):
        if not (_row_empty(ws, rr) or _row_has_total_label(ws, rr)):
            raise ExcelFormatError(f"{rr}행에 예상하지 못한 내용이 있습니다")
    return rows, last


def _style_parts():
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    side = Side(style="thin", color="D9D9D9")
    return {
        "border": Border(left=side, right=side, top=side, bottom=side),
        "title_fill": PatternFill("solid", fgColor="1F4E78"),
        "title_font": Font(name="Malgun Gothic", bold=True, color="FFFFFF", size=14),
        "bold": Font(name="Malgun Gothic", bold=True),
        "plain": Font(name="Malgun Gothic"),
        "center": Alignment(horizontal="center", vertical="center"),
        "right": Alignment(horizontal="right"),
        "head_fill": PatternFill("solid", fgColor="DDEBF7"),
    }


def new_workbook(day: _dt.date, created: _dt.datetime):
    openpyxl = _load_openpyxl()
    st = _style_parts()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_TITLE
    ws.merge_cells("A1:D1")
    ws["A1"] = f"도서기록 {day:%Y-%m-%d}"
    ws["A1"].fill = st["title_fill"]
    ws["A1"].font = st["title_font"]
    ws["A1"].alignment = st["center"]
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:D2")
    ws["A2"] = f"생성시각: {created:%Y-%m-%d %H:%M:%S}"
    ws["A2"].alignment = st["right"]
    for i, h in enumerate(HEADERS, start=1):
        c = ws.cell(HEADER_ROW, i, h)
        c.font = st["bold"]
        c.alignment = st["center"]
        c.border = st["border"]
        c.fill = st["head_fill"]
    for col, width in zip("ABCD", (22, 10, 22, 26)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = f"A{FIRST_DATA_ROW}"
    return wb


def write_rows_and_totals(ws, existing_last: int, new_rows: Sequence[Row]) -> List[Row]:
    """기록 줄 뒤(합계 위)에 새 줄을 적고, 합계 줄을 지우고 다시 쓴다. 전체 기록 줄을 돌려준다."""
    st = _style_parts()
    # 기록 줄 아래(빈 줄 · 합계 줄)를 걷어낸다 — 끼워 넣기 대신 다시 쓰므로 합계가 어긋나지 않는다
    for rng in list(ws.merged_cells.ranges):
        if rng.max_row > existing_last:  # 합계 자리에 합친 칸이 있으면 먼저 푼다
            ws.unmerge_cells(str(rng))
    if ws.max_row > existing_last:
        ws.delete_rows(existing_last + 1, ws.max_row - existing_last)
    r = existing_last
    for row in new_rows:
        r += 1
        for i, v in enumerate(row.key(), start=1):
            c = ws.cell(r, i, v)
            c.border = st["border"]
            c.font = st["plain"]
        ws.cell(r, 2).alignment = st["center"]
    all_rows, last = _scan_layout(ws)
    ws.auto_filter.ref = f"A{HEADER_ROW}:D{max(last, HEADER_ROW)}"
    counts = (
        len(all_rows),
        sum(1 for x in all_rows if x.kind == KIND_LOAN),
        sum(1 for x in all_rows if x.kind == KIND_RETURN),
    )
    tr = last + 2
    for label, n in zip(TOTAL_LABELS, counts):
        a = ws.cell(tr, 1, label)
        b = ws.cell(tr, 2, n)
        for c in (a, b):
            c.font = st["bold"]
            c.border = st["border"]
        tr += 1
    return all_rows


def read_totals(ws) -> dict:
    """합계 줄(라벨 오른쪽 칸의 숫자)을 읽는다 — 시험과 대조용."""
    out = {}
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() in TOTAL_LABELS:
                out[v.strip()] = ws.cell(r, c + 1).value
    return out


def save_atomic(wb, path: str) -> None:
    """같은 폴더의 임시 파일에 다 쓴 뒤 바꿔 끼운다 — 저장 중 끊겨도 원래 파일은 온전하다."""
    folder = os.path.dirname(path)
    fd, tmp = tempfile.mkstemp(prefix=".저장중_", suffix=".xlsx", dir=folder)
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


# ---------------------------------------------------------------- 기록 처리


@dataclass
class RecordResult:
    recorded: bool  # 일지에 남았는가 — False 면 이 스캔은 저장되지 않았다
    row: Optional[Row] = None
    excel_saved: bool = False  # 이번 스캔까지 엑셀에 옮겨 적었는가
    pending: int = 0  # 일지에만 있고 엑셀에 아직 없는 줄 수(그날)
    duplicate_seconds: Optional[int] = None  # 방금 같은 기록이 있었으면 몇 초 전인지
    missing: str = ""  # 입력 필요: "member" / "book"
    error: str = ""  # 기록 실패 · 엑셀 실패 까닭(화면용 짧은 말)


@dataclass
class DayView:
    day: _dt.date
    rows: List[Tuple[Row, bool]]  # (줄, 엑셀 대기 여부) — 오래된 것이 앞
    excel_error: str = ""


class RecordStore:
    """저장 폴더 하나에 대한 기록 처리. 날짜는 매 기록 때 시계에서 다시 읽는다."""

    def __init__(self, folder: str, clock: Clock = _dt.datetime.now, require_member_on_return: bool = True,
                 saver: Callable = save_atomic):
        self.folder = os.path.abspath(folder)
        self.clock = clock
        self.require_member_on_return = require_member_on_return
        self._saver = saver
        self._last: Optional[Tuple[Tuple[str, str, str], _dt.datetime]] = None
        self._excel_cache: dict = {}  # day -> 마지막으로 읽은 엑셀 기록 줄
        self.last_excel_error: dict = {}

    # -- 경로
    def excel_path(self, day: _dt.date) -> str:
        return os.path.join(self.folder, excel_name(day))

    def journal_path(self, day: _dt.date) -> str:
        return os.path.join(self.folder, journal_name(day))

    def today(self) -> _dt.date:
        return self.clock().date()

    # -- 기록
    def record(self, kind: str, member: str, book: str) -> RecordResult:
        if kind not in KINDS:
            raise ValueError(kind)
        member = normalize_input(member)
        book = normalize_input(book)
        if not member and (kind == KIND_LOAN or self.require_member_on_return):
            return RecordResult(False, missing="member")
        if not book:
            return RecordResult(False, missing="book")
        now = self.clock()
        day = now.date()
        row = Row(now.strftime(TIME_FORMAT), kind, member, book)
        try:
            os.makedirs(self.folder, exist_ok=True)
            _append_line(self.journal_path(day), "\t".join(("R",) + row.key()))
        except OSError as exc:
            return RecordResult(False, row=row, error=f"일지에 쓸 수 없습니다 ({exc.__class__.__name__})")

        dup = None
        sig = (kind, member, book)
        if self._last is not None and self._last[0] == sig:
            gap = (now - self._last[1]).total_seconds()
            if 0 <= gap <= DUPLICATE_WINDOW_SECONDS:
                dup = int(gap)
        self._last = (sig, now)

        ok, pending, err = self.flush(day)
        return RecordResult(True, row=row, excel_saved=ok, pending=pending, duplicate_seconds=dup, error=err)

    # -- 엑셀로 옮겨 적기
    def flush(self, day: _dt.date) -> Tuple[bool, int, str]:
        """일지에서 엑셀로 아직 안 옮긴 줄을 옮긴다. (성공, 남은 줄 수, 실패 까닭)."""
        jpath = self.journal_path(day)
        state = read_journal(jpath)
        pending_rows = state.rows[state.saved_count:]
        xpath = self.excel_path(day)
        if not pending_rows and os.path.exists(xpath):
            return True, 0, ""
        try:
            openpyxl = _load_openpyxl()
            if os.path.exists(xpath):
                wb = openpyxl.load_workbook(xpath)
                if SHEET_TITLE in wb.sheetnames:
                    ws = wb[SHEET_TITLE]
                else:
                    raise ExcelFormatError(f"「{SHEET_TITLE}」 시트가 없습니다")
            else:
                wb = new_workbook(day, self.clock())
                ws = wb[SHEET_TITLE]
            excel_rows, last = _scan_layout(ws)
            # 엑셀에 이미 있는 줄은 다시 적지 않는다(저장 뒤 S 줄을 못 쓰고 끊긴 경우).
            # 이미 옮겼다고 확인된 줄로 설명되지 않는 엑셀 줄만 「이미 적힌 대기 줄」로 본다.
            unexplained = Counter(r.key() for r in excel_rows) - Counter(r.key() for r in state.rows[:state.saved_count])
            to_add = []
            for r in pending_rows:
                if unexplained[r.key()] > 0:
                    unexplained[r.key()] -= 1
                else:
                    to_add.append(r)
            all_rows = write_rows_and_totals(ws, last, to_add)
            self._saver(wb, xpath)
        except Exception as exc:  # 엑셀 쪽 실패는 무엇이든 기록을 잃지 않는 쪽으로
            msg = _excel_error_text(exc)
            self.last_excel_error[day] = msg
            return False, len(pending_rows), msg
        try:
            _append_line(jpath, f"S\t{len(state.rows)}\t{self.clock():{TIME_FORMAT}}")
        except OSError:
            pass  # 다음 저장 때 엑셀에 이미 있는 줄로 걸러진다
        self._excel_cache[day] = all_rows
        self.last_excel_error.pop(day, None)
        return True, 0, ""

    def ensure_today_file(self) -> Tuple[bool, str]:
        day = self.today()
        ok, _, err = self.flush(day)
        return ok, err

    def recover_recent(self, days: int = 14) -> List[Tuple[_dt.date, bool, int]]:
        """시작할 때: 최근 일지 중 엑셀에 못 옮긴 줄이 있는 날을 다시 채운다."""
        out = []
        today = self.today()
        for i in range(days, -1, -1):
            d = today - _dt.timedelta(days=i)
            jp = self.journal_path(d)
            if not os.path.exists(jp):
                continue
            st = read_journal(jp)
            if st.saved_count < len(st.rows):
                ok, pending, _ = self.flush(d)
                out.append((d, ok, pending))
        return out

    # -- 화면용
    def day_view(self, day: Optional[_dt.date] = None) -> DayView:
        day = day or self.today()
        state = read_journal(self.journal_path(day))
        excel_rows = None
        err = ""
        xpath = self.excel_path(day)
        if os.path.exists(xpath):
            try:
                wb = _load_openpyxl().load_workbook(xpath, read_only=True)
                try:
                    ws = wb[SHEET_TITLE]
                    excel_rows = _read_rows_readonly(ws)
                finally:
                    wb.close()
                self._excel_cache[day] = excel_rows
            except Exception as exc:
                err = _excel_error_text(exc)
                excel_rows = self._excel_cache.get(day)
        if excel_rows is None:
            excel_rows = list(state.rows[:state.saved_count]) if not os.path.exists(xpath) else []
        pending = state.rows[state.saved_count:]
        unexplained = Counter(r.key() for r in excel_rows) - Counter(r.key() for r in state.rows[:state.saved_count])
        rows = [(r, False) for r in excel_rows]
        for r in pending:
            if unexplained[r.key()] > 0:
                unexplained[r.key()] -= 1
            else:
                rows.append((r, True))
        return DayView(day, rows, err or self.last_excel_error.get(day, ""))


def _read_rows_readonly(ws) -> List[Row]:
    rows = []
    for i, vals in enumerate(ws.iter_rows(min_row=FIRST_DATA_ROW, max_col=4, values_only=True)):
        vals = list(vals) + [None] * (4 - len(vals))
        if all(v in (None, "") for v in vals):
            break
        if any(isinstance(v, str) and v.strip() in TOTAL_LABELS for v in vals):
            break
        t = vals[0].strftime(TIME_FORMAT) if isinstance(vals[0], _dt.datetime) else ("" if vals[0] is None else str(vals[0]))
        k, m, b = ("" if v is None else str(v) for v in vals[1:4])
        rows.append(Row(t, k.strip(), m, b))
    return rows


def _excel_error_text(exc: BaseException) -> str:
    if isinstance(exc, PermissionError):
        return "엑셀 파일이 열려 있어 저장하지 못했습니다"
    if isinstance(exc, ExcelFormatError):
        return f"엑셀 파일 양식이 예상과 달라 손대지 않았습니다 — {exc}"
    if isinstance(exc, OSError):
        return f"엑셀 파일을 저장하지 못했습니다 ({exc.__class__.__name__})"
    return f"엑셀 파일을 읽거나 저장하지 못했습니다 ({exc.__class__.__name__})"


def iter_counts(rows: Iterable[Row]) -> Tuple[int, int, int]:
    rows = list(rows)
    return (len(rows), sum(r.kind == KIND_LOAN for r in rows), sum(r.kind == KIND_RETURN for r in rows))
