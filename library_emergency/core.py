"""기록 처리 — 화면 없이 돌고 시험할 수 있는 부분.

한 번의 스캔이 지나는 길:
  1) 입력값 정리(앞뒤 공백 · 바코드 기계 꼬리 글자 · 보이지 않는 글자)
  2) 그날 일지 파일(덧붙이기 전용)에 한 줄을 먼저 쓴다 — 여기까지가 「기록됨」
  3) 그날 엑셀 파일에 일지에서 아직 안 옮긴 줄을 모두 옮겨 적는다
     실패하면(엑셀이 열려 있음 · 저장 실패) 줄은 일지에 남고, 다음 저장 때 다시 채운다

화면은 1)·2)만 하고 바로 다음 입력을 받는다(그날 건수와 상관없이 같은 시간).
3)은 뒤쪽 일꾼(ExcelWorker) 하나가 맡는다 — 엑셀에 쓰는 것은 늘 그 하나뿐이다.

인터넷 · 외부 프로그램을 쓰지 않는다.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
import tempfile
import threading
import time
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


#: 끊긴 꼬리를 봉인하는 표지. 맨 앞 글자 \x1f 는 제어 글자라 입력값에 들어갈 수 없다(normalize_input 이 걷어냄).
TRUNCATED_SEAL = "\x1f끊김"
TIME_RE = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")


def _journal_line_ok(line: str) -> bool:
    """끝까지 다 쓴 줄인가: 탭 말고 제어 글자가 없어야 한다(봉인 표지 · 전원 차단 뒤 0 바이트 등)."""
    return not any(unicodedata.category(ch) == "Cc" and ch != "\t" for ch in line)


def read_journal(path: str) -> JournalState:
    """일지를 읽는다. 기록으로 인정하는 줄은 끝까지 다 쓴 R 줄만이다.

    - 줄 끝(\\n)이 없는 마지막 줄은 쓰다 끊긴 줄이다 — 믿지 않는다.
    - 끊긴 줄 뒤에 다음 줄을 덧붙일 때는 그 꼬리에 봉인 표지(제어 글자)를 붙여 닫는다(_append_line).
      제어 글자가 섞인 줄 · UTF-8 이 깨진 줄 · 칸 수나 시각 형식이 어긋난 줄은 버린다.
    - v2.0.0 · v2.0.1 이 쓴 일지도 같은 꼴(R/S 줄)이라 그대로 읽힌다.
    """
    state = JournalState()
    if not os.path.exists(path):
        return state
    with open(path, "rb") as fh:
        for raw in fh:
            if not raw.endswith(b"\n"):
                continue  # 쓰다 끊긴 마지막 줄은 믿지 않는다
            try:
                line = raw.decode("utf-8").rstrip("\r\n")
            except UnicodeDecodeError:
                continue  # 한글 글자 가운데서 끊긴 줄
            if not line or line.startswith("#") or not _journal_line_ok(line):
                continue
            parts = line.split("\t")
            if parts[0] == "R" and len(parts) == 5 and parts[2] in KINDS and TIME_RE.fullmatch(parts[1]) \
                    and parts[4]:
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
                    # 쓰다 끊긴 꼬리가 있다 — 봉인 표지를 붙여 닫는다. 그냥 줄만 바꾸면 끊긴 꼬리가
                    # 「끝까지 쓴 줄」처럼 읽혀 잘린 번호가 기록이 된다(㉸).
                    prefix = (TRUNCATED_SEAL + "\n").encode("utf-8")
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


def _width(ws) -> int:
    # ws.max_column 은 부를 때마다 칸 전체를 훑는다 — 줄마다 부르면 줄 수의 제곱으로 느려진다. 한 번만 센다.
    return max(ws.max_column, 4)


@dataclass
class Layout:
    rows: List[Row]  # 기록 줄(위에서부터)
    last: int  # 마지막 기록 줄 번호(없으면 머리글 줄)
    labels: List[Tuple[int, int, str]]  # 합계 줄 라벨 자리 (행, 열, 라벨)


def _classify(vals: Sequence) -> Tuple[str, object]:
    """한 줄(A열부터의 값들)을 가른다 → ("record", Row) / ("label", [(열, 라벨)]) / ("other", None).

    기록 줄 판정 규칙:
      A(처리시간)가 날짜 · 시각 값이거나 「YYYY-MM-DD HH:MM:SS」 글자이고,
      B(구분)가 「대출」 또는 「반납」이고(앞뒤 공백 무시),
      D(도서번호)가 비어 있지 않은 줄. C(회원번호)는 따지지 않는다(기존 판 파일).
    합계 줄: 어느 칸이든 「전체 건수 · 대출 건수 · 반납 건수」 라벨이 있는 줄.
    그 밖의 칸에 사람이 적은 것은 모두 메모 — 세지 않고 · 고치지 않고 · 쓰기를 막지 않는다.
    """
    vals = list(vals) + [None] * (4 - len(vals))
    labels = [(i + 1, v.strip()) for i, v in enumerate(vals) if isinstance(v, str) and v.strip() in TOTAL_LABELS]
    if labels:
        return "label", labels
    a, b, c, d = vals[:4]
    if isinstance(a, _dt.datetime):
        t = a.strftime(TIME_FORMAT)
    elif isinstance(a, str) and TIME_RE.fullmatch(a.strip()):
        t = a.strip()
    else:
        return "other", None
    if not (isinstance(b, str) and b.strip() in KINDS):
        return "other", None
    book = "" if d is None else str(d)
    if not book.strip():
        return "other", None
    return "record", Row(t, b.strip(), "" if c is None else str(c), book)


def _scan_layout(ws) -> Layout:
    rows: List[Row] = []
    last = HEADER_ROW
    labels: List[Tuple[int, int, str]] = []
    width = _width(ws)
    for r, vals in enumerate(ws.iter_rows(min_row=FIRST_DATA_ROW, max_col=width, values_only=True),
                             start=FIRST_DATA_ROW):
        kind, x = _classify(vals)
        if kind == "record":
            rows.append(x)
            last = r
        elif kind == "label":
            labels.extend((r, c, lab) for c, lab in x)
    return Layout(rows, last, labels)


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


def write_rows_and_totals(ws, layout, new_rows: Sequence[Row]) -> List[Row]:
    """마지막 기록 줄 바로 다음에 새 줄을 끼워 적고, 합계 줄의 숫자 칸만 다시 쓴다. 전체 기록 줄을 돌려준다.

    그 아래에 있던 것(빈 줄 · 메모 · 합계 줄 · 합계 아래 메모)은 내용과 차례 그대로 아래로 밀린다.
    합계 줄의 다른 칸 · 기록 줄 오른쪽 칸 · 머리 위 칸은 손대지 않는다.
    """
    if isinstance(layout, int):  # 예전 부름꼴(시험 준비용) — 지금 시트를 읽어 쓴다
        layout = _scan_layout(ws)
    st = _style_parts()
    k = len(new_rows)
    last = layout.last
    labels = list(layout.labels)
    if k:
        # openpyxl 의 줄 끼우기는 합친 칸을 옮기지 않는다 — 아래쪽 합친 칸은 풀었다가 옮겨 다시 합친다
        moved = []
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row > last:
                moved.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
                ws.unmerge_cells(str(rng))
        ws.insert_rows(last + 1, k)
        for r1, c1, r2, c2 in moved:
            ws.merge_cells(start_row=r1 + k, start_column=c1, end_row=r2 + k, end_column=c2)
        labels = [(r + k if r > last else r, c, lab) for r, c, lab in labels]
        for j, row in enumerate(new_rows, start=1):
            r = last + j
            for i, v in enumerate(row.key(), start=1):
                c = ws.cell(r, i, v)
                c.border = st["border"]
                c.font = st["plain"]
            ws.cell(r, 2).alignment = st["center"]
    all_rows = list(layout.rows) + list(new_rows)
    new_last = last + k
    ws.auto_filter.ref = f"A{HEADER_ROW}:D{max(new_last, HEADER_ROW)}"
    counts = {
        TOTAL_LABELS[0]: len(all_rows),
        TOTAL_LABELS[1]: sum(1 for x in all_rows if x.kind == KIND_LOAN),
        TOTAL_LABELS[2]: sum(1 for x in all_rows if x.kind == KIND_RETURN),
    }
    if labels:
        for r, c, lab in labels:  # 숫자 칸(라벨 오른쪽)만 다시 쓴다
            ws.cell(r, c + 1).value = counts[lab]
    else:  # 합계 줄이 아직 없다(새 파일 등) — 맨 아래 내용 뒤 한 줄 띄고 만든다
        tr = max(new_last, ws.max_row) + 2
        for label in TOTAL_LABELS:
            a = ws.cell(tr, 1, label)
            b = ws.cell(tr, 2, counts[label])
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
    day: Optional[_dt.date] = None  # 이 줄이 적힌 날
    index: int = -1  # 그날 일지에서 이 줄의 차례(0부터) — 엑셀에 들어갔는지 맞춰 보는 데 쓴다
    queued: bool = False  # 엑셀 옮겨 적기를 뒤쪽 일꾼에게 맡겼다(아직 시도 전)


@dataclass
class FlushResult:
    day: _dt.date
    ok: bool
    pending: int  # 일지에만 있고 엑셀에 아직 없는 줄 수
    error: str = ""
    saved_upto: int = 0  # 그날 일지 앞쪽에서 엑셀에 들어갔다고 확인된 줄 수


@dataclass
class DayView:
    day: _dt.date
    rows: List[Tuple[Row, bool]]  # (줄, 엑셀 대기 여부) — 오래된 것이 앞
    excel_error: str = ""
    indexes: List[int] = field(default_factory=list)  # rows 와 나란히: 엑셀 대기 줄의 일지 차례(아니면 -1)


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
        # 일지에 덧붙이는 것(화면의 R 줄 · 일꾼의 S 줄)이 한 줄씩만 들어가게
        self._journal_lock = threading.Lock()
        # 엑셀 파일을 읽고 쓰는 것은 한 번에 하나만(일꾼 · 시작 때 되살림 · 화면 목록 짓기)
        self._excel_lock = threading.Lock()
        self._journal_counts: dict = {}  # day -> 그날 일지 R 줄 수(처음 한 번만 읽고 그 뒤로는 셈)

    # -- 경로
    def excel_path(self, day: _dt.date) -> str:
        return os.path.join(self.folder, excel_name(day))

    def journal_path(self, day: _dt.date) -> str:
        return os.path.join(self.folder, journal_name(day))

    def today(self) -> _dt.date:
        return self.clock().date()

    # -- 기록
    def journal_count(self, day: _dt.date) -> int:
        with self._journal_lock:
            return self._journal_count_locked(day)

    def _journal_count_locked(self, day: _dt.date) -> int:
        if day not in self._journal_counts:
            self._journal_counts[day] = len(read_journal(self.journal_path(day)).rows)
        return self._journal_counts[day]

    def record(self, kind: str, member: str, book: str, flush: bool = True) -> RecordResult:
        """한 건을 기록한다. flush=False 면 일지에만 적고 돌아온다(엑셀은 뒤쪽 일꾼이 옮긴다)."""
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
            with self._journal_lock:
                index = self._journal_count_locked(day)
                _append_line(self.journal_path(day), "\t".join(("R",) + row.key()))
                self._journal_counts[day] = index + 1
        except OSError as exc:
            return RecordResult(False, row=row, error=f"일지에 쓸 수 없습니다 ({exc.__class__.__name__})")

        dup = None
        sig = (kind, member, book)
        if self._last is not None and self._last[0] == sig:
            gap = (now - self._last[1]).total_seconds()
            if 0 <= gap <= DUPLICATE_WINDOW_SECONDS:
                dup = int(gap)
        self._last = (sig, now)

        if not flush:
            return RecordResult(True, row=row, duplicate_seconds=dup, day=day, index=index, queued=True)
        f = self.flush_detail(day)
        return RecordResult(True, row=row, excel_saved=f.ok and f.saved_upto > index, pending=f.pending,
                            duplicate_seconds=dup, error=f.error, day=day, index=index)

    # -- 엑셀로 옮겨 적기
    def flush(self, day: _dt.date) -> Tuple[bool, int, str]:
        """일지에서 엑셀로 아직 안 옮긴 줄을 옮긴다. (성공, 남은 줄 수, 실패 까닭)."""
        f = self.flush_detail(day)
        return f.ok, f.pending, f.error

    def flush_detail(self, day: _dt.date) -> FlushResult:
        with self._excel_lock:
            return self._flush_locked(day)

    def _flush_locked(self, day: _dt.date) -> FlushResult:
        jpath = self.journal_path(day)
        state = read_journal(jpath)
        pending_rows = state.rows[state.saved_count:]
        xpath = self.excel_path(day)
        if not pending_rows and os.path.exists(xpath):
            return FlushResult(day, True, 0, "", len(state.rows))
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
            layout = _scan_layout(ws)
            excel_rows = layout.rows
            # 엑셀에 이미 있는 줄은 다시 적지 않는다(저장 뒤 S 줄을 못 쓰고 끊긴 경우).
            # 이미 옮겼다고 확인된 줄로 설명되지 않는 엑셀 줄만 「이미 적힌 대기 줄」로 본다.
            unexplained = Counter(r.key() for r in excel_rows) - Counter(r.key() for r in state.rows[:state.saved_count])
            to_add = []
            for r in pending_rows:
                if unexplained[r.key()] > 0:
                    unexplained[r.key()] -= 1
                else:
                    to_add.append(r)
            all_rows = write_rows_and_totals(ws, layout, to_add)
            self._saver(wb, xpath)
        except Exception as exc:  # 엑셀 쪽 실패는 무엇이든 기록을 잃지 않는 쪽으로
            msg = _excel_error_text(exc)
            self.last_excel_error[day] = msg
            return FlushResult(day, False, len(pending_rows), msg, state.saved_count)
        try:
            with self._journal_lock:
                _append_line(jpath, f"S\t{len(state.rows)}\t{self.clock():{TIME_FORMAT}}")
        except OSError:
            pass  # 다음 저장 때 엑셀에 이미 있는 줄로 걸러진다
        self._excel_cache[day] = all_rows
        self.last_excel_error.pop(day, None)
        return FlushResult(day, True, 0, "", len(state.rows))

    def ensure_today_file(self) -> Tuple[bool, str]:
        day = self.today()
        ok, _, err = self.flush(day)
        return ok, err

    def days_with_pending(self, days: int = 14) -> List[_dt.date]:
        """최근 일지 중 엑셀에 옮겼다는 표시(S 줄)가 아직 다 안 붙은 날."""
        out = []
        today = self.today()
        for i in range(days, -1, -1):
            d = today - _dt.timedelta(days=i)
            jp = self.journal_path(d)
            if os.path.exists(jp):
                st = read_journal(jp)
                if st.saved_count < len(st.rows):
                    out.append(d)
        return out

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
        """화면 목록을 처음부터 짓는다 — 시작할 때 · 날짜가 바뀔 때만 부른다(스캔마다 부르지 않는다)."""
        day = day or self.today()
        with self._excel_lock:
            return self._day_view_locked(day)

    def _day_view_locked(self, day: _dt.date) -> DayView:
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
        indexes = [-1] * len(rows)
        for j, r in enumerate(pending):
            if unexplained[r.key()] > 0:
                unexplained[r.key()] -= 1
            else:
                rows.append((r, True))
                indexes.append(state.saved_count + j)
        return DayView(day, rows, err or self.last_excel_error.get(day, ""), indexes)



# ---------------------------------------------------------------- 뒤쪽 일꾼


class ExcelWorker:
    """엑셀 옮겨 적기를 화면과 다른 줄(스레드)에서 맡는 일꾼 하나.

    - 화면은 스캔마다 poke(그날)만 부르고 바로 돌아간다.
    - 일꾼은 밀린 날을 하나씩 옮겨 적는다. 그사이 스캔이 더 오면 다음 번에 한꺼번에 옮긴다.
    - 엑셀이 열려 있어 못 쓴 날은 retry_seconds 마다 스캔 없이 다시 시도한다.
    - 엑셀 파일은 매번 다시 읽고 그 위에 적는다(기억해 둔 내용으로 덮어쓰지 않는다) —
      사람이 그사이 지운 줄 · 적은 메모가 그대로 남는다.
    - 결과는 on_result(FlushResult)로 알린다(일꾼의 줄에서 불린다 — 화면은 큐로 받아야 한다).
    """

    def __init__(self, store: "RecordStore", on_result: Callable[[FlushResult], None],
                 retry_seconds: float = 3.0):
        self.store = store
        self.on_result = on_result
        self.retry_seconds = retry_seconds
        self._lock = threading.Lock()
        self._dirty: dict = {}  # day -> 다음 시도 시각(time.monotonic)
        self._wake = threading.Event()
        self._stopping = False
        self._busy = False
        self._thread = threading.Thread(target=self._run, name="excel-worker", daemon=True)

    def start(self) -> "ExcelWorker":
        if not self._thread.is_alive():
            self._thread.start()
        return self

    def poke(self, day: _dt.date) -> None:
        """그날 일지에 새 줄이 있다 — 곧바로 옮겨 적게 한다."""
        with self._lock:
            self._dirty[day] = 0.0
        self._wake.set()

    def idle(self) -> bool:
        """밀린 날이 없고 지금 옮겨 적는 중도 아니다."""
        with self._lock:
            return not self._dirty and not self._busy

    def busy(self) -> bool:
        with self._lock:
            return self._busy

    def pending_days(self) -> List[_dt.date]:
        with self._lock:
            return sorted(self._dirty)

    def wait_idle(self, timeout: float) -> bool:
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            if self.idle():
                return True
            time.sleep(0.02)
        return self.idle()

    def stop(self, timeout: float = 2.0) -> bool:
        """멈추게 하고 timeout 초까지 기다린다. 저장 중이면 그 저장은 끝까지 간다(임시 파일 → 바꿔 끼우기).
        다 못 기다려도 엑셀 파일은 깨지지 않는다 — 못 옮긴 줄은 다음 시작 때 일지에서 채운다."""
        self._stopping = True
        self._wake.set()
        if self._thread.is_alive():
            self._thread.join(timeout)
        return not self._thread.is_alive()

    def _next_due(self) -> Tuple[Optional[_dt.date], Optional[float]]:
        """(지금 할 날, 그게 없으면 다음 시도까지 남은 초)."""
        with self._lock:
            if not self._dirty:
                return None, None
            now = time.monotonic()
            day, due = min(self._dirty.items(), key=lambda kv: kv[1])
            if due <= now:
                self._busy = True
                return day, None
            return None, due - now

    def _run(self) -> None:
        while not self._stopping:
            day, wait = self._next_due()
            if day is None:
                self._wake.wait(wait)
                self._wake.clear()
                continue
            try:
                res = self.store.flush_detail(day)
            except Exception as exc:  # 일꾼은 죽지 않는다 — 다음에 다시
                res = FlushResult(day, False, -1, _excel_error_text(exc), 0)
            with self._lock:
                self._busy = False
                if res.ok and self.store.journal_count(day) <= res.saved_upto:
                    self._dirty.pop(day, None)  # 다 옮겼다(그 뒤 스캔은 새 poke 로 다시 들어온다)
                elif res.ok:
                    self._dirty[day] = 0.0  # 옮기는 사이 새 줄이 왔다 — 곧바로 다시
                else:
                    self._dirty[day] = time.monotonic() + self.retry_seconds
            try:
                self.on_result(res)
            except Exception:
                pass


def _read_rows_readonly(ws) -> List[Row]:
    rows = []
    for vals in ws.iter_rows(min_row=FIRST_DATA_ROW, max_col=4, values_only=True):
        kind, x = _classify(vals)
        if kind == "record":
            rows.append(x)
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
