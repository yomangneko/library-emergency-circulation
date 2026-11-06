"""화면 — 기존 판(부록 A 채록)과 같은 차례 · 문구 · 단축키. 기록 처리는 core 에 맡긴다."""

from __future__ import annotations

import datetime as _dt
import os
import queue
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from . import core
from .messages import BUSY, FAIL, OK, WARN, missing_text, status_for, status_for_queued

TITLE = "도서관 비상 대출/반납 기록 프로그램"
MODE_COLORS = {core.KIND_LOAN: "#3498db", core.KIND_RETURN: "#e74c3c"}
STATUS_COLORS = {OK: "#1e7e34", WARN: "#b8860b", FAIL: "#c0392b", BUSY: "#555555"}
POLL_MS = 100  # 뒤쪽 일꾼의 결과를 화면에 옮기는 간격
FONT = "Malgun Gothic"
MONO = "Consolas"


def base_dir() -> str:
    """실행 파일(또는 스크립트) 옆 — PyInstaller 한 파일 묶음은 임시 풀림 폴더가 아니라 exe 자리."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


#: 파이썬은 한 번에 한 줄(스레드)만 돈다. 뒤쪽 일꾼이 엑셀을 정리하는 동안 화면 줄이 차례를 기다리는
#: 시간이 길어지지 않게 차례를 넘기는 간격을 짧게 한다(기본 0.005초).
SWITCH_INTERVAL = 0.0005


def tune_threads() -> None:
    sys.setswitchinterval(SWITCH_INTERVAL)


def open_path(path: str) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def choose_export_dir(root: tk.Misc, first: core.FolderCheck) -> Optional[str]:
    """실행 파일 옆 exports 를 못 쓰면 알리고 쓸 수 있는 자리를 고르게 한다."""
    check = first
    while True:
        if check.ok and not check.warning:
            return check.path
        if check.ok and check.warning:
            if messagebox.askyesno(
                "저장 폴더 확인",
                f"{check.warning}\n\n지금 자리: {check.path}\n\n"
                "다른 저장 폴더를 고르시겠습니까? (아니오: 이 자리에 그대로 저장)",
                parent=root,
            ):
                pass
            else:
                return check.path
        else:
            messagebox.showwarning(
                "저장 폴더 확인",
                f"{check.reason}.\n\n자리: {check.path}\n\n기록을 쌓을 폴더를 골라 주세요.",
                parent=root,
            )
        picked = filedialog.askdirectory(parent=root, title="기록을 저장할 폴더 고르기", mustexist=False)
        if not picked:
            if messagebox.askretrycancel("저장 폴더 확인", "저장 폴더 없이는 기록할 수 없습니다.", parent=root):
                continue
            return None
        check = core.check_export_dir(picked)


class EmergencyLibraryApp:
    def __init__(self, root: tk.Tk, store: core.RecordStore, worker: Optional[core.ExcelWorker] = None):
        self.root = root
        self.store = store
        self.mode = core.KIND_LOAN
        self.shown_day = store.today()
        # 엑셀 옮겨 적기는 뒤쪽 일꾼 하나가 맡는다. 결과는 큐로 받아 화면 줄에서만 화면을 고친다.
        self.results: "queue.Queue[core.FlushResult]" = queue.Queue()
        self.worker = worker or core.ExcelWorker(store, self.results.put)
        self._pending_items: dict = {}  # (날, 일지 차례) -> (목록 줄 id, Row)
        self._last_scan: Optional[core.RecordResult] = None
        self._last_shown: Optional[tuple] = None
        self._build()
        self.set_mode(core.KIND_LOAN, announce=False)
        self.refresh_list()
        self.set_status("준비되었습니다", OK)
        self.worker.start()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self._closed = False
        self._poll_job = self.root.after(POLL_MS, self._poll_results)
        self._tick_job = self.root.after(30_000, self._tick)

    # ------------------------------------------------------------ 화면 짓기
    def _build(self) -> None:
        r = self.root
        r.title(TITLE)
        r.geometry("980x720")
        r.minsize(760, 560)
        style = ttk.Style(r)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=(FONT, 10))
        style.configure("Treeview", rowheight=24, font=(FONT, 10))
        style.configure("Treeview.Heading", font=(FONT, 10, "bold"))
        style.configure("TLabelframe.Label", font=(FONT, 11, "bold"))

        top = ttk.Frame(r, padding=(12, 10, 12, 4))
        top.pack(fill="x")
        self.btn_loan = tk.Button(top, text="대출 (F1)", font=(FONT, 12, "bold"), width=12,
                                  command=lambda: self.set_mode(core.KIND_LOAN))
        self.btn_return = tk.Button(top, text="반납 (F2)", font=(FONT, 12, "bold"), width=12,
                                    command=lambda: self.set_mode(core.KIND_RETURN))
        self.btn_loan.pack(side="left", padx=(0, 6))
        self.btn_return.pack(side="left")

        self.mode_bar = tk.Label(r, font=(FONT, 15, "bold"), fg="white", pady=10)
        self.mode_bar.pack(fill="x", padx=12, pady=(4, 8))

        box = ttk.LabelFrame(r, text="현장 입력", padding=12)
        box.pack(fill="x", padx=12)
        ttk.Label(box, text="회원번호").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.member_var = tk.StringVar()
        self.member_entry = tk.Entry(box, textvariable=self.member_var, font=(MONO, 16), width=28)
        self.member_entry.grid(row=0, column=1, sticky="we", pady=4)
        ttk.Label(box, text="도서 바코드").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.book_var = tk.StringVar()
        self.book_entry = tk.Entry(box, textvariable=self.book_var, font=(MONO, 16), width=28)
        self.book_entry.grid(row=1, column=1, sticky="we", pady=4)
        self.btn_record = ttk.Button(box, text="즉시 기록", command=self.record)
        self.btn_record.grid(row=0, column=2, rowspan=2, sticky="ns", padx=(10, 0), pady=4)
        box.columnconfigure(1, weight=1)
        ttk.Label(box, foreground="#555555",
                  text="빠른 사용: 회원번호 스캔 → 책 스캔 → Enter. Enter와 동시에 엑셀과 화면 목록에 바로 저장됩니다."
                  ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.member_entry.bind("<Return>", self._member_enter)
        self.member_entry.bind("<KP_Enter>", self._member_enter)
        self.book_entry.bind("<Return>", lambda e: self.record())
        self.book_entry.bind("<KP_Enter>", lambda e: self.record())

        # 아래쪽(상태 줄 · 단추)을 먼저 붙여 자리를 잡고, 남는 자리를 목록이 쓴다
        self.status = tk.Label(r, anchor="w", font=(FONT, 11), padx=12, pady=6, relief="sunken", bd=1,
                               wraplength=950, justify="left")
        # 문구가 두 줄이 되어도 잘리지 않게: 상태 줄 · 아래 단추를 목록보다 먼저 아래에 붙이고 폭에 맞춰 접는다
        self.status.bind("<Configure>", lambda e: self.status.configure(wraplength=max(200, e.width - 30)))
        self.status.pack(fill="x", side="bottom")

        bottom = ttk.Frame(r, padding=(12, 0, 12, 4))
        bottom.pack(fill="x", side="bottom")
        ttk.Button(bottom, text="오늘 엑셀 열기", command=self.open_today_excel).pack(side="left")
        ttk.Button(bottom, text="저장 폴더 열기", command=self.open_folder).pack(side="left", padx=6)
        ttk.Button(bottom, text="입력 초기화 (ESC)", command=self.reset_inputs).pack(side="left")
        self.folder_label = ttk.Label(bottom, foreground="#777777", text=f"저장 폴더: {self.store.folder}")
        self.folder_label.pack(side="right")

        listbox = ttk.LabelFrame(r, text="오늘 기록", padding=8)
        listbox.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ("시각", "구분", "회원", "도서")
        self.tree = ttk.Treeview(listbox, columns=cols, show="headings", height=12)
        for c, w in zip(cols, (180, 120, 240, 280)):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="center" if c == "구분" else "w")
        self.tree.tag_configure("pending", background="#fde2e1")
        sb = ttk.Scrollbar(listbox, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")


        r.bind("<F1>", lambda e: self.set_mode(core.KIND_LOAN))
        r.bind("<F2>", lambda e: self.set_mode(core.KIND_RETURN))
        r.bind("<Escape>", lambda e: self.reset_inputs())

    # ------------------------------------------------------------ 동작
    def set_mode(self, kind: str, announce: bool = True) -> None:
        self.mode = kind
        color = MODE_COLORS[kind]
        self.mode_bar.configure(text=f"[ {kind} ] 모드 - 회원번호를 스캔하세요", bg=color)
        for btn, k in ((self.btn_loan, core.KIND_LOAN), (self.btn_return, core.KIND_RETURN)):
            if k == kind:
                btn.configure(bg=MODE_COLORS[k], fg="white", activebackground=MODE_COLORS[k], relief="sunken")
            else:
                btn.configure(bg="#ecf0f1", fg="#333333", activebackground="#dfe6e9", relief="raised")
        self.reset_inputs(announce=announce)

    def reset_inputs(self, announce: bool = True) -> None:
        self.member_var.set("")
        self.book_var.set("")
        self.member_entry.focus_set()
        if announce:
            self.set_status(f"[ {self.mode} ] 입력을 초기화했습니다", OK)

    def _member_enter(self, _event=None):
        if not core.normalize_input(self.member_var.get()):
            self._warn_missing("member")
            return "break"
        self.member_var.set(core.normalize_input(self.member_var.get()))
        self.book_entry.focus_set()
        return "break"

    def _warn_missing(self, missing: str) -> None:
        messagebox.showwarning("입력 필요", missing_text(missing), parent=self.root)
        (self.member_entry if missing == "member" else self.book_entry).focus_set()

    def record(self) -> None:
        """스캔 한 건: 일지에 한 줄 → 목록 맨 위에 그 한 줄 → 다음 입력. 엑셀은 뒤쪽 일꾼이 옮긴다.
        여기서는 엑셀 파일을 읽지도 쓰지도 않는다 — 그날 건수와 상관없이 같은 시간이 든다."""
        result = self.store.record(self.mode, self.member_var.get(), self.book_var.get(), flush=False)
        if not result.recorded and result.missing:
            self._warn_missing(result.missing)
            return
        if not result.recorded:
            text, level = status_for(result)
            self.set_status(text, level)
            # 저장되지 않았으므로 입력을 지우지 않는다 — 다시 시도할 수 있게
            messagebox.showerror("기록 실패", text, parent=self.root)
            return
        self.worker.poke(result.day)
        self._last_scan = result
        self._last_shown = None
        text, level = status_for_queued(result)
        self.set_status(text, level)
        # 같은 회원이 여러 권을 이어서 처리하도록 회원번호는 두고 책 칸만 비운다
        self.member_var.set(result.row.member)
        self.book_var.set("")
        self.book_entry.focus_set()
        if result.day != self.shown_day:
            self.refresh_list()  # 자정을 넘긴 첫 기록 — 그날 목록으로 새로 짓는다
        else:
            self._add_pending_row(result.day, result.index, result.row)

    def _row_values(self, row: core.Row, pending: bool) -> tuple:
        kind = row.kind + (" · 엑셀 대기" if pending else "")
        shown_time = row.time[11:] if len(row.time) >= 19 else row.time
        return (shown_time, kind, row.member, row.book)

    def _add_pending_row(self, day, index: int, row: core.Row) -> None:
        iid = self.tree.insert("", 0, values=self._row_values(row, True), tags=("pending",))
        self._pending_items[(day, index)] = (iid, row)

    def refresh_list(self) -> None:
        """목록을 처음부터 짓는다 — 시작할 때 · 날짜가 바뀔 때만."""
        view = self.store.day_view()
        self.shown_day = view.day
        self.tree.delete(*self.tree.get_children())
        self._pending_items = {}
        indexes = view.indexes or [-1] * len(view.rows)
        for (row, pending), idx in zip(reversed(view.rows), reversed(indexes)):  # 최근 것이 위
            iid = self.tree.insert("", "end", values=self._row_values(row, pending),
                                   tags=("pending",) if pending else ())
            if pending:
                self._pending_items[(view.day, idx)] = (iid, row)
        self.root.title(f"{TITLE} — {view.day:%Y-%m-%d}")

    def _poll_results(self) -> None:
        try:
            while True:
                try:
                    res = self.results.get_nowait()
                except queue.Empty:
                    break
                self._apply_result(res)
        finally:
            if not self._closed:
                self._poll_job = self.root.after(POLL_MS, self._poll_results)

    def _apply_result(self, res: core.FlushResult) -> None:
        # 엑셀에 들어간 줄은 「엑셀 대기」 표시를 푼다
        for key in [k for k in self._pending_items if k[0] == res.day and k[1] < res.saved_upto]:
            iid, row = self._pending_items.pop(key)
            if self.tree.exists(iid):
                self.tree.item(iid, values=self._row_values(row, False), tags=())
        last = self._last_scan
        if last is None or last.day != res.day:
            return
        if last.index < res.saved_upto:
            shown = ("saved", last.index)
            if self._last_shown != shown:
                self._last_shown = shown
                done = core.RecordResult(True, row=last.row, excel_saved=True,
                                         duplicate_seconds=last.duplicate_seconds)
                self.set_status(*status_for(done))
        elif not res.ok:
            shown = ("failed", last.index, res.pending, res.error)
            if self._last_shown != shown:
                self._last_shown = shown
                failed = core.RecordResult(True, row=last.row, excel_saved=False, pending=res.pending,
                                           error=res.error, duplicate_seconds=last.duplicate_seconds)
                self.set_status(*status_for(failed))

    def close(self) -> None:
        """창 닫기. 엑셀 정리 중이면 잠깐 기다린다 — 다 못 기다려도 엑셀 파일은 임시 파일로 바꿔 끼우므로
        깨지지 않고, 못 옮긴 줄은 다음 시작 때 일지에서 채운다."""
        self._closed = True
        for job in (self._poll_job, self._tick_job):
            try:
                self.root.after_cancel(job)
            except tk.TclError:
                pass
        try:
            self.worker.stop(timeout=2.0)
        finally:
            self.root.destroy()

    def list_count(self) -> int:
        return len(self.tree.get_children())

    def set_status(self, text: str, level: str) -> None:
        self.status.configure(text=text, fg=STATUS_COLORS.get(level, "#333333"))

    def open_today_excel(self) -> None:
        day = self.store.today()
        self.worker.poke(day)  # 엑셀에 쓰는 것은 일꾼 하나뿐 — 화면은 맡기기만 한다
        path = self.store.excel_path(day)
        try:
            if not os.path.exists(path):
                raise FileNotFoundError(path)
            open_path(path)
        except Exception:
            messagebox.showerror("오류", "엑셀 파일을 열 수 없습니다.", parent=self.root)

    def open_folder(self) -> None:
        try:
            open_path(self.store.folder)
        except Exception:
            messagebox.showerror("오류", "폴더를 열 수 없습니다.", parent=self.root)

    def _tick(self) -> None:
        """자정을 넘기면 목록과 제목을 그날 것으로 바꾼다(기록 파일은 매 기록 때 날짜로 고른다)."""
        try:
            if self.store.today() != self.shown_day:
                self.worker.poke(self.store.today())  # 그날 엑셀 만들기도 일꾼이 한다
                self.refresh_list()
                self.set_status(f"날짜가 바뀌었습니다 — {self.shown_day:%Y-%m-%d} 파일에 기록합니다", WARN)
        finally:
            if not self._closed:
                self._tick_job = self.root.after(30_000, self._tick)


def main() -> int:
    tune_threads()
    root = tk.Tk()
    root.withdraw()
    first = core.check_export_dir(os.path.join(base_dir(), "exports"))
    folder = choose_export_dir(root, first)
    if not folder:
        root.destroy()
        return 1
    store = core.RecordStore(folder)
    recovered = store.recover_recent()
    ok, err = store.ensure_today_file()
    root.deiconify()
    app = EmergencyLibraryApp(root, store)
    notes = []
    for day, done, pending in recovered:
        if done:
            notes.append(f"{day:%m-%d} 일지에서 엑셀로 되살림")
        else:
            notes.append(f"{day:%m-%d} 엑셀 미반영 {pending}건(일지에 있음)")
    if not ok:
        notes.append(err)
    if notes:
        app.set_status("준비되었습니다 — " + " · ".join(notes), WARN if not ok or any(not d for _, d, _ in recovered) else OK)
    root.mainloop()
    return 0
