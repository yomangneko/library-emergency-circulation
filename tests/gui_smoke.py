"""화면 스모크 + 사진 — 윈도우 빌드 기계에서 돈다(맥에서도 사진 없이 돈다).

    python -I tests/gui_smoke.py <사진 폴더>

  - 인터넷 차단 상태에서 화면을 띄운다(네트워크를 건드리면 실패)
  - 대출 모드 · 반납 모드 · 저장 실패 표시를 실제 단축키 · 입력 흐름으로 만들고 사진을 찍는다
  - 화면 목록 건수 = 엑셀 줄 수 = 일지 줄 수
"""

from __future__ import annotations

import datetime as dt
import os
import shutil
import socket
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

touched = []


def _deny(*a, **k):
    touched.append(a[:1])
    raise OSError("인터넷 차단(게이트)")


socket.socket.connect = _deny
socket.create_connection = _deny
socket.getaddrinfo = _deny

import tkinter as tk  # noqa: E402

from library_emergency import core  # noqa: E402
from library_emergency.app import EmergencyLibraryApp  # noqa: E402
from gates import excel_save_fault, journal_r_lines, read_sheet  # noqa: E402


def shot(root, out_dir, name):
    root.update()
    root.after(400)
    root.update()
    if not sys.platform.startswith("win"):
        return None
    from PIL import ImageGrab

    x, y = root.winfo_rootx(), root.winfo_rooty()
    w, h = root.winfo_width(), root.winfo_height()
    path = os.path.join(out_dir, name)
    ImageGrab.grab(bbox=(x - 8, y - 32, x + w + 8, y + h + 8)).save(path)
    return path


def settle(app, root, until, timeout=30.0):
    """뒤쪽 일꾼이 엑셀을 옮기는 동안 화면을 돌리며 기다린다(시험만 기다린다 — 직원은 기다리지 않는다)."""
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        root.update()
        if until():
            return True
        time.sleep(0.02)
    root.update()
    return until()


def saved(app):
    return app.worker.idle() and app.results.empty()


def scan(app, root, member, book):
    app.member_entry.focus_force()
    app.member_var.set(member)
    app.member_entry.event_generate("<Return>")
    root.update()
    app.book_var.set(book)
    app.book_entry.event_generate("<Return>")
    root.update()


def main(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    folder = tempfile.mkdtemp(prefix="le_gui_", dir=os.getcwd())
    problems = []
    try:
        store = core.RecordStore(folder)
        store.ensure_today_file()
        root = tk.Tk()
        app = EmergencyLibraryApp(root, store)
        root.geometry("980x720+20+20")
        root.lift()
        root.attributes("-topmost", True)
        root.update()

        scan(app, root, "20240001", "EM0000123456")
        scan(app, root, "20240001", "EM0000123457")
        right_after = app.status.cget("text")
        if "저장 완료" in right_after:
            # 일꾼이 벌써 옮겼다면 맞는 말이다 — 그 줄이 지금 엑셀에 실제로 있는지 대조한다
            books_now = [r[3] for r in read_sheet(store.excel_path(store.today()))[0]]
            if "EM0000123457" not in books_now:
                problems.append(f"엑셀에 들어가기 전에 저장 완료라 함: {right_after} · 엑셀 {books_now}")
        settle(app, root, lambda: saved(app) and app.status.cget("text").startswith("저장 완료"))
        if not app.status.cget("text").startswith("저장 완료"):
            problems.append(f"대출 문구: {app.status.cget('text')}")
        shots = [shot(root, out_dir, "1_대출모드.png")]

        root.event_generate("<F2>")
        root.update()
        if not app.mode_bar.cget("text").startswith("[ 반납 ]"):
            problems.append("F2 가 반납 모드로 바꾸지 않음")
        scan(app, root, "20240002", "EM0000099999")
        settle(app, root, lambda: saved(app))
        shots.append(shot(root, out_dir, "2_반납모드.png"))

        with excel_save_fault():
            scan(app, root, "20240003", "EM0000088888")
            settle(app, root, lambda: "일지에만 기록됨" in app.status.cget("text"))
            text = app.status.cget("text")
            if "일지에만 기록됨" not in text or "저장 완료" in text:
                problems.append(f"저장 실패 문구: {text}")
            shots.append(shot(root, out_dir, "3_저장실패표시.png"))

        # 엑셀이 다시 쓸 수 있게 되면 스캔 없이 채워진다(뒤쪽 일꾼이 다시 시도)
        if not settle(app, root, lambda: saved(app) and app.status.cget("text").startswith("저장 완료")):
            problems.append(f"스캔 없이 다시 채우지 않음: {app.status.cget('text')}")

        # 같은 책 연달아
        scan(app, root, "20240003", "EM0000088888")
        settle(app, root, lambda: saved(app))
        if "방금" not in app.status.cget("text"):
            problems.append("겹침 스캔 알림 없음")
        pinks = [i for i in app.tree.get_children() if "pending" in app.tree.item(i, "tags")]
        if pinks:
            problems.append(f"엑셀에 들어간 뒤에도 분홍 줄 {len(pinks)}")
        shots.append(shot(root, out_dir, "4_엑셀복구_겹침알림.png"))

        root.event_generate("<Escape>")
        root.update()
        if app.member_var.get() or app.book_var.get():
            problems.append("Esc 가 입력을 비우지 않음")

        day = store.today()
        n_list = app.list_count()
        n_xl = len(read_sheet(store.excel_path(day))[0])
        n_j = len(journal_r_lines(store.journal_path(day)))
        if not (n_list == n_xl == n_j == 5):
            problems.append(f"건수 어긋남 화면 {n_list} · 엑셀 {n_xl} · 일지 {n_j}")
        if touched:
            problems.append(f"네트워크 시도 {len(touched)}")
        app.close()
        print(f"화면 {n_list}건 · 엑셀 {n_xl}줄 · 일지 {n_j}줄 · 네트워크 시도 {len(touched)} · 사진 {[s for s in shots if s]}")
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    for p in problems:
        print("FAIL", p)
    print("PASS 화면 스모크" if not problems else "FAIL 화면 스모크")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "screenshots"))
