"""고치기 전 꼴을 흉내 낸 사본 — 게이트가 붉음을 먼저 내는지 보는 대조군.

기존 판의 원본 코드는 없다. 이 사본은 감사자 채록(지시서 부록 A)에서 읽히는 꼴을 그대로 옮긴 것이다:
  - 파일 이름을 시작할 때 한 번 정한다
  - 일지 없이 엑셀에 곧장 저장한다(저장 중 실패하면 그 스캔은 사라진다)
  - 합계 위에 줄을 끼워 넣는다(합계는 범위 수식)
  - 같은 기록이 연달아 와도 알리지 않는다
  - 입력값은 str.strip() 만
  - 저장 폴더는 만들기만 하고 확인하지 않는다
실제 기존 판이 이렇게 동작하는지는 실행해 보지 않았으므로 「미확인 · 흉내」다.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional

import openpyxl

from library_emergency.core import (FIRST_DATA_ROW, HEADER_ROW, HEADERS, SHEET_TITLE, TIME_FORMAT, TOTAL_LABELS,
                                    excel_name)


@dataclass
class LegacyResult:
    recorded: bool
    excel_saved: bool = False
    missing: str = ""
    duplicate_seconds: Optional[int] = None
    pending: int = 0
    error: str = ""
    row: object = None


def legacy_check_export_dir(path):
    os.makedirs(path, exist_ok=True)  # 못 만들면 예외로 죽는다

    class _C:
        ok = True
        warning = ""
        reason = ""

    _C.path = path
    return _C


class LegacyStore:
    def __init__(self, folder, clock, require_member_on_return=True):
        self.folder = folder
        self.clock = clock
        os.makedirs(folder, exist_ok=True)
        self.day = clock().date()
        self.path = os.path.join(folder, excel_name(self.day))  # 시작할 때 한 번
        if not os.path.exists(self.path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = SHEET_TITLE
            ws.merge_cells("A1:D1")
            ws["A1"] = f"도서기록 {self.day:%Y-%m-%d}"
            ws.merge_cells("A2:D2")
            ws["A2"] = f"생성시각: {clock():%Y-%m-%d %H:%M:%S}"
            for i, h in enumerate(HEADERS, 1):
                ws.cell(HEADER_ROW, i, h)
            ws.cell(FIRST_DATA_ROW + 1, 1, TOTAL_LABELS[0])
            ws.cell(FIRST_DATA_ROW + 1, 2, f"=COUNTA(D{FIRST_DATA_ROW}:D{FIRST_DATA_ROW})")
            ws.cell(FIRST_DATA_ROW + 2, 1, TOTAL_LABELS[1])
            ws.cell(FIRST_DATA_ROW + 2, 2, f'=COUNTIF(B{FIRST_DATA_ROW}:B{FIRST_DATA_ROW},"대출")')
            ws.cell(FIRST_DATA_ROW + 3, 1, TOTAL_LABELS[2])
            ws.cell(FIRST_DATA_ROW + 3, 2, f'=COUNTIF(B{FIRST_DATA_ROW}:B{FIRST_DATA_ROW},"반납")')
            wb.save(self.path)
        self.first_insert = True

    def record(self, kind, member, book):
        member = (member or "").strip()
        book = (book or "").strip()
        if not member:
            return LegacyResult(False, missing="member")
        if not book:
            return LegacyResult(False, missing="book")
        now = self.clock()
        try:
            wb = openpyxl.load_workbook(self.path)
            ws = wb[SHEET_TITLE]
            total_row = None
            for r in range(FIRST_DATA_ROW, ws.max_row + 1):
                if ws.cell(r, 1).value == TOTAL_LABELS[0]:
                    total_row = r
                    break
            # 첫 기록은 비어 있던 5행에, 그 뒤로는 합계 위(빈 줄 자리)에 끼워 넣는다
            target = total_row - 1
            if ws.cell(target, 1).value not in (None, ""):
                ws.insert_rows(target + 1)
                target += 1
            for i, v in enumerate((now.strftime(TIME_FORMAT), kind, member, book), 1):
                ws.cell(target, i, v)
            wb.save(self.path)
        except Exception as exc:
            return LegacyResult(False, error=str(exc))
        return LegacyResult(True, excel_saved=True)
