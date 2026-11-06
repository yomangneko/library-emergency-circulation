"""화면에 보일 문구 — 기록 결과에서 문구를 만드는 일은 화면 없이 시험한다."""

from __future__ import annotations

from typing import Tuple

from .core import RecordResult

# 상태 줄 색
OK = "ok"
WARN = "warn"
FAIL = "fail"


def status_for(result: RecordResult) -> Tuple[str, str]:
    """(문구, 등급). 엑셀에 못 옮긴 기록을 「저장 완료」로 알리지 않는다."""
    if not result.recorded:
        if result.missing:
            return ("", WARN)
        return (f"기록 실패 — 이 스캔은 저장되지 않았습니다. {result.error}", FAIL)
    row = result.row
    who = f"회원 {row.member} / 도서 {row.book}"
    if result.excel_saved:
        text = f"저장 완료: {who}"
        level = OK
    else:
        text = (
            f"일지에만 기록됨: {who} — {result.error}. "
            f"엑셀 미반영 {result.pending}건은 엑셀을 닫으면 다음 기록 때 채워집니다"
        )
        level = FAIL
    if result.duplicate_seconds is not None:
        text += f"  ※ 방금({result.duplicate_seconds}초 전) 같은 기록이 있음"
        if level == OK:
            level = WARN
    return (text, level)


def missing_text(missing: str) -> str:
    return "회원번호를 먼저 입력하세요" if missing == "member" else "도서 바코드를 입력하세요"
