-- 씨앗(real_seed.xlsx)을 Numbers 로 열어 엑셀 파일 둘로 내보낸다(2026-11-07 시험 자료를 실제로 만든 절차).
--   real_numbers_parts.xlsx   : Numbers 가 그대로 다시 쓴 것
--   real_numbers_textbox.xlsx : 위에 Numbers 에서 글상자를 하나 붙인 것
--
-- Numbers Creator Studio 15.3.1 에서는 AppleScript 의 open 이 멈추므로(실측: 오류 -609 뒤 Numbers 가 SIGTRAP 으로 끝남)
-- 여는 것은 macOS 에 맡긴다:
--   open -g -a "Numbers Creator Studio" tests/fixtures/real_seed.xlsx   (12초 기다림)
--   osascript tests/numbers_export.applescript <tests/fixtures 절대 경로>
on run argv
	set outDir to item 1 of argv
	tell application id "com.apple.Numbers"
		set d to front document
		export d to file ((POSIX file (outDir & "/real_numbers_parts.xlsx")) as text) as Microsoft Excel
		tell active sheet of d
			make new text item with properties {object text:"오늘 합계 확인함"}
		end tell
		export d to file ((POSIX file (outDir & "/real_numbers_textbox.xlsx")) as text) as Microsoft Excel
		close d saving no
	end tell
end run
