-- 씨앗(seed.xlsx)을 Numbers 로 열어 엑셀 파일 둘로 내보낸다.
--   real_numbers_parts.xlsx   : Numbers 가 그대로 다시 쓴 것(조건부 서식 · 드롭다운 · 차트)
--   real_numbers_textbox.xlsx : 위에 Numbers 에서 글상자를 하나 붙인 것
-- osascript tests/numbers_export.applescript <씨앗 경로> <내보낼 폴더>
on run argv
	set seedPath to POSIX file (item 1 of argv)
	set outDir to item 2 of argv
	tell application id "com.apple.Numbers"
		set d to open seedPath
		delay 2
		export d to POSIX file (outDir & "/real_numbers_parts.xlsx") as Microsoft Excel
		tell active sheet of d
			make new text item with properties {object text:"오늘 합계 확인함"}
		end tell
		export d to POSIX file (outDir & "/real_numbers_textbox.xlsx") as Microsoft Excel
		close d saving no
	end tell
end run
