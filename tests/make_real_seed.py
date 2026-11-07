"""실제 프로그램(Numbers) 시험 자료의 씨앗 — 이 프로그램이 만든 그날 파일(3건)에 직원이 붙였을 부품을 openpyxl 로 얹는다.

    python -I -X utf8 tests/make_real_seed.py <폴더>

씨앗을 Numbers 로 열어(글상자는 Numbers 에서 붙인다) Numbers 가 엑셀 파일로 내보낸 것이 시험 자료다
(tests/numbers_export.applescript). 시험 자료의 파일 속 내용은 Numbers 가 썼다.
"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from library_emergency import core  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.chart import BarChart, Reference  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402
from openpyxl.styles import PatternFill  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402

out = sys.argv[1]
os.makedirs(out, exist_ok=True)
t = [dt.datetime(2026, 11, 6, 9, 0, 0)]
s = core.RecordStore(out, clock=lambda: t[0])
for k, m, b in [("대출", "M1", "MB1"), ("반납", "M2", "MB2"), ("대출", "M3", "MB3")]:
    s.record(k, m, b)
    t[0] += dt.timedelta(seconds=20)
p = os.path.join(out, "도서기록_20261106.xlsx")
wb = openpyxl.load_workbook(p)
ws = wb[core.SHEET_TITLE]
ws.conditional_formatting.add("B9:B11", CellIsRule(operator="greaterThan", formula=["1"],
                                                   fill=PatternFill("solid", fgColor="FFFF00")))
dv = DataValidation(type="list", formula1='"대출,반납"')
dv.add("B5:B7")
ws.add_data_validation(dv)
ch = BarChart()
ch.add_data(Reference(ws, min_col=2, min_row=9, max_row=11))
ch.set_categories(Reference(ws, min_col=1, min_row=9, max_row=11))
ws.add_chart(ch, "F9")
seed = os.path.join(out, "seed.xlsx")
wb.save(seed)
print(seed)
