"""v2.0.4 게이트(고침 지시 FIX3 2026-11-07 §1-4)의 탐침 — 화면 없이 판 하나(--pkg)의 기록 처리를 돌리고 JSON 한 줄을 낸다.

    python -I -X utf8 tests/fix3_probe.py --pkg <판 폴더> --folder <폴더> --mode <move|refuse|real|inventory>

  move      : 부품마다 · 자리마다(합계 줄 · 기록 줄 · 머리 칸 · 걸침) — 3건 기록 → 부품을 붙인다 → 다음 기록
              → 부품이 가리키는 자리 = 붙인 직후 파일에서 읽은 자리에 「8행에 1줄 끼움」 규칙을 시험이 직접 적용한 값
  refuse    : 저장하면 지워지거나 옮길 수 없는 부품 — 3건 기록 → 부품을 붙인다 → 다음 기록
              → 파일 지문 그대로 · 일지에 남음 · 빨간 문구(까닭에 부품 이름)
  real      : 실제 프로그램(Numbers)이 만든 시험 자료(tests/fixtures)로 같은 일을 한다
  inventory : 판과 상관없이 openpyxl 자체를 잰다(1-1 표) — 읽고 저장하면 남는가 · 줄을 끼우면 옮기는가

부품이 가리키는 자리는 시험이 xlsx 속 XML 을 직접 읽어 얻는다(판의 코드도 openpyxl 의 해석도 쓰지 않는다).
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import io
import json
import os
import re
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
DAY = dt.date(2026, 11, 6)
T_SHEET = "대출반납기록"
AT, K = 8, 1  # 기록 5~7행 · 빈 줄 8행 · 합계 9~11행 → 다음 기록은 8행에 1줄 끼워진다

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
      "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
RID = "{%s}id" % NS["r"]


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


def load(pkg):
    sys.path.insert(0, pkg)
    from library_emergency import core

    assert os.path.realpath(core.__file__).startswith(os.path.realpath(pkg)), core.__file__
    return core


def sha(p):
    with open(p, "rb") as fh:
        return hashlib.sha256(fh.read()).hexdigest()


# ------------------------------------------------------------------ 시험 쪽 규칙(판의 코드를 쓰지 않는다)


def r1(n):
    """1부터 세는 줄 번호: 끼운 자리 이상은 K 줄 아래로."""
    return n + K if n >= AT else n


def shift_text(text):
    """칸 이름(A1 · $A$1 · A1:B2 · 4:4)이 든 글자의 줄 번호를 규칙대로 고친다(열 전체 A:D 는 그대로)."""
    def cell(m):
        return f"{m.group(1)}{m.group(2)}{m.group(3)}{r1(int(m.group(4)))}"

    def rows(m):
        return f"{m.group(1)}{r1(int(m.group(2)))}:{m.group(3)}{r1(int(m.group(4)))}"

    if re.fullmatch(r"\$?\d+:\$?\d+", text.split("!")[-1]):
        head, _, tail = text.rpartition("!")
        return (head + "!" if head else "") + re.sub(r"(\$?)(\d+):(\$?)(\d+)", rows, tail)
    return re.sub(r"(?<![A-Za-z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])", cell, text)


def expect(tokens):
    """읽은 자리 꼬리표 → 규칙을 적용한 기대값."""
    out = []
    for tag, v in tokens:
        if tag == "ref":
            out.append((tag, shift_text(v)))
        elif tag == "row1":  # 1부터 세는 줄 · 고정한 마지막 줄
            out.append((tag, r1(v)))
        elif tag == "row0":  # 0부터 세는 줄(그림 자리 · 페이지 나누기의 새 쪽 첫 줄)
            out.append((tag, r1(v + 1) - 1))
        else:
            out.append((tag, v))
    return sorted(out, key=repr)


# ------------------------------------------------------------------ XML 직접 읽기


class X:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.f = {n: z.read(n) for n in z.namelist()}

    def xml(self, n):
        return ET.fromstring(self.f[n])

    def rels(self, part):
        d, b = os.path.split(part)
        rp = f"{d}/_rels/{b}.rels" if d else f"_rels/{b}.rels"
        if rp not in self.f:
            return {}
        out = {}
        for e in ET.fromstring(self.f[rp]):
            t = e.get("Target")
            if e.get("TargetMode") != "External":
                t = t.lstrip("/") if t.startswith("/") else os.path.normpath(os.path.join(d, t)).replace("\\", "/")
            out[e.get("Id")] = (e.get("Type").rsplit("/", 1)[-1], t)
        return out

    def workbook(self):
        return [t for _, (ty, t) in self.rels("").items() if ty == "officeDocument"][0]

    def sheet(self):
        wbp = self.workbook()
        rels = self.rels(wbp)
        for s in self.xml(wbp).iter("{%s}sheet" % NS["m"]):
            if s.get("name") == T_SHEET:
                return rels[s.get(RID)][1]

    def sheet_xml(self):
        return self.xml(self.sheet())

    def related(self, part, typ):
        return [t for ty, t in self.rels(part).values() if ty == typ]

    def defined(self, name):
        out = []
        for d in self.xml(self.workbook()).iter("{%s}definedName" % NS["m"]):
            if d.get("name") == name:
                out.append(d.text)
        return out


def read_cf(x):
    out = []
    for cf in x.sheet_xml().iter("{%s}conditionalFormatting" % NS["m"]):
        out.append(("ref", cf.get("sqref")))
        for f in cf.iter("{%s}formula" % NS["m"]):
            out.append(("ref", f.text))
    return out


def read_dv(x):
    out = []
    for dv in x.sheet_xml().iter("{%s}dataValidation" % NS["m"]):
        out.append(("ref", dv.get("sqref")))
        for f in dv.iter("{%s}formula1" % NS["m"]):
            out.append(("ref", f.text))
    return out


def _drawings(x):
    return x.related(x.sheet(), "drawing")


def read_chart_refs(x):
    out = []
    for d in _drawings(x):
        for c in x.related(d, "chart"):
            for f in x.xml(c).iter("{%s}f" % NS["c"]):
                out.append(("ref", f.text))
    return out


def _anchor_rows(x, want):
    out = []
    for d in _drawings(x):
        for anc in x.xml(d):
            has = {e.tag.rsplit("}", 1)[-1] for e in anc}
            if want not in has:
                continue
            for side in ("from", "to"):
                for e in anc.findall(f"xdr:{side}/xdr:row", NS):
                    out.append(("row0", int(e.text)))
    return out


def read_chart_anchor(x):
    return _anchor_rows(x, "graphicFrame")


def read_pic_anchor(x):
    return _anchor_rows(x, "pic")


def read_print_area(x):
    return [("ref", v) for v in x.defined("_xlnm.Print_Area")]


def read_print_titles(x):
    return [("ref", v) for v in x.defined("_xlnm.Print_Titles")]


def read_named(x):
    return [("ref", v) for v in x.defined("직원범위")]


def read_hyperlink(x):
    out = []
    sp = x.sheet()
    rels = x.rels(sp)
    for h in x.sheet_xml().iter("{%s}hyperlink" % NS["m"]):
        out.append(("ref", h.get("ref")))
        loc = h.get("location") or (rels.get(h.get(RID), ("", ""))[1].lstrip("#") if h.get(RID) else None)
        if loc:
            out.append(("ref", loc))
    return out


def read_comment(x):
    out = []
    for c in x.related(x.sheet(), "comments"):
        for e in x.xml(c).iter("{%s}comment" % NS["m"]):
            out.append(("ref", e.get("ref")))
    return out


def read_sort(x):
    out = []
    for ss in x.sheet_xml().iter("{%s}sortCondition" % NS["m"]):
        out.append(("ref", ss.get("ref")))
    return out


def read_freeze(x):
    out = []
    for p in x.sheet_xml().iter("{%s}pane" % NS["m"]):
        if p.get("state") == "frozen" and p.get("ySplit"):
            y = int(float(p.get("ySplit")))
            out.append(("row1", y))
            out.append(("keep", p.get("topLeftCell") == f"A{y + 1}"))
    return out


def _rows_with(x, attr, val):
    return [("row1", int(r.get("r"))) for r in x.sheet_xml().iter("{%s}row" % NS["m"]) if r.get(attr) == val]


def read_height(x):
    return _rows_with(x, "ht", "33")


def read_hidden(x):
    return _rows_with(x, "hidden", "1")


def read_outline(x):
    return _rows_with(x, "outlineLevel", "1")


def read_breaks(x):
    return [("row0", int(b.get("id"))) for rb in x.sheet_xml().iter("{%s}rowBreaks" % NS["m"]) for b in rb]


def read_table(x):
    out = []
    for t in x.related(x.sheet(), "table"):
        out.append(("ref", x.xml(t).get("ref")))
    return out


def read_merge(x):
    return [("ref", m.get("ref")) for m in x.sheet_xml().iter("{%s}mergeCell" % NS["m"])
            if m.get("ref").startswith("E")]


def read_rich(x):
    """부분 굵게 한 글자 칸(꾸밈 유지)이 있는 줄."""
    sst = []
    for n in x.f:
        if n.endswith("sharedStrings.xml"):
            for si in x.xml(n).iter("{%s}si" % NS["m"]):
                sst.append(si.find("m:r", NS) is not None)
    out = []
    for c in x.sheet_xml().iter("{%s}c" % NS["m"]):
        rich = False
        if c.get("t") == "s":
            v = c.find("m:v", NS)
            rich = v is not None and sst[int(v.text)]
        elif c.get("t") == "inlineStr":
            rich = c.find("m:is/m:r", NS) is not None
        if rich:
            out.append(("row1", int(re.sub(r"[A-Z]+", "", c.get("r")))))
    return out


# ------------------------------------------------------------------ 부품 붙이기(시험 준비 — 시험 쪽 openpyxl)

POS = {  # 자리 이름 → (대표 줄, 범위 시작, 범위 끝)
    "합계 줄": (10, 9, 11),
    "기록 줄": (6, 5, 7),
    "머리 칸": (2, 1, 3),
    "걸침": (6, 6, 10),
}


def _chart(ws, r, a, b):
    from openpyxl.chart import BarChart, Reference, Series

    ch = BarChart()
    s = Series(Reference(ws, min_col=2, min_row=a, max_row=b), title_from_data=False)
    from openpyxl.chart.series import SeriesLabel, StrRef

    s.tx = SeriesLabel(strRef=StrRef(f"'{T_SHEET}'!$A${a}"))
    ch.series.append(s)
    ch.set_categories(Reference(ws, min_col=1, min_row=a, max_row=b))
    return ch


def _two(a, b, edit=None):
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor

    return TwoCellAnchor(editAs=edit, _from=AnchorMarker(col=5, row=a - 1), to=AnchorMarker(col=8, row=b))


def _png(folder):
    from PIL import Image as PI

    p = os.path.join(folder, "그림.png")
    PI.new("RGB", (12, 9), (200, 30, 30)).save(p)
    return p


def attach(part, wb, ws, folder, r, a, b):
    from openpyxl.comments import Comment
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.hyperlink import Hyperlink
    from openpyxl.worksheet.pagebreak import Break

    if part == "조건부 서식":
        ws.conditional_formatting.add(f"A{a}:B{b}", FormulaRule(formula=[f'$B{a}<>""'],
                                                                 fill=PatternFill("solid", fgColor="FFFF00")))
    elif part == "드롭다운(데이터 유효성)":
        dv = DataValidation(type="list", formula1=f"$C${a}:$C${b}")
        dv.add(f"B{a}:B{b}")
        ws.add_data_validation(dv)
    elif part == "차트 계열 · 이름표":
        ws.add_chart(_chart(ws, r, a, b), "H2")
    elif part == "차트 자리":
        ch = _chart(ws, 1, 1, 3)
        ch.anchor = _two(a, b)
        ws.add_chart(ch)
    elif part == "그림 자리":
        from openpyxl.drawing.image import Image

        ws.add_image(Image(_png(folder)), f"F{r}")
    elif part == "인쇄 영역":
        ws.print_area = f"A{a}:D{b}"
    elif part == "인쇄 제목":
        ws.print_title_rows = f"{a}:{b}"
    elif part == "하이퍼링크":
        c = ws[f"E{r}"]
        c.value = "링크"
        c.hyperlink = Hyperlink(ref=f"E{r}", location=f"'{T_SHEET}'!A{b}")
    elif part == "칸 메모":
        ws[f"E{r}"].comment = Comment("직원 메모", "직원")
    elif part == "자동 거르개 정렬":
        ws.auto_filter.add_sort_condition(f"B{a}:B{b}")
    elif part == "틀 고정":
        ws.freeze_panes = f"A{r + 1}"
    elif part == "줄 높이":
        ws.row_dimensions[r].height = 33
    elif part == "숨긴 줄":
        ws.row_dimensions[r].hidden = True
    elif part == "줄 묶음":
        for i in range(a, b + 1):
            ws.row_dimensions[i].outlineLevel = 1
    elif part == "페이지 나누기":
        ws.row_breaks.append(Break(id=r))
    elif part == "표":
        from openpyxl.worksheet.table import Table

        ws[f"F{a}"], ws[f"G{a}"] = "이름", "값"
        ws.add_table(Table(displayName="직원표", ref=f"F{a}:G{b}"))
    elif part == "합친 칸":
        ws.merge_cells(f"E{a}:E{b}")
    elif part == "정의된 이름":
        from openpyxl.workbook.defined_name import DefinedName

        wb.defined_names["직원범위"] = DefinedName("직원범위", attr_text=f"'{T_SHEET}'!$E${a}:$E${b}")
    elif part == "꾸민 글자":
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont

        ws[f"E{r}"] = CellRichText([TextBlock(InlineFont(b=True), "굵게"), " 보통"])
    else:
        raise KeyError(part)


MOVE_PARTS = {  # 부품 → (읽는 법, 붙일 자리들)
    "조건부 서식": (read_cf, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "드롭다운(데이터 유효성)": (read_dv, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "차트 계열 · 이름표": (read_chart_refs, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "차트 자리": (read_chart_anchor, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "그림 자리": (read_pic_anchor, ("합계 줄", "기록 줄", "머리 칸")),
    "인쇄 영역": (read_print_area, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "인쇄 제목": (read_print_titles, ("합계 줄", "기록 줄", "머리 칸")),
    "하이퍼링크": (read_hyperlink, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "칸 메모": (read_comment, ("합계 줄", "기록 줄", "머리 칸")),
    "자동 거르개 정렬": (read_sort, ("합계 줄", "기록 줄", "걸침")),
    "틀 고정": (read_freeze, ("합계 줄", "기록 줄", "머리 칸")),
    "줄 높이": (read_height, ("합계 줄", "기록 줄", "머리 칸")),
    "숨긴 줄": (read_hidden, ("합계 줄", "기록 줄", "머리 칸")),
    "줄 묶음": (read_outline, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "페이지 나누기": (read_breaks, ("합계 줄", "기록 줄", "머리 칸")),
    "표": (read_table, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "합친 칸": (read_merge, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "정의된 이름": (read_named, ("합계 줄", "기록 줄", "머리 칸", "걸침")),
    "꾸민 글자": (read_rich, ("합계 줄", "기록 줄", "머리 칸")),
}


# ------------------------------------------------------------------ 저장하면 지워지는 부품 붙이기(XML 을 직접 넣는다)

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _rewrite(path, edit):
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    edit(files)
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    os.replace(tmp, path)


def _add_ct(files, part=None, ext=None, ctype=""):
    s = files["[Content_Types].xml"].decode()
    if part:
        s = s.replace("</Types>", f'<Override PartName="/{part}" ContentType="{ctype}"/></Types>')
    if ext and f'Extension="{ext}"' not in s:
        s = s.replace("</Types>", f'<Default Extension="{ext}" ContentType="{ctype}"/></Types>')
    files["[Content_Types].xml"] = s.encode()


def _add_rel(files, part, rid, typ, target):
    d, b = os.path.split(part)
    rp = f"{d}/_rels/{b}.rels"
    s = files.get(rp, b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="%s"></Relationships>'
                  % NS["pr"].encode()).decode()
    s = s.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{typ}" Target="{target}"/></Relationships>')
    files[rp] = s.encode()


def _sheet_insert(files, before_tags, snippet):
    """시트 XML 에서 before_tags 중 처음 나오는 요소 앞(없으면 끝)에 snippet 을 넣는다(스키마 차례 지킴)."""
    sp = "xl/worksheets/sheet1.xml"
    s = files[sp].decode()
    pos = len(s) - len("</worksheet>")
    for t in before_tags:
        m = re.search(r"<%s[ />]" % t, s)
        if m:
            pos = min(pos, m.start())
    files[sp] = (s[:pos] + snippet + s[pos:]).encode()


RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
SP_TEXTBOX = ('<xdr:twoCellAnchor xmlns:xdr="%(xdr)s" xmlns:a="%(a)s"><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0'
              '</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col>'
              '<xdr:colOff>0</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" '
              'textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="TextBox 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>'
              '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/>'
              '<a:p><a:r><a:t>오늘 합계 확인함</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/>'
              '</xdr:twoCellAnchor>') % NS
SP_RECT = SP_TEXTBOX.replace('<xdr:cNvSpPr txBox="1"/>', "<xdr:cNvSpPr/>").replace(
    '<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>오늘 합계 확인함</a:t></a:r></a:p></xdr:txBody>', "").replace(
    'name="TextBox 1"', 'name="Rectangle 1"').replace('prst="rect"', 'prst="ellipse"')


def _drawing_with(files, anchor):
    sp = "xl/worksheets/sheet1.xml"
    rels = files.get("xl/worksheets/_rels/sheet1.xml.rels", b"").decode()
    m = re.search(r'Target="([^"]*drawing[^"]*\.xml)"', rels)
    if m:
        dp = m.group(1).lstrip("/") if m.group(1).startswith("/") else os.path.normpath(
            os.path.join("xl/worksheets", m.group(1)))
        s = files[dp].decode()
        files[dp] = re.sub(r"(</(?:xdr:)?wsDr>)", anchor.replace("\\", "\\\\") + r"\1", s).encode()
        return
    dp = "xl/drawings/drawing99.xml"
    files[dp] = (f'<?xml version="1.0" encoding="UTF-8"?><xdr:wsDr xmlns:xdr="{NS["xdr"]}" xmlns:a="{NS["a"]}">'
                 f"{anchor}</xdr:wsDr>").encode()
    _add_ct(files, part=dp, ctype="application/vnd.openxmlformats-officedocument.drawing+xml")
    _add_rel(files, sp, "rIdD99", RT + "drawing", "/" + dp)
    _sheet_insert(files, ("legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
                          "webPublishItems", "tableParts", "extLst"), f'<drawing xmlns:r="{NS["r"]}" r:id="rIdD99"/>')


def _ext(uri, body):
    return (f'<extLst><ext uri="{uri}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
            f'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">{body}</ext></extLst>')


SPARK = _ext("{05C60535-1F16-4fd2-B633-F4F36F0B64E0}",
             '<x14:sparklineGroups><x14:sparklineGroup displayEmptyCellsAs="gap"><x14:colorSeries rgb="FF376092"/>'
             f"<x14:sparklines><x14:sparkline><xm:f>'{T_SHEET}'!B9:B11</xm:f><xm:sqref>E9</xm:sqref></x14:sparkline>"
             "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups>")
DATABAR = _ext("{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
               '<x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" '
               'id="{00000000-0000-0000-0000-000000000001}"><x14:dataBar minLength="0" maxLength="100">'
               '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>'
               "<xm:sqref>B9:B11</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings>")
XL_COMMENT_RUNS = ('<text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/>'
                   '</rPr><t>직원:</t></r><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/>'
                   '<family val="2"/></rPr><t xml:space="preserve">\n합계 확인</t></r></text>')
CHECKBOX = ('<v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel" '
            'id="_x0000_s2049" type="#_x0000_t201" style="position:absolute;width:60pt;height:15pt">'
            '<x:ClientData ObjectType="Checkbox"><x:Anchor>5, 0, 9, 0, 6, 0, 9, 18</x:Anchor><x:Checked>1</x:Checked>'
            "</x:ClientData></v:shape>")


def refuse_attach(part, path, folder):
    """반환: 까닭 문구에 들어 있어야 할 부품 이름."""
    import openpyxl

    if part == "글상자":
        _rewrite(path, lambda f: _drawing_with(f, SP_TEXTBOX))
        return "글상자"
    if part == "도형":
        _rewrite(path, lambda f: _drawing_with(f, SP_RECT))
        return "도형"
    if part == "스파크라인":
        _rewrite(path, lambda f: _sheet_insert(f, (), SPARK))
        return "스파크라인"
    if part == "새 꼴 조건부 서식(데이터 막대)":
        _rewrite(path, lambda f: _sheet_insert(f, (), DATABAR))
        return "조건부 서식"
    if part == "인쇄 설정(프린터 정보)":
        def ed(f):
            f["xl/printerSettings/printerSettings1.bin"] = b"\x00" * 64
            _add_ct(f, ext="bin", ctype="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings")
            _add_rel(f, "xl/worksheets/sheet1.xml", "rIdP1", RT + "printerSettings",
                     "../printerSettings/printerSettings1.bin")
        _rewrite(path, ed)
        return "인쇄 설정"
    if part in ("차트 스타일", "크기 고정 그림이 끼운 줄에 걸침", "위치 고정 차트가 끼운 줄 아래"):
        wb = openpyxl.load_workbook(path)
        ws = wb[T_SHEET]
        if part == "크기 고정 그림이 끼운 줄에 걸침":
            from openpyxl.drawing.image import Image

            img = Image(_png(folder))
            img.anchor = _two(6, 10, "oneCell")
            ws.add_image(img)
        else:
            ch = _chart(ws, 10, 9, 11)
            ch.anchor = _two(9, 11, "absolute" if part.startswith("위치") else None)
            ws.add_chart(ch)
        wb.save(path)
        if part == "차트 스타일":
            def ed(f):
                cp = [n for n in f if re.fullmatch(r"xl/charts/chart\d+\.xml", n)][0]
                f["xl/charts/style1.xml"] = (b'<?xml version="1.0" encoding="UTF-8"?><cs:chartStyle '
                                             b'xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"'
                                             b' id="201"/>')
                _add_ct(f, part="xl/charts/style1.xml", ctype="application/vnd.ms-office.chartstyle+xml")
                _add_rel(f, cp, "rIdS1", "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
                         "style1.xml")
            _rewrite(path, ed)
            return "차트 스타일"
        return "크기 고정" if part.startswith("크기") else "위치 고정"
    if part in ("엑셀에서 단 메모(꾸민 글)", "늘 보이는 메모", "양식 컨트롤(체크 상자)"):
        from openpyxl.comments import Comment

        wb = openpyxl.load_workbook(path)
        wb[T_SHEET]["E10"].comment = Comment("합계 확인", "직원")
        wb.save(path)

        def ed(f):
            cp = [n for n in f if "comments" in n and n.endswith(".xml")][0]
            vp = [n for n in f if n.endswith(".vml")][0]
            if part.startswith("엑셀"):
                f[cp] = re.sub(rb"<text>.*?</text>", XL_COMMENT_RUNS.encode(), f[cp], flags=re.S)
            elif part.startswith("늘"):
                f[vp] = re.sub(rb"<(\w+):AutoFill>", rb"<\1:Visible/><\1:AutoFill>", f[vp]).replace(
                    b"visibility:hidden", b"visibility:visible")
            else:
                f[vp] = f[vp].replace(b"</xml>", CHECKBOX.encode() + b"</xml>")
        _rewrite(path, ed)
        return {"엑": "메모", "늘": "메모", "양": "양식 컨트롤"}[part[0]]
    raise KeyError(part)


REFUSE_PARTS = ("글상자", "도형", "스파크라인", "새 꼴 조건부 서식(데이터 막대)", "인쇄 설정(프린터 정보)", "차트 스타일",
                "엑셀에서 단 메모(꾸민 글)", "늘 보이는 메모", "양식 컨트롤(체크 상자)", "크기 고정 그림이 끼운 줄에 걸침",
                "위치 고정 차트가 끼운 줄 아래")


# ------------------------------------------------------------------ 기록 흐름


def base(core, d):
    clk = Clock()
    s = core.RecordStore(d, clock=clk)
    for k, m, b in [("대출", "M1", "MB1"), ("반납", "M2", "MB2"), ("대출", "M3", "MB3")]:
        s.record(k, m, b)
        clk.tick()
    p = os.path.join(d, xname(DAY))
    import openpyxl

    ws = openpyxl.load_workbook(p)[T_SHEET]
    assert [ws[f"A{i}"].value for i in (9, 10, 11)] == ["전체 건수", "대출 건수", "반납 건수"], "합계 자리가 예상과 다름"
    assert ws["A8"].value is None and ws["A7"].value, "기록 자리가 예상과 다름"
    return s, clk, p


def totals(p):
    import openpyxl

    ws = openpyxl.load_workbook(p)[T_SHEET]
    out = {}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.strip() in ("전체 건수", "대출 건수", "반납 건수"):
                out[c.value.strip()] = ws.cell(c.row, c.column + 1).value
    return out


def move(core, root):
    import openpyxl

    out = {}
    for part, (reader, positions) in MOVE_PARTS.items():
        for pos in positions:
            name = f"{part} · {pos}"
            d = os.path.join(root, f"m{len(out)}")
            os.makedirs(d)
            try:
                s, clk, p = base(core, d)
                wb = openpyxl.load_workbook(p)
                attach(part, wb, wb[T_SHEET], d, *POS[pos])
                wb.save(p)
                before = reader(X(p))
                assert before, "붙인 부품을 읽지 못함(시험 준비 실패)"
                want = expect(before)
                r = s.record("대출", "M4", "MB4")
                got = sorted(reader(X(p)), key=repr)
                tot = totals(p)
                ok = bool(r.excel_saved) and got == want and tot == {"전체 건수": 4, "대출 건수": 3, "반납 건수": 1}
                out[name] = {"ok": ok, "before": before, "want": want, "got": got, "saved": bool(r.excel_saved),
                             "error": "" if r.excel_saved else r.error}
            except Exception as exc:
                out[name] = {"ok": False, "exception": f"{exc.__class__.__name__}: {exc}"}
    return out


def refuse(core, root):
    from library_emergency.messages import status_for

    out = {}
    for part in REFUSE_PARTS:
        d = os.path.join(root, f"r{len(out)}")
        os.makedirs(d)
        try:
            s, clk, p = base(core, d)
            key = refuse_attach(part, p, d)
            before = sha(p)
            r = s.record("대출", "M4", "MB4")
            text = status_for(r)[0]
            in_journal = any(row.book == "MB4" for row in core.read_journal(os.path.join(d, jname(DAY))).rows)
            res = {"saved": bool(r.excel_saved), "file_unchanged": sha(p) == before, "in_journal": in_journal,
                   "status": text, "key": key}
            res["ok"] = (not res["saved"]) and res["file_unchanged"] and in_journal and text.startswith(
                "일지에만 기록됨") and key in text
            out[part] = res
        except Exception as exc:
            out[part] = {"ok": False, "exception": f"{exc.__class__.__name__}: {exc}"}
    return out


# ------------------------------------------------------------------ 실제 프로그램이 만든 시험 자료

REAL = os.path.join(HERE, "fixtures")


def last_record_row(p):
    import openpyxl

    wb = openpyxl.load_workbook(p, read_only=True)
    try:  # 읽기 전용은 파일을 열어 둔다 — 닫지 않으면 윈도우에서 판의 저장이 「열려 있음」으로 막힌다
        last = 4
        for i, row in enumerate(wb[T_SHEET].iter_rows(min_row=5, max_col=4, values_only=True), start=5):
            a, bb, _, d = (list(row) + [None] * 4)[:4]
            if a not in (None, "") and isinstance(bb, str) and bb.strip() in ("대출", "반납") and d not in (None, ""):
                last = i
    finally:
        wb.close()
    return last


def real(core, root):
    """tests/fixtures/real_*.xlsx — 프로그램(Numbers)이 내보낸 파일을 그날 엑셀 파일 자리에 두고 다음 기록."""
    from library_emergency.messages import status_for

    with open(os.path.join(REAL, "real_fixtures.json"), encoding="utf-8") as fh:
        spec = json.load(fh)
    out = {}
    for item in spec["files"]:
        d = os.path.join(root, f"x{len(out)}")
        os.makedirs(d)
        try:
            src = os.path.join(REAL, item["file"])
            assert sha(src) == item["sha256"], "시험 자료 지문이 다름"
            clk = Clock()
            s = core.RecordStore(d, clock=clk)
            p = os.path.join(d, xname(DAY))
            shutil.copyfile(src, p)
            # 이 시험 자료에는 일지가 없다 — 엑셀 속 기록 줄은 손으로 넣은 줄로 읽힌다(판정 규칙 「나」)
            # 끼울 자리는 시험이 직접 센다: 구분 칸이 대출/반납인 마지막 줄 다음
            global AT
            AT = last_record_row(p) + 1
            before = sha(p)
            x0 = X(p)
            readers = {n: MOVE_PARTS[n][0] for n in item.get("parts", [])}
            want = {n: expect(f(x0)) for n, f in readers.items()}
            r = s.record("대출", "M4", "MB4")
            text = status_for(r)[0]
            in_journal = any(row.book == "MB4" for row in core.read_journal(os.path.join(d, jname(DAY))).rows)
            res = {"saved": bool(r.excel_saved), "file_unchanged": sha(p) == before, "in_journal": in_journal,
                   "status": text, "expect": item["expect"]}
            if item["expect"] == "refuse":
                res["ok"] = (not res["saved"]) and res["file_unchanged"] and in_journal and text.startswith(
                    "일지에만 기록됨") and item["key"] in text
            else:
                x1 = X(p)
                got = {n: sorted(f(x1), key=repr) for n, f in readers.items()}
                res["want"], res["got"] = want, got
                res["ok"] = res["saved"] and got == want and in_journal
            res["at"] = AT
            out[item["file"]] = res
        except Exception as exc:
            out[item["file"]] = {"ok": False, "exception": f"{exc.__class__.__name__}: {exc}"}
    return out


# ------------------------------------------------------------------ 1-1 표: openpyxl 자체 실측


def inventory(core, root):
    """부품마다 openpyxl 3.1.5 로 ① 읽고 그대로 저장 ② 8행에 1줄 끼우고 저장 — 결과를 XML 로 직접 잰다."""
    import openpyxl

    rows = []
    for part, (reader, _) in MOVE_PARTS.items():
        d = os.path.join(root, f"i{len(rows)}")
        os.makedirs(d)
        s, clk, p = base(core, d)
        wb = openpyxl.load_workbook(p)
        attach(part, wb, wb[T_SHEET], d, *POS["합계 줄"])
        wb.save(p)
        before = sorted(reader(X(p)), key=repr)
        p1, p2 = p + ".1.xlsx", p + ".2.xlsx"
        openpyxl.load_workbook(p).save(p1)
        kept = sorted(reader(X(p1)), key=repr) == before
        wb = openpyxl.load_workbook(p)
        wb[T_SHEET].insert_rows(AT, K)
        wb.save(p2)
        moved = sorted(reader(X(p2)), key=repr) == expect(before)
        p3 = p + ".3.xlsx"
        openpyxl.load_workbook(p, rich_text=True).save(p3)  # 이 프로그램이 v2.0.4 부터 쓰는 읽기 방식
        kept_rich = sorted(reader(X(p3)), key=repr) == before
        rows.append({"part": part, "kept": kept, "kept_rich_text": kept_rich, "moved": moved, "before": before,
                     "after_insert": sorted(reader(X(p2)), key=repr)})
    for part in REFUSE_PARTS:
        d = os.path.join(root, f"i{len(rows)}")
        os.makedirs(d)
        s, clk, p = base(core, d)
        refuse_attach(part, p, d)
        p1 = p + ".1.xlsx"
        import warnings

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            openpyxl.load_workbook(p).save(p1)
        kept = _marker(part, X(p)) == _marker(part, X(p1))
        rows.append({"part": part, "kept": kept, "moved": None, "marker_before": _marker(part, X(p)),
                     "marker_after": _marker(part, X(p1)), "warnings": [str(x.message)[:120] for x in w]})
    return rows


def _marker(part, x):
    """저장하면 지워지는 부품의 표지 — 파일 속에 그 부품을 이루는 XML 이 몇 개 있는가."""
    alltext = b"".join(v for n, v in x.f.items() if n.endswith((".xml", ".vml", ".rels", ".bin")))
    pats = {"글상자": rb"txBox=\"1\"", "도형": rb"prst=\"ellipse\"", "스파크라인": rb"sparkline>",
            "새 꼴 조건부 서식(데이터 막대)": rb"x14:dataBar", "인쇄 설정(프린터 정보)": rb"printerSettings1\.bin",
            "차트 스타일": rb"chartStyle", "엑셀에서 단 메모(꾸민 글)": rb"<rFont val=\"Tahoma\"",
            "늘 보이는 메모": rb":Visible/>", "양식 컨트롤(체크 상자)": rb"ObjectType=\"Checkbox\"",
            "크기 고정 그림이 끼운 줄에 걸침": rb"editAs=\"oneCell\"", "위치 고정 차트가 끼운 줄 아래": rb"editAs=\"absolute\""}
    return len(re.findall(pats[part], alltext))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pkg", required=True)
    ap.add_argument("--folder", required=True)
    ap.add_argument("--mode", required=True)
    a = ap.parse_args()
    core = load(a.pkg)
    fn = {"move": move, "refuse": refuse, "real": real, "inventory": inventory}[a.mode]
    print("RESULT " + json.dumps(fn(core, a.folder), ensure_ascii=False, default=str), flush=True)


if __name__ == "__main__":
    main()
