"""직원이 엑셀에서 붙인 부품 지키기 (v2.0.4 · 고침 지시 2026-11-07 FIX3).

줄을 끼울 때(core.write_rows_and_totals):
  move_parts(wb, ws, at, k)  — openpyxl 이 읽고 저장하지만 줄 끼우기 때 옮기지 않는 부품을
                               엑셀에서 같은 자리에 줄을 끼웠을 때와 같게 옮긴다(1-2).
                               끼운 자리 이상은 k 줄 아래로 · 위는 그대로 · 범위는 양 끝 따로.
저장 앞뒤(core.RecordStore._flush_locked):
  guard = Guard(원래 파일)     — 닫힌 목록 검문: 이 프로그램이 저장하면 지워지거나, 옮기는 법을 모르는
                               부품이 하나라도 있으면 PartError — 그 파일에는 쓰지 않는다(1-3).
  guard.check_after(새 파일, 시트, at, k)
                             — 저장한 새 파일을 다시 열어, 원래 파일에서 따로 센 기대값과 맞춰 본다.
                               지워진 부품 · 기대와 다른 자리가 있으면 PartError(바꿔 끼우지 않는다).

검문은 openpyxl 을 쓰지 않고 파일 속 XML 을 직접 읽는다(만드는 쪽의 해석을 기대값으로 쓰지 않는다).
"""

from __future__ import annotations

import hashlib
import posixpath
import re
import zipfile
from collections import Counter
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from .core import ExcelFormatError, _shift_ref, _split_sheet

# ------------------------------------------------------------------ 이름표

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_X = "urn:schemas-microsoft-com:office:excel"

SHORT = {NS_MAIN: "", NS_A: "a:", NS_C: "c:", NS_XDR: "xdr:", NS_MC: "mc:", NS_R: "r:",
         "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties": "",
         "http://schemas.openxmlformats.org/package/2006/metadata/core-properties": "cp:",
         "http://purl.org/dc/elements/1.1/": "dc:", "http://purl.org/dc/terms/": "dcterms:",
         "http://schemas.microsoft.com/office/spreadsheetml/2014/revision": "xr:",
         "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac": "x15ac:"}

CT = "application/vnd.openxmlformats-officedocument."
#: 이 프로그램이 다룰 줄 아는 부품 종류(닫힌 목록). 그 밖의 종류가 파일에 있으면 쓰지 않는다.
KINDS = {
    CT + "spreadsheetml.sheet.main+xml": "workbook",
    CT + "spreadsheetml.worksheet+xml": "worksheet",
    CT + "spreadsheetml.styles+xml": "styles",
    CT + "spreadsheetml.sharedStrings+xml": "sharedStrings",
    CT + "theme+xml": "theme",
    "application/vnd.openxmlformats-package.core-properties+xml": "core",
    CT + "extended-properties+xml": "app",
    CT + "custom-properties+xml": "custom",
    CT + "spreadsheetml.calcChain+xml": "calcChain",
    CT + "drawing+xml": "drawing",
    CT + "drawingml.chart+xml": "chart",
    CT + "spreadsheetml.comments+xml": "comments",
    CT + "vmlDrawing": "vml",
    CT + "spreadsheetml.table+xml": "table",
    "application/vnd.openxmlformats-package.relationships+xml": "rels",
    "image/png": "image",
    "image/jpeg": "image",
    "image/gif": "image",
}

#: 다룰 줄 모르는 부품의 이름(까닭 문구용). 목록에 없는 것은 「알 수 없는 부품」.
REFUSED_KIND_NAMES = {
    CT + "spreadsheetml.printerSettings": "인쇄 설정(프린터 정보)",
    "application/vnd.ms-office.chartstyle+xml": "차트 스타일",
    "application/vnd.ms-office.chartcolorstyle+xml": "차트 색 구성",
    CT + "drawingml.chartshapes+xml": "차트 안 도형",
    "application/vnd.ms-excel.threadedcomments+xml": "대화형 메모",
    "application/vnd.ms-excel.person+xml": "대화형 메모",
    CT + "spreadsheetml.pivotTable+xml": "피벗 테이블",
    CT + "spreadsheetml.pivotCacheDefinition+xml": "피벗 테이블",
    CT + "spreadsheetml.pivotCacheRecords+xml": "피벗 테이블",
    CT + "spreadsheetml.externalLink+xml": "다른 파일 연결",
    "application/vnd.ms-excel.slicer+xml": "슬라이서",
    "application/vnd.ms-excel.slicerCache+xml": "슬라이서",
    "application/vnd.ms-excel.timeline+xml": "시간 표시 막대",
    "application/vnd.ms-excel.timelineCache+xml": "시간 표시 막대",
    "application/vnd.ms-excel.controlproperties+xml": "양식 컨트롤",
    "application/vnd.ms-office.activeX+xml": "ActiveX 컨트롤",
    "application/vnd.ms-office.activeX": "ActiveX 컨트롤",
    CT + "oleObject": "삽입한 개체",
    CT + "spreadsheetml.chartsheet+xml": "차트 시트",
    CT + "spreadsheetml.dialogsheet+xml": "대화 상자 시트",
    "application/vnd.ms-excel.macrosheet+xml": "매크로 시트",
    "application/vnd.ms-office.vbaProject": "매크로",
    CT + "spreadsheetml.connections+xml": "외부 데이터 연결",
    CT + "spreadsheetml.queryTable+xml": "외부 데이터 연결",
    CT + "spreadsheetml.sheetMetadata+xml": "동적 배열 수식 · 칸 속 그림",
    "application/vnd.ms-excel.rdrichvalue+xml": "칸 속 그림",
    "application/vnd.ms-excel.rdrichvaluestructure+xml": "칸 속 그림",
    "application/vnd.ms-excel.rdrichvaluetypes+xml": "칸 속 그림",
    "application/vnd.ms-excel.rdarray+xml": "칸 속 그림",
    "application/vnd.ms-excel.rdsupportingpropertybag+xml": "칸 속 그림",
    "application/vnd.ms-excel.rdsupportingpropertybagstructure+xml": "칸 속 그림",
    CT + "spreadsheetml.volatileDependencies+xml": "실시간 데이터 수식",
    CT + "spreadsheetml.revisionHeaders+xml": "변경 내용 추적",
    CT + "spreadsheetml.revisionLog+xml": "변경 내용 추적",
    CT + "spreadsheetml.userNames+xml": "통합 문서 공유",
    CT + "drawingml.diagramData+xml": "스마트아트",
    CT + "drawingml.diagramLayout+xml": "스마트아트",
    CT + "drawingml.diagramStyle+xml": "스마트아트",
    CT + "drawingml.diagramColors+xml": "스마트아트",
    "application/vnd.ms-office.drawingml.diagramDrawing+xml": "스마트아트",
    "application/vnd.ms-office.webextension+xml": "추가 기능",
    "application/vnd.ms-office.webextensiontaskpanes+xml": "추가 기능",
    "application/vnd.ms-excel.attachedToolbars": "도구 모음",
    "image/x-emf": "그림(EMF)",
    "image/x-wmf": "그림(WMF)",
    "image/svg+xml": "그림(SVG)",
    "image/tiff": "그림(TIFF)",
    "image/bmp": "그림(BMP)",
    "application/vnd.ms-excel.customDataProperties+xml": "사용자 데이터",
    "application/vnd.ms-excel.featurepropertybag+xml": "칸 체크 상자 등 새 기능",
    "application/vnd.ms-excel.namedsheetviews+xml": "시트 보기",
    "application/vnd.ms-office.classificationlabels+xml": "민감도 레이블",
    CT + "customXmlProperties+xml": "사용자 XML",
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml": "사용자 XML",
}

#: 시트에 있어도 되는 것(옮기는 법을 알거나 줄과 상관없는 것). 그 밖은 쓰지 않는다.
SHEET_CHILDREN = {"sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection",
                  "autoFilter", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
                  "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
                  "colBreaks", "drawing", "legacyDrawing", "tableParts"}
SHEET_CHILD_NAMES = {
    "sheetCalcPr": "시트 계산 설정", "sortState": "시트 정렬 상태", "dataConsolidate": "통합",
    "customSheetViews": "사용자 지정 보기", "scenarios": "시나리오", "protectedRanges": "편집 허용 범위",
    "cellWatches": "조사식 창", "ignoredErrors": "오류 무시 표시", "smartTags": "스마트 태그",
    "legacyDrawingHF": "머리글 · 바닥글 그림", "picture": "시트 배경 그림", "oleObjects": "삽입한 개체",
    "controls": "양식 컨트롤", "webPublishItems": "웹 게시", "AlternateContent": "양식 컨트롤 · 개체",
    "extLst": "엑셀 새 기능 부품",
}
WORKBOOK_CHILDREN = {"fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
                     "definedNames", "calcPr", "AlternateContent", "revisionPtr", "extLst", "fileRecoveryPr"}
WORKBOOK_CHILD_NAMES = {"externalReferences": "다른 파일 연결", "pivotCaches": "피벗 테이블",
                        "customWorkbookViews": "사용자 지정 보기", "functionGroups": "사용자 함수 묶음",
                        "smartTagPr": "스마트 태그", "smartTagTypes": "스마트 태그", "webPublishing": "웹 게시",
                        "webPublishObjects": "웹 게시", "oleSize": "삽입한 개체"}

#: 엑셀 새 기능 꾸러미(extLst/ext uri) 가운데 있어도 되는 것 — 뜻이 없는 엑셀 꼬리표뿐.
EXT_OK = {
    "workbook": {"{140A7094-0E35-4892-8432-C4D2E57EDEB5}",  # x15 workbookPr(차트 추적 기준)
                 "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}",  # 계산 기능 목록
                 "{79F54976-1DA5-4618-B147-4CDE4B953A38}"},  # x14 workbookPr
    "styles": {"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",  # 기본 슬라이서 스타일 이름
               "{9260A510-F301-46a8-8635-F512D64BE5F5}"},  # 기본 시간 막대 스타일 이름
    "worksheet": set(),
    "drawing": {"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}",  # 개체 생성 번호
                "{28A0092B-C50C-407E-A947-70E740481C1C}"},  # 그림 압축 안 함 표시
    "chart": {"{C3380CC4-5D6E-409C-BE32-E72D297353CC}"},  # 계열 고유 번호
}
EXT_NAMES = {
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}": "조건부 서식(데이터 막대 · 아이콘 등 새 꼴)",
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}": "다른 시트를 가리키는 드롭다운",
    "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}": "스파크라인",
    "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}": "슬라이서",
    "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}": "슬라이서",
    "{7E03D99C-DC04-49d9-9315-930204A7B6E9}": "시간 표시 막대",
    "{01252117-D84E-4E92-8308-4BE1C098FCBB}": "추가 기능",
    "{96DAC541-7B7A-43D3-8B79-37D633B846F1}": "그림(SVG)",
    "{63B3BB69-23CF-44E3-9099-C40C66FF867C}": "양식 컨트롤",
    "{02D57815-91ED-43cb-92C2-25804820EDAC}": "차트(숨긴 계열 · 범위 이름표)",
    "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}": "차트(셀 값 데이터 레이블)",
    "{56B9EC1D-385E-4148-901F-78D8002777C0}": "차트(#N/A 표시 설정)",
    "{FCE2AD5D-F65C-4FA6-A056-5C36A1767C68}": "데이터 모델",
    "{BBE1A952-AA13-448e-AADC-164F8A28A991}": "슬라이서",
    "{46BE6895-7355-4a93-B00E-2C351335B9C9}": "슬라이서",
    "{876F7934-8845-4945-9796-88D515C7AA90}": "피벗 테이블",
}


class PartError(ExcelFormatError):
    """직원이 붙인 부품을 이 프로그램이 지키지 못한다 — 그 파일에는 쓰지 않는다."""


def lost(name: str) -> PartError:
    return PartError(f"{name}: 이 프로그램이 저장하면 지워집니다")


def unmovable(name: str) -> PartError:
    return PartError(f"{name}: 새 줄에 맞춰 옮길 수 없습니다")


def _short(tag: str) -> str:
    if tag.startswith("{"):
        uri, local = tag[1:].split("}", 1)
        return SHORT.get(uri, "{" + uri + "}") + local
    return tag


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


# ------------------------------------------------------------------ 옮기기(1-2)


def _row_shift(n: int, at: int, k: int) -> int:
    return n + k if n >= at else n


def shift_sqref(sq, at: int, k: int) -> str:
    return " ".join(_shift_ref(x, at, k) for x in str(sq).split())


def shift_bare(text: str, home: str, target: str, at: int, k: int, name: str) -> str:
    """「=」 없는 수식 글자(조건부 서식 · 드롭다운 · 차트 · 이름)를 고친다."""
    from .core import _shift_formula_text

    if text is None or text == "":
        return text
    try:
        return _shift_formula_text("=" + str(text), home, target, at, k)[1:]
    except Exception:
        raise unmovable(name)


def shift_location(loc: str, home: str, target: str, at: int, k: int) -> str:
    """하이퍼링크가 가리키는 파일 안 자리(「'시트'!A10」 · 「A10」 · 이름)."""
    try:
        sheet, ref = _split_sheet(loc)
    except ValueError:
        raise unmovable("하이퍼링크(가리키는 자리)")
    if (sheet if sheet is not None else home) != target:
        return loc
    try:
        new = _shift_ref(ref, at, k)
    except ValueError:
        if sheet is None and re.fullmatch(r"[A-Za-z_\\\u0080-￿][\w.\u0080-￿]*", ref):
            return loc  # 정의된 이름 — 이름 쪽에서 고친다
        raise unmovable("하이퍼링크(가리키는 자리)")
    return loc[: len(loc) - len(ref)] + new


def _shift_anchor(anchor, at: int, k: int, name: str) -> None:
    """그림 · 차트가 놓인 자리. 엑셀의 「셀에 맞춰 이동 · 크기 조정」(editAs) 별로."""
    from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor, OneCellAnchor, TwoCellAnchor

    if isinstance(anchor, AbsoluteAnchor):
        return  # 칸이 아니라 쪽 위 자리로 놓였다 — 엑셀도 옮기지 않는다
    if isinstance(anchor, OneCellAnchor):
        anchor._from.row = _row_shift(anchor._from.row + 1, at, k) - 1
        return
    if isinstance(anchor, TwoCellAnchor):
        top, bottom = anchor._from.row + 1, anchor.to.row + 1
        mode = anchor.editAs or "twoCell"
        if mode == "twoCell":
            anchor._from.row = _row_shift(top, at, k) - 1
            anchor.to.row = _row_shift(bottom, at, k) - 1
        elif mode == "oneCell":
            if top >= at:
                anchor._from.row += k
                anchor.to.row += k
            elif bottom >= at:
                raise unmovable(name + "(크기 고정 · 끼운 줄에 걸침)")
        else:
            if bottom >= at:
                raise unmovable(name + "(위치 고정)")
        return
    raise unmovable(name)


def _walk_chart_refs(obj, home: str, target: str, at: int, k: int, seen: set) -> None:
    from openpyxl.descriptors.serialisable import Serialisable

    if id(obj) in seen:
        return
    seen.add(id(obj))
    if isinstance(obj, (list, tuple)):
        for x in obj:
            _walk_chart_refs(x, home, target, at, k, seen)
        return
    if not isinstance(obj, Serialisable) and not hasattr(obj, "__dict__"):
        return
    for key, val in list(vars(obj).items()):
        if key == "f" and isinstance(val, str):
            setattr(obj, key, shift_bare(val, home, target, at, k, "차트(계열 · 이름표)"))
        elif key in ("anchor", "_parent"):
            continue
        elif isinstance(val, (list, tuple)) or isinstance(val, Serialisable):
            _walk_chart_refs(val, home, target, at, k, seen)


def move_parts(wb, ws, at: int, k: int) -> None:
    """ws 의 at 자리에 k 줄을 끼운 직후 부른다(칸 · 합친 칸 · 수식 · 이름은 core 가 이미 옮겼다)."""
    from openpyxl.formatting.formatting import ConditionalFormattingList
    from openpyxl.worksheet.cell_range import MultiCellRange

    T = ws.title
    for sh in wb.worksheets:
        mine = sh.title == T
        # 조건부 서식 — 자리와 규칙 안 수식
        old = sh.conditional_formatting
        if len(old):
            new = ConditionalFormattingList()
            for cf in old:
                rng = shift_sqref(cf.sqref, at, k) if mine else str(cf.sqref)
                for rule in cf.rules:
                    if rule.formula:
                        rule.formula = [shift_bare(f, sh.title, T, at, k, "조건부 서식") for f in rule.formula]
                    for sub in (rule.colorScale, rule.dataBar, rule.iconSet):
                        for cfvo in (getattr(sub, "cfvo", None) or []):
                            if cfvo.val is not None and cfvo.type in ("num", "percent", "percentile", "formula"):
                                cfvo.val = shift_bare(str(cfvo.val), sh.title, T, at, k, "조건부 서식")
                    new.add(rng, rule)
            sh.conditional_formatting = new
        # 데이터 유효성(드롭다운) — 자리와 목록 수식
        for dv in sh.data_validations.dataValidation:
            if mine:
                dv.sqref = MultiCellRange(shift_sqref(dv.sqref, at, k))
            for a in ("formula1", "formula2"):
                v = getattr(dv, a)
                if v:
                    setattr(dv, a, shift_bare(v, sh.title, T, at, k, "드롭다운(데이터 유효성)"))
        # 차트 — 계열 · 이름표가 가리키는 자리, 차트가 놓인 자리
        for ch in sh._charts:
            _walk_chart_refs(ch, sh.title, T, at, k, set())
            if mine:
                _shift_anchor(ch.anchor, at, k, "차트")
        # 그림이 놓인 자리
        if mine:
            for img in sh._images:
                _shift_anchor(img.anchor, at, k, "그림")
        # 하이퍼링크 — 붙은 칸, 가리키는 파일 안 자리
        for c in list(sh._cells.values()):
            h = c.hyperlink
            if h is None:
                continue
            if mine and h.ref:
                h.ref = shift_sqref(h.ref, at, k)
            if h.location:
                h.location = shift_location(h.location, sh.title, T, at, k)
            if h.target and h.target.startswith("#"):
                h.target = "#" + shift_location(h.target[1:], sh.title, T, at, k)
        if not mine:
            continue
        # 줄 높이 · 숨긴 줄 · 줄 묶음(줄 속성 전체)
        rd = sh.row_dimensions
        keys = sorted((i for i in list(rd.keys()) if i >= at), reverse=True)
        items = [(i, rd[i]) for i in keys]
        for i in keys:
            del rd[i]
        for i, d in items:
            d.index = i + k
            rd[i + k] = d
        # 페이지 나누기(id 는 새 쪽이 시작하는 줄의 0부터 번호)
        for b in sh.row_breaks.brk:
            b.id = _row_shift(b.id + 1, at, k) - 1
        # 인쇄 영역 · 인쇄 제목
        pa = sh.print_area
        if pa:
            refs = []
            for part in pa.split(","):
                _, ref = _split_sheet(part)
                refs.append(_shift_ref(ref.replace("$", ""), at, k))
            sh.print_area = refs
        if sh.print_title_rows:
            sh.print_title_rows = _shift_ref(sh.print_title_rows.replace("$", ""), at, k)
        # 틀 고정(고정한 줄 수) — 경계 위 마지막 고정 줄 번호로 따진다
        for view in sh.views.sheetView:
            p = view.pane
            if p is not None and p.state in ("frozen", "frozenSplit") and p.ySplit:
                y = int(p.ySplit)
                ny = _row_shift(y, at, k)
                if ny != y:
                    p.ySplit = ny
                    if p.topLeftCell:
                        p.topLeftCell = _shift_ref(p.topLeftCell, at, k)
        # 자동 거르개의 정렬 상태
        ss = sh.auto_filter.sortState
        if ss is not None:
            ss.ref = shift_sqref(ss.ref, at, k)
            for sc in ss.sortCondition:
                sc.ref = shift_sqref(sc.ref, at, k)
        # 표
        for t in sh.tables.values():
            t.ref = _shift_ref(t.ref, at, k)
            if t.autoFilter is not None and t.autoFilter.ref:
                t.autoFilter.ref = _shift_ref(t.autoFilter.ref, at, k)
            if t.sortState is not None:
                t.sortState.ref = shift_sqref(t.sortState.ref, at, k)
                for sc in t.sortState.sortCondition:
                    sc.ref = shift_sqref(sc.ref, at, k)


# ------------------------------------------------------------------ 검문(1-3)


class _Pkg:
    """xlsx 꾸러미를 XML 로 직접 읽는다(openpyxl 을 쓰지 않는다)."""

    def __init__(self, path: str):
        with zipfile.ZipFile(path) as z:
            self.files = {n: z.read(n) for n in z.namelist() if not n.endswith("/")}
        ct = ET.fromstring(self.files["[Content_Types].xml"])
        defaults = {e.get("Extension", "").lower(): e.get("ContentType") for e in ct if _local(e.tag) == "Default"}
        overrides = {e.get("PartName", "").lstrip("/"): e.get("ContentType") for e in ct
                     if _local(e.tag) == "Override"}
        self.ctype = {}
        for n in self.files:
            if n == "[Content_Types].xml":
                continue
            self.ctype[n] = overrides.get(n) or defaults.get(n.rsplit(".", 1)[-1].lower(), "")
        self._xml: Dict[str, ET.Element] = {}

    def kind(self, name: str) -> Optional[str]:
        return KINDS.get(self.ctype.get(name, ""))

    def xml(self, name: str) -> ET.Element:
        if name not in self._xml:
            self._xml[name] = ET.fromstring(self.files[name])
        return self._xml[name]

    def rels(self, name: str) -> Dict[str, Tuple[str, str, str]]:
        """{rId: (type 끝 이름, 대상 경로, TargetMode)}"""
        d, b = posixpath.split(name)
        rp = posixpath.join(d, "_rels", b + ".rels")
        out = {}
        if rp not in self.files:
            return out
        for e in ET.fromstring(self.files[rp]):
            tgt = e.get("Target", "")
            mode = e.get("TargetMode", "")
            if mode != "External":
                tgt = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(d, tgt))
            out[e.get("Id")] = (e.get("Type", "").rsplit("/", 1)[-1], tgt, mode)
        return out

    def workbook_name(self) -> str:
        for rid, (t, tgt, _) in self.rels("").items():
            if t == "officeDocument":
                return tgt
        raise PartError("통합 문서 본문을 찾지 못했습니다")

    def sheets(self) -> Dict[str, str]:
        """{시트 이름: 시트 부품 경로}"""
        wbn = self.workbook_name()
        rels = self.rels(wbn)
        out = {}
        for s in self.xml(wbn).iter("{%s}sheet" % NS_MAIN):
            rid = s.get("{%s}id" % NS_R)
            if rid in rels:
                out[s.get("name")] = rels[rid][1]
        return out

    def sheet_parts(self, sheet_path: str, typ: str) -> List[str]:
        return [tgt for t, tgt, _ in self.rels(sheet_path).values() if t == typ]


def _rels_name(name: str) -> str:
    d, b = posixpath.split(name)
    return posixpath.join(d, "_rels", b + ".rels")


def _ext_uris(root: ET.Element) -> List[str]:
    return [x.get("uri", "") for e in root.iter() if _local(e.tag) == "extLst" for x in e]


VML_NOTE_CHILDREN = Counter(["fill", "shadow", "path", "textbox", "div", "ClientData", "MoveWithCells",
                             "SizeWithCells", "AutoFill", "Row", "Column"])


def _openpyxl_note_style() -> str:
    from openpyxl.comments.shape_writer import _shape_factory

    return _shape_factory(0, 0, 79, 144).get("style")


class Guard:
    """원래 파일을 검문하고(쓰기 전), 저장한 새 파일을 대조한다(바꿔 끼우기 전)."""

    def __init__(self, path: str, have_pil: Optional[bool] = None):
        self.pkg = _Pkg(path)
        self._have_pil = have_pil
        self.check_before()

    # ---- 쓰기 전: 닫힌 목록
    def check_before(self) -> None:
        p = self.pkg
        for n, ct in p.ctype.items():
            if KINDS.get(ct) is None:
                raise lost(REFUSED_KIND_NAMES.get(ct, f"알 수 없는 부품({posixpath.basename(n)})"))
        if any(p.kind(n) == "image" for n in p.files) and not self._pil():
            raise lost("그림")
        for n in p.files:
            kind = p.kind(n)
            if kind == "worksheet":
                for ch in p.xml(n):
                    if _local(ch.tag) == "extLst":
                        continue  # 아래 꾸러미 검문에서 이름으로 가린다
                    if _local(ch.tag) not in SHEET_CHILDREN:
                        raise lost(SHEET_CHILD_NAMES.get(_local(ch.tag), f"시트 부품({_local(ch.tag)})"))
            elif kind == "workbook":
                for ch in p.xml(n):
                    if _local(ch.tag) not in WORKBOOK_CHILDREN:
                        raise lost(WORKBOOK_CHILD_NAMES.get(_local(ch.tag), f"통합 문서 부품({_local(ch.tag)})"))
                    if _local(ch.tag) == "AlternateContent":
                        inner = [_local(e.tag) for e in ch.iter()]
                        if set(inner) - {"AlternateContent", "Choice", "Fallback", "absPath"}:
                            raise lost("통합 문서 부품(호환 묶음)")
            elif kind == "drawing":
                self._check_drawing(n)
            elif kind == "vml":
                self._check_vml(n)
            elif kind == "sharedStrings":
                if any(_local(e.tag) == "rPh" for e in p.xml(n).iter()):
                    raise lost("윗주(발음 표시)")
            elif kind == "chart":
                for e in p.xml(n).iter():
                    if _local(e.tag) in ("userShapes", "externalData"):
                        raise lost("차트 안 도형 · 차트 데이터 통합 문서")
        # 이름 있는 부품을 먼저 가린 뒤, 남은 엑셀 새 기능 꾸러미(extLst)를 본다
        for n in p.files:
            kind = p.kind(n)
            if kind in ("workbook", "worksheet", "styles", "drawing", "chart"):
                for uri in _ext_uris(p.xml(n)):
                    if uri not in EXT_OK[kind]:
                        raise lost(EXT_NAMES.get(uri, "엑셀 새 기능 부품"))

    def _pil(self) -> bool:
        if self._have_pil is not None:
            return self._have_pil
        try:
            from openpyxl.drawing.image import PILImage
        except Exception:
            return False
        return bool(PILImage)

    def _check_drawing(self, n: str) -> None:
        root = self.pkg.xml(n)
        for anc in root:
            if _local(anc.tag) not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                raise lost("글상자 · 도형")
            for ch in anc:
                t = _local(ch.tag)
                if t in ("from", "to", "pos", "ext", "clientData"):
                    continue
                if t == "pic":
                    continue
                if t == "graphicFrame":
                    gd = [e for e in ch.iter() if _local(e.tag) == "graphicData"]
                    if not gd or gd[0].get("uri") != NS_C:
                        raise lost("스마트아트 · 그 밖의 개체")
                    continue
                if t in ("sp", "grpSp", "cxnSp"):
                    raise lost("글상자 · 도형")
                raise lost("글상자 · 도형 · 그 밖의 개체")

    def _check_vml(self, n: str) -> None:
        try:
            root = ET.fromstring(self.pkg.files[n].decode("utf-8", "replace").replace("<br>", "<br/>"))
        except ET.ParseError:
            raise lost("메모 상자 · 양식 컨트롤")
        style = _openpyxl_note_style()
        for sh in root:
            t = _local(sh.tag)
            if t in ("shapelayout", "shapetype"):
                continue
            if t != "shape":
                raise lost("양식 컨트롤 · 도형")
            cd = [e for e in sh if _local(e.tag) == "ClientData"]
            if not cd or cd[0].get("ObjectType") != "Note":
                raise lost("양식 컨트롤")
            kids = Counter(_local(e.tag) for e in sh.iter() if e is not sh)
            if sh.get("style") != style or kids - VML_NOTE_CHILDREN:
                raise lost("메모 상자 모양(크기 · 위치 · 늘 보이기)")

    # ---- 저장 뒤: 새 파일을 다시 읽어 대조
    def check_after(self, new_path: str, target_sheet: str, at: int, k: int) -> None:
        a, b = self.pkg, _Pkg(new_path)
        # 1) 종류별 부품 수가 줄지 않았나
        # sharedStrings: openpyxl 3.1.5 는 글자를 칸 안에(inlineStr) 적는다 — 공유 글자 부품이 없어져도 글자는 칸에 있다(실측)
        skip = {"calcChain", "app", "rels", "sharedStrings"}
        ca = Counter(a.kind(n) for n in a.files if a.kind(n) not in skip)
        cb = Counter(b.kind(n) for n in b.files if b.kind(n) not in skip)
        for kind, n in (ca - cb).items():
            raise lost(KIND_LOST_NAMES.get(kind, kind))
        # 2) 시트 · 통합 문서 · 서식 · 표 · 속성: 요소가 사라지지 않았나(값 칸 sheetData 는 core 가 다룬다)
        for kind in ("workbook", "worksheet", "styles", "table", "core", "custom"):
            ea = _inventory(a, kind)
            eb = _inventory(b, kind)
            gone = ea - eb
            if gone:
                path = sorted(gone)[0]
                raise lost(ELEMENT_LOST_NAMES.get(path.rsplit("/", 1)[-1], f"부품({path})"))
        # 3) 그림 · 차트 · 칸 메모: 원래 파일에 규칙대로 줄을 끼운 기대값과 새 파일이 같은가
        exp, got = _drawing_sigs(a, target_sheet, at, k), _drawing_sigs(b, target_sheet, 0, 0)
        if exp != got:
            raise lost(_which(exp, got, "그림 · 차트"))
        exp, got = _comment_sigs(a, target_sheet, at, k), _comment_sigs(b, target_sheet, 0, 0)
        if exp != got:
            raise lost("칸 메모(글 꾸밈 · 자리)")


KIND_LOST_NAMES = {"drawing": "그림 · 차트 · 도형", "chart": "차트", "image": "그림", "comments": "칸 메모",
                   "vml": "칸 메모 상자", "table": "표", "custom": "문서 사용자 속성", "theme": "테마",
                   "sharedStrings": "글자 칸", "styles": "서식"}
ELEMENT_LOST_NAMES = {"conditionalFormatting": "조건부 서식", "cfRule": "조건부 서식", "dataValidation": "드롭다운",
                      "hyperlink": "하이퍼링크", "mergeCell": "합친 칸", "brk": "페이지 나누기",
                      "definedName": "정의된 이름", "autoFilter": "자동 거르개", "filterColumn": "자동 거르개 조건",
                      "sortState": "정렬 상태", "pane": "틀 고정", "selection": "선택 영역",
                      "sheetProtection": "시트 보호", "workbookProtection": "통합 문서 보호",
                      "fileSharing": "쓰기 암호", "headerFooter": "머리글 · 바닥글", "pageSetup": "인쇄 설정",
                      "dxf": "조건부 서식의 서식", "numFmt": "표시 형식", "cellStyle": "셀 스타일",
                      "tableStyle": "표 스타일", "tabColor": "시트 탭 색", "col": "열 너비"}

#: 요소 대조에서 빼는 것 — 엑셀이 다시 만드는 꼬리표 · 화면 상태뿐(내용 없음)
INV_SKIP = {
    "workbook": ("workbook/fileVersion", "workbook/mc:AlternateContent", "workbook/xr:revisionPtr",
                 "workbook/extLst", "workbook/fileRecoveryPr", "workbook/bookViews"),
    "worksheet": ("worksheet/sheetData", "worksheet/dimension", "worksheet/phoneticPr",
                  "worksheet/sheetViews/sheetView/selection"),
    "styles": ("styleSheet/extLst", "styleSheet/colors/mruColors"),
    "table": (),
    "core": ("cp:coreProperties/cp:lastModifiedBy", "cp:coreProperties/dcterms:modified"),
    "custom": (),
}


def _inventory(p: _Pkg, kind: str) -> Counter:
    out = Counter()
    skip = INV_SKIP[kind]

    def walk(e, path):
        if path.startswith(skip):
            return
        if len(e) == 0 and set(e.attrib) <= {"count"} and e.get("count", "0") == "0" and not (e.text or "").strip() \
                and path.count("/") == 1:
            return  # 빈 목록(<dxfs count="0"/> 등) — openpyxl 은 빈 목록을 적지 않는다(Numbers 파일 실측). 내용 없음
        if path == "styleSheet/numFmts/numFmt" and e.get("numFmtId", "").isdigit() and int(e.get("numFmtId")) < 164:
            return  # 엑셀 내장 표시 형식(번호 0~163)을 다시 적은 것 — openpyxl 은 내장 형식을 적지 않는다(Numbers 「General」 실측)
        if path != "styleSheet/numFmts":  # 목록 그릇은 세지 않고 안의 형식만 센다
            out[path] += 1
        for ch in e:
            if isinstance(ch.tag, str):
                walk(ch, path + "/" + _short(ch.tag))

    for n in p.files:
        if p.kind(n) == kind:
            r = p.xml(n)
            walk(r, _short(r.tag))
    return out


# ---- 그림 · 차트 서명(원래 파일 쪽은 시험과 같은 규칙으로 따로 옮긴다)

_IGN_ATTR = {"id", "name", "{%s}id" % NS_R, "{%s}embed" % NS_R, "{%s}link" % NS_R, "macro", "fPublished",
             "{http://www.w3.org/XML/1998/namespace}space"}


def _canon(e: ET.Element, f_shift=None) -> tuple:
    tag = _short(e.tag)
    if tag in ("c:extLst", "mc:AlternateContent", "a:extLst"):
        return ()  # 검문(EXT_OK)에서 뜻 없는 꼬리표만 남겼다
    if tag == "a:ln" and [(_short(c.tag), c.attrib) for c in e] == [("a:prstDash", {"val": "solid"})] \
            and not e.attrib:
        return ()  # openpyxl 이 다시 저장할 때 그림에 붙이는 빈 선(색 · 굵기 없음 — 그려지지 않는다)
    attrs = tuple(sorted((_short(k), v) for k, v in e.attrib.items() if k not in _IGN_ATTR and v != ""))
    text = (e.text or "").strip()
    if f_shift and tag == "c:f":
        text = f_shift(text)
    kids = tuple(x for x in (_canon(c, f_shift) for c in e if isinstance(c.tag, str)) if x)
    return (tag, attrs, text, kids)


def _drawing_sigs(p: _Pkg, target: str, at: int, k: int) -> Counter:
    out = Counter()
    for sname, spath in p.sheets().items():
        mine = sname == target and k
        for dpath in p.sheet_parts(spath, "drawing"):
            drels = p.rels(dpath)
            for anc in p.xml(dpath):
                rows = []
                kind = _local(anc.tag)
                frm = [e for e in anc if _local(e.tag) == "from"]
                to = [e for e in anc if _local(e.tag) == "to"]
                r1 = int([e for e in frm[0] if _local(e.tag) == "row"][0].text) if frm else None
                r2 = int([e for e in to[0] if _local(e.tag) == "row"][0].text) if to else None
                if mine and r1 is not None:
                    mode = anc.get("editAs", "twoCell") if kind == "twoCellAnchor" else "oneCell"
                    if kind == "oneCellAnchor" or mode == "twoCell":
                        r1 = _row_shift(r1 + 1, at, k) - 1
                        if r2 is not None:
                            r2 = _row_shift(r2 + 1, at, k) - 1
                    elif mode == "oneCell" and r1 + 1 >= at:
                        r1, r2 = r1 + k, r2 + k
                rows = (r1, r2)
                body = []
                for ch in anc:
                    t = _local(ch.tag)
                    if t in ("from", "to"):
                        body.append(_canon_anchor_point(ch))
                    elif t == "graphicFrame":
                        rid = [e.get("{%s}id" % NS_R) for e in ch.iter() if _local(e.tag) == "chart"][0]
                        cpath = drels[rid][1]
                        fs = (lambda f, home=sname: shift_bare(f, home, target, at, k, "차트")) if k else None
                        body.append(("graphicFrame", _canon(ch), _canon(p.xml(cpath), fs)))
                    elif t == "pic":
                        rid = [e.get("{%s}embed" % NS_R) for e in ch.iter() if _local(e.tag) == "blip"][0]
                        img = p.files[drels[rid][1]]
                        body.append(("pic", _canon(ch), _image_sig(img)))
                    else:
                        body.append(_canon(ch))
                out[(sname, kind, anc.get("editAs", ""), rows, tuple(body))] += 1
    return out


def _canon_anchor_point(e: ET.Element) -> tuple:
    return tuple((_local(c.tag), (c.text or "").strip()) for c in e if _local(c.tag) != "row")


def _image_sig(data: bytes) -> tuple:
    try:
        from PIL import Image

        import io

        im = Image.open(io.BytesIO(data))
        return (im.size, hashlib.sha256(im.convert("RGBA").tobytes()).hexdigest())
    except Exception:
        return (len(data),)


def _comment_sigs(p: _Pkg, target: str, at: int, k: int) -> Counter:
    out = Counter()
    for sname, spath in p.sheets().items():
        for cpath in p.sheet_parts(spath, "comments"):
            root = p.xml(cpath)
            authors = [(a.text or "") for a in root.iter("{%s}author" % NS_MAIN)]
            for c in root.iter("{%s}comment" % NS_MAIN):
                ref = c.get("ref")
                if sname == target and k:
                    ref = _shift_ref(ref, at, k)
                text = [x for x in c if _local(x.tag) == "text"]
                aid = int(c.get("authorId", "0"))
                out[(sname, ref, authors[aid] if aid < len(authors) else "",
                     _canon(text[0]) if text else ())] += 1
    return out


def _which(exp: Counter, got: Counter, default: str) -> str:
    for key in (exp - got):
        body = key[-1]
        names = [b[0] for b in body if isinstance(b, tuple) and b and isinstance(b[0], str)]
        if "graphicFrame" in names:
            return "차트(모양 · 자리)"
        if "pic" in names:
            return "그림(모양 · 자리)"
    return default
