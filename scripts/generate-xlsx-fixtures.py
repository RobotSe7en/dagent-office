"""Independent deterministic SpreadsheetML inputs, using Python ZIP/XML only.

No SDK importer/exporter, spreadsheet library or formula engine is used. Shared
DrawingML fixture primitives come from this project's independent PPTX generator.
"""
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.sax.saxutils import escape, quoteattr
import argparse
import hashlib
import importlib.util
import json
import posixpath
import re
import struct
import zlib
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
MR = "http://schemas.microsoft.com/office/2017/10/relationships/"
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
PREFIX = "application/vnd.openxmlformats-officedocument."
TYPES = {"workbook": PREFIX + "spreadsheetml.sheet.main+xml", "worksheet": PREFIX + "spreadsheetml.worksheet+xml",
         "styles": PREFIX + "spreadsheetml.styles+xml", "sharedStrings": PREFIX + "spreadsheetml.sharedStrings+xml",
         "table": PREFIX + "spreadsheetml.table+xml", "comments": PREFIX + "spreadsheetml.comments+xml",
         "theme": PREFIX + "theme+xml", "drawing": PREFIX + "drawing+xml", "chart": PREFIX + "drawingml.chart+xml"}
spec = importlib.util.spec_from_file_location("independent_drawingml", ROOT / "scripts/generate-pptx-fixtures.py")
drawing = importlib.util.module_from_spec(spec)
spec.loader.exec_module(drawing)


def attrs(values):
    return "".join(" " + key + "=" + quoteattr(str(value)) for key, value in values.items() if value is not None)


def column(index):
    result = ""
    while index:
        index, digit = divmod(index - 1, 26)
        result = chr(65 + digit) + result
    return result


def cell(address, value=None, kind=None, style=None, formula=None, formula_attrs=None, rich=None):
    kind = kind or ("b" if isinstance(value, bool) else "inlineStr" if isinstance(value, str) else "n")
    props = {"r": address, "s": style, "t": kind if kind != "n" else None}
    content = ""
    if formula is not None:
        content += "<f" + attrs(formula_attrs or {}) + ">" + escape(formula) + "</f>"
    elif formula_attrs:
        content += "<f" + attrs(formula_attrs) + "/>"
    if rich is not None:
        content += "<is>" + rich + "</is>"
    elif value is not None:
        if kind == "inlineStr":
            content += '<is><t xml:space="preserve">' + escape(value) + "</t></is>"
        else:
            content += "<v>" + escape("1" if value is True else "0" if value is False else str(value)) + "</v>"
    return "<c" + attrs(props) + ">" + content + "</c>"


def styles():
    formats = [(164, '#,##0.00;[Red]-#,##0.00'), (165, 'yyyy-mm-dd'), (166, '0.00%'), (167, 'hh:mm:ss'),
               (168, '0.00E+00'), (169, '"¥"#,##0.00'), (170, '0.00;[Red](0.00);"—";@')]
    fmt = "".join('<numFmt' + attrs({"numFmtId": i, "formatCode": code}) + '/>' for i, code in formats)
    fonts = '<font><sz val="11"/><color theme="1"/><name val="Liberation Sans"/><family val="2"/><scheme val="minor"/></font>'
    fonts += '<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Noto Sans CJK SC"/><family val="2"/></font>'
    fonts += '<font><i/><u/><sz val="11"/><color theme="4" tint="-0.2"/><name val="Noto Sans CJK SC"/></font>'
    fills = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4"/><bgColor indexed="64"/></patternFill></fill>'
    borders = '<border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF4472C4"/></left><right style="thin"><color rgb="FF4472C4"/></right><top style="thin"><color rgb="FF4472C4"/></top><bottom style="thin"><color rgb="FF4472C4"/></bottom><diagonal/></border>'
    xfs = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    xfs += '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    for i, _ in formats:
        xfs += f'<xf numFmtId="{i}" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
    xfs += '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>'
    return XML + f'<styleSheet xmlns="{S}"><numFmts count="7">{fmt}</numFmts><fonts count="3">{fonts}</fonts><fills count="3">{fills}</fills><borders count="2">{borders}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="10">{xfs}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="2"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill></dxf></dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'


class Sheet:
    def __init__(self, name, visibility="visible"):
        self.name, self.visibility = name, visibility
        self.rows, self.row_props = {}, {}
        self.before = '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" defaultColWidth="12"/>'
        self.after = ""
        self.page = '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="1"/>'
        self.tail = ""
        self.print_area = "$A$1:$L$30"

    def add(self, address, value=None, **options):
        row = int(re.search(r"\d+$", address)[0])
        self.rows.setdefault(row, []).append(cell(address, value, **options))

    def xml(self):
        def order(cell_xml):
            letters = re.search(r'r="([A-Z]+)', cell_xml)[1]
            return len(letters), letters
        data = "".join('<row' + attrs({"r": row, **self.row_props.get(row, {})}) + '>' + "".join(sorted(cells, key=order)) + '</row>' for row, cells in sorted(self.rows.items()))
        return XML + f'<worksheet xmlns="{S}" xmlns:r="{R}"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>{self.before}<sheetData>{data}</sheetData>{self.after}{self.page}{self.tail}</worksheet>'


class Book:
    def __init__(self, date1904=False, mode="auto"):
        self.date1904, self.mode = date1904, mode
        self.sheets, self.parts, self.types, self.relations, self.names, self.strings = [], {}, {}, {}, [], []
        self.part("xl/styles.xml", styles(), "styles")
        self.part("xl/theme/theme1.xml", drawing.theme(), "theme")
        self.relate("", "office", R + "/officeDocument", "xl/workbook.xml")
        self.relate("xl/workbook.xml", "styles", R + "/styles", "xl/styles.xml")
        self.relate("xl/workbook.xml", "theme", R + "/theme", "xl/theme/theme1.xml")

    def part(self, path, content, kind):
        self.parts[path] = content
        self.types[path] = TYPES.get(kind, kind)

    def relate(self, source, rid, kind, target, external=False):
        self.relations.setdefault(source, []).append({"Id": rid, "Type": kind, "Target": target if external else posixpath.relpath(target, posixpath.dirname(source) or "."), "TargetMode": "External" if external else None})

    def sheet(self, name="Data", visibility="visible"):
        sheet = Sheet(name, visibility)
        self.sheets.append(sheet)
        sheet.path = f"xl/worksheets/sheet{len(self.sheets)}.xml"
        self.relate("xl/workbook.xml", "sheet" + str(len(self.sheets)), R + "/worksheet", sheet.path)
        return sheet

    def finish(self):
        for sheet in self.sheets:
            self.part(sheet.path, sheet.xml(), "worksheet")
        sheets = "".join('<sheet' + attrs({"name": s.name, "sheetId": i + 1, "state": s.visibility, "r:id": "sheet" + str(i + 1)}) + '/>' for i, s in enumerate(self.sheets))
        names = list(self.names)
        for i, sheet in enumerate(self.sheets):
            names.append(({"name": "_xlnm.Print_Area", "localSheetId": i}, "'" + sheet.name.replace("'", "''") + "'!" + sheet.print_area))
        defined = "".join('<definedName' + attrs(props) + '>' + escape(formula) + '</definedName>' for props, formula in names)
        self.part("xl/workbook.xml", XML + f'<workbook xmlns="{S}" xmlns:r="{R}"><workbookPr date1904="{int(self.date1904)}"/><bookViews><workbookView activeTab="0"/></bookViews><sheets>{sheets}</sheets><definedNames>{defined}</definedNames><calcPr calcId="191029" calcMode="{self.mode}" fullCalcOnLoad="0"/></workbook>', "workbook")
        if self.strings:
            self.part("xl/sharedStrings.xml", XML + f'<sst xmlns="{S}" uniqueCount="{len(self.strings)}">' + "".join('<si>' + s + '</si>' for s in self.strings) + '</sst>', "sharedStrings")
            self.relate("xl/workbook.xml", "strings", R + "/sharedStrings", "xl/sharedStrings.xml")
        for source, relations in self.relations.items():
            path = "_rels/.rels" if not source else posixpath.join(posixpath.dirname(source), "_rels", posixpath.basename(source) + ".rels")
            self.parts[path] = XML + f'<Relationships xmlns="{P}">' + "".join('<Relationship' + attrs(r) + '/>' for r in relations) + '</Relationships>'
        self.parts["[Content_Types].xml"] = XML + f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + "".join('<Override' + attrs({"PartName": "/" + path, "ContentType": value}) + '/>' for path, value in sorted(self.types.items())) + '</Types>'
        return deepcopy(self.parts)


def zipped(parts):
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED, compresslevel=9) as archive:
        for name, content in sorted(parts.items()):
            info = ZipInfo(name, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            archive.writestr(info, content.encode("utf-8") if isinstance(content, str) else content, compresslevel=9)
    return output.getvalue()


def png():
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    pixels = b"".join(b"\x00" + bytes([68, 114, 196, 255]) * 32 for _ in range(24))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 32, 24, 8, 6, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(pixels)) + chunk(b"IEND", b"")


def marker(tag, col, row, dx=0, dy=0):
    return f'<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>{drawing.emu(dx)}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{drawing.emu(dy)}</xdr:rowOff></xdr:{tag}>'


def one_anchor(body, col, row, width, height):
    return '<xdr:oneCellAnchor>' + marker("from", col, row) + f'<xdr:ext cx="{drawing.emu(width)}" cy="{drawing.emu(height)}"/>{body}<xdr:clientData/></xdr:oneCellAnchor>'


def graphic_frame(index, rid):
    return f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{index}" name="Chart {index}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="{C}"><c:chart xmlns:c="{C}" r:id="{rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>'


def add_drawing(book, sheet, content):
    index = book.sheets.index(sheet) + 1
    path = f"xl/drawings/drawing{index}.xml"
    book.part(path, XML + f'<xdr:wsDr xmlns:xdr="{D}" xmlns:a="{A}" xmlns:r="{R}">{content}</xdr:wsDr>', "drawing")
    book.relate(sheet.path, "drawing", R + "/drawing", path)
    sheet.tail += '<drawing r:id="drawing"/>'
    return path


def basic_shape(index, text, preset="rect", textbox=False):
    return (f'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="{index}" name="Shape {index}" descr="独立测试形状"/><xdr:cNvSpPr' + (' txBox="1"' if textbox else '') + '/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm rot="600000" flipH="0"><a:off x="0" y="0"/><a:ext cx="{drawing.emu(180)}" cy="{drawing.emu(90)}"/></a:xfrm><a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="E8F0E6"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="287454"/></a:solidFill></a:ln></xdr:spPr>'
            '<xdr:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="1400"><a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/></a:rPr><a:t>' + escape(text) + '</a:t></a:r><a:endParaRPr lang="zh-CN"/></a:p></xdr:txBody></xdr:sp>')


def add_classic(book, sheet):
    book.part("xl/comments/comment1.xml", XML + f'<comments xmlns="{S}"><authors><author>独立备注作者</author></authors><commentList><comment ref="B2" authorId="0"><text><r><rPr><b/><rFont val="Noto Sans CJK SC"/></rPr><t>请确认：</t></r><r><t>预算😀禰󠄀</t></r></text></comment></commentList></comments>', "comments")
    book.relate(sheet.path, "notes", R + "/comments", "xl/comments/comment1.xml")
    vml = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
           '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
           '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:30pt;width:150pt;height:75pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2, 10, 1, 2, 5, 20, 5, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>')
    book.part("xl/drawings/comments.vml", vml, "application/vnd.openxmlformats-officedocument.vmlDrawing")
    book.relate(sheet.path, "legacy", R + "/vmlDrawing", "xl/drawings/comments.vml")
    sheet.tail += '<legacyDrawing r:id="legacy"/>'


def add_modern(book, sheet):
    person1, person2 = "{11111111-1111-4111-8111-111111111111}", "{22222222-2222-4222-8222-222222222222}"
    root, reply, second = "{33333333-3333-4333-8333-333333333333}", "{44444444-4444-4444-8444-444444444444}", "{55555555-5555-4555-8555-555555555555}"
    book.part("xl/persons/person.xml", XML + f'<personList xmlns="{TC}"><person displayName="预算负责人" id="{person1}" userId="opaque-owner" providerId="None"/><person displayName="复核人" id="{person2}" userId="opaque-reviewer" providerId="None"/></personList>', "application/vnd.ms-excel.person+xml")
    book.relate("xl/workbook.xml", "persons", MR + "person", "xl/persons/person.xml")
    body = f'<threadedComment ref="B2" dT="2026-11-04T08:00:00Z" personId="{person1}" id="{root}" done="1"><text>请复核预算😀禰󠄀</text></threadedComment>'
    body += f'<threadedComment ref="B2" dT="2026-11-04T08:01:00Z" personId="{person2}" id="{reply}" parentId="{root}"><text>已复核并补充。</text></threadedComment>'
    body += f'<threadedComment ref="C3" dT="2026-11-04T08:02:00Z" personId="{person1}" id="{second}" done="0"><text>待确认。</text></threadedComment>'
    book.part("xl/threadedComments/threadedComment1.xml", XML + f'<ThreadedComments xmlns="{TC}">{body}</ThreadedComments>', "application/vnd.ms-excel.threadedcomments+xml")
    book.relate(sheet.path, "threads", MR + "threadedComment", "xl/threadedComments/threadedComment1.xml")


def cases():
    result = []
    reference_text = ["类型与中文", "日期", "冻结、隐藏与跨块合并", "名称与跨表", "产品", "比较", "第一季度", "图片、形状与文本框", "经典备注", "现代线程", "独立打印", "多表组合", "1904 日期系统", "安全保留未实现函数", "客户 0", "稀疏远端坐标"]
    def save(name, book, **expected):
        result.append((name + ".xlsx", book.finish(), {"sheets": [s.name for s in book.sheets], "referenceText": [reference_text[len(result)]], **expected}))

    book = Book(); sheet = book.sheet()
    values = [("A1", "类型与中文😀禰󠄀", {"style": 1}), ("A2", "编号 001", {}), ("B2", "001234567890123456789", {}),
              ("A3", 12.5, {}), ("B3", -2.25, {}), ("C3", True, {}), ("D3", False, {}),
              ("A4", "", {}), ("B4", None, {"style": 2}), ("C4", "#DIV/0!", {"kind": "e"}), ("D4", "#N/A", {"kind": "e"})]
    for address, value, options in values: sheet.add(address, value, **options)
    rich = '<r><rPr><b/><rFont val="Noto Sans CJK SC"/></rPr><t>中文</t></r><r><rPr><i/><rFont val="Liberation Sans"/></rPr><t xml:space="preserve"> Latin 😀禰󠄀\n下一行</t></r>'
    sheet.add("A6", kind="inlineStr", rich=rich, style=9)
    book.strings = ['<t xml:space="preserve">共享字符串 &amp; 保留空格 </t>', rich]
    sheet.add("B6", 0, kind="s"); sheet.add("C6", 0, kind="s"); sheet.add("D6", 1, kind="s")
    sheet.after = '<hyperlinks><hyperlink ref="A2" r:id="link" display="文档链接" tooltip="外部链接，不自动访问"/></hyperlinks>'
    book.relate(sheet.path, "link", R + "/hyperlink", "https://example.com/authorized-link", True)
    save("01-values-text", book, cellCount=15, formulaCount=0, text=["类型与中文😀禰󠄀", "001234567890123456789", "共享字符串 & 保留空格 "], values={"Data!A3": 12.5, "Data!C3": True, "Data!C4": "#DIV/0!"})

    book = Book(); sheet = book.sheet("格式")
    for c, label in enumerate(["数值", "日期", "百分比", "时间", "科学计数", "货币", "分段格式"], 1): sheet.add(column(c) + "1", label, style=1)
    for c, value in enumerate([12345.678, 46000, .125, .5, 1.25e8, -1234.5, 0], 1): sheet.add(column(c) + "2", value, style=c + 1)
    for r, serial in enumerate([0, 1, 59, 60, 61], 4): sheet.add("B" + str(r), serial, style=3)
    sheet.row_props[1] = {"ht": 28, "customHeight": 1}
    sheet.before += '<cols><col min="1" max="7" width="16" customWidth="1"/></cols>'
    save("02-styles-dates", book, cellCount=19, formulaCount=0, text=["日期", "分段格式"], dateSystem="1900", values={"格式!B2": 46000, "格式!B7": 60})

    book = Book(); sheet = book.sheet("结构")
    sheet.before = '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B3" sqref="B3"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" defaultColWidth="12"/><cols><col min="1" max="1" width="18" customWidth="1"/><col min="4" max="4" width="10" hidden="1" customWidth="1"/></cols>'
    sheet.add("A1", "冻结、隐藏与跨块合并", style=1); sheet.add("B2", "合并标题", style=1)
    sheet.add("A3", 3); sheet.add("A4", 4); sheet.row_props[4] = {"hidden": 1}
    sheet.add("A256", "跨分块合并", style=1); sheet.add("B260", 260)
    sheet.after = '<mergeCells count="2"><mergeCell ref="B2:D2"/><mergeCell ref="A256:A258"/></mergeCells>'
    sheet.row_props[2] = {"ht": 30, "customHeight": 1}
    save("03-structure", book, cellCount=6, formulaCount=0, text=["冻结、隐藏与跨块合并", "跨分块合并"], merges=["B2:D2", "A256:A258"], frozen={"rows": 2, "columns": 1})

    book = Book(); sheet = book.sheet("Data"); base = book.sheet("基础 2")
    base.add("A1", "名称与跨表"); base.add("B2", 5)
    sheet.add("A1", "数据", style=1); sheet.add("B1", "公式", style=1)
    for r in range(2, 6):
        sheet.add("A" + str(r), r)
        sheet.add("B" + str(r), r * 2, formula="A2*2" if r == 2 else None, formula_attrs={"t": "shared", "si": 0, **({"ref": "B2:B5"} if r == 2 else {})})
    formulas = [("D1", "SUM(A2:A5)", 14), ("D2", "'基础 2'!$B$2+A2", 7), ("D3", "BaseRate*A3", 15),
                ("D4", 'IF(A4>3,"通过","待定")', "通过"), ("D5", 'IFERROR(1/0,99)', 99),
                ("F1", "SUM(A:A)", 14), ("F2", "SUM(5:5)", 114), ("F3", "SUM($A2:A$5)", 14)]
    for address, formula, value in formulas: sheet.add(address, value, formula=formula, kind="str" if isinstance(value, str) else "n")
    book.names += [({"name": "BaseRate"}, "'基础 2'!$B$2"), ({"name": "LocalNumbers", "localSheetId": 0}, "Data!$A$2:$A$5")]
    save("04-formulas-names", book, cellCount=20, formulaCount=12, text=["名称与跨表", "通过"], values={"Data!D1": 14, "Data!D2": 7, "Data!F2": 114}, sharedFormulas=4, names=["BaseRate", "LocalNumbers"])

    book = Book(); sheet = book.sheet("Sales")
    for c, label in enumerate(["产品", "数量", "单价", "金额"], 1): sheet.add(column(c) + "1", label, style=1)
    for r, (name, quantity, price) in enumerate([("甲", 2, 10), ("乙", 3, 20), ("丙", 4, 30), ("丁", 1, 40)], 2):
        sheet.add("A" + str(r), name); sheet.add("B" + str(r), quantity); sheet.add("C" + str(r), price, style=2)
        sheet.add("D" + str(r), quantity * price, style=2, formula="SalesTable[[#This Row],[数量]]*SalesTable[[#This Row],[单价]]")
    sheet.add("A6", "合计", style=1); sheet.add("D6", 240, style=2, formula="SUBTOTAL(109,SalesTable[金额])")
    table = f'<table xmlns="{S}" id="1" name="SalesTable" displayName="SalesTable" ref="A1:D6" totalsRowCount="1"><autoFilter ref="A1:D5"/><tableColumns count="4"><tableColumn id="1" name="产品" totalsRowLabel="合计"/><tableColumn id="2" name="数量"/><tableColumn id="3" name="单价"/><tableColumn id="4" name="金额" totalsRowFunction="sum"><calculatedColumnFormula>SalesTable[[#This Row],[数量]]*SalesTable[[#This Row],[单价]]</calculatedColumnFormula></tableColumn></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
    book.part("xl/tables/table1.xml", XML + table, "table"); book.relate(sheet.path, "table", R + "/table", "xl/tables/table1.xml")
    sheet.tail = '<tableParts count="1"><tablePart r:id="table"/></tableParts>'
    save("05-native-table", book, cellCount=22, formulaCount=5, text=["产品", "金额", "合计"], tables=["SalesTable"], values={"Sales!D6": 240})

    book = Book(); sheet = book.sheet("Rules")
    for c, label in enumerate(["比较", "标量公式", "双色", "三色", "验证"], 1): sheet.add(column(c) + "1", label, style=1)
    for r, value in enumerate([10, 50, 100, 150, 200], 2):
        for c in range(1, 5): sheet.add(column(c) + str(r), value)
    sheet.add("E2", "甲"); sheet.add("E3", 99); sheet.add("E4", .5); sheet.add("E5", 46000, style=3); sheet.add("E6", .5, style=5); sheet.add("E7", "短文"); sheet.add("E8", 5)
    sheet.after = ('<conditionalFormatting sqref="A2:A6"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>'
                   '<conditionalFormatting sqref="B2:B6"><cfRule type="expression" dxfId="1" priority="2"><formula>AND(B2&gt;=50,B2&lt;=150)</formula></cfRule></conditionalFormatting>'
                   '<conditionalFormatting sqref="C2:C6"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
                   '<conditionalFormatting sqref="D2:D6"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>')
    rules = [("E2", "list", '"甲,乙,丙"', None), ("E3", "whole", "1", "10"), ("E4", "decimal", "0", "1"),
             ("E5", "date", "45000", "47000"), ("E6", "time", "0", "1"), ("E7", "textLength", "1", "5"), ("E8", "custom", "E8&gt;0", None)]
    sheet.after += '<dataValidations count="7">' + "".join('<dataValidation' + attrs({"type": kind, "errorStyle": "warning" if address == "E4" else "stop", "operator": "between", "allowBlank": 1, "showErrorMessage": 1, "showInputMessage": 1, "sqref": address, "errorTitle": "输入不符", "error": "请检查范围", "promptTitle": "填写提示", "prompt": "保留来源违规值"}) + '><formula1>' + formula + '</formula1>' + ('<formula2>' + second + '</formula2>' if second else '') + '</dataValidation>' for address, kind, formula, second in rules) + '</dataValidations>'
    save("06-rules-validation", book, cellCount=32, formulaCount=0, text=["比较", "三色", "短文"], conditionalFormats=4, validations=7, importedViolations=["Rules!E3"])

    book = Book(); sheet = book.sheet()
    for c, label in enumerate(["季度", "交付", "计划"], 1): sheet.add(column(c) + "1", label, style=1)
    for r, (name, actual, planned) in enumerate([("第一季度", 2, 3), ("第二季度", 5, 6), ("第三季度", 8, 9)], 2):
        for c, value in enumerate([name, actual, planned], 1): sheet.add(column(c) + str(r), value)
    anchors = "".join(one_anchor(graphic_frame(i, "chart" + str(i)), 0 if i != 2 else 5, 6 if i != 3 else 20, 300, 190) for i in range(1, 4))
    path = add_drawing(book, sheet, anchors)
    for i, kind in enumerate(["column", "line", "pie"], 1):
        source = drawing.chart_xml(kind, ["第一季度", "第二季度", "第三季度"], [("交付", [2, 5, 8])] + ([] if kind == "pie" else [("计划", [3, 6, 9])]))
        source = re.sub(r'<c:externalData\b.*?</c:externalData>', '', source)
        chart = f"xl/charts/chart{i}.xml"
        book.part(chart, source, "chart"); book.relate(path, "chart" + str(i), R + "/chart", chart)
    sheet.print_area = "$A$1:$K$35"
    save("07-three-charts", book, cellCount=12, formulaCount=0, text=["第一季度", "交付", "计划"], charts=["column", "line", "pie"], drawings=3)

    book = Book(); sheet = book.sheet("绘图")
    sheet.add("A1", "图片、形状与文本框", style=1)
    picture = ('<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="独立蓝色位图" descr="32 × 24 蓝色矩形"/><xdr:cNvPicPr/></xdr:nvPicPr>'
               '<xdr:blipFill><a:blip r:embed="image"/><a:srcRect l="10000" t="5000" r="10000" b="5000"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
               '<xdr:spPr><a:xfrm rot="900000"><a:off x="0" y="0"/><a:ext cx="1524000" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>')
    anchors = one_anchor(picture, 1, 2, 160, 120)
    anchors += '<xdr:twoCellAnchor editAs="twoCell">' + marker("from", 4, 2) + marker("to", 8, 9) + basic_shape(2, "随单元格移动缩放", "roundRect") + '<xdr:clientData/></xdr:twoCellAnchor>'
    anchors += '<xdr:absoluteAnchor><xdr:pos x="762000" y="2857500"/><xdr:ext cx="3810000" cy="1143000"/>' + basic_shape(3, "绝对位置文本框 ABC 中文😀", textbox=True) + '<xdr:clientData/></xdr:absoluteAnchor>'
    path = add_drawing(book, sheet, anchors)
    book.part("xl/media/blue.png", png(), "image/png"); book.relate(path, "image", R + "/image", "xl/media/blue.png")
    save("08-images-shapes", book, cellCount=1, formulaCount=0, text=["图片、形状与文本框", "随单元格移动缩放", "绝对位置文本框 ABC 中文😀"], drawings=3, images=1, shapes=2, anchors=["oneCellAnchor", "twoCellAnchor", "absoluteAnchor"])

    book = Book(); sheet = book.sheet("备注"); sheet.add("A1", "经典备注", style=1); sheet.add("B2", 1200, style=2)
    add_classic(book, sheet)
    save("09-classic-notes", book, cellCount=2, formulaCount=0, text=["经典备注", "请确认：预算😀禰󠄀"], classicNotes=1, commentAnchors=["B2"])

    book = Book(); sheet = book.sheet("线程"); sheet.add("A1", "现代线程", style=1); sheet.add("B2", 1200, style=2); sheet.add("C3", "待处理")
    add_modern(book, sheet)
    save("10-modern-threads", book, cellCount=3, formulaCount=0, text=["请复核预算😀禰󠄀", "已复核并补充。", "待确认。"], threads=2, replies=1, resolvedThreads=1, commentAnchors=["B2", "C3"])

    book = Book(); sheet = book.sheet("打印")
    for r in range(1, 46):
        sheet.add("A" + str(r), "标题" if r == 1 else "记录 " + str(r), style=1 if r == 1 else None)
        sheet.add("B" + str(r), "金额" if r == 1 else r * 100, style=1 if r == 1 else 2)
    sheet.print_area = "$A$1:$F$45"
    book.names.append(({"name": "_xlnm.Print_Titles", "localSheetId": 0}, "'打印'!$1:$1,'打印'!$A:$A"))
    sheet.page = ('<printOptions horizontalCentered="1" headings="0" gridLines="1"/><pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                  '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
                  '<headerFooter differentOddEven="0" differentFirst="0"><oddHeader>&amp;L独立打印&amp;R&amp;A</oddHeader><oddFooter>&amp;C第 &amp;P 页 / &amp;N 页</oddFooter></headerFooter>'
                  '<rowBreaks count="1" manualBreakCount="1"><brk id="23" min="0" max="16383" man="1"/></rowBreaks>')
    save("11-print-layout", book, cellCount=90, formulaCount=0, text=["标题", "记录 45", "独立打印"], printTitles=True, manualRowBreaks=[23], orientation="landscape")

    book = Book(); main = book.sheet("Overview"); data = book.sheet("Input"); hidden = book.sheet("Hidden", "hidden")
    main.add("A1", "多表组合", style=1); main.add("B2", 30, formula="SUM(Input!B2:B3)")
    for address, value in [("A1", "来源"), ("B1", "金额"), ("A2", "甲"), ("B2", 10), ("A3", "乙"), ("B3", 20)]: data.add(address, value)
    hidden.add("A1", "保留隐藏表"); data.row_props[3] = {"hidden": 1}
    data.after = '<autoFilter ref="A1:B3"><filterColumn colId="0"><filters><filter val="甲"/></filters></filterColumn></autoFilter>'
    book.names.append(({"name": "Amounts"}, "Input!$B$2:$B$3"))
    save("12-multi-sheet-filter", book, cellCount=9, formulaCount=1, text=["多表组合", "保留隐藏表"], hiddenSheets=["Hidden"], autoFilters=1, values={"Overview!B2": 30})

    book = Book(date1904=True); sheet = book.sheet("日期1904")
    sheet.add("A1", "1904 日期系统", style=1)
    for r, serial in enumerate([0, 1, 59, 60, 1462], 2): sheet.add("A" + str(r), serial, style=3)
    save("13-date-1904", book, cellCount=6, formulaCount=0, text=["1904 日期系统"], dateSystem="1904", values={"日期1904!A2": 0, "日期1904!A6": 1462})

    book = Book(mode="manual"); sheet = book.sheet()
    sheet.add("A1", "安全保留未实现函数", style=1); sheet.add("A2", 4)
    sheet.add("B2", 24, formula="FACT(A2)"); sheet.add("C2", 25, formula="B2+1")
    sheet.add("B3", None, formula="FACT(A2)")
    save("14-preserved-formulas", book, cellCount=5, formulaCount=3, text=["安全保留未实现函数"], calculationMode="manual", cachedUnsupported=["Data!B2"], unavailableCache=["Data!B3"], values={"Data!B2": 24, "Data!C2": 25})

    book = Book(); sheet = book.sheet("Baseline")
    for r in range(1, 10001):
        sheet.add("A" + str(r), r * 20); sheet.add("B" + str(r), "客户 " + str(r - 1) + " 😀禰󠄀")
        for c in range(3, 19): sheet.add(column(c) + str(r), r * 20 + c - 1)
        subtotal = sum(r * 20 + c for c in range(2, 18))
        sheet.add("S" + str(r), subtotal, formula=f"SUM(C{r}:R{r})")
        sheet.add("T" + str(r), r * 20 + subtotal, formula=f"A{r}+S{r}")
    sheet.print_area = "$A$1:$T$20"
    save("15-ten-thousand-rows", book, cellCount=200000, formulaCount=20000, text=["客户 0 😀禰󠄀", "客户 9999 😀禰󠄀"], rows=10000, columns=20, values={"Baseline!S10000": 3200152, "Baseline!T10000": 3400152})

    book = Book(); sheet = book.sheet("Sparse")
    sheet.add("A1", "稀疏远端坐标", style=1); sheet.add("XFD1048576", 42)
    sheet.add("B2", 42, formula="XFD1048576")
    save("16-sparse-far-cell", book, cellCount=3, formulaCount=1, text=["稀疏远端坐标"], values={"Sparse!XFD1048576": 42, "Sparse!B2": 42})

    return result

def negative_cases(valid):
    sources = {name: parts for name, parts, _ in valid}
    result = []
    def edit(name, source, path, old, new, error, reason):
        parts = deepcopy(sources[source + ".xlsx"])
        assert old in parts[path], name
        parts[path] = parts[path].replace(old, new, 1)
        result.append((name, zipped(parts), {"error": error, "reason": reason, "part": path, "marker": new}))
    ws = "xl/worksheets/sheet1.xml"
    edit("duplicate-cell", "01-values-text", ws, '<c r="A3">', '<c r="A2">', "InvalidDocument", "Duplicate coordinates cannot overwrite an earlier source cell.")
    edit("cell-out-of-bounds", "01-values-text", ws, '<c r="A3">', '<c r="XFE1048577">', "InvalidDocument", "Coordinates exceed native worksheet bounds.")
    edit("invalid-shared-string", "01-values-text", ws, '<c r="B6" t="s"><v>0</v>', '<c r="B6" t="s"><v>99</v>', "InvalidDocument", "Shared string index is unresolved.")
    edit("invalid-style", "01-values-text", ws, 's="1"', 's="99"', "InvalidDocument", "Cell style index is unresolved.")
    edit("doctype", "01-values-text", ws, XML, XML + '<!DOCTYPE worksheet [<!ENTITY secret SYSTEM "file:///not-authorized">]>', "InvalidDocument", "Document types and external entities are prohibited.")
    edit("array-formula", "04-formulas-names", ws, '<f>SUM(A2:A5)</f>', '<f t="array" ref="D1:D2">SUM(A2:A5)</f>', "UnsupportedFeature", "Array formulas are outside the declared scope.")
    edit("orphan-shared-formula", "04-formulas-names", ws, '<f t="shared" si="0"/>', '<f t="shared" si="99"/>', "InvalidDocument", "A shared follower must have an unambiguous origin.")
    edit("external-formula", "04-formulas-names", ws, '<f>SUM(A2:A5)</f>', '<f>SUM([External.xlsx]Data!A1:A5)</f>', "UnsupportedFeature", "External workbook dependencies cannot be safely preserved.")
    edit("overlapping-merges", "03-structure", ws, 'ref="A256:A258"', 'ref="C2:E3"', "InvalidDocument", "Merged rectangles overlap.")
    edit("protected-sheet", "01-values-text", ws, '</sheetData>', '</sheetData><sheetProtection sheet="1" password="FFFF"/>', "UnsupportedFeature", "Protection is not a supported editing or import contract.")
    edit("unknown-comment-author", "10-modern-threads", "xl/threadedComments/threadedComment1.xml", 'personId="{11111111-1111-4111-8111-111111111111}"', 'personId="{99999999-9999-4999-8999-999999999999}"', "InvalidDocument", "A threaded comment references an undefined author.")
    edit("thread-parent-cycle", "10-modern-threads", "xl/threadedComments/threadedComment1.xml", 'done="1"', 'done="1" parentId="{44444444-4444-4444-8444-444444444444}"', "InvalidDocument", "Thread parent links contain a cycle.")
    edit("unsupported-conditional-format", "06-rules-validation", ws, '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule>', '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule>', "UnsupportedFeature", "Data bars cannot be silently dropped.")
    edit("external-image", "08-images-shapes", "xl/drawings/_rels/drawing1.xml.rels", 'Target="../media/blue.png"', 'Target="https://example.com/unauthorized.png" TargetMode="External"', "UnsupportedFeature", "An image relationship cannot grant network access.")
    edit("unsupported-chart", "07-three-charts", "xl/charts/chart3.xml", '<c:pieChart>', '<c:pie3DChart>', "UnsupportedFeature", "Three-dimensional charts are not in scope.")
    # Keep the unsupported chart well-formed so rejection tests the feature, not XML syntax.
    name, data, expected = result.pop()
    with ZipFile(BytesIO(data)) as archive: parts = {p: archive.read(p) for p in archive.namelist()}
    parts[expected["part"]] = parts[expected["part"]].replace(b'</c:pieChart>', b'</c:pie3DChart>')
    result.append((name, zipped(parts), expected))
    parts = deepcopy(sources["01-values-text.xlsx"]); del parts[ws]
    result.append(("missing-worksheet", zipped(parts), {"error": "InvalidDocument", "reason": "Internal relationship target is missing.", "missingPart": ws}))
    result.append(("truncated-zip", zipped(sources["01-values-text.xlsx"])[:-22], {"error": "InvalidDocument", "reason": "ZIP directory is missing."}))
    return result


def generate(output):
    output.mkdir(parents=True, exist_ok=True)
    (output / "invalid").mkdir(exist_ok=True)
    manifest = {"format": "doffi-xlsx-fixtures", "version": 1, "generator": "Independent Python ZIP/XML, no SDK exporter or spreadsheet library.", "cases": [], "negativeCases": []}
    valid = cases()
    for name, parts, expected in valid:
        for path, data in parts.items():
            if path.endswith((".xml", ".rels", ".vml")): ET.fromstring(data)
        data = zipped(parts); (output / name).write_bytes(data)
        manifest["cases"].append({"file": name, "sha256": hashlib.sha256(data).hexdigest(), **expected})
    for name, data, expected in negative_cases(valid):
        filename = "invalid/" + name + ".xlsx"; (output / filename).write_bytes(data)
        manifest["negativeCases"].append({"file": filename, "sha256": hashlib.sha256(data).hexdigest(), **expected})
    (output / "expected.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
    return manifest


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "examples/xlsx")
    manifest = generate(parser.parse_args().output)
    print(f'Generated {len(manifest["cases"])} independent XLSX sources and {len(manifest["negativeCases"])} negative inputs.')
