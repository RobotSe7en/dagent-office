"""Independent deterministic ECMA-376/MS-PPTX fixtures, using only Python ZIP/XML.

This generator never imports an Office/editor library or the SDK exporter.
LibreOffice producer files and PDF references are recorded by pptx-reference.py.
"""
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_STORED, ZipFile, ZipInfo
from xml.sax.saxutils import escape, quoteattr
import argparse
import hashlib
import json
import posixpath
import re
from urllib.parse import quote, unquote
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
MODERN = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"
AC = "http://schemas.microsoft.com/office/drawing/2013/main/command"
MODERN_REL = "http://schemas.microsoft.com/office/2018/10/relationships/"
MODERN_EXT = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"
NS = f'xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"'
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
OFFICE_CT = "application/vnd.openxmlformats-officedocument."
CTYPE = {
    "presentation": OFFICE_CT + "presentationml.presentation.main+xml",
    "slide": OFFICE_CT + "presentationml.slide+xml",
    "master": OFFICE_CT + "presentationml.slideMaster+xml",
    "layout": OFFICE_CT + "presentationml.slideLayout+xml",
    "theme": OFFICE_CT + "theme+xml",
    "notes": OFFICE_CT + "presentationml.notesSlide+xml",
    "notesMaster": OFFICE_CT + "presentationml.notesMaster+xml",
    "comments": OFFICE_CT + "presentationml.comments+xml",
    "commentAuthors": OFFICE_CT + "presentationml.commentAuthors+xml",
    "chart": OFFICE_CT + "drawingml.chart+xml",
    "modernComments": "application/vnd.ms-powerpoint.comments+xml",
    "authors": "application/vnd.ms-powerpoint.authors+xml",
}
STAMP = "2026-10-31T08:00:00Z"


def emu(points):
    return round(points * 12700)


def zipped(parts):
    output = BytesIO()
    with ZipFile(output, "w", ZIP_STORED) as archive:
        for name, data in sorted(parts.items()):
            info = ZipInfo(name, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_STORED
            info.external_attr = 0o100644 << 16
            archive.writestr(info, data.encode("utf-8") if isinstance(data, str) else data)
    return output.getvalue()


def relationships(values):
    content = "".join(
        f'<Relationship Id={quoteattr(rid)} Type={quoteattr(kind if "://" in kind else R + "/" + kind)}'
        f' Target={quoteattr(target)}' + (' TargetMode="External"' if external else "") + "/>"
        for rid, kind, target, external in values)
    return XML + f'<Relationships xmlns="{PKG}">{content}</Relationships>'


def transform(x, y, width, height, rotation=0, flip_h=False, flip_v=False, tag="a:xfrm", child=None):
    attrs = (f' rot="{round(rotation * 60000)}"' if rotation else "")
    attrs += ' flipH="1"' if flip_h else ""
    attrs += ' flipV="1"' if flip_v else ""
    inner = f'<a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(width)}" cy="{emu(height)}"/>'
    if child is not None:
        cx, cy, cw, ch = child
        inner += f'<a:chOff x="{emu(cx)}" y="{emu(cy)}"/><a:chExt cx="{emu(cw)}" cy="{emu(ch)}"/>'
    return f"<{tag}{attrs}>{inner}</{tag}>"


def solid(color="4472C4", alpha=None, scheme=False):
    name = "schemeClr" if scheme else "srgbClr"
    opacity = f'<a:alpha val="{alpha}"/>' if alpha is not None else ""
    return f'<a:solidFill><a:{name} val="{color}">{opacity}</a:{name}></a:solidFill>'


def line(color="44546A", width=1, dashed=False, arrow=False):
    return f'<a:ln w="{emu(width)}">{solid(color)}<a:prstDash val="{"dash" if dashed else "solid"}"/><a:headEnd type="{"triangle" if arrow else "none"}"/><a:tailEnd type="none"/></a:ln>'


def run(value, props="", bold=False, italic=False, size=2400):
    attrs = (' b="1"' if bold else "") + (' i="1"' if italic else "")
    return f'<a:r><a:rPr lang="zh-CN" sz="{size}"{attrs}>{props}</a:rPr><a:t xml:space="preserve">{escape(value)}</a:t></a:r>'


def paragraph(value="", props="", content=None):
    return f'<a:p><a:pPr>{props}</a:pPr>{run(value) if content is None else content}<a:endParaRPr lang="zh-CN" sz="2400"/></a:p>'


def text_body(value="", paragraphs=None, tag="p:txBody", body_props='wrap="square" anchor="t" lIns="0" tIns="0" rIns="0" bIns="0"', autofit="noAutofit"):
    return f'<{tag}><a:bodyPr {body_props}><a:{autofit}/></a:bodyPr><a:lstStyle/>{paragraph(value) if paragraphs is None else paragraphs}</{tag}>'


def shape(index, text="", x=48, y=48, width=600, height=80, preset="rect", fill=None, outline=None,
          text_box=True, placeholder=None, no_transform=False, rotation=0, flip_h=False, flip_v=False,
          paragraphs=None, effects="", name=None, adjustments=""):
    ph = "" if placeholder is None else f'<p:ph type="{placeholder[0]}" idx="{placeholder[1]}"/>'
    xfrm = "" if no_transform else transform(x, y, width, height, rotation, flip_h, flip_v)
    body = "" if text is None and paragraphs is None else text_body(text or "", paragraphs=paragraphs)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{index}" name={quoteattr(name or "Shape " + str(index))}/>'
            f'<p:cNvSpPr' + (' txBox="1"' if text_box else "") + f'/><p:nvPr>{ph}</p:nvPr></p:nvSpPr>'
            f'<p:spPr>{xfrm}<a:prstGeom prst="{preset}"><a:avLst>{adjustments}</a:avLst></a:prstGeom>'
            + (fill if fill is not None else "<a:noFill/>")
            + (outline if outline is not None else '<a:ln><a:noFill/></a:ln>')
            + effects + f"</p:spPr>{body}</p:sp>")


def group_header(index=1):
    return f'<p:nvGrpSpPr><p:cNvPr id="{index}" name="Group {index}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'


def tree(content):
    return "<p:spTree>" + group_header() + "<p:grpSpPr>" + transform(0, 0, 0, 0, child=(0, 0, 0, 0)) + "</p:grpSpPr>" + content + "</p:spTree>"


def group(index, content, x=48, y=120, width=600, height=300, child=(0, 0, 600, 300), rotation=0, flip_h=False):
    return "<p:grpSp>" + group_header(index) + "<p:grpSpPr>" + transform(x, y, width, height, rotation, flip_h, child=child) + "</p:grpSpPr>" + content + "</p:grpSp>"


def connector(index, start, end, x=200, y=180, width=240, height=160, elbow=False):
    return (f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{index}" name="Connector {index}"/>'
            f'<p:cNvCxnSpPr><a:stCxn id="{start}" idx="1"/><a:endCxn id="{end}" idx="3"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>'
            "<p:spPr>" + transform(x, y, width, height) + f'<a:prstGeom prst="{"bentConnector3" if elbow else "line"}"><a:avLst/></a:prstGeom>' + line(arrow=True) + "</p:spPr></p:cxnSp>")


def picture(index, rid, x=48, y=120, width=180, height=180, crop=None, rotation=0, flip_h=False):
    rect = "" if crop is None else '<a:srcRect ' + " ".join(f'{key}="{round(value * 100000)}"' for key, value in zip(["l", "t", "r", "b"], crop)) + "/>"
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{index}" name="Picture {index}" descr="公开合成图片"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/>{rect}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>'
            + transform(x, y, width, height, rotation, flip_h)
            + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')


def graphic_frame(index, content, uri, x=48, y=120, width=800, height=300):
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{index}" name="Frame {index}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            + transform(x, y, width, height, tag="p:xfrm")
            + f'<a:graphic><a:graphicData uri="{uri}">{content}</a:graphicData></a:graphic></p:graphicFrame>')


def theme(accent="4472C4"):
    colors = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6", "accent1": accent,
              "accent2": "ED7D31", "accent3": "A5A5A5", "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47", "hlink": "0563C1", "folHlink": "954F72"}
    scheme = "".join(f'<a:{key}><a:srgbClr val="{value}"/></a:{key}>' for key, value in colors.items())
    font = '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/><a:cs typeface="Liberation Sans"/><a:font script="Hans" typeface="Noto Sans CJK SC"/><a:font script="Hant" typeface="Noto Sans CJK TC"/>'
    fills = solid("phClr", scheme=True) * 3
    lines = "".join(f'<a:ln w="{emu(w)}">{solid("phClr", scheme=True)}<a:prstDash val="solid"/></a:ln>' for w in [1, 2, 3])
    effects = "<a:effectStyle><a:effectLst/></a:effectStyle>" * 3
    return XML + f'<a:theme xmlns:a="{A}" name="Pinned fixture"><a:themeElements><a:clrScheme name="Fixture">{scheme}</a:clrScheme><a:fontScheme name="Pinned"><a:majorFont>{font}</a:majorFont><a:minorFont>{font}</a:minorFont></a:fontScheme><a:fmtScheme name="Fixture"><a:fillStyleLst>{fills}</a:fillStyleLst><a:lnStyleLst>{lines}</a:lnStyleLst><a:effectStyleLst>{effects}</a:effectStyleLst><a:bgFillStyleLst>{fills}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'


COLOR_MAP = '<p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>'
COLOR_OVERRIDE = "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
DEFAULT_STYLE = '<a:lvl1pPr><a:defRPr sz="2400"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'


class Deck:
    def __init__(self, title):
        self.title = title
        self.parts = {}
        self.types = {}
        self.rels = {}
        self.slides = []
        self.masters = []
        self.layouts = []
        self.notes_master = None
        self.add_master()

    def part(self, name, data, kind):
        self.parts[name] = data if isinstance(data, bytes) else XML + data.removeprefix(XML)
        self.types[name] = CTYPE.get(kind, kind)

    def relate(self, source, rid, kind, target, external=False):
        relative = target if external else posixpath.relpath(target, posixpath.dirname(source) or ".")
        self.rels.setdefault(source, []).append((rid, kind, relative, external))

    def add_master(self, accent="4472C4", content="", layout_content="", layout_type="blank"):
        index = len(self.masters) + 1
        master = f"ppt/slideMasters/slideMaster{index}.xml"
        layout = f"ppt/slideLayouts/slideLayout{index}.xml"
        theme_path = f"ppt/theme/theme{index}.xml"
        self.masters.append(master)
        self.layouts.append(layout)
        self.part(theme_path, theme(accent), "theme")
        styles = f"<p:txStyles><p:titleStyle>{DEFAULT_STYLE}</p:titleStyle><p:bodyStyle>{DEFAULT_STYLE}</p:bodyStyle><p:otherStyle>{DEFAULT_STYLE}</p:otherStyle></p:txStyles>"
        self.part(master, f'<p:sldMaster {NS}><p:cSld name="Master {index}">{tree(content)}</p:cSld>{COLOR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id="{2147483648 + index}" r:id="layout"/></p:sldLayoutIdLst>{styles}</p:sldMaster>', "master")
        self.part(layout, f'<p:sldLayout {NS} type="{layout_type}" preserve="1"><p:cSld name="Layout {index}">{tree(layout_content)}</p:cSld>{COLOR_OVERRIDE}</p:sldLayout>', "layout")
        self.relate(master, "theme", "theme", theme_path)
        self.relate(master, "layout", "slideLayout", layout)
        self.relate(layout, "master", "slideMaster", master)
        return index

    def slide(self, content, layout=1, hidden=False, extension="", name=None):
        index = len(self.slides) + 1
        path = f"ppt/slides/slide{index}.xml"
        self.slides.append(path)
        self.part(path, f'<p:sld {NS}' + (' show="0"' if hidden else "") + f'><p:cSld name={quoteattr(name or "Slide " + str(index))}>{tree(content)}</p:cSld>{COLOR_OVERRIDE}{extension}</p:sld>', "slide")
        self.relate(path, "layout", "slideLayout", self.layouts[layout - 1])
        return path

    def files(self):
        master_ids = "".join(f'<p:sldMasterId id="{2147483648 + i}" r:id="master{i + 1}"/>' for i in range(len(self.masters)))
        slide_ids = "".join(f'<p:sldId id="{256 + i}" r:id="slide{i + 1}"/>' for i in range(len(self.slides)))
        notes_ids = '<p:notesMasterIdLst><p:notesMasterId r:id="notes-master"/></p:notesMasterIdLst>' if self.notes_master else ""
        self.part("ppt/presentation.xml", f'<p:presentation {NS}><p:sldMasterIdLst>{master_ids}</p:sldMasterIdLst>{notes_ids}<p:sldIdLst>{slide_ids}</p:sldIdLst><p:sldSz cx="{emu(960)}" cy="{emu(540)}" type="custom"/><p:notesSz cx="{emu(540)}" cy="{emu(720)}"/></p:presentation>', "presentation")
        self.part("docProps/core.xml", '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>' + escape(self.title) + '</dc:title><dc:language>zh-CN</dc:language></cp:coreProperties>', "application/vnd.openxmlformats-package.core-properties+xml")
        links = deepcopy(self.rels)
        def link(source, rid, kind, target):
            links.setdefault(source, []).append((rid, kind, posixpath.relpath(target, posixpath.dirname(source) or "."), False))
        for i, path in enumerate(self.masters):
            link("ppt/presentation.xml", f"master{i + 1}", "slideMaster", path)
        for i, path in enumerate(self.slides):
            link("ppt/presentation.xml", f"slide{i + 1}", "slide", path)
        if self.notes_master:
            link("ppt/presentation.xml", "notes-master", "notesMaster", self.notes_master)
        link("", "main", "officeDocument", "ppt/presentation.xml")
        link("", "core", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml")
        files = dict(self.parts)
        for source, values in links.items():
            directory, name = posixpath.split(source)
            path = f"{directory + '/' if directory else ''}_rels/{name}.rels" if source else "_rels/.rels"
            # Each call remains deterministic and does not accumulate duplicate IDs.
            ids = [value[0] for value in values]
            if len(set(ids)) != len(ids):
                raise ValueError("Duplicate fixture relationship IDs")
            files[path] = relationships(values)
        overrides = "".join(f'<Override PartName="/{name}" ContentType={quoteattr(kind)}/>' for name, kind in sorted(self.types.items()))
        files["[Content_Types].xml"] = XML + f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{overrides}</Types>'
        return files


def table_cell(value, attributes=""):
    borders = "".join(f"<a:{side}>{solid('44546A')}<a:prstDash val=\"solid\"/></a:{side}>" for side in ["lnL", "lnR", "lnT", "lnB"])
    return f"<a:tc{attributes}>" + text_body(value, tag="a:txBody") + f'<a:tcPr marL="0" marR="0" marT="0" marB="0">{borders}{solid("FFFFFF")}</a:tcPr></a:tc>'


def table():
    rows = [
        table_cell("合并标题", ' gridSpan="2"') + table_cell("", ' hMerge="1"') + table_cell("第三列"),
        table_cell("跨行", ' rowSpan="2"') + table_cell("甲") + table_cell("乙"),
        table_cell("", ' vMerge="1"') + table_cell("丙") + table_cell("丁"),
    ]
    grid = "".join(f'<a:gridCol w="{emu(240)}"/>' for _ in range(3))
    return '<a:tbl><a:tblPr firstRow="0" bandRow="0"/><a:tblGrid>' + grid + "</a:tblGrid>" + "".join(f'<a:tr h="{emu(80)}">{row}</a:tr>' for row in rows) + "</a:tbl>"


def workbook(categories, series):
    parts = {}
    overrides = (
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    parts["[Content_Types].xml"] = XML + f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{overrides}</Types>'
    parts["_rels/.rels"] = relationships([("workbook", "officeDocument", "xl/workbook.xml", False)])
    parts["xl/workbook.xml"] = XML + f'<workbook xmlns="{SS}" xmlns:r="{R}"><sheets><sheet name="Data" sheetId="1" r:id="sheet"/></sheets></workbook>'
    parts["xl/_rels/workbook.xml.rels"] = relationships([("sheet", "worksheet", "worksheets/sheet1.xml", False)])
    def string(address, value):
        return f'<c r="{address}" t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    rows = [string("A1", "分类") + "".join(string(chr(66 + i) + "1", item[0]) for i, item in enumerate(series))]
    for row, category in enumerate(categories, 2):
        cells = string("A" + str(row), category)
        for index, (_, values) in enumerate(series):
            if values[row - 2] is not None:
                cells += f'<c r="{chr(66 + index)}{row}"><v>{values[row - 2]}</v></c>'
        rows.append(cells)
    parts["xl/worksheets/sheet1.xml"] = XML + f'<worksheet xmlns="{SS}"><sheetData>' + "".join(f'<row r="{i}">{row}</row>' for i, row in enumerate(rows, 1)) + "</sheetData></worksheet>"
    return zipped(parts)


def chart_xml(kind, categories, series):
    def cache(values, number=False):
        item = "numCache" if number else "strCache"
        points = "".join(f'<c:pt idx="{i}"><c:v>{escape(str(value))}</c:v></c:pt>' for i, value in enumerate(values) if value is not None)
        return f"<c:{item}>" + ("<c:formatCode>General</c:formatCode>" if number else "") + f'<c:ptCount val="{len(values)}"/>{points}</c:{item}>'
    body = ""
    end = str(len(categories) + 1)
    for i, (name, values) in enumerate(series):
        column = chr(66 + i)
        name_formula = "Data!$" + column + "$1"
        cat_formula = "Data!$A$2:$A$" + end
        val_formula = "Data!$" + column + "$2:$" + column + "$" + end
        body += f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:strRef><c:f>{name_formula}</c:f>{cache([name])}</c:strRef></c:tx>'
        body += '<c:spPr>' + (line("4472C4", 2) if kind == "line" else solid("4472C4")) + '</c:spPr>'
        if kind == "line":
            body += '<c:marker><c:symbol val="none"/></c:marker>'
        body += f"<c:cat><c:strRef><c:f>{cat_formula}</c:f>{cache(categories)}</c:strRef></c:cat><c:val><c:numRef><c:f>{val_formula}</c:f>{cache(values, True)}</c:numRef></c:val>"
        if kind == "line":
            body += '<c:smooth val="0"/>'
        body += "</c:ser>"
    tags = {"column": "barChart", "line": "lineChart", "pie": "pieChart"}
    settings = '<c:barDir val="col"/><c:grouping val="clustered"/>' if kind == "column" else '<c:grouping val="standard"/>' if kind == "line" else ""
    settings += '<c:varyColors val="0"/>'
    axes = ""
    if kind != "pie":
        body += '<c:axId val="100"/><c:axId val="200"/>'
        axes = ('<c:catAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="200"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>'
                '<c:valAx><c:axId val="200"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="0" sourceLinked="0"/><c:crossAx val="100"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')
    title = "<c:title><c:tx>" + text_body("图表 " + kind, tag="c:rich") + "</c:tx><c:overlay val=\"0\"/></c:title>"
    return XML + f'<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{R}"><c:lang val="zh-CN"/><c:chart>{title}<c:plotArea><c:layout/><c:{tags[kind]}>{settings}{body}</c:{tags[kind]}>{axes}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:externalData r:id="data"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>'


def add_chart(deck, slide, index, kind):
    categories, series = ["第一季度", "第二季度", "第三季度"], [("交付", [2, 5, 8])]
    chart_path, data_path = f"ppt/charts/chart{index}.xml", f"ppt/embeddings/data{index}.xlsx"
    deck.part(chart_path, chart_xml(kind, categories, series), "chart")
    deck.part(data_path, workbook(categories, series), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    deck.relate(chart_path, "data", "package", data_path)
    deck.relate(slide, f"chart{index}", "chart", chart_path)


def add_notes(deck, slide, index=1):
    if deck.notes_master is None:
        deck.notes_master = "ppt/notesMasters/notesMaster1.xml"
        deck.part(deck.notes_master, f'<p:notesMaster {NS}><p:cSld>{tree("")}</p:cSld>{COLOR_MAP}<p:notesStyle>{DEFAULT_STYLE}</p:notesStyle></p:notesMaster>', "notesMaster")
        deck.relate(deck.notes_master, "theme", "theme", "ppt/theme/theme1.xml")
    path = f"ppt/notesSlides/notesSlide{index}.xml"
    body = shape(2, None, x=48, y=360, width=444, height=240, placeholder=("body", 1),
                 paragraphs=paragraph(content=run("演讲者备注：", bold=True) + run("介绍上线、风险与培训。")))
    deck.part(path, f'<p:notes {NS}><p:cSld>{tree(body)}</p:cSld>{COLOR_OVERRIDE}</p:notes>', "notes")
    deck.relate(path, "slide", "slide", slide)
    deck.relate(path, "master", "notesMaster", deck.notes_master)
    deck.relate(slide, "notes", "notesSlide", path)


def add_classic_comments(deck, slide):
    authors = "ppt/commentAuthors.xml"
    deck.part(authors, f'<p:cmAuthorLst xmlns:p="{P}"><p:cmAuthor id="0" name="公开合成作者" initials="SY" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>', "commentAuthors")
    deck.relate("ppt/presentation.xml", "authors-classic", "commentAuthors", authors)
    path = "ppt/comments/comment1.xml"
    deck.part(path, f'<p:cmLst xmlns:p="{P}"><p:cm authorId="0" dt="{STAMP}" idx="1"><p:pos x="576" y="576"/><p:text>经典批注：补充交付细节。</p:text></p:cm></p:cmLst>', "comments")
    deck.relate(slide, "comments-classic", "comments", path)


def add_modern_comments(deck, slide, index, anchor="slide", resolved=False):
    author_id = "{11111111-1111-4111-8111-111111111111}"
    reply_author_id = "{22222222-2222-4222-8222-222222222222}"
    authors = "ppt/authors.xml"
    if authors not in deck.parts:
        deck.part(authors, f'<p188:authorLst xmlns:p188="{MODERN}"><p188:author id="{author_id}" name="公开审阅者" initials="SY" userId="synthetic-reviewer" providerId="synthetic"/><p188:author id="{reply_author_id}" name="公开作者" initials="ZZ" userId="synthetic-author" providerId="synthetic"/></p188:authorLst>', "authors")
        deck.relate("ppt/presentation.xml", "authors-modern", MODERN_REL + "authors", authors)
    document = f'<pc:docMk/><pc:sldMk sldId="{255 + index}"/>'
    moniker = f"<pc:sldMkLst>{document}</pc:sldMkLst>"
    if anchor == "object":
        moniker = f'<ac:deMkLst>{document}<ac:spMk id="2"/></ac:deMkLst>'
    if anchor == "text":
        moniker = f'<ac:txMkLst>{document}<ac:spMk id="2"/><ac:txMk cp="2" len="4"/></ac:txMkLst>'
    body = text_body("现代批注：" + anchor, tag="p188:txBody")
    reply = f'<p188:reply id="{{44444444-4444-4444-8444-{index:012d}}}" authorId="{reply_author_id}" created="{STAMP}">' + text_body("回复：已记录。", tag="p188:txBody") + "</p188:reply>"
    status = "resolved" if resolved else "active"
    comment = f'<p188:cm id="{{33333333-3333-4333-8333-{index:012d}}}" authorId="{author_id}" created="{STAMP}" status="{status}">{moniker}<p188:pos x="0" y="0"/><p188:replyLst>{reply}</p188:replyLst>{body}</p188:cm>'
    path = f"ppt/comments/modernComment{index}.xml"
    deck.part(path, f'<p188:cmLst xmlns:p188="{MODERN}" xmlns:a="{A}" xmlns:pc="{PC}" xmlns:ac="{AC}">{comment}</p188:cmLst>', "modernComments")
    deck.relate(slide, "comments-modern", MODERN_REL + "comments", path)
    extension = f'<p:extLst><p:ext uri="{MODERN_EXT}"><p188:commentRel xmlns:p188="{MODERN}" r:id="comments-modern"/></p:ext></p:extLst>'
    deck.parts[slide] = deck.parts[slide].replace("</p:sld>", extension + "</p:sld>")


def cases():
    result = []
    def add(name, deck, text, **extra):
        result.append((name, deck.files(), {"text": text, "slides": len(deck.slides), **extra}))

    deck = Deck("中文与富文本")
    fonts = '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/><a:cs typeface="Liberation Sans"/>'
    paragraphs = (paragraph(content=run("中文与 Latin 😀，", fonts) + run("粗体", bold=True) + run("和斜体", italic=True))
                  + paragraph("第一项", '<a:buAutoNum type="arabicPeriod" startAt="1"/>')
                  + paragraph("第二级", '<a:buChar char="•"/>')
                  + paragraph(content=run("访问示例", '<a:hlinkClick r:id="link"/>'))
                  + paragraph(content=run("换行前") + "<a:br/>" + run("换行后"))
                  + paragraph("扩展字𠀀与变体葛󠄀"))
    slide = deck.slide(shape(2, None, paragraphs=paragraphs, height=420))
    deck.relate(slide, "link", "hyperlink", "https://example.com/pptx", True)
    add("01-rich-text", deck, ["中文与 Latin 😀，粗体和斜体", "第一项", "第二级", "访问示例", "换行前", "换行后", "扩展字𠀀与变体葛󠄀"], objects=1)

    deck = Deck("主题与占位符")
    parent = shape(2, "母版标题占位符", x=48, y=30, width=864, height=90, placeholder=("title", 0))
    layout = shape(2, "布局标题占位符", x=60, y=40, width=840, height=100, placeholder=("title", 0))
    index = deck.add_master("C0504D", parent, layout, "title")
    deck.slide(shape(2, "继承主题标题", placeholder=("title", 0), no_transform=True,
                     paragraphs=paragraph(content=run("继承主题标题", '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/>'))), layout=index)
    add("02-theme-placeholders", deck, ["继承主题标题"], inheritedGeometry={"x": 60, "y": 40, "width": 840, "height": 100}, masters=2)

    deck = Deck("多母版和布局")
    second = deck.add_master("70AD47", shape(2, "第二母版", x=48, y=460, height=40))
    deck.slide(shape(2, "第一主题页面"))
    deck.slide(shape(2, "第二主题页面", paragraphs=paragraph(content=run("第二主题页面", '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'))), layout=second)
    add("03-masters-layouts", deck, ["第一主题页面", "第二主题页面"], masters=2)

    deck = Deck("基本形状与外观")
    presets = ["rect", "roundRect", "ellipse", "triangle", "diamond", "parallelogram", "trapezoid", "pentagon", "hexagon", "rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "chevron", "line"]
    shapes = ""
    for i, preset in enumerate(presets):
        fill = solid("4472C4", 75000) if i % 2 == 0 else solid("accent2", scheme=True)
        if i == 1:
            fill = '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="4472C4"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>'
        effects = '<a:effectLst><a:outerShdw blurRad="38100" dist="25400" dir="2700000" rotWithShape="1"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>' if i == 0 else ""
        shapes += shape(i + 2, preset, x=40 + i % 4 * 225, y=30 + i // 4 * 125, width=195, height=95,
                        preset=preset, fill=fill, outline=line(dashed=i % 2 == 1), effects=effects, text_box=False)
    deck.slide(shapes)
    add("04-shapes", deck, presets, objects=len(presets), presets=presets)

    deck = Deck("嵌套组合")
    nested = group(4, shape(5, "组内文字", x=-20, y=0, width=160, height=70), x=180, y=50, width=240, height=140, child=(0, 0, 240, 140), rotation=15)
    deck.slide(group(2, shape(3, "局部坐标", x=0, y=0, width=160, height=80) + nested, child=(-50, -25, 900, 450), rotation=10, flip_h=True))
    add("05-groups", deck, ["局部坐标", "组内文字"], objects=4)

    deck = Deck("连接线")
    content = shape(2, "起点", x=48, y=100, width=180, height=100, preset="roundRect", fill=solid(), text_box=False)
    content += shape(3, "终点", x=600, y=320, width=180, height=100, preset="ellipse", fill=solid("70AD47"), text_box=False)
    deck.slide(content + connector(4, 2, 3) + connector(5, 2, 3, x=230, y=210, width=360, height=160, elbow=True))
    add("06-connectors", deck, ["起点", "终点"], objects=4)

    deck = Deck("图片与裁剪")
    slide = deck.slide(shape(2, "PNG JPEG WebP") + "".join(picture(3 + i, "image" + str(i), x=48 + 270 * i, y=180,
                      crop=(.1, .1, .1, .1) if i == 0 else None, rotation=10 if i == 1 else 0, flip_h=i == 2) for i in range(3)))
    for i, (suffix, mime) in enumerate([("png", "image/png"), ("jpg", "image/jpeg"), ("webp", "image/webp")]):
        path = "ppt/media/sample." + suffix
        deck.part(path, (ROOT / "examples/fidelity/assets" / ("sample." + suffix)).read_bytes(), mime)
        deck.relate(slide, "image" + str(i), "image", path)
    add("07-images", deck, ["PNG JPEG WebP"], objects=4, images=3)

    deck = Deck("原生合并表格")
    deck.slide(shape(2, "合并表格") + graphic_frame(3, table(), A.removesuffix("/main") + "/table", width=720, height=240))
    add("08-table-merges", deck, ["合并表格", "合并标题", "第三列", "跨行", "甲", "乙", "丙", "丁"], objects=2)

    deck = Deck("三类基础图表")
    for i, kind in enumerate(["column", "line", "pie"], 1):
        slide = deck.slide(shape(2, "图表 " + kind) + graphic_frame(3, f'<c:chart xmlns:c="{C}" r:id="chart{i}"/>', C, width=760, height=340))
        add_chart(deck, slide, i, kind)
    add("09-charts", deck, ["图表 column", "图表 line", "图表 pie"], charts=3, data={"categories": ["第一季度", "第二季度", "第三季度"], "values": [2, 5, 8]})

    deck = Deck("演讲者备注")
    slide = deck.slide(shape(2, "带备注的幻灯片"))
    add_notes(deck, slide)
    add("10-notes", deck, ["带备注的幻灯片"], notes=["演讲者备注：介绍上线、风险与培训。"])

    deck = Deck("经典批注")
    slide = deck.slide(shape(2, "经典批注来源"))
    add_classic_comments(deck, slide)
    add("11-classic-comments", deck, ["经典批注来源"], classicComments=1, commentText=["经典批注：补充交付细节。"])

    deck = Deck("现代批注与回复")
    for i, anchor in enumerate(["slide", "object", "text"], 1):
        value = "中文𠀀😀批注锚点" if anchor == "text" else "现代批注 " + anchor
        slide = deck.slide(shape(2, paragraphs=paragraph(content=run(value, fonts))))
        add_modern_comments(deck, slide, i, anchor, resolved=i == 2)
    add("12-modern-comments", deck, ["现代批注 slide", "现代批注 object", "中文𠀀😀批注锚点"],
        modernComments=3, replies=3, textAnchor={"cp": 2, "length": 4, "quote": "𠀀😀"},
        referenceText=["现代批注 slide", "现代批注 object", "中文", "😀", "批注锚点"],
        referenceLimits=["LibreOffice 25.8 omits U+20000 from PDF text and reorders the emoji text run; original XML and UTF-16 anchor remain exact."])

    deck = Deck("多页组合文稿")
    first = deck.slide(shape(2, "项目进展报告"))
    add_notes(deck, first)
    add_classic_comments(deck, first)
    chart_slide = deck.slide(shape(2, "交付数据") + graphic_frame(3, '<c:chart xmlns:c="' + C + '" r:id="chart1"/>', C, width=720, height=320))
    add_chart(deck, chart_slide, 1, "column")
    add_modern_comments(deck, chart_slide, 2, "object")
    deck.slide(shape(2, "隐藏的补充页"), hidden=True)
    add("13-multi-slide", deck, ["项目进展报告", "交付数据", "隐藏的补充页"], hiddenSlides=1, charts=1, classicComments=1, modernComments=1, replies=1)

    base = Deck("命名空间与关系")
    base.slide(shape(2, "命名空间不依赖前缀"))
    strict = {}
    aliases = {P: "http://purl.oclc.org/ooxml/presentationml/main", A: "http://purl.oclc.org/ooxml/drawingml/main", R: "http://purl.oclc.org/ooxml/officeDocument/relationships"}
    for path, value in base.files().items():
        if isinstance(value, str):
            for old, new in aliases.items():
                value = value.replace(old, new)
        strict[path] = value
    result.append(("15-strict-namespaces", strict, {"text": ["命名空间不依赖前缀"], "slides": 1, "objects": 1}))
    prefixed = {}
    for path, value in base.files().items():
        if isinstance(value, str):
            value = value.replace("xmlns:p=", "xmlns:slide=").replace("xmlns:a=", "xmlns:draw=")
            value = re.sub(r"(<\/?|\s)p:", r"\1slide:", value)
            value = re.sub(r"(<\/?|\s)a:", r"\1draw:", value)
        prefixed[path] = value
    result.append(("16-alternate-prefixes", prefixed, {"text": ["命名空间不依赖前缀"], "slides": 1, "objects": 1}))
    renamed = rename_parts(base.files(), {
        "ppt/presentation.xml": "office/main.xml", "ppt/slides/slide1.xml": "canvases/page-one.xml",
        "ppt/slideMasters/slideMaster1.xml": "definitions/master.xml",
        "ppt/slideLayouts/slideLayout1.xml": "definitions/layout.xml", "ppt/theme/theme1.xml": "definitions/theme.xml",
    })
    result.append(("17-renamed-parts", renamed, {"text": ["命名空间不依赖前缀"], "slides": 1, "objects": 1}))
    return result


def rename_parts(files, mapping):
    result = {}
    for path, value in files.items():
        new_path = mapping.get(path, path)
        if path.endswith(".rels"):
            source = "" if path == "_rels/.rels" else posixpath.dirname(path)[:-5] + posixpath.basename(path)[:-5]
            target_source = mapping.get(source, source)
            if target_source:
                directory, name = posixpath.split(target_source)
                new_path = (directory + "/" if directory else "") + "_rels/" + name + ".rels"
            root = ET.fromstring(value)
            for relationship in root:
                if relationship.get("TargetMode") == "External":
                    continue
                old_target = unquote(relationship.get("Target", ""))
                absolute = old_target.lstrip("/") if old_target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(source), old_target))
                target = mapping.get(absolute, absolute)
                relative = posixpath.relpath(target, posixpath.dirname(target_source) or ".")
                relationship.set("Target", "/" + quote(target, safe="/") if source == "" else quote(relative, safe="/"))
            ET.register_namespace("", PKG)
            value = XML + ET.tostring(root, encoding="unicode")
        elif path == "[Content_Types].xml":
            root = ET.fromstring(value)
            for override in root:
                if override.tag == "{" + CT + "}Override":
                    original = unquote(override.attrib["PartName"].lstrip("/"))
                    override.set("PartName", "/" + quote(mapping.get(original, original), safe="/"))
            ET.register_namespace("", CT)
            value = XML + ET.tostring(root, encoding="unicode")
        result[new_path] = value
    return result


def negative_cases(valid):
    sources = {name: files for name, files, _ in valid}
    base = deepcopy(sources["01-rich-text"])
    result = []
    def add(name, parts, code, reason):
        result.append((name, zipped(parts), {"error": code, "reason": reason}))
    missing = deepcopy(base)
    del missing["ppt/slides/slide1.xml"]
    add("missing-slide", missing, "InvalidDocument", "An internal relationship targets a missing part.")
    duplicate = deepcopy(base)
    duplicate["ppt/slides/slide1.xml"] = duplicate["ppt/slides/slide1.xml"].replace("</p:spTree>", shape(2, "重复原生身份") + "</p:spTree>")
    add("duplicate-native-id", duplicate, "InvalidDocument", "Native drawing identities repeat inside one slide.")
    unsafe_xml = deepcopy(base)
    unsafe_xml["ppt/slides/slide1.xml"] = unsafe_xml["ppt/slides/slide1.xml"].replace(XML, XML + '<!DOCTYPE p:sld [<!ENTITY private SYSTEM "file:///not-authorized">]>')
    add("doctype", unsafe_xml, "InvalidDocument", "DOCTYPE and external entities are prohibited.")
    scripted = deepcopy(base)
    scripted["ppt/slides/slide1.xml"] = scripted["ppt/slides/slide1.xml"].replace("</p:spTree>", '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="unsupported"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' + transform(0, 0, 100, 100, tag="p:xfrm") + '<a:graphic><a:graphicData uri="unsupported-smartart"/></a:graphic></p:graphicFrame></p:spTree>')
    add("unsupported-object", scripted, "UnsupportedFeature", "An unimplemented visible graphic cannot be silently dropped.")
    animated = deepcopy(base)
    animated["ppt/slides/slide1.xml"] = animated["ppt/slides/slide1.xml"].replace("</p:sld>", '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite"/></p:par></p:tnLst></p:timing></p:sld>')
    add("animation", animated, "UnsupportedFeature", "Animation is outside the declared static scope.")
    external = deepcopy(sources["07-images"])
    external["ppt/slides/slide1.xml"] = external["ppt/slides/slide1.xml"].replace('r:embed="image0"', 'r:link="image0"')
    relation = "ppt/slides/_rels/slide1.xml.rels"
    root = ET.fromstring(external[relation])
    for item in root:
        if item.get("Id") == "image0":
            item.set("Target", "https://example.com/unauthorized.png")
            item.set("TargetMode", "External")
    external[relation] = XML + ET.tostring(root, encoding="unicode")
    add("external-image", external, "UnsupportedFeature", "External images cannot trigger resource or network access.")
    formula = deepcopy(sources["09-charts"])
    with ZipFile(BytesIO(formula["ppt/embeddings/data1.xlsx"])) as archive:
        workbook_files = {name: archive.read(name) for name in archive.namelist()}
    worksheet = workbook_files["xl/worksheets/sheet1.xml"].decode()
    workbook_files["xl/worksheets/sheet1.xml"] = worksheet.replace('<c r="B2"><v>2</v></c>', '<c r="B2"><f>1+1</f><v>2</v></c>')
    formula["ppt/embeddings/data1.xlsx"] = zipped(workbook_files)
    add("chart-formula", formula, "UnsupportedFeature", "Cell formulas are not graph data constants; range references remain supported.")
    external_chart = deepcopy(sources["09-charts"])
    relation = "ppt/charts/_rels/chart1.xml.rels"
    external_chart[relation] = relationships([("data", "package", "https://example.com/unauthorized.xlsx", True)])
    add("external-chart-data", external_chart, "UnsupportedFeature", "Chart data cannot load an external workbook.")
    unknown_author = deepcopy(sources["12-modern-comments"])
    unknown_author["ppt/comments/modernComment1.xml"] = unknown_author["ppt/comments/modernComment1.xml"].replace('authorId="{11111111-1111-4111-8111-111111111111}"', 'authorId="{99999999-9999-4999-8999-999999999999}"')
    add("unknown-comment-author", unknown_author, "InvalidDocument", "A modern comment references an undefined author.")
    result.append(("truncated-zip", zipped(base)[:-22], {"error": "InvalidDocument", "reason": "ZIP directory is missing."}))
    return result


def generate(output):
    output.mkdir(parents=True, exist_ok=True)
    manifest = {"format": "dagent-office-pptx-fixtures", "version": 1, "generator": "Independent Python ZIP/XML, no SDK exporter or third-party Office library.", "cases": [], "negativeCases": []}
    valid = cases()
    for name, files, expected in valid:
        for path, data in files.items():
            if path.endswith(".xml") or path.endswith(".rels"):
                ET.fromstring(data)
        data = zipped(files)
        filename = name + ".pptx"
        (output / filename).write_bytes(data)
        manifest["cases"].append({"file": filename, "sha256": hashlib.sha256(data).hexdigest(), **expected})
    invalid = output / "invalid"
    invalid.mkdir(exist_ok=True)
    for name, data, expected in negative_cases(valid):
        filename = name + ".pptx"
        (invalid / filename).write_bytes(data)
        manifest["negativeCases"].append({"file": "invalid/" + filename, "sha256": hashlib.sha256(data).hexdigest(), **expected})
    (output / "expected.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
    return manifest


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "examples/pptx")
    args = parser.parse_args()
    manifest = generate(args.output)
    print(f'Generated {len(manifest["cases"])} independent PPTX sources and {len(manifest["negativeCases"])} negative inputs.')
