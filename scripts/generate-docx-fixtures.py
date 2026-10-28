"""Independent, deterministic ECMA-376 fixtures; does not call the SDK exporter.
Reference PDFs are generated separately by LibreOffice, never by this generator.
"""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.sax.saxutils import escape
import json

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'examples/docx'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
NS = f'xmlns:w="{W}" xmlns:r="{R}" xmlns:a="{A}" xmlns:wp="{WP}" xmlns:pic="{PIC}"'

def run(value, props=''):
    return f'<w:r>{"<w:rPr>"+props+"</w:rPr>" if props else ""}<w:t xml:space="preserve">{escape(value)}</w:t></w:r>'

def para(value='', props='', content=None):
    return f'<w:p>{"<w:pPr>"+props+"</w:pPr>" if props else ""}{run(value) if content is None else content}</w:p>'

PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
STYLES = f'''<w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Noto Sans CJK SC" w:hAnsi="Noto Sans CJK SC" w:eastAsia="Noto Sans CJK SC"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1f4b99"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="240"/></w:pPr><w:rPr><w:b/><w:color w:val="123456"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Base"/><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style></w:styles>'''
NUMBERING = f'''<w:numbering {NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/><w:suff w:val="space"/><w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride></w:num></w:numbering>'''

def numbering(level=0, instance=1):
    return f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{instance}"/></w:numPr>'

def cell(value, props=''):
    return f'<w:tc><w:tcPr>{props}</w:tcPr>{para(value)}</w:tc>'

def picture(rid, index):
    return f'''<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="{index}" name="Sample image {index}" descr="测试图片"/><a:graphic><a:graphicData uri="{PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="sample"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'''

CASES=[]
def case(filename, body, expected, extras=None, rels=None, final_section=PAGE):
    CASES.append((filename,body,expected,extras or {},rels or [],final_section))

case('01-text-numbering', para('实施报告', '<w:pStyle w:val="Heading1"/>') + para(content=run('中文与 emoji 😀，')+run('粗体', '<w:b/>')+run('和斜体','<w:rStyle w:val="Emphasis"/>')) + para('第一项', numbering()) + para('子项',numbering(1)) + para('第二项',numbering()) + para('从四开始',numbering(0,2)), ['实施报告','中文与 emoji 😀，粗体和斜体','第一项','子项','第二项','从四开始'])
borders=''.join(f'<w:{side} w:val="single" w:sz="8" w:color="334455"/>' for side in ['top','left','bottom','right','insideH','insideV'])
table=f'<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>{cell("合并标题", "<w:gridSpan w:val=\"2\"/>")}{cell("第三列")}</w:tr><w:tr>{cell("纵向合并", "<w:vMerge w:val=\"restart\"/>")}{cell("甲")}{cell("乙")}</w:tr><w:tr>{cell("", "<w:vMerge/>")}{cell("丙")}{cell("丁")}</w:tr></w:tbl>'
case('02-table-merges',para('合并表格')+table,['合并标题','第三列','纵向合并','甲','乙','丙','丁'])
images={f'word/media/sample.{ext}':(ROOT/f'examples/fidelity/assets/sample.{ext}').read_bytes() for ext in ['png','jpg','webp']}
case('03-images',para('图片样例')+para(content=''.join(picture(f'image{i}',i) for i in range(1,4))),['图片样例'],images,[(f'image{i}','image',f'media/sample.{ext}',False) for i,ext in enumerate(['png','jpg','webp'],1)])
header=f'<w:hdr {NS}>{para("报告页眉")}</w:hdr>'
footer=f'<w:ftr {NS}>{para("报告页脚")}</w:ftr>'
section='<w:headerReference w:type="default" r:id="header1"/><w:footerReference w:type="default" r:id="footer1"/>'+PAGE
case('04-styles-sections',para('继承样式','<w:pStyle w:val="Derived"/>')+para(content=run('直接覆盖颜色','<w:color w:val="cc0000"/>'),props='<w:pStyle w:val="Derived"/>')+para('第一节结束',f'<w:sectPr>{section}</w:sectPr>')+para('第二节正文'),['继承样式','直接覆盖颜色','第一节结束','第二节正文','报告页眉','报告页脚'],{'word/header1.xml':header,'word/footer1.xml':footer},[('header1','header','header1.xml',False),('footer1','footer','footer1.xml',False)],PAGE)
case('05-long-page-break',para('长文档')+''.join(para(f'第{i}段：固定字体与明确分页验证。') for i in range(1,41))+para(content=run('明确分页前')+'<w:r><w:br w:type="page"/></w:r>'+run('明确分页后')),['长文档','第40段：固定字体与明确分页验证。','明确分页前','明确分页后'])
links='<w:bookmarkStart w:id="0" w:name="目标"/>'+run('书签目标')+'<w:bookmarkEnd w:id="0"/>'
links+='<w:hyperlink w:anchor="目标">'+run('内部链接')+'</w:hyperlink><w:hyperlink r:id="link1">'+run('外部链接')+'</w:hyperlink>'
tabcontent=run('左侧')+'<w:r><w:tab/></w:r>'+run('右侧')
case('06-links-tabs',para(content=links)+para(content=tabcontent,props='<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="5000"/></w:tabs>'),['书签目标','内部链接','外部链接','左侧','右侧'],rels=[('link1','hyperlink','https://example.com/document',True)])
def control(kind):return f'<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>'
def instruction(value):return '<w:r><w:instrText xml:space="preserve">'+escape(value)+'</w:instrText></w:r>'
fields=para(content='<w:fldSimple w:instr=" PAGE ">'+run('3')+'</w:fldSimple>')
fields+=para(content=control('begin')+instruction(' NUMPAGES ')+control('separate')+run('7')+control('end'))
fields+=para(content=control('begin')+instruction(' IF ')+control('begin')+instruction(' PAGE ')+control('separate')+run('3')+control('end')+instruction(' = 3 "正确" "错误" ')+control('separate')+run('正确')+control('end'))
fields+=para(content=control('begin')+instruction(' TOC \\o "1-3" ')+control('separate')+run('缓存目录第一项'))+para(content=run('缓存目录第二项')+control('end'))
case('07-fields',fields,['3','7','正确','缓存目录第一项','缓存目录第二项'])
comment=f'<w:comments {NS}><w:comment w:id="0" w:author="审阅者" w:initials="SY" w:date="2026-01-01T08:00:00Z">{para(content=run("请修改", "<w:b/>")+run("时间安排"))}</w:comment></w:comments>'
review=para(content='<w:commentRangeStart w:id="0"/>'+run('第三季度')+run('完成项目')+'<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>')
review+=para(content='<w:del w:id="1" w:author="编辑者" w:date="2026-01-02T08:00:00Z"><w:r><w:delText>旧安排</w:delText></w:r></w:del><w:ins w:id="2" w:author="编辑者" w:date="2026-01-02T08:00:00Z">'+run('新安排')+'</w:ins>')
case('08-comments-revisions',review,['第三季度完成项目','旧安排','新安排','请修改时间安排'],{'word/comments.xml':comment},[('comments','comments','comments.xml',False)])

case('09-chinese-fonts',para(content=run('中西文混排 ABCxyz，全角标点。（测试）𠀀', '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:eastAsia="Noto Serif CJK SC"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/>'))+para(content=run('缺失中文字体回退 ABC', '<w:rFonts w:ascii="Liberation Serif" w:eastAsia="Dagent Office Missing CJK Test Font"/>')),['中西文混排 ABCxyz，全角标点。（测试）𠀀','缺失中文字体回退 ABC'])

OUT.mkdir(exist_ok=True)
manifest=[]
for filename,body,expected,extras,rels,section in CASES:
    files={'word/document.xml':f'<?xml version="1.0" encoding="UTF-8"?><w:document {NS}><w:body>{body}<w:sectPr>{section}</w:sectPr></w:body></w:document>', 'word/styles.xml':STYLES,'word/numbering.xml':NUMBERING,'word/settings.xml':f'<w:settings {NS}><w:defaultTabStop w:val="720"/></w:settings>',**extras}
    relationships=[('styles','styles','styles.xml',False),('numbering','numbering','numbering.xml',False),('settings','settings','settings.xml',False),*rels]
    files['word/_rels/document.xml.rels']='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{escape(target)}"'+(' TargetMode="External"' if external else '')+'/>' for rid,kind,target,external in relationships)+'</Relationships>'
    files['_rels/.rels']=f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>'
    type_overrides={'word/document.xml':'document.main','word/styles.xml':'styles','word/numbering.xml':'numbering','word/settings.xml':'settings'}
    for path in extras:
        if path.endswith('.xml'):
            type_overrides[path]='header' if '/header' in path else 'footer' if '/footer' in path else 'comments'
    files['[Content_Types].xml']='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="webp" ContentType="image/webp"/>'+''.join(f'<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>' for path,kind in type_overrides.items())+'</Types>'
    with ZipFile(OUT/(filename+'.docx'),'w') as archive:
        for path,data in sorted(files.items()):
            info=ZipInfo(path,date_time=(2026,1,1,0,0,0));info.compress_type=ZIP_DEFLATED
            archive.writestr(info,data.encode() if isinstance(data,str) else data)
    manifest.append({'file':filename+'.docx','requiredText':expected})
(OUT/'expected.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
print(f'Generated {len(CASES)} independent DOCX fixtures.')
