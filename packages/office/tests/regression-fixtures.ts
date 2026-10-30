import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocumentSession, openDocument } from "@dagent-office/core";
import type { DocumentMetadata } from "@dagent-office/core";
import { importDocx, writeDocumentZip } from "@dagent-office/office";

export function mutateDocx(
  name: string,
  changes: Record<string, (source: string) => string>,
): Uint8Array {
  const files = unzipSync(new Uint8Array(readFileSync(
    new URL("../../../examples/docx/" + name, import.meta.url),
  )));
  for (const [part, change] of Object.entries(changes))
    files[part] = strToU8(change(files[part] ? strFromU8(files[part]) : ""));
  return zipSync(files);
}

export function editedChineseFontFixture(): Uint8Array {
  const model = openDocument(importDocx(mutateDocx("09-chinese-fonts.docx", {}), { documentId: "edited-fonts" }));
  const ids = [...model.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => match[1]!);
  const target = ids.map((id) => model.getNode(id)!).find((node) => node.attributes["data-word-run"] === "true" && node.text.includes("ABC"))!;
  const session = createDocumentSession(model);
  session.commit(session.prepare({ protocolVersion: 1, documentId: "edited-fonts", baseRevision: 0, operations: [
    { op: "replaceText", nodeId: target.nodeId, from: 0, to: 0, expectedText: "", text: "新增仿宋ABC𠀀\u{e0100}" },
    { op: "setStyle", nodeId: target.nodeId, property: "--word-font-eastAsia", value: '"FandolFang"' },
  ] }));
  const result = writeDocumentZip(session.serialize());
  session.dispose();
  return result;
}

export const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const styles = (content: string): string =>
  '<w:styles xmlns:w="' + wordNamespace + '">' + content + "</w:styles>";
export const body = (source: string, content: string): string =>
  source.replace(/<w:body>[\s\S]*<\/w:body>/, "<w:body>" + content + "</w:body>");
export const complexField = (instruction: string, result: string): string =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  "<w:r><w:instrText>" + instruction + "</w:instrText></w:r>" +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + result +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

export function crossSectionRangesFixture(instructionOpen = false): Uint8Array {
  const auxiliary = (name: string, commentId: number): string =>
    '<w:bookmarkStart w:id="0" w:name="' + name + 'Mark"/>' +
    '<w:commentRangeStart w:id="' + commentId + '"/>' +
    complexField(" PAGE ", '<w:r><w:t>' + name + ' cache</w:t></w:r>') +
    '<w:commentRangeEnd w:id="' + commentId + '"/><w:bookmarkEnd w:id="0"/>';
  return mutateDocx("04-styles-sections.docx", {
    "word/document.xml": (source) => {
      const section = source.match(/<w:sectPr>[\s\S]*?<\/w:sectPr>/)![0];
      return body(source,
        '<w:p><w:pPr>' + section + '</w:pPr>' +
        '<w:bookmarkStart w:id="0" w:name="mainMark"/><w:commentRangeStart w:id="0"/>' +
        '<w:r><w:fldChar w:fldCharType="begin"/><w:instrText> REF </w:instrText>' +
        (instructionOpen ? '' : '<w:fldChar w:fldCharType="separate"/><w:t>main before</w:t>') +
        '</w:r></w:p><w:p><w:r>' +
        (instructionOpen ? '<w:instrText> mainMark </w:instrText><w:fldChar w:fldCharType="separate"/>' : '') +
        '<w:t>main after</w:t><w:fldChar w:fldCharType="end"/></w:r>' +
        '<w:commentRangeEnd w:id="0"/><w:bookmarkEnd w:id="0"/></w:p>');
    },
    "word/header1.xml": () => '<w:hdr xmlns:w="' + wordNamespace + '"><w:p>' + auxiliary("header", 1) + '</w:p></w:hdr>',
    "word/footer1.xml": () => '<w:ftr xmlns:w="' + wordNamespace + '"><w:p>' + auxiliary("footer", 2) + '</w:p></w:ftr>',
    "word/comments.xml": () => '<w:comments xmlns:w="' + wordNamespace + '">' +
      [0, 1, 2].map((id) => '<w:comment w:id="' + id + '"><w:p><w:r><w:t>review</w:t></w:r></w:p></w:comment>').join('') +
      '</w:comments>',
    "word/_rels/document.xml.rels": (source) => source.replace('</Relationships>',
      '<Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'),
    "[Content_Types].xml": (source) => source.replace('</Types>',
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'),
  });
}

export function stylesheetPackageFixture(css: string): Uint8Array {
  const files = new Map(importDocx(mutateDocx("06-links-tabs.docx", {}), { documentId: "stylesheet" }));
  files.set("styles.css", strToU8(strFromU8(files.get("styles.css")!) + "\n" + css));
  return writeDocumentZip(files);
}

export function indentedTabsFixture(): Uint8Array {
  const stops = '<w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="left" w:pos="2880"/></w:tabs>';
  const paragraph = (indent: string, text: string, tabs = stops): string =>
    '<w:p><w:pPr><w:ind ' + indent + '/>' + tabs + '</w:pPr><w:r><w:tab/><w:t>' + text + '</w:t></w:r></w:p>';
  return mutateDocx("06-links-tabs.docx", {
    "word/styles.xml": () => styles('<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults>'),
    "word/document.xml": (source) => body(source,
      paragraph('w:left="720"', 'custom') +
      paragraph('w:left="720"', 'default', '') +
      paragraph('w:left="1800"', 'skip passed stop') +
      paragraph('w:left="720" w:firstLine="360"', 'first line') +
      paragraph('w:left="720" w:hanging="360"', 'hanging') +
      paragraph('w:left="-360"', 'negative indent')),
  });
}

export function nestedFieldFixture(simple = false): Uint8Array {
  return mutateDocx("03-images.docx", {
    "word/document.xml": (source) => {
      const drawing = source.match(/<w:drawing>[\s\S]*?<\/w:drawing>/)![0];
      const cached = "<w:r><w:t>cache</w:t>" +
        (simple ? "<w:br/>" : '<w:br w:type="page"/>') + drawing + "<w:tab/></w:r>" +
        '<w:fldSimple w:instr=" REF child "><w:r><w:t>child cache</w:t></w:r></w:fldSimple>';
      const inner = simple
        ? '<w:fldSimple w:instr=" REF cache ">' + cached + "</w:fldSimple>"
        : complexField(" REF cache ", cached);
      const outer = '<w:r><w:fldChar w:fldCharType="begin"/><w:instrText> IF </w:instrText></w:r>' +
        inner + '<w:r><w:instrText> = 1 </w:instrText><w:fldChar w:fldCharType="separate"/>' +
        '<w:t>outer result</w:t><w:fldChar w:fldCharType="end"/></w:r>';
      return body(source, "<w:p>" + outer + "</w:p>");
    },
  });
}

export function styleInheritanceFixture(): Uint8Array {
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () => styles(
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Noto Sans CJK SC" w:eastAsia="Noto Sans CJK SC"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:pPr><w:tabs><w:tab w:pos="2000" w:val="left"/></w:tabs></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/><w:rPr><w:rFonts w:eastAsia="Noto Serif CJK SC"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Character"><w:rPr><w:highlight w:val="green"/></w:rPr></w:style>' +
      '<w:style w:type="table" w:styleId="TableBase"><w:tblPr><w:tblInd w:w="600" w:type="dxa"/><w:tblCellMar><w:left w:w="600"/><w:top w:w="40"/></w:tblCellMar></w:tblPr></w:style>' +
      '<w:style w:type="table" w:styleId="TableDerived"><w:basedOn w:val="TableBase"/><w:tblPr><w:tblCellMar><w:right w:w="800"/></w:tblCellMar></w:tblPr></w:style>' +
      '<w:style w:type="character" w:styleId="Shaded"><w:rPr><w:highlight w:val="yellow"/><w:shd w:val="clear" w:fill="ddeeff"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="ShadedClear"><w:basedOn w:val="Shaded"/><w:rPr><w:highlight w:val="none"/></w:rPr></w:style>',
    ),
    "word/document.xml": (source) => {
      const table = source.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)![0]
        .replace("<w:tblPr>", '<w:tblPr><w:tblStyle w:val="TableDerived"/><w:tblCellMar><w:bottom w:w="60"/></w:tblCellMar>');
      return body(source,
        '<w:p><w:pPr><w:pStyle w:val="Derived"/><w:tabs><w:tab w:pos="4000" w:val="left"/></w:tabs></w:pPr>' +
        '<w:r><w:rPr><w:rFonts w:eastAsia="Noto Serif CJK SC"/><w:highlight w:val="none"/></w:rPr><w:t>ABCé中文</w:t></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="Character"/><w:highlight w:val="none"/></w:rPr><w:t>cleared</w:t></w:r>' +
        '<w:r><w:tab/><w:t>tab result</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:rStyle w:val="Shaded"/><w:highlight w:val="none"/></w:rPr><w:t>shaded direct</w:t></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="ShadedClear"/></w:rPr><w:t>shaded style</w:t></w:r></w:p>' + table);
    },
  });
}

export function sectionPackageFixture(articleRoot: boolean): Uint8Array {
  const files = new Map(importDocx(mutateDocx("04-styles-sections.docx", {}), {
    documentId: "section-package",
  }));
  const metadata = JSON.parse(strFromU8(files.get("metadata.json")!)) as DocumentMetadata;
  const setup = { ...metadata.pageSetup, width: 841.9 };
  metadata.sections[0]!.pageSetup = setup;
  delete metadata.sections[1]!.pageSetup;
  if (articleRoot) {
    delete metadata.word;
    const html = strFromU8(files.get("content.html")!);
    metadata.sections = [{
      sectionId: "article",
      rootNodeId: html.match(/<article[^>]*data-node-id="([^"]+)"/)![1]!,
      pageSetup: setup,
    }];
  }
  files.set("metadata.json", strToU8(JSON.stringify(metadata)));
  return writeDocumentZip(files);
}
