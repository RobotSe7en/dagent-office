import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
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
    files[part] = strToU8(change(strFromU8(files[part]!)));
  return zipSync(files);
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
