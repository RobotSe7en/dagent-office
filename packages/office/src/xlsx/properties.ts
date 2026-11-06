import { DocumentError } from "@dagent-office/core";
import type { WorkbookDocumentMetadata } from "@dagent-office/core";
import type { OpcPackage } from "../ooxml/package.js";
import { attr, check, elements, fail, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, integer, one, value } from "./values.js";

const CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", DC = "http://purl.org/dc/elements/1.1/", DCT = "http://purl.org/dc/terms/";
const EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", XSI = "http://www.w3.org/2001/XMLSchema-instance";
type Properties = NonNullable<WorkbookDocumentMetadata["properties"]>;
const coreFields: Readonly<Record<string, readonly string[]>> = {
  [CP]: ["lastModifiedBy", "revision", "category", "contentStatus", "contentType", "version", "keywords", "lastPrinted"],
  [DC]: ["title", "creator", "subject", "description", "identifier", "language"], [DCT]: ["created", "modified"],
};
const appFields: Readonly<Record<string, keyof Properties>> = { Template: "template", Manager: "manager", Company: "company", Application: "application", AppVersion: "applicationVersion" };
const counts = new Set(["Pages", "Words", "Characters", "Lines", "Paragraphs", "Slides", "Notes", "TotalTime", "HiddenSlides", "MMClips", "CharactersWithSpaces", "DocSecurity"]);
const flags = new Set(["ScaleCrop", "LinksUpToDate", "SharedDoc", "HyperlinksChanged"]);
function leaf(node: XmlNode): string {
  if (node.content.some(item => typeof item !== "string")) fail(node, "A document property must contain only text.");
  const result = text(node); if (result.length > 32767) throw new DocumentError("ResourceLimitExceeded", "A document property exceeds its text budget."); return result;
}
function numberValue(node: XmlNode): number { return integer({ ...node, attrs: new Map([[" val", leaf(node)]]) }, "val", 0, 0x7fff_ffff); }
/** App heading/title vectors are derived document statistics. Audit the
 * supported primitive shape before discarding those replaceable caches. */
function vector(parent: XmlNode, headings: boolean): void {
  check(parent, ["vector"], VT); const source = one(parent, "vector", VT);
  if (!source) fail(parent, "An application property is missing its vector.");
  attributes(source, ["size", "baseType"]); check(source, [headings ? "variant" : "lpstr"], VT);
  const items = elements(source, undefined, VT);
  if (integer(source, "size") !== items.length || value(source, "baseType") !== (headings ? "variant" : "lpstr") || headings && items.length % 2) fail(source, "An application-property vector has an invalid size or type.");
  for (const [index, item] of items.entries()) {
    attributes(item, []);
    if (!headings) { leaf(item); continue; }
    const name = index % 2 ? "i4" : "lpstr"; check(item, [name], VT); const value = one(item, name, VT);
    if (!value) fail(item, "A heading pair has the wrong primitive value."); attributes(value, []);
    if (index % 2) numberValue(value); else leaf(value);
  }
}
export function readWorkbookProperties(opc: OpcPackage): Omit<WorkbookDocumentMetadata, "workbook" | "word" | "presentation"> {
  const result: Omit<WorkbookDocumentMetadata, "workbook" | "word" | "presentation"> = {}, properties: Properties = {};
  const core = opc.related("", "metadata/core-properties");
  if (core) {
    const root = opc.read(core); if (root.uri !== CP || root.name !== "coreProperties") fail(root, "The core-properties root is invalid."); attributes(root, []);
    const seen = new Set<string>();
    for (const item of root.content) {
      if (typeof item === "string") { if (item.trim()) fail(root, "Unexpected core-property text."); continue; }
      if (!coreFields[item.uri]?.includes(item.name)) unsupported(item, "This core document property is not supported.");
      if (seen.has(item.name)) fail(item, "A core document property is repeated."); seen.add(item.name);
      if (item.uri === DCT) {
        attributes(item, ["type"], XSI);
        const type = attr(item, "type", XSI), parts = type?.split(":");
        if (!parts || parts.length !== 2 || parts[1] !== "W3CDTF" || item.namespaces[parts[0]!] !== DCT) fail(item, "A core-property date requires the W3CDTF type declaration.");
      } else attributes(item, []);
      const value = leaf(item);
      if (item.name === "title") result.title = value;
      else if (item.name === "language") { if (value) result.locale = value; }
      else properties[item.name as Exclude<keyof Properties, "totalEditingMinutes">] = value;
    }
  }
  const app = opc.related("", "extended-properties");
  if (app) {
    const root = opc.read(app); if (root.uri !== EP || root.name !== "Properties") fail(root, "The application-properties root is invalid."); attributes(root, []);
    check(root, [...Object.keys(appFields), ...counts, ...flags, "HeadingPairs", "TitlesOfParts", "HyperlinkBase", "PresentationFormat"], EP);
    for (const node of elements(root, undefined, EP)) {
      one(root, node.name, EP); attributes(node, []);
      if (node.name === "HeadingPairs" || node.name === "TitlesOfParts") vector(node, node.name === "HeadingPairs");
      else if (appFields[node.name]) properties[appFields[node.name] as Exclude<keyof Properties, "totalEditingMinutes">] = leaf(node);
      else if (counts.has(node.name)) { const n = numberValue(node); if (node.name === "TotalTime") properties.totalEditingMinutes = n; }
      else if (flags.has(node.name)) boolean({ ...node, attrs: new Map([[" val", leaf(node).trim()]]) }, "val");
      else if (leaf(node)) unsupported(node, "Nonempty hyperlink bases and presentation layout properties are not supported in a workbook.");
    }
  }
  if (Object.keys(properties).length) result.properties = properties;
  return result;
}
