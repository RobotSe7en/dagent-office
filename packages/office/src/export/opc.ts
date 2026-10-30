import { empty, declaration, attributes, XmlParts } from "./xml.js";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
interface Relation { id: string; type: string; target: string; external: boolean }
/** Output part names are generated here, never copied from host paths or IDs. */
export class OutputOpc {
  private readonly relations = new Map<string, Relation[]>();
  private readonly types = new Map<string, string>();
  constructor(readonly parts: XmlParts) {}
  binary(path: string, type: string, bytes: Uint8Array): void { this.types.set(path, type); this.parts.put(path, bytes); }
  part(path: string, type: string): ReturnType<XmlParts["buffer"]> { this.types.set(path, type); return this.parts.buffer(path); }
  relate(part: string, type: string, target: string, external = false): string {
    const list = this.relations.get(part) ?? []; this.relations.set(part, list);
    const fullType = type.includes(":") ? type : REL + type;
    const existing = list.find((entry) => entry.type === fullType && entry.target === target && entry.external === external);
    if (existing) return existing.id;
    const id = "rId" + (list.length + 1); list.push({ id, type: fullType, target, external }); return id;
  }
  finish(): void {
    for (const [part, list] of this.relations) {
      const split = part.lastIndexOf("/"), path = part ? part.slice(0, split + 1) + "_rels/" + part.slice(split + 1) + ".rels" : "_rels/.rels";
      const buffer = this.parts.buffer(path); buffer.add(declaration + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
      for (const relation of list) buffer.add(empty("Relationship", { Id: relation.id, Type: relation.type, Target: relation.target, ...(relation.external ? { TargetMode: "External" } : {}) }));
      buffer.add("</Relationships>"); buffer.finish();
    }
    const buffer = this.parts.buffer("[Content_Types].xml"); buffer.add(declaration + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">');
    buffer.add(empty("Default", { Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" }));
    for (const [path, type] of this.types) buffer.add("<Override" + attributes({ PartName: "/" + path, ContentType: type }) + "/>");
    buffer.add("</Types>"); buffer.finish();
  }
}
export const contentType = (part: string): string => "application/vnd.openxmlformats-officedocument.wordprocessingml." + part + "+xml";
