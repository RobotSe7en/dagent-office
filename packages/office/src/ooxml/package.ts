import {
  DEFAULT_DOCUMENT_READ_LIMITS,
  DocumentError,
  PackagePathSchema,
} from "@dagent-office/core";
import type { DocumentPackage, DocumentReadLimits } from "@dagent-office/core";
import { readDocumentZip } from "../document-zip.js";
import type { DocumentZipOptions } from "../document-zip.js";
import { PKG, attr, elements, fail, xml } from "./xml.js";
import type { XmlBudget, XmlNode } from "./xml.js";

export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}
export interface OpcPackage {
  readonly files: DocumentPackage;
  readonly limits: DocumentReadLimits;
  read(part: string): XmlNode;
  relationships(part: string): ReadonlyMap<string, Relationship>;
  related(part: string, type: string): string | undefined;
  contentType(part: string): string | undefined;
}
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
/** Resolve an OPC part URI; this deliberately differs from HTML URL handling. */
export function resolvePart(source: string, target: string): string {
  if (
    !target ||
    /[\\\u0000-\u0020\u007f?#]/.test(target) ||
    /^[a-z][a-z0-9+.-]*:/i.test(target)
  )
    throw new DocumentError(
      "InvalidDocument",
      "An internal OPC target is invalid.",
    );
  const result = target.startsWith("/") ? [] : source.split("/").slice(0, -1);
  for (const encoded of target.replace(/^\//, "").split("/")) {
    let value: string;
    try {
      value = decodeURIComponent(encoded);
    } catch {
      throw new DocumentError(
        "InvalidDocument",
        "An OPC target has invalid escaping.",
      );
    }
    if (value === ".") continue;
    if (value === "..") {
      if (!result.length)
        throw new DocumentError(
          "InvalidDocument",
          "An OPC target escapes its package.",
        );
      result.pop();
    } else {
      if (!value || /[/\\:\u0000-\u001f\u007f?#]/.test(value))
        throw new DocumentError(
          "InvalidDocument",
          "An OPC target contains an invalid segment.",
        );
      result.push(value);
    }
  }
  const path = result.join("/");
  if (!PackagePathSchema.safeParse(path).success)
    throw new DocumentError(
      "InvalidDocument",
      "An OPC target is not a package path.",
    );
  return path;
}
export function openOpc(
  bytes: Uint8Array,
  options: DocumentZipOptions,
): OpcPackage {
  const files = readDocumentZip(bytes, options);
  const limits: DocumentReadLimits = {
    ...DEFAULT_DOCUMENT_READ_LIMITS,
    ...options.limits,
  };
  const budget: XmlBudget = { nodes: 0 };
  const parsed = new Map<string, XmlNode>();
  const relationCache = new Map<string, ReadonlyMap<string, Relationship>>();
  const read = (part: string): XmlNode => {
    const cached = parsed.get(part);
    if (cached) return cached;
    const data = files.get(part);
    if (!data)
      throw new DocumentError(
        "InvalidDocument",
        "A required OPC part is missing.",
        [
          {
            code: "MissingPart",
            path: [part.slice(0, 128)],
            message: "A required OPC part is missing.",
          },
        ],
      );
    const result = xml(data, part, limits, budget);
    parsed.set(part, result);
    return result;
  };
  const types = read("[Content_Types].xml");
  if (types.uri !== CT || types.name !== "Types")
    fail(types, "The OPC content types part is invalid.");
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const item of elements(types, undefined, CT)) {
    const type = attr(item, "ContentType", "");
    if (!type) fail(item, "An OPC content type is missing.");
    if (item.name === "Default") {
      const extension = attr(item, "Extension", "")?.toLowerCase();
      if (!extension || defaults.has(extension))
        fail(item, "OPC default content types must be unique.");
      defaults.set(extension, type);
    } else if (item.name === "Override") {
      const name = attr(item, "PartName", "");
      if (!name?.startsWith("/"))
        fail(item, "An OPC override requires an absolute part URI.");
      const path = resolvePart("", name);
      if (overrides.has(path) || !files.has(path))
        fail(item, "An OPC override is duplicate or missing its part.");
      overrides.set(path, type);
    } else fail(item, "An OPC content type element is unknown.");
  }
  const relationships = (part: string): ReadonlyMap<string, Relationship> => {
    const cached = relationCache.get(part);
    if (cached) return cached;
    const split = part.lastIndexOf("/");
    const path = part
      ? part.slice(0, split + 1) + "_rels/" + part.slice(split + 1) + ".rels"
      : "_rels/.rels";
    const result = new Map<string, Relationship>();
    if (files.has(path)) {
      const root = read(path);
      if (root.uri !== PKG || root.name !== "Relationships")
        fail(root, "An OPC relationship part is invalid.");
      for (const relation of elements(root, undefined, PKG)) {
        const id = attr(relation, "Id", "");
        const type = attr(relation, "Type", "");
        const target = attr(relation, "Target", "");
        const mode = attr(relation, "TargetMode", "");
        if (
          relation.name !== "Relationship" ||
          !id ||
          !type ||
          !target ||
          result.has(id) ||
          (mode !== undefined && !["External", "Internal"].includes(mode))
        )
          fail(relation, "An OPC relationship is invalid or duplicate.");
        const external = mode === "External";
        const resolved = external ? target : resolvePart(part, target);
        if (!external && !files.has(resolved))
          fail(relation, "An OPC relationship targets a missing part.");
        result.set(id, {
          id,
          type: type.replace(
            "http://purl.oclc.org/ooxml/officeDocument/relationships/",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
          ),
          target: resolved,
          external,
        });
      }
    }
    relationCache.set(part, result);
    return result;
  };
  return {
    files,
    limits,
    read,
    relationships,
    related: (part, type) => {
      const matches = [...relationships(part).values()].filter((item) =>
        item.type.endsWith("/" + type),
      );
      if (matches.length > 1 || matches.some((item) => item.external))
        throw new DocumentError(
          "InvalidDocument",
          "A singleton OPC relationship is ambiguous or external.",
        );
      return matches[0]?.target;
    },
    contentType: (part) =>
      overrides.get(part) ?? defaults.get(part.split(".").pop()!.toLowerCase()),
  };
}
