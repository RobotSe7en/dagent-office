import { DocumentError } from "@dagent-office/core";
import type { DocumentReadLimits } from "@dagent-office/core";
import { SaxesParser } from "saxes";
import { fail, normalizeXmlNamespace } from "./xml.js";
import type { XmlBudget, XmlNode } from "./xml.js";
import { ImportControl } from "../import-control.js";
import type { OfficeImportProgress } from "../import-control.js";

interface XmlStreamOptions {
  readonly limits: DocumentReadLimits;
  readonly budget: XmlBudget;
  readonly maxRetainedNodes: number;
  readonly maxTextCodeUnits: number;
  readonly control: ImportControl;
  readonly progress: (bytesRead: number, totalBytes: number) => OfficeImportProgress;
  /** false releases a completed subtree; the parent is still available. */
  readonly complete: (node: XmlNode, parent: XmlNode | undefined) => boolean;
  readonly open?: (node: XmlNode, parent: XmlNode | undefined) => void;
  /** The format adapter identifies structural containers with ignorable spaces. */
  readonly ignoreWhitespace?: (parent: XmlNode) => boolean;
}

/** Stream UTF-8/16 XML into bounded records and an optional small header tree. */
export async function streamXml(bytes: Uint8Array, part: string, options: XmlStreamOptions): Promise<XmlNode> {
  const { limits, budget, control } = options;
  if (bytes.byteLength > limits.maxFileBytes) throw new DocumentError("ResourceLimitExceeded", "An XML part exceeds its byte budget.");
  if (![options.maxRetainedNodes, options.maxTextCodeUnits].every(value => Number.isSafeInteger(value) && value > 0)) fail({ part, name: "XML" }, "Invalid streaming XML budgets.");
  const codecs = globalThis as unknown as { TextDecoder: new (label: string, options: { fatal: boolean }) => { decode(value?: Uint8Array, options?: { stream: boolean }): string } };
  const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? "utf-16le" : bytes[0] === 0xfe && bytes[1] === 0xff ? "utf-16be" : "utf-8";
  const decoder = new codecs.TextDecoder(encoding, { fatal: true }), parser = new SaxesParser({ xmlns: true });
  const stack: { node: XmlNode; retained: number }[] = [];
  let root: XmlNode | undefined, retained = 0;
  const allocate = (): void => {
    if (++budget.nodes > limits.maxNodes || ++retained > options.maxRetainedNodes) throw new DocumentError("ResourceLimitExceeded", "XML exceeds its cumulative or retained node budget.");
  };
  parser.on("doctype", () => fail({ part, name: "DOCTYPE" }, "DTD and external entities are not allowed."));
  parser.on("processinginstruction", () => fail({ part, name: "XML" }, "XML processing instructions are not supported.", "UnsupportedFeature"));
  parser.on("error", () => fail({ part, name: "XML" }, "An XML part is malformed."));
  parser.on("opentag", tag => {
    allocate();
    if (stack.length + 1 > limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "XML exceeds the configured depth budget.");
    const parent = stack[stack.length - 1]?.node;
    const node: XmlNode = { name: tag.local, uri: normalizeXmlNamespace(tag.uri),
      attrs: new Map(Object.values(tag.attributes).map(item => [normalizeXmlNamespace(item.uri) + " " + item.local, item.value])),
      namespaces: { ...(parent?.namespaces ?? {}), ...Object.fromEntries(Object.entries(tag.ns).map(([key, value]) => [key, normalizeXmlNamespace(value)])) }, content: [], part };
    options.open?.(node, parent); stack.push({ node, retained: 1 });
  });
  parser.on("closetag", () => {
    const frame = stack.pop()!, parent = stack[stack.length - 1];
    if (options.complete(frame.node, parent?.node)) {
      if (parent) { parent.node.content.push(frame.node); parent.retained += frame.retained; }
      else { if (root) fail(frame.node, "An XML part has multiple roots."); root = frame.node; }
    } else {
      retained -= frame.retained;
      if (!parent) fail(frame.node, "The streaming XML root cannot be discarded.");
    }
  });
  const addText = (value: string): void => {
    const parent = stack[stack.length - 1]; if (!parent || !value) return;
    if (!/[^\t\r\n ]/.test(value) && options.ignoreWhitespace?.(parent.node)) return;
    const content = parent.node.content, last = content[content.length - 1];
    const length = value.length + (typeof last === "string" ? last.length : 0);
    if (length > options.maxTextCodeUnits) throw new DocumentError("ResourceLimitExceeded", "An XML text record exceeds its budget.");
    if (typeof last === "string") content[content.length - 1] = last + value;
    else { allocate(); parent.retained++; content.push(value); }
  };
  parser.on("text", addText); parser.on("cdata", addText);
  control.check();
  try {
    for (let cursor = 0; cursor < bytes.length; cursor += 65536) {
      control.check(); parser.write(decoder.decode(bytes.subarray(cursor, cursor + 65536), { stream: true }));
      await control.checkpoint(options.progress(Math.min(cursor + 65536, bytes.length), bytes.length));
    }
    parser.write(decoder.decode()).close();
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return fail({ part, name: "XML" }, "An XML part could not be parsed.");
  }
  if (!root) return fail({ part, name: "XML" }, "An XML part has no root.");
  return root;
}
