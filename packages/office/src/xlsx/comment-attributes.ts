import { MC, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes } from "./values.js";

export const COMMENT_REVISION = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

/** Understand only the selected revision UID. Ignorable declarations never
 * authorize dropping another namespace's attributes, elements or payload. */
export function classicCommentAttributes(node: XmlNode, names: readonly string[], uid = false): void {
  const ignorable = node.attrs.get(MC + " Ignorable");
  if (ignorable !== undefined) {
    const prefixes = ignorable.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "").split(/[\t\r\n ]+/);
    if (prefixes.some(prefix => !prefix || node.namespaces[prefix] !== COMMENT_REVISION)) unsupported(node, "A classic comment declares an unsupported ignorable namespace.");
  }
  attributes({ ...node, attrs: new Map([...node.attrs].filter(([key]) => key !== MC + " Ignorable" && !(uid && key === COMMENT_REVISION + " uid"))) }, names);
}
