import { defaultTreeAdapter, Parser } from "parse5";
import type { DefaultTreeAdapterMap, DefaultTreeAdapterTypes, Token, TreeAdapter } from "parse5";
import type { DocumentReadLimits } from "../runtime.js";
import { checkLimit, invalid, unsupported } from "./validation.js";

type Node = DefaultTreeAdapterTypes.Node;
type ParentNode = DefaultTreeAdapterTypes.ParentNode;
type ChildNode = DefaultTreeAdapterTypes.ChildNode;

/** Keep parser-specific token and construction checks behind the HTML adapter. */
export function parseHtmlTree(source: string, limits: DocumentReadLimits, supportedTags: ReadonlySet<string>): DefaultTreeAdapterTypes.Document {
  const depths = new WeakMap<Node, number>();
  const sourceElements = new Set<number>();
  let count = 0;

  function checkNode(node: Node, depth: number): void {
    if (!depths.has(node)) checkLimit(++count, limits.maxNodes);
    checkLimit(depth, limits.maxDepth);
    depths.set(node, depth);
  }

  function checkInsertion(parent: ParentNode, child: ChildNode): void {
    const stack = [{ node: child, depth: (depths.get(parent) ?? 0) + 1 }];
    while (stack.length) {
      const { node, depth } = stack.pop()!;
      checkNode(node, depth);
      // HTML tree corrections can move whole subtrees, changing every descendant's depth.
      if ("childNodes" in node) {
        for (const descendant of node.childNodes) stack.push({ node: descendant, depth: depth + 1 });
      }
    }
  }

  const treeAdapter: TreeAdapter<DefaultTreeAdapterMap> = {
    ...defaultTreeAdapter,
    createDocument() {
      const document = defaultTreeAdapter.createDocument();
      checkNode(document, 0);
      return document;
    },
    appendChild(parent, child) {
      checkInsertion(parent, child);
      defaultTreeAdapter.appendChild(parent, child);
    },
    insertBefore(parent, child, reference) {
      checkInsertion(parent, child);
      defaultTreeAdapter.insertBefore(parent, child, reference);
    },
    // The default adapter inserts text and doctypes through its own methods,
    // bypassing our appendChild/insertBefore hooks. Check those insertions too.
    insertText(parent, text) {
      defaultTreeAdapter.insertText(parent, text);
      checkInsertion(parent, parent.childNodes[parent.childNodes.length - 1]!);
    },
    insertTextBefore(parent, text, reference) {
      defaultTreeAdapter.insertTextBefore(parent, text, reference);
      checkInsertion(parent, parent.childNodes[parent.childNodes.indexOf(reference) - 1]!);
    },
    setDocumentType(document, name, publicId, systemId) {
      defaultTreeAdapter.setDocumentType(document, name, publicId, systemId);
      checkInsertion(document, document.childNodes.find((node) => node.nodeName === "#documentType")!);
    },
    setNodeSourceCodeLocation(node, location) {
      if (defaultTreeAdapter.isElementNode(node) && location?.startTag) {
        sourceElements.add(location.startTag.startOffset);
      }
      defaultTreeAdapter.setNodeSourceCodeLocation(node, location);
    },
  };

  // parse5 does not report every ignored start tag as a parse error. Observe
  // actual tokenizer output, preserving its HTML context (e.g. title text).
  class DocumentParser extends Parser<DefaultTreeAdapterMap> {
    override onStartTag(token: Token.TagToken): void {
      if (!supportedTags.has(token.tagName)) unsupported("The HTML element is outside the supported document structure.");
      super.onStartTag(token);
      if (!token.location || !sourceElements.has(token.location.startOffset)) {
        invalid("An HTML source element was discarded by the parser.");
      }
    }
  }

  return DocumentParser.parse(source, {
    treeAdapter,
    sourceCodeLocationInfo: true,
    onParseError: () => invalid("The HTML source is malformed."),
  });
}
