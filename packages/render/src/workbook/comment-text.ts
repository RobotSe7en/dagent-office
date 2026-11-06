import { DocumentError } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import { workbookRichTextFonts } from "./text-fonts.js";
import { copyWorkbookTextStyle, decorateWorkbookText } from "./text-style.js";

export interface WorkbookCommentText {
  readonly element: HTMLElement;
  readonly nodes: number;
  readonly sourceCodeUnits: number;
  readonly projections: readonly { readonly node: Text; readonly sourceNodeId: string; readonly from: number; readonly to: number }[];
}

/** Materialize one validated comment story into ordinary rich-text flow.
 * Native spreadsheet run decorations and Chinese font roles remain distinct
 * from DrawingML text layout. Source offsets address the entire body, with no
 * fabricated paragraph separators. The owning panel resolves local fonts and
 * supplies width/scrolling; this detached view never changes the story. */
export function createWorkbookCommentText(source: HTMLElement, maxNodes: number = 50_000, maxCodeUnits: number = 8 * 1024 * 1024): WorkbookCommentText {
  const document = source.ownerDocument;
  if (!source.isConnected || !document.defaultView || source.dataset.officeKind !== "comment" || !source.dataset.nodeId
    || ![maxNodes, maxCodeUnits].every(value => Number.isSafeInteger(value) && value > 0))
    throw new DocumentError("InvalidOperation", "Comment text requires its own connected source story and positive budgets.");
  if ((source.textContent?.length ?? 0) > maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Comment text exceeds its source-text budget.");
  const root = markViewNode(document.createElement("div")), projections: { node: Text; sourceNodeId: string; from: number; to: number }[] = [];
  root.dataset.officeView = "workbook-comment-text"; root.dataset.sourceNodeId = source.dataset.nodeId;
  copyWorkbookTextStyle(source, root); root.style.whiteSpace = "pre-wrap";
  const language = source.closest("[lang]")?.getAttribute("lang"); if (language !== undefined && language !== null) root.lang = language;
  let nodes = 1, offset = 0;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Comment text exceeds its node budget."); };
  const advance = (amount: number): void => { offset += amount; if (offset > maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Comment text exceeds its source-text budget."); };
  const inline = new Set(["span", "strong", "em", "u", "s", "sub", "sup"]);
  for (const child of source.childNodes) {
    if (child.nodeType === 3 && !child.textContent?.trim()) { advance(child.textContent?.length ?? 0); continue; }
    if (child.nodeType !== 1 || (child as Element).localName !== "p") throw new DocumentError("UnsupportedFeature", "Comment display requires paragraph-based source content.");
    account(1); const paragraph = child as HTMLElement, target = document.createElement("p"), runs = workbookRichTextFonts(paragraph, maxNodes - nodes + 1), base = offset;
    let local = 0, runIndex = 0;
    const appendText = (text: string, original: HTMLElement, output: HTMLElement): void => {
      const computed = document.defaultView!.getComputedStyle(original), end = local + text.length;
      while (local < end) {
        const run = runs[runIndex++];
        if (!run || run.from !== local || run.to > end) throw new DocumentError("InvalidDocument", "Comment font runs disagree with their source text.");
        account(2); advance(run.text.length);
        const span = document.createElement("span"), node = document.createTextNode(run.text);
        span.dataset.officeScript = run.role; span.style.fontFamily = run.font.family;
        decorateWorkbookText(computed, span, account, false).append(node); output.append(span);
        projections.push(Object.freeze({ node, sourceNodeId: source.dataset.nodeId!, from: base + run.from, to: base + run.to })); local = run.to;
      }
    };
    const append = (original: HTMLElement, output: HTMLElement, depth: number): void => {
      if (depth > 64) throw new DocumentError("ResourceLimitExceeded", "Comment text exceeds its nesting budget.");
      const computed = copyWorkbookTextStyle(original, output, true);
      if (original.hasAttribute("lang")) output.lang = original.lang;
      for (const key of ["direction", "unicode-bidi", "white-space"]) output.style.setProperty(key, computed.getPropertyValue(key));
      for (const node of original.childNodes) {
        if (node.nodeType === 3) appendText(node.textContent ?? "", original, output);
        else if (node.nodeType === 1 && (node as Element).localName === "br") {
          account(1); const span = document.createElement("span"); copyWorkbookTextStyle(node as HTMLElement, span); span.style.whiteSpace = "pre";
          appendText("\n", node as HTMLElement, span); output.append(span);
        } else if (node.nodeType === 1 && inline.has((node as Element).localName)) {
          account(1); const span = document.createElement("span"); append(node as HTMLElement, span, depth + 1); output.append(span);
        } else throw new DocumentError("UnsupportedFeature", "Comment display contains unsupported inline source content.");
      }
    };
    append(paragraph, target, 0);
    const computed = document.defaultView!.getComputedStyle(paragraph);
    for (const key of ["text-align", "text-indent", "margin-left", "margin-right", "margin-top", "margin-bottom"]) target.style.setProperty(key, computed.getPropertyValue(key));
    if (runIndex !== runs.length) throw new DocumentError("InvalidDocument", "Comment text did not consume its source font runs.");
    root.append(target);
  }
  return Object.freeze({ element: root, nodes, sourceCodeUnits: offset, projections: Object.freeze(projections) });
}
