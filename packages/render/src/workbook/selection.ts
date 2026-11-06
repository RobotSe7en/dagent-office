import { DocumentError } from "@doffi/core";

export interface WorkbookViewPoint { readonly node: Node; readonly offset: number }

function nativeSelection(view: HTMLElement): Selection {
  const selection = view.ownerDocument.defaultView?.getSelection();
  if (!view.isConnected || !selection) throw new DocumentError("InvalidOperation", "Workbook selection requires a connected display surface.");
  return selection;
}

/** Read actual view endpoints, including engines that retarget getRangeAt()
 * to the shadow host. Never map a host/source/other-view selection by guessing
 * its text: an owned range is required for precise UTF-16 source projection. */
export function readWorkbookViewRange(view: HTMLElement): Range | undefined {
  const selection = nativeSelection(view), root = view.getRootNode();
  if (selection.rangeCount > 1) return undefined;
  const ranges = typeof selection.getComposedRanges === "function" && root.nodeType === 11 && "host" in root
    ? selection.getComposedRanges({ shadowRoots: [root as ShadowRoot] })
    : selection.rangeCount === 1 ? [selection.getRangeAt(0)] : [];
  if (ranges.length !== 1) return undefined;
  const range = ranges[0]!;
  if (!view.contains(range.startContainer) || !view.contains(range.endContainer)) return undefined;
  const result = view.ownerDocument.createRange();
  result.setStart(range.startContainer, range.startOffset); result.setEnd(range.endContainer, range.endOffset);
  return result;
}

/** Use endpoints instead of addRange(), which does not enter a shadow view in
 * all engines. Validate both points before changing the browser selection. */
export function selectWorkbookView(view: HTMLElement, anchor: WorkbookViewPoint, focus: WorkbookViewPoint): void {
  const selection = nativeSelection(view);
  for (const point of [anchor, focus]) {
    const length = point.node.nodeType === 3 ? (point.node as Text).length : point.node.nodeType === 1 ? point.node.childNodes.length : -1;
    if (!view.contains(point.node) || !Number.isSafeInteger(point.offset) || point.offset < 0 || point.offset > length) throw new DocumentError("InvalidOperation", "Workbook selection endpoints must be valid positions in this view.");
  }
  if (typeof selection.setBaseAndExtent !== "function") throw new DocumentError("UnsupportedFeature", "This browser cannot select precise workbook view endpoints.");
  selection.setBaseAndExtent(anchor.node, anchor.offset, focus.node, focus.offset);
}
