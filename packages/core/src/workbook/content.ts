import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookManifest } from "../contracts/document.js";
import type { WorkbookMetadata } from "../contracts/workbook.js";
import type { ParsedHtml } from "../document/html.js";
import { checkLimit, invalid, unsupported } from "../document/validation.js";
import { validateWorkbookComments } from "./comments.js";
import { checkNonOverlappingRanges } from "./ranges.js";

/** Bounded rich comment content is separate from compact cell chunks, while
 * both remain parts of the same persisted HTML package and revision. */
export function validateWorkbookContent(workbook: WorkbookMetadata, manifest: WorkbookManifest, annotations: readonly Annotation[], html: ParsedHtml | undefined, classes: ReadonlySet<string>, maxItems: number, priorItems = 0): void {
  const sheets = new Set(workbook.sheets.map(sheet => sheet.sheetId)), stories = new Map<string, string>(), storyIds = new Set<string>();
  let items = priorItems;
  for (const story of workbook.stories ?? []) {
    if (storyIds.has(story.storyId) || stories.has(story.rootNodeId)) invalid("Workbook stories require unique identities.");
    const node = html?.peekNode(story.rootNodeId), parent = node?.parentNodeId ? html?.peekNode(node.parentNodeId) : undefined;
    if (!node || node.tagName !== "section" || node.attributes["data-office-kind"] !== "comment" || node.attributes["data-office-story"] !== "comment" || parent?.attributes["data-office-kind"] !== "workbook-content" || parent.parentNodeId !== null) invalid("A workbook comment story requires an independent HTML body.");
    storyIds.add(story.storyId); stories.set(story.rootNodeId, story.sheetId); items++;
  }
  if (html) {
    const allowed = new Set(["p", "span", "strong", "em", "u", "s", "sub", "sup", "br"]);
    for (const id of html.nodeIds()) {
      const node = html.peekNode(id)!;
      for (const name of node.attributes.class?.split(/\s+/).filter(Boolean) ?? []) if (!classes.has(name)) invalid("A workbook content style is undefined.");
      if (node.parentNodeId === null) {
        if (node.attributes["data-office-story"] !== undefined || html.directText(id)?.trim()) invalid("Workbook content text must belong to a declared story.");
        continue;
      }
      if (stories.has(id)) continue;
      if (node.attributes["data-office-story"] !== undefined || node.attributes["data-office-kind"] !== undefined) invalid("Only declared story roots may define workbook rich-text semantics.");
      if (!allowed.has(node.tagName)) unsupported("This workbook content node has no declared comment or object representation.");
      let parent: string | null = node.parentNodeId;
      while (parent && !stories.has(parent)) parent = html.peekNode(parent)!.parentNodeId;
      if (!parent) invalid("Workbook rich text must belong to a declared story.");
    }
  }
  const ids = new Set<string>(), usedBodies = new Set<string>();
  const identity = (id: string): void => { if (ids.has(id)) invalid("Workbook annotation and reply identities must be unique."); ids.add(id); };
  const body = (id: string | undefined, sheetId: string): void => {
    if (id === undefined) return;
    if (stories.get(id) !== sheetId || usedBodies.has(id)) invalid("An annotation body must refer to its own comment story on the anchored worksheet.");
    usedBodies.add(id);
  };
  for (const annotation of annotations) {
    identity(annotation.annotationId);
    if (annotation.documentId !== manifest.documentId || annotation.baseRevision > manifest.revision) invalid("A workbook annotation belongs to a different document or future revision.");
    const anchor = annotation.anchor;
    if (anchor.kind !== "cells") unsupported("This workbook annotation anchor kind is not implemented yet.");
    if (annotation.status !== "needs_reanchor" && !sheets.has(anchor.sheetId)) invalid("An active workbook annotation must target an existing worksheet.");
    checkNonOverlappingRanges(anchor.ranges); items += anchor.ranges.length + 1;
    body(annotation.bodyNodeId, anchor.sheetId);
    for (const reply of annotation.replies ?? []) { identity(reply.replyId); body(reply.bodyNodeId, anchor.sheetId); items++; }
    checkLimit(items, maxItems);
  }
  for (const id of stories.keys()) if (!usedBodies.has(id)) invalid("A workbook comment story is not owned by an annotation or reply.");
  checkLimit(items + validateWorkbookComments(workbook, annotations, maxItems, classes), maxItems);
}
