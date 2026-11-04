import { DocumentError } from "@dagent-office/core";
import type { DocumentReadLimits, WorkbookMetadata } from "@dagent-office/core";
import { utf8Bytes } from "./chunks.js";

/** Only bounded auxiliary object/comment HTML is retained as a tree-like
 * string list. Actual cells stay in the independent streaming chunk builder. */
export class WorkbookContent {
  readonly stories: NonNullable<WorkbookMetadata["stories"]> = [];
  private readonly bodies: string[] = [];
  private nodes = 1;
  private bytes: number;
  private readonly prefix = '<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body><section data-node-id="workbook-content" data-office-kind="workbook-content">';
  private readonly suffix = '</section></body></html>';
  constructor(readonly limits: DocumentReadLimits) { this.bytes = this.prefix.length + this.suffix.length; }
  id(): string {
    if (++this.nodes > this.limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook auxiliary content exceeds its node budget.");
    return `workbook-node-${this.nodes}`;
  }
  comment(sheetId: string, richHtml: string): string {
    const rootNodeId = this.id(), paragraph = this.id();
    const html = `<section data-node-id="${rootNodeId}" data-office-kind="comment" data-office-story="comment"><p data-node-id="${paragraph}" style="white-space:pre-wrap">${richHtml}</p></section>`;
    this.bytes += utf8Bytes(html);
    if (this.bytes > this.limits.maxFileBytes || this.bytes > this.limits.maxPackageBytes) throw new DocumentError("ResourceLimitExceeded", "Workbook auxiliary content exceeds its byte budget.");
    this.bodies.push(html); this.stories.push({ storyId: `story-${rootNodeId}`, rootNodeId, kind: "comment", sheetId }); return rootNodeId;
  }
  html(): string | undefined { return this.bodies.length ? this.prefix + this.bodies.join("") + this.suffix : undefined; }
}
