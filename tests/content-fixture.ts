import { openDocument as openAnyDocument } from "@dagent-office/core";
import type { DocumentPackage, OpenDocumentOptions, ContentDocumentModel } from "@dagent-office/core";

/** Legacy Word/PPT fixtures intentionally exercise only the content branch.
 * Workbook and dispatch tests call the public unified reader directly. */
export function openContentFixture(files: DocumentPackage, options?: OpenDocumentOptions): ContentDocumentModel {
  const model = openAnyDocument(files, options);
  if (model.kind === "workbook") throw new Error("This fixture expects a Word or PPT model.");
  return model;
}
