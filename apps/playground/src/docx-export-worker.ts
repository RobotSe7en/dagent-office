import { DocumentError, openDocument } from "@dagent-office/core";
import { exportDocx } from "@dagent-office/office";
import type { DocxWorkerRequest, DocxWorkerResponse } from "./export-docx.js";

const scope = globalThis as unknown as {
  onmessage: ((event: MessageEvent<DocxWorkerRequest>) => void) | null;
  postMessage(message: DocxWorkerResponse, transfer?: Transferable[]): void;
};
scope.onmessage = (event) => {
  const { taskId, files } = event.data;
  try {
    const output = exportDocx(openDocument(new Map(files))), bytes = new Uint8Array(output.bytes);
    scope.postMessage({ taskId, ok: true, bytes, report: output.report }, [bytes.buffer]);
  } catch (reason) {
    const error = reason instanceof DocumentError ? reason : new DocumentError("ExportFailed", "无法生成 DOCX，请重试。");
    scope.postMessage({ taskId, ok: false, error: error.toJSON() });
  }
};
