import { DocumentError, openDocument } from "@dagent-office/core";
import { exportPptx } from "@dagent-office/office";
import type { PptxWorkerRequest, PptxWorkerResponse } from "./export-pptx.js";
const scope = globalThis as unknown as {
    onmessage: ((event: MessageEvent<PptxWorkerRequest>) => void) | null;
    postMessage(message: PptxWorkerResponse, transfer?: Transferable[]): void;
};
scope.onmessage = (event) => {
    const { taskId, files } = event.data;
    try {
        const output = exportPptx(openDocument(new Map(files))), bytes = new Uint8Array(output.bytes);
        scope.postMessage({ taskId, ok: true, bytes, report: output.report }, [bytes.buffer]);
    }
    catch (reason) {
        const error = reason instanceof DocumentError ? reason : new DocumentError("ExportFailed", "无法生成 PPTX，请重试。");
        scope.postMessage({ taskId, ok: false, error: error.toJSON() });
    }
};
