import { DocumentError, parseContract } from "@dagent-office/core";
import type { DocumentErrorData, DocumentPackage } from "@dagent-office/core";
import type { PptxExportReport, PptxExportResult } from "@dagent-office/office";
export interface PptxWorkerRequest {
    readonly taskId: string;
    readonly files: [
        string,
        Uint8Array
    ][];
}
export type PptxWorkerResponse = {
    readonly taskId: string;
    readonly ok: true;
    readonly bytes: Uint8Array;
    readonly report: PptxExportReport;
} | {
    readonly taskId: string;
    readonly ok: false;
    readonly error: DocumentErrorData;
};
/** The host owns this short-lived worker and terminates it on document changes,
 * cancellation or completion. No mutable editing session crosses the boundary. */
export function exportCurrentPptx(files: DocumentPackage, taskId: string, signal: AbortSignal): Promise<PptxExportResult> {
    if (signal.aborted)
        return Promise.reject(signal.reason ?? new Error("已取消 PPTX 导出。"));
    return new Promise((resolve, reject) => {
        const worker = new Worker(new URL("./pptx-export-worker.ts", import.meta.url), { type: "module" });
        const finish = (): void => { clearTimeout(timer); signal.removeEventListener("abort", cancel); worker.terminate(); };
        const cancel = (): void => { finish(); reject(signal.reason instanceof Error ? signal.reason : new Error("已取消 PPTX 导出。")); };
        const timer = setTimeout(() => { finish(); reject(new Error("PPTX 导出超时，请重试。")); }, 30000);
        worker.onerror = () => { finish(); reject(new Error("无法生成 PPTX，请重试。")); };
        worker.onmessageerror = () => { finish(); reject(new Error("PPTX 导出结果无法读取。")); };
        worker.onmessage = (event: MessageEvent<PptxWorkerResponse>) => {
            finish();
            try {
                const output = event.data;
                if (output.taskId !== taskId)
                    throw new Error("PPTX 导出结果与请求不匹配。");
                if (!output.ok) {
                    const error = parseContract("document-error", output.error);
                    throw new DocumentError(error.code, error.message, error.details);
                }
                if (!(output.bytes instanceof Uint8Array) || !output.bytes.length || !output.report)
                    throw new Error("PPTX 导出结果无法读取。");
                resolve({ bytes: output.bytes, report: output.report });
            }
            catch (reason) {
                reject(reason);
            }
        };
        signal.addEventListener("abort", cancel, { once: true });
        if (signal.aborted) {
            cancel();
            return;
        }
        try {
            const entries = [...files].map(([path, bytes]) => [path, new Uint8Array(bytes)] as [
                string,
                Uint8Array<ArrayBuffer>
            ]);
            worker.postMessage({ taskId, files: entries } satisfies PptxWorkerRequest, entries.map(([, bytes]) => bytes.buffer));
        }
        catch (reason) {
            finish();
            reject(reason);
        }
    });
}
