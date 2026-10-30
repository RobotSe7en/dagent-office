import { useEffect, useRef, useState } from "react";
import { createDocumentSession, createOpaqueId, DocumentError, openDocument } from "@dagent-office/core";
import type { DocumentModel, DocumentSession, DocumentSnapshot } from "@dagent-office/core";
import {
  importDocx,
  readDocumentZip,
  writeDocumentZip,
} from "@dagent-office/office";
import type { DocumentWorkbenchView } from "@dagent-office/editor";
import { DocumentWorkbench } from "@dagent-office/editor/react";
import { reportFiles, sampleProposal } from "./samples.js";
const sample = import.meta.env.BASE_URL + "01-text-numbering.docx";
function base64(bytes: Uint8Array): string {
  let value = ""; for (let offset = 0; offset < bytes.length; offset += 32768) value += String.fromCharCode(...bytes.subarray(offset, offset + 32768)); return btoa(value);
}

export function App() {
  const [model, setModel] = useState<DocumentModel>();
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [layout, setLayout] = useState<"pages" | "continuous">(() => {
    try { return localStorage.getItem("dagent-office:layout") === "continuous" ? "continuous" : "pages"; } catch { return "pages"; }
  });
  const [snapshot, setSnapshot] = useState<DocumentSnapshot>();
  const sessionRef = useRef<DocumentSession>(undefined);
  const workbenchRef = useRef<DocumentWorkbenchView>(undefined);
  const [session, setSession] = useState<DocumentSession>();
  const [submitted, setSubmitted] = useState(false);
  const generation = useRef(0);
  const [exporting, setExporting] = useState(false);
  const exportRef = useRef<{ controller: AbortController; session: DocumentSession; revision: number }>(undefined);
  useEffect(() => {
    setSubmitted(false);
    if (!model) { setSession(undefined); return; }
    const session = createDocumentSession(model);
    sessionRef.current = session;
    setSession(session);
    setSnapshot(session.snapshot());
    const unsubscribe = session.subscribe((event) => {
      if (event.type !== "error") {
        setSnapshot(session.snapshot());
        if (exportRef.current?.session === session && exportRef.current.revision !== session.snapshot().manifest.revision) exportRef.current.controller.abort();
      }
    });
    return () => {
      unsubscribe(); session.dispose();
      if (sessionRef.current === session) sessionRef.current = undefined;
    };
  }, [model]);
  useEffect(
    () => () => {
      generation.current++;
      exportRef.current?.controller.abort(); exportRef.current = undefined;
    },
    [],
  );
  async function load(
    bytes: Promise<ArrayBuffer>,
    kind: "docx" | "package",
  ): Promise<void> {
    const current = ++generation.current;
    exportRef.current?.controller.abort(); exportRef.current = undefined; setExporting(false);
    setLoading(true);
    setError("");
    setModel(undefined);
    setSnapshot(undefined);
    try {
      const value = new Uint8Array(await bytes);
      const files =
        kind === "docx"
          ? importDocx(value, { documentId: createOpaqueId() })
          : readDocumentZip(value);
      const next = openDocument(files);
      if (current === generation.current) setModel(next);
    } catch (reason) {
      if (current === generation.current)
        setError(
          reason instanceof DocumentError
            ? [
                reason.code,
                reason.message,
                ...reason.details.map((item) => item.path.join(" / ")),
              ].join(" · ")
            : "无法读取文件。",
        );
    } finally {
      if (current === generation.current) setLoading(false);
    }
  }
  function download(): void {
    const session = sessionRef.current;
    if (!session) return;
    const bytes = writeDocumentZip(session.serialize());
    const buffer = new ArrayBuffer(bytes.length);
    new Uint8Array(buffer).set(bytes);
    const url = URL.createObjectURL(
      new Blob([buffer], { type: "application/zip" }),
    );
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "document.zip";
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }
  async function openReport(): Promise<void> {
    const current = ++generation.current;
    exportRef.current?.controller.abort(); exportRef.current = undefined; setExporting(false);
    setLoading(true); setError(""); setModel(undefined);
    try { const files = await reportFiles(); if (current === generation.current) setModel(openDocument(files)); }
    catch (reason) { if (current === generation.current) setError(reason instanceof Error ? reason.message : "报告读取失败"); }
    finally { if (current === generation.current) setLoading(false); }
  }
  async function loadSampleProposal(): Promise<void> {
    const current = generation.current, view = workbenchRef.current;
    if (!view) return;
    try {
      const proposal = await sampleProposal();
      if (current !== generation.current || workbenchRef.current !== view) return;
      const request = view.request() ?? view.buildRequest();
      view.loadProposal(request, { ...proposal, requestId: request.requestId, documentId: request.documentId,
        baseRevision: request.baseRevision, annotationIds: request.annotations.map((annotation) => annotation.annotationId) });
    } catch (reason) { if (current === generation.current) setError(reason instanceof Error ? reason.message : "固定提案导入失败"); }
  }
  async function exportPdf(): Promise<void> {
    const active = sessionRef.current; if (!active || exportRef.current) return;
    const current = generation.current, identity = active.snapshot().manifest;
    const job = { controller: new AbortController(), session: active, revision: identity.revision };
    exportRef.current = job; setExporting(true); setError("");
    try {
      const taskId = createOpaqueId(), bytes = writeDocumentZip(active.serialize());
      const response = await fetch("/__dagent_office/convert", { method: "POST", signal: job.controller.signal,
        headers: { "Content-Type": "application/json", "X-Dagent-Office-Dev": "1" },
        body: JSON.stringify({ protocolVersion: 1, taskId, format: "pdf", input: { format: "document-package-zip", documentId: identity.documentId, bytesBase64: base64(bytes) } }),
      });
      const result = await response.json();
      if (current !== generation.current || exportRef.current !== job) return;
      if (!response.ok || !result.ok) throw new Error(typeof result.error?.message === "string" ? result.error.message.slice(0, 256) : "PDF 导出失败。");
      if (result.protocolVersion !== 1 || result.taskId !== taskId || result.output?.mimeType !== "application/pdf"
        || result.output.report?.documentId !== identity.documentId || result.output.report?.revision !== identity.revision) throw new Error("PDF 导出结果与当前请求不匹配。");
      if (active.snapshot().manifest.revision !== identity.revision) throw new Error("导出期间文档已变化，请重新导出。");
      const content = Uint8Array.from(atob(result.output.bytesBase64), (character) => character.charCodeAt(0));
      const url = URL.createObjectURL(new Blob([content.buffer], { type: "application/pdf" })), anchor = document.createElement("a");
      anchor.href = url; anchor.download = "document.pdf"; anchor.click(); setTimeout(() => URL.revokeObjectURL(url), 0);
    } catch (reason) {
      if (current === generation.current && exportRef.current === job) setError(active.snapshot().manifest.revision !== identity.revision ? "导出期间文档已变化，请重新导出。"
        : job.controller.signal.aborted ? "已取消 PDF 导出。" : reason instanceof Error ? reason.message : "PDF 导出失败。");
    } finally { if (exportRef.current === job) { exportRef.current = undefined; setExporting(false); } }
  }
  return (
    <main>
      <p className="eyebrow">DOCX 预览 · 批注请求 · 提案审阅</p>
      <h1>dagent-office Playground</h1>
      <div className="toolbar">
        <label>
          打开 DOCX 或文档包
          <input
            aria-label="打开文件"
            type="file"
            accept=".docx,.zip"
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              if (file)
                void load(
                  file.arrayBuffer(),
                  file.name.toLowerCase().endsWith(".docx")
                    ? "docx"
                    : "package",
                );
              event.currentTarget.value = "";
            }}
          />
        </label>
        <button
          disabled={loading}
          onClick={() =>
            void load(
              fetch(sample).then((response) => {
                if (!response.ok) throw new Error("样例读取失败");
                return response.arrayBuffer();
              }),
              "docx",
            )
          }
        >
          打开 DOCX 样例
        </button>
        <button disabled={loading} onClick={() => { void openReport(); }}>打开报告协作样例</button>
        <button disabled={!session} onClick={() => { void loadSampleProposal(); }}>载入固定示例提案</button>
        <button disabled={!model} onClick={download}>
          下载文档包
        </button>
        <button disabled={!session || exporting} onClick={() => { void exportPdf(); }}>{exporting ? "正在导出 PDF…" : "导出 PDF"}</button>
        {exporting && <button onClick={() => exportRef.current?.controller.abort()}>取消导出</button>}
        <label>显示方式 <select aria-label="显示方式" value={layout} onChange={(event) => {
          const value = event.currentTarget.value === "continuous" ? "continuous" : "pages"; setLayout(value);
          try { localStorage.setItem("dagent-office:layout", value); } catch { /* The host preference is optional. */ }
        }}><option value="pages">页面视图</option><option value="continuous">连续视图</option></select></label>
      </div>
      <p>{layout === "pages" ? "页面视图 · 支持的页码与目录随布局更新" : "连续视图 · 显示字段缓存"} · 文字修订保留标记</p>
      {loading && <p role="status">正在读取…</p>}
      {error && <p role="alert">{error}</p>}
      {!model && !loading && !error && (
        <p>选择文件，查看读取与样式映射结果。</p>
      )}
      {model && (
        <p role="status">
          已打开 · 修订 {(snapshot ?? model.snapshot()).manifest.revision} · 批注{" "}
          {(snapshot ?? model.snapshot()).annotations.length}
        </p>
      )}
      {submitted && <p>Playground 已接收批注请求；使用固定提案验证返回与应用。</p>}
      {model && session && <DocumentWorkbench session={session} layout={layout} className="preview"
        onViewChange={(view) => { workbenchRef.current = view; }}
        onSubmitAnnotations={() => { setSubmitted(true); }}
        onError={(reason) => setError(reason.message)} />}

    </main>
  );
}
