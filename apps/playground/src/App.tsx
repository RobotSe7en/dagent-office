import { useEffect, useRef, useState } from "react";
import { DocumentError, openDocument } from "@dagent-office/core";
import type { DocumentModel } from "@dagent-office/core";
import {
  importDocx,
  readDocumentZip,
  writeDocumentZip,
} from "@dagent-office/office";
import { mountDocumentPreview } from "@dagent-office/editor";
const sample = import.meta.env.BASE_URL + "01-text-numbering.docx";

export function App() {
  const [model, setModel] = useState<DocumentModel>();
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const container = useRef<HTMLDivElement>(null);
  const generation = useRef(0);
  useEffect(() => {
    if (!model || !container.current) return;
    const preview = mountDocumentPreview(container.current, model);
    let active = true;
    void preview.ready.catch((reason) => {
      if (active) {
        setError(reason instanceof Error ? reason.message : "预览失败");
        setModel(undefined);
      }
    });
    return () => {
      active = false;
      preview.dispose();
    };
  }, [model]);
  useEffect(
    () => () => {
      generation.current++;
    },
    [],
  );
  async function load(
    bytes: Promise<ArrayBuffer>,
    kind: "docx" | "package",
  ): Promise<void> {
    const current = ++generation.current;
    setLoading(true);
    setError("");
    setModel(undefined);
    try {
      const value = new Uint8Array(await bytes);
      const files =
        kind === "docx"
          ? importDocx(value, { documentId: crypto.randomUUID() })
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
    if (!model) return;
    const bytes = writeDocumentZip(model.serialize());
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
  return (
    <main>
      <p className="eyebrow">DOCX 读取与映射 · 只读预览</p>
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
        <button disabled={!model} onClick={download}>
          下载文档包
        </button>
      </div>
      <p>连续视图 · 字段保留原文件缓存 · 支持的文字修订保留标记</p>
      {loading && <p role="status">正在读取…</p>}
      {error && <p role="alert">{error}</p>}
      {!model && !loading && !error && (
        <p>选择文件，查看读取与样式映射结果。</p>
      )}
      {model && (
        <p role="status">
          已打开 · 修订 {model.snapshot().manifest.revision} · 批注{" "}
          {model.snapshot().annotations.length}
        </p>
      )}
      <div ref={container} className="preview" />
    </main>
  );
}
