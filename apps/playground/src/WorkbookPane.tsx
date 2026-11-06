import { useEffect, useRef } from "react";
import { mountDocumentPreview } from "@doffi/render";
import type { WorkbookPreview, WorkbookPreviewSource } from "@doffi/render";

export function WorkbookPane({ source, preparing, activation, onReady, onError }: {
  source: WorkbookPreviewSource; preparing: boolean; activation?: object | undefined;
  onReady(source: WorkbookPreviewSource): void; onError(source: WorkbookPreviewSource, error: Error): void;
}) {
  const container = useRef<HTMLDivElement>(null), preview = useRef<WorkbookPreview>(undefined), changing = useRef(false);
  const boundSource = useRef<WorkbookPreviewSource>(undefined);
  const callbacks = useRef({ onReady, onError, source }); callbacks.current = { onReady, onError, source };
  const resize = useRef<() => Promise<void>>(async () => {});
  useEffect(() => {
    const element = container.current!;
    resize.current = async () => {
      const view = preview.current; if (!view || changing.current || !element.clientWidth) return;
      const widthPt = Math.max(180, element.clientWidth * .75 - 49.5), viewport = view.layout().viewport;
      if (Math.abs(viewport.widthPt - widthPt) > .5) await view.setViewport({ ...viewport, widthPt });
    };
    const observer = new ResizeObserver(() => { void resize.current().catch(error => {
      if (!changing.current) callbacks.current.onError(callbacks.current.source, error instanceof Error ? error : new Error("工作簿尺寸更新失败。"));
    }); });
    observer.observe(element);
    return () => { observer.disconnect(); preview.current?.dispose(); preview.current = undefined; boundSource.current = undefined; };
  }, []);
  useEffect(() => {
    let cancelled = false; changing.current = true;
    try {
      if (preview.current && boundSource.current !== source) preview.current.update(source);
      else if (!preview.current) preview.current = mountDocumentPreview(container.current!, source, {
        widthPt: Math.max(180, container.current!.clientWidth * .75 - 49.5), heightPt: 450,
        onError(error) { if (!changing.current) callbacks.current.onError(callbacks.current.source, new Error(error.message)); },
      });
      boundSource.current = source;
      const view = preview.current;
      void view.settled().then(report => {
        if (cancelled || report.documentId !== source.cells.snapshot.manifest.documentId || report.revision !== source.cells.snapshot.manifest.revision) return;
        changing.current = false; callbacks.current.onReady(source);
        void resize.current().catch(error => { if (!cancelled) callbacks.current.onError(source, error instanceof Error ? error : new Error("工作簿尺寸更新失败。")); });
      }).catch(error => {
        if (cancelled) return; changing.current = false; boundSource.current = undefined;
        // A failed replacement keeps the previous published window. An initial
        // failure has no retained view and can be retried by a new mount.
        try { view.layout(); } catch { view.dispose(); if (preview.current === view) preview.current = undefined; }
        callbacks.current.onError(source, error instanceof Error ? error : new Error("工作簿显示失败。"));
      });
    } catch (error) { changing.current = false; boundSource.current = undefined; callbacks.current.onError(source, error instanceof Error ? error : new Error("工作簿显示失败。")); }
    return () => { cancelled = true; };
  }, [source, activation]);
  return <div className={"workbook-pane" + (preparing ? " preparing" : "")} aria-label="Excel 只读预览" aria-hidden={preparing || undefined} ref={container} />;
}
