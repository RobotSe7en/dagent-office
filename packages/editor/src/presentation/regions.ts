import type { Anchor, DeepReadonly, DocumentEditingSession } from "@dagent-office/core";
import { isPresentationCommentPoint } from "@dagent-office/core";
import type { PreviewSource } from "@dagent-office/render";
export interface RegionHost {
    readonly session: DocumentEditingSession;
    readonly source: PreviewSource;
    enabled(): boolean;
    selected(): DeepReadonly<Anchor> | undefined;
    focus(): string | undefined;
    select(anchor: Anchor): void;
    error(reason: unknown): void;
    stale(value: boolean): void;
}
export interface PresentationRegions {
    paint(): void;
    cancel(): void;
    dispose(): void;
}
/** All outlines are outside the source tree and leave no source or history state. */
export function bindPresentationRegions(frame: Document, host: RegionHost): PresentationRegions {
    const controls = new AbortController(), overlays: HTMLElement[] = [];
    let drag: {
        surface: HTMLElement;
        pointerId: number;
        x: number;
        y: number;
        revision: number;
        target: Element;
        rect: {
            x: number;
            y: number;
            width: number;
            height: number;
        };
    } | undefined, disposed = false;
    const clear = (): void => { for (const node of overlays)
        node.remove(); overlays.length = 0; };
    const draw = (rect: {
        left: number;
        top: number;
        width: number;
        height: number;
    }, color: string, dashed = false): void => { const node = frame.createElement("div"); node.dataset.sdkPptAnnotation = ""; node.style.cssText = `all:initial!important;position:fixed!important;pointer-events:none!important;z-index:2147483645!important;box-sizing:border-box!important;left:${rect.left}px!important;top:${rect.top}px!important;width:${Math.max(1, rect.width)}px!important;height:${Math.max(1, rect.height)}px!important;border:2px ${dashed ? "dashed" : "solid"} ${color}!important`; frame.documentElement.append(node); overlays.push(node); };
    const paint = (): void => {
        clear();
        if (disposed)
            return;
        let stale = false;
        const keys = new Map<string, string>();
        for (const annotation of host.session.getAnnotations())
            if (annotation.status !== "needs_reanchor") {
                let changed = false;
                if (annotation.anchor.kind === "region" && !isPresentationCommentPoint(annotation.anchor)) {
                    const id = annotation.anchor.surfaceId;
                    let key = keys.get(id);
                    if (!key) {
                        key = host.source.layoutKey(id);
                        keys.set(id, key);
                    }
                    changed = annotation.anchor.layoutKey !== key;
                }
                stale ||= changed;
                if (host.focus() === annotation.annotationId)
                    for (const rect of host.source.rectangles(annotation.anchor))
                        draw(rect, "#a15b00", changed);
            }
        const anchor = host.selected();
        if (anchor?.kind === "region")
            for (const rect of host.source.rectangles(anchor))
                draw(rect, "#245aa5");
        if (drag) {
            const box = drag.surface.getBoundingClientRect(), r = drag.rect;
            draw({ left: box.left + r.x * 4 / 3, top: box.top + r.y * 4 / 3, width: r.width * 4 / 3, height: r.height * 4 / 3 }, "#245aa5", true);
        }
        host.stale(stale);
    };
    const cancel = (): void => { const old = drag; drag = undefined; if (old)
        try {
            old.target.releasePointerCapture(old.pointerId);
        }
        catch { /* Capture can already be gone. */ } paint(); };
    frame.addEventListener("pointerdown", event => {
        if (!host.enabled() || event.button !== 0)
            return;
        const target = event.target as Element | null, surface = target?.closest<HTMLElement>('[data-office-kind="slide"][data-node-id]');
        if (!surface || host.source.elements.get(surface.dataset.nodeId!)?.element !== surface)
            return;
        event.preventDefault();
        const box = surface.getBoundingClientRect(), x = Math.min(box.width, Math.max(0, event.clientX - box.left)) * .75, y = Math.min(box.height, Math.max(0, event.clientY - box.top)) * .75;
        drag = { surface, pointerId: event.pointerId, x, y, revision: host.session.snapshot().manifest.revision, target: surface, rect: { x, y, width: 0, height: 0 } };
        surface.setPointerCapture(event.pointerId);
        paint();
    }, { signal: controls.signal });
    frame.addEventListener("pointermove", event => { if (!drag || event.pointerId !== drag.pointerId)
        return; const box = drag.surface.getBoundingClientRect(), x = Math.min(box.width, Math.max(0, event.clientX - box.left)) * .75, y = Math.min(box.height, Math.max(0, event.clientY - box.top)) * .75; drag.rect = { x: Math.min(x, drag.x), y: Math.min(y, drag.y), width: Math.abs(x - drag.x), height: Math.abs(y - drag.y) }; paint(); }, { signal: controls.signal });
    frame.addEventListener("pointerup", event => { if (!drag || event.pointerId !== drag.pointerId)
        return; const old = drag; cancel(); if (old.revision !== host.session.snapshot().manifest.revision || old.rect.width < 1 || old.rect.height < 1)
        return; try {
        host.select(host.source.slideRegion(old.surface.dataset.nodeId!, old.rect));
    }
    catch (error) {
        host.error(error);
    } }, { signal: controls.signal });
    frame.addEventListener("pointercancel", cancel, { signal: controls.signal });
    frame.addEventListener("lostpointercapture", () => { if (drag)
        cancel(); }, { signal: controls.signal });
    frame.addEventListener("keydown", event => { if (event.key === "Escape")
        cancel(); }, { signal: controls.signal });
    frame.defaultView!.addEventListener("blur", cancel, { signal: controls.signal });
    frame.defaultView!.addEventListener("resize", paint, { signal: controls.signal });
    frame.addEventListener("scroll", paint, { capture: true, signal: controls.signal });
    paint();
    return { paint, cancel, dispose() { if (disposed)
            return; disposed = true; controls.abort(); const old = drag; drag = undefined; if (old)
            try {
                old.target.releasePointerCapture(old.pointerId);
            }
            catch { /* Already released. */ } clear(); } };
}
