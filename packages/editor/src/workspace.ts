/** Shared host chrome; never injected into Office document frames. */
export const workspaceControls = `
:host{color:#344b3d;font:12px "PingFang SC","Microsoft YaHei",system-ui,sans-serif}
*{box-sizing:border-box}[hidden]{display:none!important}
button,input,select,textarea{font:inherit}
button,select{border:1px solid #dce5dd;border-radius:5px;background:#fff;color:#627467;padding:6px 9px}
button{cursor:pointer}button:hover:not(:disabled){background:#edf4ed;border-color:#c4d6c6}
button:disabled{opacity:.4;cursor:default}button[aria-pressed=true]{background:#e7f0e7;color:#28714b;border-color:#c6dcc9}
button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible{outline:2px solid #2c7a56;outline-offset:2px}
input,textarea{border:1px solid #dce5dd;border-radius:5px;padding:7px;background:white;color:#425c49}
.error{color:#a54040;font-size:12px}
`;

export const annotationPopoverStyles = `
.annotation-popover{position:fixed;z-index:1000;width:300px;max-width:calc(100vw - 24px);padding:15px;background:white;border:1px solid #dbe6dc;border-radius:10px;box-shadow:0 12px 42px #243f2926,0 2px 8px #243f2912;color:#3c5543}
.annotation-popover[hidden]{display:none!important}
.annotation-popover .popover-heading{display:flex;justify-content:space-between;align-items:center;margin-bottom:11px;font-size:12px;font-weight:600}
.annotation-popover .popover-avatar{display:inline-grid;place-items:center;width:25px;height:25px;margin-right:7px;border-radius:50%;background:#e8f0e6;color:#6b8e68;font-size:11px}
.annotation-popover .popover-close{border:0;background:none;padding:2px 6px;font-size:18px;color:#93a18f}
.annotation-popover .popover-quote{margin:0 0 11px;padding:4px 9px;border-left:2px solid #a8c6a3;font-size:11px;color:#8a9986;background:#f6f9f4;max-height:46px;overflow:hidden;overflow-wrap:anywhere}
.annotation-popover textarea,.annotation-popover input[type=text]{display:block;width:100%;resize:vertical;min-height:80px;padding:10px;border:1px solid #dce5d9;border-radius:6px;background:#fcfdfb;font:inherit;line-height:1.7}
.annotation-popover input[type=text]{min-height:0;resize:none;padding:6px 8px}
.annotation-popover .composer{display:flex;align-items:center;justify-content:flex-end;flex-wrap:wrap;gap:7px;margin-top:12px}
.annotation-popover .composer button{font-size:11px;padding:6px 11px}
.annotation-popover .composer button:first-child{background:#287454;border-color:#287454;color:white}
.annotation-popover .popover-hint{margin:10px 0 0;font-size:9px;color:#a2ad9d}
`;

export interface AnnotationPopover {
  open(): void;
  close(): void;
  dispose(): void;
}

/** Owns transient UI only. Callers keep revision checks and commits at their transaction boundary. */
export function createAnnotationPopover(
  root: ShadowRoot,
  input: HTMLTextAreaElement,
  content: readonly HTMLElement[],
  options: { rect(): DOMRect | undefined; bounds(): DOMRect; quote(): string; submit(): void; cancel(): void },
): AnnotationPopover {
  const document = root.ownerDocument, window = document.defaultView!;
  const controller = new AbortController(), dialog = document.createElement("div");
  dialog.className = "annotation-popover";
  dialog.setAttribute("role", "dialog");
  dialog.setAttribute("aria-label", "新建批注");
  dialog.hidden = true;
  const heading = document.createElement("div"), author = document.createElement("span"), avatar = document.createElement("span");
  heading.className = "popover-heading"; avatar.className = "popover-avatar"; avatar.textContent = "我";
  author.append(avatar, "添加批注");
  const closeButton = document.createElement("button"); closeButton.type = "button"; closeButton.textContent = "×"; closeButton.className = "popover-close"; closeButton.setAttribute("aria-label", "取消批注");
  const quote = document.createElement("p"); quote.className = "popover-quote";
  const hint = document.createElement("p"); hint.className = "popover-hint"; hint.textContent = "Ctrl / ⌘ + Enter 发送 · Esc 取消";
  heading.append(author, closeButton); dialog.append(heading, quote, ...content, hint); root.append(dialog);
  let frame: number | undefined;
  function position(): void {
    if (dialog.hidden) return;
    const rect = options.rect();
    const width = document.documentElement.clientWidth, height = document.documentElement.clientHeight;
    const bounds = options.bounds();
    const leftEdge = Math.max(12, bounds.left + 8), rightEdge = Math.min(width - 12, bounds.right - 8);
    const topEdge = Math.max(12, bounds.top + 8), bottomEdge = Math.min(height - 12, bounds.bottom - 8);
    dialog.style.width = Math.min(300, Math.max(180, rightEdge - leftEdge)) + "px";
    dialog.style.maxHeight = Math.max(120, bottomEdge - topEdge) + "px";
    dialog.style.overflow = "auto";
    const box = dialog.getBoundingClientRect();
    const left = rect ? (rect.right + box.width + 12 < rightEdge ? rect.right + 12 : rect.left) : rightEdge - box.width;
    const top = rect ? (rect.bottom + box.height + 10 < bottomEdge ? rect.bottom + 10 : rect.top - box.height - 10) : topEdge;
    dialog.style.left = Math.max(12, Math.min(width - box.width - 12, Math.max(leftEdge, Math.min(left, rightEdge - box.width)))) + "px";
    dialog.style.top = Math.max(12, Math.min(height - box.height - 12, Math.max(topEdge, Math.min(top, bottomEdge - box.height)))) + "px";
  }
  function schedule(): void {
    if (dialog.hidden || frame !== undefined) return;
    frame = window.requestAnimationFrame(() => { frame = undefined; position(); });
  }
  function close(): void { dialog.hidden = true; }
  function cancel(): void { close(); options.cancel(); }
  closeButton.addEventListener("click", cancel, { signal: controller.signal });
  dialog.addEventListener("keydown", event => {
    if (event.isComposing) return;
    if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); cancel(); }
    else if ((event.ctrlKey || event.metaKey) && event.key === "Enter") { event.preventDefault(); options.submit(); }
  }, { signal: controller.signal });
  document.addEventListener("scroll", schedule, { capture: true, signal: controller.signal });
  // Scroll events within the editor's shadow tree do not cross its boundary.
  root.addEventListener("scroll", schedule, { capture: true, signal: controller.signal });
  window.addEventListener("resize", schedule, { signal: controller.signal });
  return {
    open() { quote.textContent = options.quote(); dialog.hidden = false; position(); },
    close,
    dispose() { controller.abort(); if (frame !== undefined) window.cancelAnimationFrame(frame); input.value = ""; dialog.remove(); },
  };
}
