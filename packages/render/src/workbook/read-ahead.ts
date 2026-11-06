/** One optional background job. Cancelling an uncooperative read never starts
 * another background read until it settles. Foreground work never waits here. */
export function workbookIdle(view: Window, signal: AbortSignal): Promise<boolean> {
  return new Promise(resolve => {
    if (signal.aborted) { resolve(false); return; }
    let handle: number;
    const idle = typeof view.requestIdleCallback === "function";
    const finish = (): void => { signal.removeEventListener("abort", cancel); resolve(!signal.aborted); };
    const cancel = (): void => { if (idle) view.cancelIdleCallback(handle); else view.clearTimeout(handle); finish(); };
    signal.addEventListener("abort", cancel, { once: true });
    handle = idle ? view.requestIdleCallback(finish, { timeout: 200 }) : view.setTimeout(finish, 32);
  });
}
export class WorkbookReadAhead {
  private work: ((signal: AbortSignal) => Promise<boolean>) | undefined;
  private controller: AbortController | undefined;
  private busy = false;
  private disposed = false;
  constructor(private readonly view: Window) {}
  set(work: (signal: AbortSignal) => Promise<boolean>): void { this.stop(); this.work = work; this.start(); }
  stop(): void { this.work = undefined; this.controller?.abort(); }
  interrupt(): void { this.controller?.abort(); this.start(); }
  private start(): void {
    if (this.busy || this.disposed || !this.work) return;
    const work = this.work, controller = new AbortController(); this.controller = controller; this.busy = true;
    void (async () => {
      try {
        while (this.work === work && !controller.signal.aborted && await workbookIdle(this.view, controller.signal)) {
          if (controller.signal.aborted || this.work !== work) break;
          if (!await work(controller.signal)) { if (this.work === work) this.work = undefined; break; }
        }
      } catch { if (!controller.signal.aborted && this.work === work) this.work = undefined; }
      finally {
        this.busy = false; if (this.controller === controller) this.controller = undefined;
        if (!this.disposed && this.work && (controller.signal.aborted || this.work !== work)) this.start();
      }
    })();
  }
  dispose(): void { this.disposed = true; this.stop(); }
}
