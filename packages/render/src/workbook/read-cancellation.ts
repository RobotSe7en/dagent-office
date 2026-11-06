import { DocumentError } from '@doffi/core';

/** One removable abort waiter per read; late replies never retain a shared
 * cancellation promise's completed responses or delay caller cleanup. */
export function waitForWorkbookRead<Value>(work: Promise<Value>, signal: AbortSignal): Promise<Value> {
  return new Promise((resolve, reject) => {
    const stopped = (): void => {
      signal.removeEventListener('abort', stopped);
      reject(new DocumentError('OperationCancelled', 'Workbook reading was cancelled.'));
    };
    if (signal.aborted) stopped(); else signal.addEventListener('abort', stopped, { once: true });
    void work.then(value => { signal.removeEventListener('abort', stopped); resolve(value); }, error => {
      signal.removeEventListener('abort', stopped); reject(error);
    });
  });
}
