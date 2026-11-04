import { DocumentError } from "@dagent-office/core";
import type { WorkbookPreparedAssets } from "./assets.js";
import { checkWindowCancellation } from "./window-source.js";
import { markViewNode } from "../view-nodes.js";

/** Only the raster payload. The object compositor owns its fill, outline,
 * shadow, orientation and pane clipping, independently of this source crop. */
export function createWorkbookImageContent(source: HTMLElement, path: string, assets: WorkbookPreparedAssets, maxCropScale = 1_000_000): { readonly element: HTMLElement; readonly image: HTMLImageElement } {
  if (!Number.isFinite(maxCropScale) || maxCropScale < 1) throw new DocumentError("InvalidOperation", "Image cropping requires a finite scale budget of at least one.");
  const sourceImage = source.children[0];
  if (source.dataset.officeKind !== "image" || source.children.length !== 1 || sourceImage?.localName !== "img") throw new DocumentError("InvalidDocument", "A workbook image requires exactly one source raster element.");
  const crop = (side: string): number => {
    const value = source.getAttribute("data-office-crop-" + side), number = value === null ? 0 : Number(value);
    if (value !== null && !/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(value) || !Number.isFinite(number) || number < 0 || number >= 1) throw new DocumentError("InvalidDocument", "Image crop fractions must be finite values in [0,1).");
    return number;
  };
  const left = crop("left"), right = crop("right"), top = crop("top"), bottom = crop("bottom"), width = 1 - left - right, height = 1 - top - bottom;
  if (width <= 0 || height <= 0) throw new DocumentError("InvalidDocument", "An image crop must retain a visible area.");
  if (1 / width > maxCropScale || 1 / height > maxCropScale) throw new DocumentError("ResourceLimitExceeded", "Workbook image cropping exceeds its scale budget.");
  const url = assets.url(path); if (!url) throw new DocumentError("AssetUnavailable", "This image is not present in the candidate's resource lease.");
  const document = source.ownerDocument, root = markViewNode(document.createElement("div")), image = document.createElement("img");
  root.style.cssText = "position:absolute;left:0;top:0;width:100%;height:100%;overflow:hidden";
  image.dataset.sourceNodeId = sourceImage.getAttribute("data-node-id")!; image.alt = sourceImage.getAttribute("alt") ?? source.getAttribute("data-office-alt") ?? "";
  image.style.cssText = `position:absolute;max-width:none;max-height:none;left:${-left / width * 100}%;top:${-top / height * 100}%;width:${100 / width}%;height:${100 / height}%;object-fit:fill`;
  image.src = url; root.append(image); return Object.freeze({ element: root, image });
}

/** Decode a detached candidate before an atomic window replacement. Failure or
 * cancellation clears only this candidate's sources; its owner releases the
 * asset lease, while the currently visible window remains untouched. */
export async function decodeWorkbookImages(images: readonly HTMLImageElement[], signal: AbortSignal, maxImages = 5000): Promise<void> {
  if (!Number.isSafeInteger(maxImages) || maxImages < 1 || images.some(image => image.isConnected)) throw new DocumentError("InvalidOperation", "Image decoding requires a detached candidate and a positive image budget.");
  if (images.length > maxImages) throw new DocumentError("ResourceLimitExceeded", "Workbook raster instances exceed their decode budget.");
  let abort = (): void => {};
  const cancelled = new Promise<never>((_, reject) => { abort = () => reject(new DocumentError("OperationCancelled", "Workbook image decoding was cancelled.")); });
  signal.addEventListener("abort", abort, { once: true });
  try {
    checkWindowCancellation(signal);
    // Bound simultaneous decode requests as well as the complete candidate.
    for (let index = 0; index < images.length; index += 8) {
      await Promise.race([Promise.all(images.slice(index, index + 8).map(image => image.decode())), cancelled]); checkWindowCancellation(signal);
    }
  } catch (error) {
    for (const image of images) image.removeAttribute("src");
    if (error instanceof DocumentError) throw error;
    throw new DocumentError("InvalidDocument", "A workbook raster could not be decoded.");
  } finally { signal.removeEventListener("abort", abort); }
}
