import { DocumentError } from "@dagent-office/core";

export interface WorkbookImageFootprint { readonly width: number; readonly height: number; readonly frames: number; readonly pixels: number }
const invalid = (): never => { throw new DocumentError("InvalidDocument", "An image has invalid or contradictory raster geometry."); };
const unsupported = (): never => { throw new DocumentError("UnsupportedFeature", "This raster encoding has no supported pre-decode geometry."); };

/** Read bounded container/frame headers before giving bytes to a browser image
 * decoder. This is not pixel decoding or a replacement for core's asset hash
 * validation. Animation uses full-canvas pixels per frame conservatively; loop
 * count does not allocate additional distinct frames. */
export function workbookImageFootprint(bytes: Uint8Array, mimeType: string, maxPixels: number, maxFrames: number): WorkbookImageFootprint {
  if (![maxPixels, maxFrames].every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Image pixel/frame limits must be positive safe integers.");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const signature = (offset: number, values: readonly number[]): boolean => values.every((value, index) => bytes[offset + index] === value);
  const tag = (offset: number): string => String.fromCharCode(...bytes.subarray(offset, offset + 4));
  const uint24 = (offset: number): number => bytes[offset]! + bytes[offset + 1]! * 256 + bytes[offset + 2]! * 65536;
  const footprint = (width: number, height: number, frames: number): WorkbookImageFootprint => {
    if (![width, height, frames].every(value => Number.isSafeInteger(value) && value > 0)) invalid();
    const pixels = width * height * frames;
    if (!Number.isSafeInteger(pixels) || pixels > maxPixels || frames > maxFrames) throw new DocumentError("ResourceLimitExceeded", "Workbook images exceed their decoded pixel or frame budget.");
    return Object.freeze({ width, height, frames, pixels });
  };
  if (mimeType === "image/png") {
    if (!signature(0, [137, 80, 78, 71, 13, 10, 26, 10])) invalid();
    let width = 0, height = 0, declared = 0, controls = 0, hasData = false, defaultInAnimation = false;
    for (let offset = 8; offset < bytes.length;) {
      if (offset + 12 > bytes.length) invalid();
      const size = view.getUint32(offset), name = tag(offset + 4), data = offset + 8, end = offset + 12 + size;
      if (end > bytes.length) invalid();
      if (offset === 8 && name !== "IHDR") invalid();
      if (name === "IHDR") {
        if (offset !== 8 || size !== 13) invalid();
        width = view.getUint32(data); height = view.getUint32(data + 4);
        if (width > 0x7fff_ffff || height > 0x7fff_ffff) invalid(); footprint(width, height, 1);
      } else if (name === "acTL") {
        if (size !== 8 || declared || hasData) invalid();
        declared = view.getUint32(data); footprint(width, height, declared);
      } else if (name === "fcTL") {
        if (!declared || size !== 26) invalid();
        const w = view.getUint32(data + 4), h = view.getUint32(data + 8), x = view.getUint32(data + 12), y = view.getUint32(data + 16);
        if (!w || !h || x + w > width || y + h > height) invalid();
        if (!hasData) { if (controls || w !== width || h !== height || x || y) invalid(); defaultInAnimation = true; }
        if (++controls > declared) invalid();
      } else if (name === "IDAT") hasData = true;
      else if (name === "fdAT" && (!declared || !controls || size < 4)) invalid();
      else if (name === "IEND") {
        if (size || end !== bytes.length || !hasData || declared !== controls) invalid();
        return footprint(width, height, declared ? declared + (defaultInAnimation ? 0 : 1) : 1);
      }
      offset = end;
    }
    return invalid();
  }
  if (mimeType === "image/jpeg") {
    if (!signature(0, [255, 216]) || !signature(bytes.length - 2, [255, 217])) invalid();
    let result: WorkbookImageFootprint | undefined;
    for (let offset = 2; offset < bytes.length;) {
      if (bytes[offset++] !== 255) invalid();
      while (bytes[offset] === 255) offset++;
      const marker = bytes[offset++];
      if (marker === undefined) return invalid();
      if (marker === 0xda) return result ?? invalid();
      if (marker === 0xd9 || marker >= 0xd0 && marker <= 0xd8 || marker === 0 || marker === 1 || offset + 2 > bytes.length) invalid();
      if (marker === 0xde) unsupported(); // Hierarchical, potentially multi-frame JPEG.
      const size = view.getUint16(offset); if (size < 2 || offset + size > bytes.length) invalid();
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
        if (![0xc0, 0xc1, 0xc2].includes(marker)) unsupported();
        if (result || size < 8) invalid(); result = footprint(view.getUint16(offset + 5), view.getUint16(offset + 3), 1);
      }
      offset += size;
    }
    return invalid();
  }
  if (mimeType !== "image/webp") return unsupported();
  if (bytes.length < 20 || tag(0) !== "RIFF" || tag(8) !== "WEBP" || view.getUint32(4, true) !== bytes.length - 8) invalid();
  const bitstream = (name: string, data: number, size: number): readonly [number, number] | undefined => {
    if (name === "VP8 ") {
      if (size < 10 || (bytes[data]! & 1) !== 0 || !signature(data + 3, [157, 1, 42])) invalid();
      return [view.getUint16(data + 6, true) & 0x3fff, view.getUint16(data + 8, true) & 0x3fff];
    }
    if (name === "VP8L") {
      if (size < 5 || bytes[data] !== 47) invalid(); const packed = view.getUint32(data + 1, true);
      if (packed >>> 29) invalid(); return [(packed & 0x3fff) + 1, (packed >>> 14 & 0x3fff) + 1];
    }
    return undefined;
  };
  const chunk = (offset: number, limit: number) => {
    if (offset + 8 > limit) invalid();
    const size = view.getUint32(offset + 4, true), end = offset + 8 + size + (size & 1);
    if (end > limit) invalid(); return { name: tag(offset), data: offset + 8, size, end };
  };
  let width = 0, height = 0, extended = false, animated = false, control = false, frames = 0, still: readonly [number, number] | undefined;
  for (let offset = 12; offset < bytes.length;) {
    const { name, data, size, end } = chunk(offset, bytes.length);
    if (name === "VP8X") {
      if (offset !== 12 || size !== 10) invalid(); extended = true; animated = Boolean(bytes[data]! & 2);
      width = uint24(data + 4) + 1; height = uint24(data + 7) + 1; footprint(width, height, 1);
    } else if (name === "ANIM") {
      if (!animated || control || frames || size !== 6) invalid(); control = true;
    } else if (name === "ANMF") {
      if (!animated || !control || size < 24) invalid();
      const x = uint24(data) * 2, y = uint24(data + 3) * 2, w = uint24(data + 6) + 1, h = uint24(data + 9) + 1;
      if (x + w > width || y + h > height) invalid(); footprint(width, height, ++frames);
      let encoded: readonly [number, number] | undefined;
      for (let inner = data + 16; inner < data + size;) {
        const nested = chunk(inner, data + size), dimensions = bitstream(nested.name, nested.data, nested.size);
        if (dimensions) { if (encoded) invalid(); encoded = dimensions; }
        inner = nested.end;
      }
      if (!encoded || encoded[0] !== w || encoded[1] !== h) invalid();
    } else {
      const dimensions = bitstream(name, data, size);
      if (dimensions) {
        if (animated || still) invalid(); still = dimensions;
        if (!extended) { width = dimensions[0]; height = dimensions[1]; }
        if (width !== dimensions[0] || height !== dimensions[1]) invalid(); footprint(width, height, 1);
      }
    }
    offset = end;
  }
  if (animated ? !frames : !still) invalid();
  return footprint(width, height, animated ? frames : 1);
}
