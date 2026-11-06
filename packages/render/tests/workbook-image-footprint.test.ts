import { describe, expect, it } from 'vitest';
import { animatedPng, animatedWebp, pngChunk, webpChunk, webpFile, workbookRasterCases } from '../../../tests/workbook-image-fixture.js';
import { workbookImageFootprint } from '../src/workbook/image-footprint.js';

const inspect = (bytes: Uint8Array, mime = 'image/png', pixels = 1024, frames = 10) => workbookImageFootprint(bytes, mime, pixels, frames);
describe('image geometry before workbook raster decoding', () => {
  for (const sample of workbookRasterCases()) it(`bounds ${sample.name} by the whole canvas and all distinct frames`, () => {
    expect(inspect(sample.bytes, sample.mimeType, 64 * sample.frames, sample.frames)).toEqual({ width: 8, height: 8, frames: sample.frames, pixels: 64 * sample.frames });
    expect(() => inspect(sample.bytes, sample.mimeType, 64 * sample.frames - 1)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    if (sample.frames > 1) expect(() => inspect(sample.bytes, sample.mimeType, 1024, sample.frames - 1)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    for (let length = 0; length < sample.bytes.length; length++) expect(() => inspect(sample.bytes.subarray(0, length), sample.mimeType)).toThrow();
    const padded = Buffer.concat([Buffer.from([0]), sample.bytes, Buffer.from([0])]);
    expect(inspect(padded.subarray(1, -1), sample.mimeType).pixels).toBe(64 * sample.frames);
  });
  it('rejects enormous PNG/APNG canvases, frame counts and contradictory subframes', () => {
    const original = animatedPng(false);
    for (const [offset, value, code] of [[16, 1_000_000, 'ResourceLimitExceeded'], [41, 1000, 'ResourceLimitExceeded'], [65, 9, 'InvalidDocument'], [69, 0, 'InvalidDocument'], [73, 1, 'InvalidDocument'], [41, 1, 'InvalidDocument']] as const) {
      const bytes = Buffer.from(original); bytes.writeUInt32BE(value, offset);
      expect(() => inspect(bytes)).toThrowError(expect.objectContaining({ code }));
    }
    const duplicate = Buffer.concat([original.subarray(0, 33), pngChunk('IHDR', original.subarray(16, 29)), original.subarray(33)]);
    expect(() => inspect(duplicate)).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
  });
  it('checks WebP encoded frame geometry instead of trusting small outer frame declarations', () => {
    const original = animatedWebp();
    // VP8X canvas (24-bit minus one), first ANMF width, first VP8 bitstream width.
    for (const [offset, value, code] of [[24, 250, 'ResourceLimitExceeded'], [58, 8, 'InvalidDocument'], [82, 9, 'InvalidDocument']] as const) {
      const bytes = Buffer.from(original); bytes[offset] = value;
      expect(() => inspect(bytes, 'image/webp')).toThrowError(expect.objectContaining({ code }));
    }
    const lossless = Buffer.from([47, 7, 192, 1, 0]); // 8 × 8 VP8L header, no compressed body: geometry only.
    expect(inspect(webpFile(webpChunk('VP8L', lossless)), 'image/webp').pixels).toBe(64);
    lossless[4] = 224;
    expect(() => inspect(webpFile(webpChunk('VP8L', lossless)), 'image/webp')).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
  });
  it('bounds supported JPEG frame headers and rejects unsupported encodings and invalid budgets', () => {
    for (const marker of [0xc0, 0xc1, 0xc2]) {
      const bytes = Buffer.from([255, 216, 255, marker, 0, 11, 8, 0, 8, 0, 8, 1, 1, 0x11, 0, 255, 218, 0, 8, 1, 1, 0, 0, 63, 0, 0, 255, 217]);
      expect(inspect(bytes, 'image/jpeg').pixels).toBe(64); bytes[9] = 255;
      expect(() => inspect(bytes, 'image/jpeg')).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
      bytes[3] = 0xc3;
      expect(() => inspect(bytes, 'image/jpeg')).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
    }
    for (const limit of [0, -1, 1.5, Infinity, Number.MAX_SAFE_INTEGER + 1]) {
      expect(() => inspect(new Uint8Array(), 'image/png', limit)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
      expect(() => inspect(new Uint8Array(), 'image/png', 64, limit)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
    }
  });
});
