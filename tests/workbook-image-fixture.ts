import { readFileSync } from 'node:fs';

// Independently assemble PNG/APNG and WebP containers from the repository's
// existing 8 × 8 raster payloads. Production header readers are not used here.
export function pngChunk(name: string, data: Uint8Array): Buffer {
  const result = Buffer.alloc(data.length + 12); result.writeUInt32BE(data.length); result.write(name, 4); result.set(data, 8);
  let crc = 0xffff_ffff;
  for (const byte of result.subarray(4, -4)) { crc ^= byte; for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb8_8320 : 0); }
  result.writeUInt32BE((crc ^ 0xffff_ffff) >>> 0, result.length - 4); return result;
}
export function webpChunk(name: string, data: Uint8Array): Buffer {
  const result = Buffer.alloc(8 + data.length + (data.length & 1)); result.write(name); result.writeUInt32LE(data.length, 4); result.set(data, 8); return result;
}
export function webpFile(...chunks: Uint8Array[]): Buffer {
  const body = Buffer.concat([Buffer.from('WEBP'), ...chunks]), header = Buffer.alloc(8); header.write('RIFF'); header.writeUInt32LE(body.length, 4); return Buffer.concat([header, body]);
}
export function animatedPng(fallback: boolean): Buffer {
  const source = readFileSync('examples/presentation/assets/sample.png'), data: Buffer[] = [];
  for (let offset = 8; offset < source.length;) { const size = source.readUInt32BE(offset); if (source.toString('ascii', offset + 4, offset + 8) === 'IDAT') data.push(source.subarray(offset + 8, offset + 8 + size)); offset += size + 12; }
  const compressed = Buffer.concat(data), animation = Buffer.alloc(8); animation.writeUInt32BE(2);
  let sequence = 0;
  const control = () => { const bytes = Buffer.alloc(26); bytes.writeUInt32BE(sequence++); bytes.writeUInt32BE(8, 4); bytes.writeUInt32BE(8, 8); bytes.writeUInt16BE(1, 20); bytes.writeUInt16BE(10, 22); return pngChunk('fcTL', bytes); };
  const frame = () => { const bytes = Buffer.alloc(4 + compressed.length); bytes.writeUInt32BE(sequence++); bytes.set(compressed, 4); return pngChunk('fdAT', bytes); };
  return Buffer.concat([source.subarray(0, 33), pngChunk('acTL', animation), ...(!fallback ? [control()] : []), pngChunk('IDAT', compressed), ...(fallback ? [control(), frame()] : []), control(), frame(), pngChunk('IEND', new Uint8Array())]);
}
export function animatedWebp(): Buffer {
  const source = readFileSync('examples/fidelity/assets/sample.webp'); let encoded: Buffer | undefined;
  for (let offset = 12; offset < source.length;) { const size = source.readUInt32LE(offset + 4), end = offset + 8 + size + (size & 1); if (source.toString('ascii', offset, offset + 4) === 'VP8 ') encoded = source.subarray(offset, end); offset = end; }
  if (!encoded) throw Error('Missing independent VP8 fixture');
  const extended = Buffer.alloc(10); extended[0] = 2; extended[4] = 7; extended[7] = 7;
  const frame = Buffer.alloc(16); frame[6] = 7; frame[9] = 7; frame[12] = 100;
  return webpFile(webpChunk('VP8X', extended), webpChunk('ANIM', new Uint8Array(6)), ...Array.from({ length: 2 }, () => webpChunk('ANMF', Buffer.concat([frame, encoded!]))));
}
export function workbookRasterCases() {
  return [
    { name: 'png', mimeType: 'image/png', bytes: readFileSync('examples/fidelity/assets/sample.png'), frames: 1 },
    { name: 'jpeg', mimeType: 'image/jpeg', bytes: readFileSync('examples/fidelity/assets/sample.jpg'), frames: 1 },
    { name: 'webp', mimeType: 'image/webp', bytes: readFileSync('examples/fidelity/assets/sample.webp'), frames: 1 },
    { name: 'apng', mimeType: 'image/png', bytes: animatedPng(false), frames: 2 },
    { name: 'apng-fallback', mimeType: 'image/png', bytes: animatedPng(true), frames: 3 },
    { name: 'animated-webp', mimeType: 'image/webp', bytes: animatedWebp(), frames: 2 },
  ];
}
