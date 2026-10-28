import { DocumentError, PackagePathSchema } from "@dagent-office/core";
import type { DocumentReadLimits } from "@dagent-office/core";

export interface ZipEntry {
  readonly path: string;
  readonly directory: boolean;
  readonly method: number;
  readonly crc: number;
  readonly size: number;
  readonly compressedSize: number;
  readonly dataOffset: number;
}

export function invalidZip(): never {
  throw new DocumentError("InvalidDocument", "The ZIP archive is malformed or contains ambiguous entries.");
}

export function zipLimit(value: number, maximum: number): void {
  if (value > maximum) throw new DocumentError("ResourceLimitExceeded", "The ZIP archive exceeds the configured limit.");
}

function unsupportedZip(): never {
  throw new DocumentError("UnsupportedFeature", "The ZIP archive uses an unsupported container feature.");
}

const crcTable = Uint32Array.from({ length: 256 }, (_, value) => {
  for (let bit = 0; bit < 8; bit++) value = (value >>> 1) ^ (value & 1 ? 0xedb88320 : 0);
  return value >>> 0;
});

export function crc32(bytes: Uint8Array): number {
  let value = 0xffffffff;
  for (const byte of bytes) value = (value >>> 8) ^ crcTable[(value ^ byte) & 255]!;
  return (value ^ 0xffffffff) >>> 0;
}

const codecs = globalThis as unknown as {
  TextDecoder: new (label: string, options: { fatal: boolean; ignoreBOM: boolean }) => { decode(bytes: Uint8Array): string };
};

function decodeName(bytes: Uint8Array, flags: number): string {
  // Do not guess an OS code page. ASCII and explicitly marked UTF-8 are portable.
  if (!(flags & 0x800) && bytes.some((byte) => byte >= 128)) unsupportedZip();
  try {
    // A leading U+FEFF is part of a filename, not a text-file encoding marker.
    return new codecs.TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return invalidZip();
  }
}

/** Inspect unmodified central/local records before a decompressor can normalize names or allocate output. */
export function inspectZip(bytes: Uint8Array, limits: DocumentReadLimits): ZipEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset: number): number => view.getUint16(offset, true);
  const u32 = (offset: number): number => view.getUint32(offset, true);
  function range(offset: number, length: number, end = bytes.length): void {
    if (offset < 0 || length < 0 || offset + length > end) invalidZip();
  }
  let end = bytes.length - 22;
  for (; end >= Math.max(0, bytes.length - 65557); end--) {
    if (u32(end) === 0x06054b50 && end + 22 + u16(end + 20) === bytes.length) break;
  }
  if (end < 0 || end < bytes.length - 65557) invalidZip();
  if (u16(end + 4) || u16(end + 6) || u16(end + 8) !== u16(end + 10)) unsupportedZip();
  const count = u16(end + 10);
  const centralSize = u32(end + 12);
  const centralStart = u32(end + 16);
  if (count === 0xffff || centralSize === 0xffffffff || centralStart === 0xffffffff) unsupportedZip();
  zipLimit(count, limits.maxFiles);
  if (centralStart + centralSize !== end) invalidZip();
  const names = new Set<string>();
  const entries: (ZipEntry & { localOffset: number; endOffset: number })[] = [];
  let total = 0;
  let offset = centralStart;

  function extras(start: number, length: number, name: string, rawName: Uint8Array): void {
    const stop = start + length;
    range(start, length, centralStart > start ? centralStart : end);
    for (let cursor = start; cursor < stop;) {
      range(cursor, 4, stop);
      const id = u16(cursor);
      const size = u16(cursor + 2);
      cursor += 4;
      range(cursor, size, stop);
      if (id === 1) unsupportedZip(); // ZIP64 is outside this bounded classic-ZIP adapter.
      if (id === 0x7075) {
        if (size < 5 || bytes[cursor] !== 1 || u32(cursor + 1) !== crc32(rawName)
          || decodeName(bytes.subarray(cursor + 5, cursor + size), 0x800) !== name) invalidZip();
      }
      cursor += size;
    }
  }

  for (let index = 0; index < count; index++) {
    range(offset, 46, end);
    if (u32(offset) !== 0x02014b50) invalidZip();
    const flags = u16(offset + 8);
    const method = u16(offset + 10);
    const crc = u32(offset + 16);
    const compressedSize = u32(offset + 20);
    const size = u32(offset + 24);
    const nameLength = u16(offset + 28);
    const extraLength = u16(offset + 30);
    const commentLength = u16(offset + 32);
    const localOffset = u32(offset + 42);
    if ((flags & ~0x80e) || ![0, 8].includes(method) || u16(offset + 34)) unsupportedZip();
    if ([compressedSize, size, localOffset].includes(0xffffffff)) unsupportedZip();
    if (((u32(offset + 38) >>> 16) & 0xf000) === 0xa000) unsupportedZip();
    range(offset + 46, nameLength + extraLength + commentLength, end);
    const rawName = bytes.subarray(offset + 46, offset + 46 + nameLength);
    const name = decodeName(rawName, flags);
    const directory = name.endsWith("/");
    const path = directory ? name.slice(0, -1) : name;
    if (!PackagePathSchema.safeParse(path).success || names.has(name)) invalidZip();
    names.add(name);
    zipLimit(size, limits.maxFileBytes);
    total += size;
    zipLimit(total, limits.maxPackageBytes);
    if (directory && size !== 0) invalidZip();
    extras(offset + 46 + nameLength, extraLength, name, rawName);
    range(localOffset, 30, centralStart);
    if (u32(localOffset) !== 0x04034b50 || u16(localOffset + 6) !== flags || u16(localOffset + 8) !== method) invalidZip();
    const localNameLength = u16(localOffset + 26);
    const localExtraLength = u16(localOffset + 28);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    range(localOffset + 30, localNameLength + localExtraLength, centralStart);
    const localName = bytes.subarray(localOffset + 30, localOffset + 30 + localNameLength);
    if (localName.length !== rawName.length || localName.some((byte, position) => byte !== rawName[position])) invalidZip();
    extras(localOffset + 30 + localNameLength, localExtraLength, name, rawName);
    for (const [field, expected] of [[14, crc], [18, compressedSize], [22, size]] as const) {
      const actual = u32(localOffset + field);
      if (actual !== expected && (!(flags & 8) || actual !== 0)) invalidZip();
    }
    range(dataOffset, compressedSize, centralStart);
    let endOffset = dataOffset + compressedSize;
    if (flags & 8) {
      const descriptor = (start: number): boolean => start + 12 <= centralStart
        && u32(start) === crc && u32(start + 4) === compressedSize && u32(start + 8) === size;
      if (endOffset + 16 <= centralStart && u32(endOffset) === 0x08074b50 && descriptor(endOffset + 4)) endOffset += 16;
      else if (descriptor(endOffset)) endOffset += 12;
      else invalidZip();
    }
    entries.push({ path, directory, method, crc, size, compressedSize, dataOffset, localOffset, endOffset });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  if (offset !== end) invalidZip();
  let physicalOffset = 0;
  for (const entry of entries.sort((left, right) => left.localOffset - right.localOffset)) {
    if (entry.localOffset !== physicalOffset) invalidZip();
    physicalOffset = entry.endOffset;
  }
  if (physicalOffset !== centralStart) invalidZip();
  const files = new Set(entries.filter((entry) => !entry.directory).map((entry) => entry.path));
  for (const entry of entries) {
    if (entry.directory && files.has(entry.path)) invalidZip();
    const parts = entry.path.split("/");
    for (let index = 1; index < parts.length; index++) if (files.has(parts.slice(0, index).join("/"))) invalidZip();
  }
  return entries;
}
