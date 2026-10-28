import { DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, openDocument } from "@dagent-office/core";
import type { DocumentPackage, DocumentReadLimits, OpenDocumentOptions } from "@dagent-office/core";
import { Inflate, Zip, ZipDeflate } from "fflate/browser";
import { crc32, inspectZip, invalidZip, zipLimit } from "./zip-format.js";
import type { ZipEntry } from "./zip-format.js";

export interface DocumentZipOptions extends OpenDocumentOptions {
  readonly maxZipBytes?: number;
}

export const DEFAULT_MAX_ZIP_BYTES: number = 64 * 1024 * 1024;

function zipOptions(options: DocumentZipOptions): { limits: DocumentReadLimits; maxZipBytes: number } {
  const limits = { ...DEFAULT_DOCUMENT_READ_LIMITS, ...options.limits };
  for (const [name, value] of Object.entries(options.limits ?? {})) {
    if (!Object.hasOwn(DEFAULT_DOCUMENT_READ_LIMITS, name) || !Number.isSafeInteger(value) || value <= 0) {
      throw new DocumentError("InvalidDocument", "Document limits must be positive safe integers with supported names.");
    }
  }
  const maxZipBytes = options.maxZipBytes ?? DEFAULT_MAX_ZIP_BYTES;
  if (!Number.isSafeInteger(maxZipBytes) || maxZipBytes <= 0) throw new DocumentError("InvalidDocument", "The ZIP byte limit must be a positive safe integer.");
  return { limits, maxZipBytes };
}

function extract(bytes: Uint8Array, entry: ZipEntry, limits: DocumentReadLimits): Uint8Array {
  const compressed = bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize);
  let result: Uint8Array;
  if (entry.method === 0) {
    if (entry.compressedSize !== entry.size) invalidZip();
    result = new Uint8Array(compressed);
  } else {
    result = new Uint8Array(entry.size);
    let written = 0;
    const inflater = new Inflate((chunk) => {
      written += chunk.length;
      zipLimit(written, limits.maxFileBytes);
      if (written > entry.size) invalidZip();
      result.set(chunk, written - chunk.length);
    });
    // Feed bounded compressed chunks; forged sizes cannot trigger an unbounded full inflate.
    if (!compressed.length) invalidZip();
    for (let cursor = 0; cursor < compressed.length; cursor += 128) {
      inflater.push(compressed.subarray(cursor, cursor + 128), cursor + 128 >= compressed.length);
    }
    if (written !== entry.size) invalidZip();
  }
  if (crc32(result) !== entry.crc) invalidZip();
  return result;
}

/** Read a bounded ZIP into independent package files. Use core.openDocument to validate the document. */
export function readDocumentZip(input: Uint8Array, options: DocumentZipOptions = {}): DocumentPackage {
  const { limits, maxZipBytes } = zipOptions(options);
  if (!(input instanceof Uint8Array)) invalidZip();
  zipLimit(input.byteLength, maxZipBytes);
  try {
    const entries = inspectZip(input, limits);
    const files = new Map<string, Uint8Array>();
    for (const entry of entries) {
      const bytes = extract(input, entry, limits);
      if (!entry.directory) files.set(entry.path, bytes);
    }
    return files;
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return invalidZip();
  }
}

/** Write a complete validated document package, retaining source files and resource bytes. */
export function writeDocumentZip(input: DocumentPackage, options: DocumentZipOptions = {}): Uint8Array {
  const { limits, maxZipBytes } = zipOptions(options);
  const files = new Map<string, Uint8Array>();
  let inputSize = 0;
  for (const [path, bytes] of input) {
    if (files.has(path) || !(bytes instanceof Uint8Array)) invalidZip();
    zipLimit(files.size + 1, limits.maxFiles);
    zipLimit(bytes.length, limits.maxFileBytes);
    inputSize += bytes.length;
    zipLimit(inputSize, limits.maxPackageBytes);
    files.set(path, new Uint8Array(bytes));
  }
  // Validate stable copied bytes, including every declared resource. Normalization
  // remains an explicit document-model serialization step rather than an archive side effect.
  openDocument(files, { limits });
  const chunks: Uint8Array[] = [];
  let length = 0;
  try {
    const archive = new Zip((error, chunk) => {
      if (error) throw error;
      length += chunk.length;
      zipLimit(length, maxZipBytes);
      chunks.push(chunk);
    });
    for (const [path, bytes] of [...files].sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)) {
      const file = new ZipDeflate(path, { level: 6 });
      file.mtime = new Date(1980, 0, 1);
      archive.add(file);
      if (!bytes.length) file.push(bytes, true);
      for (let cursor = 0; cursor < bytes.length; cursor += 65536) {
        file.push(bytes.subarray(cursor, cursor + 65536), cursor + 65536 >= bytes.length);
      }
    }
    archive.end();
    const result = new Uint8Array(length);
    let offset = 0;
    for (const chunk of chunks) { result.set(chunk, offset); offset += chunk.length; }
    // Enforce the same raw-path/container rules on output, including path collisions.
    inspectZip(result, limits);
    return result;
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    throw new DocumentError("ExportFailed", "The document package could not be written as ZIP.");
  }
}
