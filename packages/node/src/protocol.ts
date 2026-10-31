import { DocumentError, openDocument } from "@dagent-office/core";
import type { ValidationIssue } from "@dagent-office/core";
import { importDocx, importPptx, readDocumentZip } from "@dagent-office/office";
import type { DocumentRenderReport, RenderLimits } from "@dagent-office/render/contracts";
import { ConversionError, DEFAULT_CONVERTER_LIMITS } from "./render.js";
import type { ConversionErrorCode, DocumentConverter, PageRenderOptions, RenderedOutput } from "./render.js";

export interface ConversionRequestV1 {
  readonly protocolVersion: 1;
  readonly taskId: string;
  readonly format: "pdf" | "png";
  readonly input: { readonly format: "document-package-zip" | "docx" | "pptx"; readonly bytesBase64: string; readonly documentId?: string };
  readonly options?: Omit<PageRenderOptions, "signal">;
}
export type ConversionResponseV1 = {
  readonly protocolVersion: 1; readonly taskId: string; readonly ok: true;
  readonly output: { readonly mimeType: RenderedOutput["mimeType"]; readonly bytesBase64: string; readonly report: DocumentRenderReport; readonly engine: RenderedOutput["engine"] };
} | {
  readonly protocolVersion: 1; readonly taskId: string | null; readonly ok: false;
  readonly error: { readonly code: ConversionErrorCode; readonly message: string; readonly details?: readonly ValidationIssue[] };
};
export interface ProtocolLimits { readonly maxInputBytes?: number; readonly maxMessageBytes?: number; readonly maxOutputMessageBytes?: number }
function object(value: unknown, keys: readonly string[]): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || Object.keys(value).some((key) => !keys.includes(key))) throw new ConversionError("InvalidRequest", "Conversion object has an invalid shape or unknown fields.");
  return value as Record<string, unknown>;
}
function identifier(value: unknown): value is string { return typeof value === "string" && value.length > 0 && value.length <= 256; }
function integer(value: unknown, minimum: number): value is number { return typeof value === "number" && Number.isSafeInteger(value) && value >= minimum; }
export function parseConversionRequest(input: unknown): ConversionRequestV1 {
  const request = object(input, ["protocolVersion", "taskId", "format", "input", "options"]);
  if (request.protocolVersion !== 1) throw new ConversionError("UnsupportedVersion", "Unknown conversion protocol version.");
  if (!identifier(request.taskId)) throw new ConversionError("InvalidRequest", "A bounded task identifier is required.");
  if (!["pdf", "png"].includes(request.format as string)) throw new ConversionError("UnsupportedFormat", "Unknown conversion output format.");
  const source = object(request.input, ["format", "bytesBase64", "documentId"]);
  if (!["document-package-zip", "docx", "pptx"].includes(source.format as string)) throw new ConversionError("UnsupportedFormat", "Unknown conversion input format.");
  if (typeof source.bytesBase64 !== "string" || !source.bytesBase64.length) throw new ConversionError("InvalidRequest", "Authorized input bytes are required.");
  if (source.documentId !== undefined && !identifier(source.documentId) || ["docx", "pptx"].includes(source.format as string) && !identifier(source.documentId)) throw new ConversionError("InvalidRequest", "Native Office conversion requires a bounded document identifier.");
  if (request.options !== undefined) {
    const options = object(request.options, ["pageIndex", "scale", "timeoutMs", "limits"]);
    if (options.pageIndex !== undefined && !integer(options.pageIndex, 0) || options.timeoutMs !== undefined && !integer(options.timeoutMs, 1)
      || options.scale !== undefined && (typeof options.scale !== "number" || !Number.isFinite(options.scale) || options.scale < 1 || options.scale > 4)) throw new ConversionError("InvalidRequest", "Conversion output options are invalid.");
    if (request.format === "pdf" && (options.pageIndex !== undefined || options.scale !== undefined)) throw new ConversionError("InvalidRequest", "PDF output uses all measured pages at document scale.");
    if (options.limits !== undefined) {
      const limits = object(options.limits, ["maxNodes", "maxPages", "maxLayoutPasses", "maxPrintNodes"] satisfies (keyof RenderLimits)[]);
      if (Object.values(limits).some((value) => !integer(value, 1))) throw new ConversionError("InvalidRequest", "Conversion render limits are invalid.");
    }
  }
  return request as unknown as ConversionRequestV1;
}
function failure(reason: unknown, taskId: string | null): ConversionResponseV1 {
  const error = reason instanceof ConversionError || reason instanceof DocumentError ? reason.toJSON() : { code: "ExportFailed" as const, message: "Conversion failed." };
  return { protocolVersion: 1, taskId, ok: false, error };
}
export async function executeConversionRequest(converter: DocumentConverter, input: unknown, limits: ProtocolLimits = {}): Promise<ConversionResponseV1> {
  let taskId: string | null = input && typeof input === "object" && Object.hasOwn(input, "taskId") && identifier((input as Record<string,unknown>).taskId) ? (input as { taskId: string }).taskId : null;
  try {
    const request = parseConversionRequest(input); taskId = request.taskId;
    const maximum = limits.maxInputBytes ?? DEFAULT_CONVERTER_LIMITS.maxInputBytes, encoded = request.input.bytesBase64;
    if (!integer(maximum, 1)) throw new ConversionError("InvalidRequest", "Protocol input limit must be a positive safe integer.");
    if (encoded.length > Math.ceil(maximum / 3) * 4) throw new ConversionError("ResourceLimitExceeded", "Conversion input exceeds its byte budget.");
    if (encoded.length % 4 || /[^A-Za-z0-9+/=]/.test(encoded)) throw new ConversionError("InvalidRequest", "Input bytes must use canonical base64.");
    const bytes = Buffer.from(encoded, "base64");
    if (bytes.length > maximum) throw new ConversionError("ResourceLimitExceeded", "Conversion input exceeds its byte budget.");
    if (bytes.toString("base64") !== encoded) throw new ConversionError("InvalidRequest", "Input bytes must use canonical base64.");
    const model = openDocument(request.input.format === "docx" ? importDocx(bytes, { documentId: request.input.documentId! }) : request.input.format === "pptx" ? importPptx(bytes, { documentId: request.input.documentId! }) : readDocumentZip(bytes));
    if (request.input.documentId !== undefined && model.snapshot().manifest.documentId !== request.input.documentId) throw new ConversionError("InvalidDocument", "Package identity differs from the asserted document identifier.");
    const output = request.format === "pdf" ? await converter.renderPdf(model, request.options) : await converter.renderPage(model, request.options);
    return { protocolVersion: 1, taskId, ok: true, output: { bytesBase64: Buffer.from(output.bytes).toString("base64"), mimeType: output.mimeType, report: output.report, engine: output.engine } };
  } catch (reason) { return failure(reason, taskId); }
}

/** NDJSON byte transport with backpressure and bounded frames; no file paths or URLs. */
export async function serveConversionProtocol(input: AsyncIterable<Uint8Array>, write: (line: string) => Promise<void>, converter: DocumentConverter, limits: ProtocolLimits = {}): Promise<void> {
  const maximum = limits.maxMessageBytes ?? 96 * 1024 * 1024, outputMaximum = limits.maxOutputMessageBytes ?? 96 * 1024 * 1024;
  if (!integer(maximum, 1) || !integer(outputMaximum, 1)) throw new ConversionError("InvalidRequest", "Protocol frame limits must be positive safe integers.");
  const decoder = new TextDecoder("utf-8", { fatal: true }); let pending = "", pendingBytes = 0;
  async function emit(response: ConversionResponseV1): Promise<void> {
    let serialized = JSON.stringify(response);
    if (Buffer.byteLength(serialized) > outputMaximum) {
      response = failure(new ConversionError("ResourceLimitExceeded", "Conversion response exceeds its frame budget."), response.taskId); serialized = JSON.stringify(response);
      if (Buffer.byteLength(serialized) > outputMaximum) throw new ConversionError("ResourceLimitExceeded", "Response budget is too small for a bounded error frame.");
    }
    await write(serialized + "\n");
  }
  async function respond(line: string): Promise<void> {
    let value: unknown;
    try { value = JSON.parse(line); } catch { await emit(failure(new ConversionError("InvalidRequest", "Invalid conversion JSON."), null)); return; }
    await emit(await executeConversionRequest(converter, value, limits));
  }
  for await (const chunk of input) {
    let offset = 0;
    for (let index = 0; index < chunk.length; index++) {
      if (chunk[index] !== 10) continue;
      pendingBytes += index - offset;
      if (pendingBytes > maximum) throw new ConversionError("ResourceLimitExceeded", "Conversion request exceeds its frame budget.");
      pending += decoder.decode(chunk.subarray(offset, index), { stream: true });
      pending += decoder.decode(); await respond(pending); pending = ""; pendingBytes = 0; offset = index + 1;
    }
    pendingBytes += chunk.length - offset;
    if (pendingBytes > maximum) throw new ConversionError("ResourceLimitExceeded", "Conversion request exceeds its frame budget.");
    pending += decoder.decode(chunk.subarray(offset), { stream: true });
  }
  pending += decoder.decode(); if (pending.length) await respond(pending);
}
