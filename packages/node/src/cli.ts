#!/usr/bin/env node
import { createDocumentConverter } from "./render.js";
import type { ConverterOptions, DocumentConverter } from "./render.js";
import { serveConversionProtocol } from "./protocol.js";
import type { ProtocolLimits } from "./protocol.js";

async function run(): Promise<void> {
  let converter: DocumentConverter | undefined, stopping = false;
  const stop = (): void => { stopping = true; process.stdin.destroy(); void converter?.close(); };
  try {
    const raw = process.env.DAGENT_OFFICE_CONVERSION_CONFIG;
    const config = raw ? JSON.parse(raw) as { converter?: ConverterOptions; protocol?: ProtocolLimits } : {};
    converter = createDocumentConverter({ ...config.converter,
      ...(process.env.DAGENT_OFFICE_CHROMIUM_EXECUTABLE ? { browserExecutable: process.env.DAGENT_OFFICE_CHROMIUM_EXECUTABLE } : {}),
      ...(process.env.DAGENT_OFFICE_CONVERSION_TIMEOUT ? { timeoutMs: Number(process.env.DAGENT_OFFICE_CONVERSION_TIMEOUT) } : {}),
    });
    process.once("SIGTERM", stop); process.once("SIGINT", stop);
    await serveConversionProtocol(process.stdin, async (line) => {
      await new Promise<void>((resolve, reject) => process.stdout.write(line, (error) => error ? reject(error) : resolve()));
    }, converter, config.protocol);
  } catch {
    if (!stopping) { process.stderr.write("Conversion process failed or exceeded its transport budget.\n"); process.exitCode = 1; }
  } finally { process.off("SIGTERM", stop); process.off("SIGINT", stop); await converter?.close(); }
}
if (process.argv.slice(2).includes("--help")) process.stdout.write("dagent-office: one versioned conversion JSON request per stdin line; results on stdout. Input and output contain bytes, never file paths.\n");
else await run();
