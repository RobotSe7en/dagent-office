import { createServer, request } from "node:http";
import type { ClientRequest, Server } from "node:http";
import { readFileSync } from "node:fs";
import { afterEach, expect, test } from "vitest";
import { writeDocumentZip } from "@dagent-office/office";
import { createDevelopmentConversionMiddleware, runConversionProcess } from "@dagent-office/node/dev";

const servers: Server[] = [];
async function server(timeoutMs = 200, allowedHost?: string, allowedPort?: number): Promise<{ port: number; started: Promise<void> }> {
  let started: () => void = () => {};
  const seen = new Promise<void>((resolve) => { started = resolve; });
  let middleware: ReturnType<typeof createDevelopmentConversionMiddleware>;
  const instance = createServer((req, res) => { started(); middleware(req, res, () => { res.statusCode = 404; res.end(); }); }); servers.push(instance);
  await new Promise<void>((resolve) => instance.listen(0, "127.0.0.1", resolve));
  const address = instance.address(); if (!address || typeof address === "string") throw new Error("Missing test address.");
  middleware = createDevelopmentConversionMiddleware({ timeoutMs, maxConcurrency: 1, ...(allowedHost ? { allowedOrigins: ["http://" + allowedHost + ":" + (allowedPort ?? address.port)] } : {}) });
  return { port: address.port, started: seen };
}
function send(port: number, headers: Record<string,string> = {}, body = "{bad}", end = true): { client: ClientRequest; response: Promise<number> } {
  let client: ClientRequest;
  const response = new Promise<number>((resolve, reject) => {
    client = request({ hostname: "127.0.0.1", port, path: "/__dagent_office/convert", method: "POST",
      headers: { Host: "localhost:" + port, Origin: "http://localhost:" + port, "Content-Type": "application/json", "X-Dagent-Office-Dev": "1", ...headers },
    }, (value) => { value.resume(); value.once("end", () => resolve(value.statusCode!)); });
    client.once("error", reject); client.write(body); if (end) client.end();
  });
  return { client: client!, response };
}
afterEach(async () => { for (const item of servers.splice(0)) { item.closeAllConnections(); await new Promise<void>((resolve) => item.close(() => resolve())); } });

test("the standalone development adapter rejects rebinding hosts, wrong scheme and cross-origin requests", async () => {
  const { port } = await server();
  expect(await send(port, { Host: "attacker.example:" + port, Origin: "http://attacker.example:" + port }).response).toBe(403);
  expect(await send(port, { Origin: "https://localhost:" + port }).response).toBe(403);
  expect(await send(port, { Origin: "http://localhost:1" }).response).toBe(403);
  expect(await send(port).response).toBe(400);
});
test("upload deadlines release capacity even when the request body never ends", async () => {
  const { port, started } = await server(), upload = send(port, {}, "{", false);
  try {
    await started;
    expect(await send(port).response).toBe(429);
    expect(await upload.response).toBe(504);
    expect(await send(port).response).toBe(400);
  } finally { upload.client.destroy(); }
});
test("worker processes receive the host byte quota and exit without launching a browser", async () => {
  const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path,
    new Uint8Array(readFileSync(new URL("../../../examples/report/" + path, import.meta.url))),
  ]));
  const result = await runConversionProcess({ protocolVersion: 1, taskId: "quota", format: "pdf", input: { format: "document-package-zip", bytesBase64: Buffer.from(writeDocumentZip(files)).toString("base64") } }, { maxInputBytes: 1 });
  expect(result).toMatchObject({ taskId: "quota", ok: false, error: { code: "ResourceLimitExceeded" } });
});
test("remote development hosts require an exact configured origin and the actual listening port", async () => {
  const { port } = await server(200, "office.test");
  expect(await send(port, { Host: "office.test:" + port, Origin: "http://office.test:" + port }).response).toBe(400);
  expect(await send(port, { Host: "office.test:" + port, Origin: "" }).response).toBe(403);
  expect(await send(port, { Host: "office.test:" + port, Origin: "http://attacker.test:" + port }).response).toBe(403);
  expect(await send(port, { Host: "attacker.test:" + port, Origin: "http://attacker.test:" + port }).response).toBe(403);
  const wrong = await server(200, "office.test", 1);
  expect(await send(wrong.port, { Host: "office.test:1", Origin: "http://office.test:1" }).response).toBe(403);
});
test.each(["*", "http://*.office.test", "file:///tmp/doc", "https://user:password@office.test", "http://office.test/private", "http://office.test?next=evil"])("development origins reject ambiguous authority: %s", (origin) => {
  expect(() => createDevelopmentConversionMiddleware({ allowedOrigins: [origin] })).toThrowError(expect.objectContaining({ code: "InvalidRequest" }));
});
