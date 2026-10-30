export function downloadBytes(bytes: Uint8Array, name: string, mimeType: string): void {
  const copy = new Uint8Array(bytes), url = URL.createObjectURL(new Blob([copy.buffer], { type: mimeType })), anchor = document.createElement("a");
  anchor.href = url; anchor.download = name; anchor.click(); setTimeout(() => URL.revokeObjectURL(url), 0);
}
