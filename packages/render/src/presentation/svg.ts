import type { PresentationPresetGeometry } from "@dagent-office/core";

const SVG = "http://www.w3.org/2000/svg";
export function svgElement<K extends keyof SVGElementTagNameMap>(document: Document, tag: K, attributes: Readonly<Record<string, string | number>> = {}): SVGElementTagNameMap[K] {
  const node = document.createElementNS(SVG, tag);
  for (const [key, value] of Object.entries(attributes)) node.setAttribute(key, String(value));
  return node;
}
const clean = (n: number): string => String(Math.abs(n) < 1e-10 ? 0 : Number(n.toFixed(6)));
export function svgPath(path: PresentationPresetGeometry["paths"][number]): string {
  let x = 0, y = 0, startX = 0, startY = 0;
  const commands: string[] = [];
  for (const command of path.commands) {
    if (command.kind === "move" || command.kind === "line") {
      x = command.x; y = command.y;
      if (command.kind === "move") { startX = x; startY = y; }
      commands.push((command.kind === "move" ? "M" : "L") + clean(x) + " " + clean(y));
    } else if (command.kind === "close") { commands.push("Z"); x = startX; y = startY; }
    else if (command.kind === "arc") {
      const angle = command.startAngle * Math.PI / 180, rx = command.radiusX, ry = command.radiusY;
      const cx = x - rx * Math.cos(angle), cy = y - ry * Math.sin(angle), pieces = Math.max(1, Math.ceil(Math.abs(command.sweepAngle) / 180));
      const step = command.sweepAngle / pieces;
      for (let piece = 1; piece <= pieces; piece++) {
        const theta = (command.startAngle + step * piece) * Math.PI / 180;
        x = cx + rx * Math.cos(theta); y = cy + ry * Math.sin(theta);
        commands.push(rx && ry ? "A" + clean(rx) + " " + clean(ry) + " 0 0 " + (step >= 0 ? 1 : 0) + " " + clean(x) + " " + clean(y) : "L" + clean(x) + " " + clean(y));
      }
    }
  }
  return commands.join(" ");
}
