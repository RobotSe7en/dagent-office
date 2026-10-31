import { A, check, child, elements, fail, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { bool, Context, integer, points, value } from "./context.js";
import type { Html } from "./context.js";
import { fill, outline } from "./styles.js";
import { importText } from "./text.js";

export function importTable(context: Context, source: XmlNode, frame: Html, part: string): number[] {
  check(source, ["tblPr", "tblGrid", "tr"], A);
  const properties = child(source, "tblPr", A);
  if (properties) {
    check(properties, ["noFill", "solidFill", "effectLst", "tableStyleId", "extLst"], A);
    const style = child(properties, "tableStyleId", A);
    // Built-in table styles need their style definition mapped, not discarded.
    if (style && style.content.some(x => typeof x === "string" && x.trim() && !/^\{?00000000-0000-0000-0000-000000000000\}?$/.test(x.trim()))) unsupported(style, "This presentation table style is not supported.");
  }
  const columns = elements(child(source, "tblGrid", A), "gridCol", A);
  const grid = columns.map(column => points(column, "w"));
  if (!grid.length || grid.some(width => width <= 0)) fail(source, "A presentation table grid is invalid.");
  const rows = elements(source, "tr", A);
  if (!rows.length) fail(source, "A presentation table requires rows.");
  context.count(rows.length * grid.length);
  const occupied = Array.from({ length: rows.length }, () => Array<number>(grid.length).fill(0));
  const table = context.make("table");
  context.style(table, { "border-collapse": "collapse", "table-layout": "fixed", width: "100%", ...fill(properties, "--office-table-fill") });
  const tbody = context.make("tbody");
  let identity = 0;
  rows.forEach((sourceRow, rowIndex) => {
    check(sourceRow, ["tc", "extLst"], A);
    const height = points(sourceRow, "h");
    if (height <= 0) fail(sourceRow, "A table row height must be positive.");
    const row = context.make("tr"); context.style(row, { height: height + "pt" });
    const cells = elements(sourceRow, "tc", A);
    if (cells.length !== grid.length) fail(sourceRow, "Each DrawingML row must have one cell record per grid column.");
    cells.forEach((sourceCell, columnIndex) => {
      check(sourceCell, ["txBody", "tcPr", "extLst"], A);
      const continuation = bool(sourceCell, "hMerge") || bool(sourceCell, "vMerge");
      const rowSpan = integer(sourceCell, "rowSpan", 1), colSpan = integer(sourceCell, "gridSpan", 1);
      if (rowSpan < 1 || colSpan < 1 || rowIndex + rowSpan > rows.length || columnIndex + colSpan > grid.length) fail(sourceCell, "A cell merge exceeds its table.");
      if (continuation) {
        if (!occupied[rowIndex]![columnIndex]) fail(sourceCell, "A merge continuation has no origin.");
        const body = child(sourceCell, "txBody", A);
        if (body && elements(body, "p", A).some(p => elements(p, undefined, A).some(run => run.name === "br" || ["r", "fld"].includes(run.name) && text(child(run, "t", A)) !== ""))) unsupported(sourceCell, "Merged continuation cells cannot discard their text.");
        return;
      }
      const cellIdentity = ++identity;
      for (let y = rowIndex; y < rowIndex + rowSpan; y++) for (let x = columnIndex; x < columnIndex + colSpan; x++) {
        context.count();
        if (occupied[y]![x]) fail(sourceCell, "Presentation table merges overlap.");
        occupied[y]![x] = cellIdentity;
      }
      const cell = context.make("td", { ...(rowSpan > 1 ? { rowspan: String(rowSpan) } : {}), ...(colSpan > 1 ? { colspan: String(colSpan) } : {}) });
      const p = child(sourceCell, "tcPr", A);
      const css: Css = { ...fill(p, "--office-cell-fill") };
      if (p) {
        check(p, ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill", "gradFill", "headers", "extLst"], A);
        if (child(p, "lnTlToBr", A) || child(p, "lnBlToTr", A) || child(p, "cell3D", A)) unsupported(p, "Diagonal cell borders and 3D cells are not supported.");
        for (const [tag, side] of [["lnL", "left"], ["lnR", "right"], ["lnT", "top"], ["lnB", "bottom"]]) {
          const line = child(p, tag!, A), declarations = outline(line, "--office-cell-border-" + side);
          Object.assign(css, declarations);
          if (line) {
            const width = points(line, "w", 12700), color = declarations["--office-cell-border-" + side + "-color"] ?? "transparent";
            css["border-" + side] = width + "pt " + (declarations["--office-cell-border-" + side + "-dash"]?.includes("dot") ? "dotted" : declarations["--office-cell-border-" + side + "-dash"]?.includes("ash") ? "dashed" : "solid") + " " + color;
          }
        }
        for (const [name, side] of [["marL", "left"], ["marR", "right"], ["marT", "top"], ["marB", "bottom"]]) {
          const amount = points(p, name!, side === "left" || side === "right" ? 91440 : 45720);
          if (amount < 0) fail(p, "Cell margins cannot be negative.");
          css["padding-" + side] = amount + "pt";
        }
        const anchor = value(p, "anchor") ?? "t";
        if (!["t", "ctr", "b"].includes(anchor)) unsupported(p, "This cell alignment is not supported.");
        css["vertical-align"] = anchor === "t" ? "top" : anchor === "b" ? "bottom" : "middle";
        if (value(p, "vert") && value(p, "vert") !== "horz") unsupported(p, "Vertical table text is not supported.");
      }
      const textBody = child(sourceCell, "txBody", A);
      if (!textBody) fail(sourceCell, "A table cell text body is missing.");
      const body = context.make("div"); importText(context, textBody, body, part);
      cell.children.push(body); context.style(cell, css); row.children.push(cell);
    });
    tbody.children.push(row);
  });
  if (occupied.some(row => row.some(v => !v))) fail(source, "A presentation table grid is not fully covered.");
  table.children.push(tbody); frame.children.push(table);
  return grid;
}
