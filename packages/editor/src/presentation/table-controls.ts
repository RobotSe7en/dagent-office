import type { DocumentEditingSession, DocumentOperation } from "@dagent-office/core";
export interface TableControlsHost {
    readonly session: DocumentEditingSession;
    target(): string | undefined;
    perform(operations: DocumentOperation[]): void;
    attempt(action: () => void): void;
}
export interface TableControls {
    refresh(): void;
    dispose(): void;
}
export function createTableControls(document: Document, parent: HTMLElement, host: TableControlsHost): TableControls {
    const controls = new AbortController(), pane = document.createElement("fieldset"), legend = document.createElement("legend");
    legend.textContent = "表格结构";
    pane.append(legend);
    parent.append(pane);
    const field = (label: string, initial = 1): HTMLInputElement => { const wrap = document.createElement("label"), input = document.createElement("input"); input.type = "number"; input.min = "1"; input.step = "1"; input.value = String(initial); input.setAttribute("aria-label", label); wrap.append(label, input); pane.append(wrap); return input; };
    const row = field("表格行"), column = field("表格列"), rows = field("选区行数"), columns = field("选区列数"), width = field("列宽（pt）", 120), height = field("行高（pt）", 40);
    const index = (input: HTMLInputElement): number => {
        const n = Number(input.value);
        if (!Number.isInteger(n) || n < 1)
            throw new Error("请输入从 1 开始的行列位置或数量。");
        return n - 1;
    };
    const button = (label: string, change: () => Extract<DocumentOperation, {
        op: "editPresentationTable";
    }>["change"]): void => {
        const node = document.createElement("button");
        node.type = "button";
        node.textContent = label;
        pane.append(node);
        node.addEventListener("click", () => host.attempt(() => {
            const nodeId = host.target();
            if (nodeId)
                host.perform([{ op: "editPresentationTable", nodeId, change: change() }]);
        }), { signal: controls.signal });
    };
    for (const [label, kind] of [["前插行", "insertRows"], ["删除行", "deleteRows"], ["前插列", "insertColumns"], ["删除列", "deleteColumns"]] as const)
        button(label, () => ({ kind, at: index(kind.endsWith("Rows") ? row : column), count: index(kind.endsWith("Rows") ? rows : columns) + 1 }));
    button("合并单元格", () => ({ kind: "mergeCells", row: index(row), column: index(column), rows: index(rows) + 1, columns: index(columns) + 1 }));
    button("拆分单元格", () => ({ kind: "splitCell", row: index(row), column: index(column) }));
    button("设置列宽", () => {
        const object = host.session.snapshot().metadata.presentation!.objects.find(o => o.nodeId === host.target())!;
        if (object.kind !== "table")
            throw new Error("请选择表格。");
        const grid = [...object.grid], c = index(column);
        if (c >= grid.length)
            throw new Error("列位置超出表格。");
        grid[c] = Number(width.value);
        return { kind: "resize", grid };
    });
    button("设置行高", () => {
        const s = host.session, node = s.getNode(host.target()!)!, table = s.getNode(node.childNodeIds[0]!)!, ids = table.childNodeIds.flatMap(id => { const row = s.getNode(id)!; return row.tagName === "tr" ? [id] : row.childNodeIds; }), r = index(row);
        if (r >= ids.length)
            throw new Error("行位置超出表格。");
        const rowHeights = ids.map(id => { const style = s.getNode(id)!.attributes.style ?? ""; return Number(/(?:^|;)height:([\d.]+)pt/.exec(style)?.[1] ?? Number(node.attributes["data-office-height"]) / ids.length); });
        rowHeights[r] = Number(height.value);
        return { kind: "resize", rowHeights };
    });
    const valueField = (label: string, type: string, value: string): HTMLInputElement => { const wrap = document.createElement("label"), input = document.createElement("input"); input.type = type; input.value = value; input.setAttribute("aria-label", label); wrap.append(label, input); pane.append(wrap); return input; };
    const choice = (label: string, values: readonly (readonly [
        string,
        string
    ])[]): HTMLSelectElement => { const wrap = document.createElement("label"), input = document.createElement("select"); input.setAttribute("aria-label", label); for (const [value, text] of values) {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = text;
        input.append(option);
    } wrap.append(label, input); pane.append(wrap); return input; };
    const fillColor = valueField("单元格填充颜色", "color", "#ffffff"), borderColor = valueField("单元格边框颜色", "color", "#44546a"), borderWidth = valueField("单元格边框宽度（pt）", "number", "1"), borderStyle = choice("单元格边框线型", [["solid", "实线"], ["dashed", "虚线"], ["dotted", "点线"], ["none", "无边框"]]), horizontal = choice("单元格水平对齐", [["left", "左对齐"], ["center", "居中"], ["right", "右对齐"]]), vertical = choice("单元格垂直对齐", [["top", "顶部"], ["middle", "中部"], ["bottom", "底部"]]);
    borderWidth.min = "0";
    borderWidth.step = "any";
    const selectedCells = (): string[] => { const s = host.session, id = host.target(), object = s.snapshot().metadata.presentation!.objects.find(o => o.nodeId === id); if (object?.kind !== "table")
        throw new Error("请选择表格。"); const table = s.getNode(s.getNode(id!)!.childNodeIds[0]!)!, rowIds = table.childNodeIds.flatMap(id => { const n = s.getNode(id)!; return n.tagName === "tr" ? [id] : n.childNodeIds; }), occupied: string[][] = Array.from({ length: rowIds.length }, () => []); for (const [r, id] of rowIds.entries())
        for (const cellId of s.getNode(id)!.childNodeIds) {
            const a = s.getNode(cellId)!.attributes;
            let c = 0;
            while (occupied[r]![c])
                c++;
            const height = Number(a.rowspan ?? 1) || rowIds.length - r, width = Number(a.colspan ?? 1);
            for (let y = r; y < r + height; y++)
                for (let x = c; x < c + width; x++)
                    occupied[y]![x] = cellId;
        } const r = index(row), c = index(column), h = index(rows) + 1, w = index(columns) + 1; if (r + h > rowIds.length || c + w > object.grid.length)
        throw new Error("选区超出表格网格。"); return [...new Set(occupied.slice(r, r + h).flatMap(line => line.slice(c, c + w)))]; };
    const formatButton = (label: string, properties: () => Readonly<Record<string, string>>, paragraphs = false): void => { const button = document.createElement("button"); button.type = "button"; button.textContent = label; pane.append(button); button.addEventListener("click", () => host.attempt(() => { const ids = selectedCells(), targets = new Set(ids); if (paragraphs) {
        const visit = (id: string): void => { const node = host.session.getNode(id)!; if (/^(?:p|h[1-6])$/.test(node.tagName))
            targets.add(id); for (const child of node.childNodeIds)
            visit(child); };
        for (const id of ids)
            visit(id);
    } const values = properties(); host.perform([...targets].flatMap(nodeId => Object.entries(values).map(([property, value]) => ({ op: "setStyle" as const, nodeId, property, value })))); }), { signal: controls.signal }); };
    const directColor = (prefix: string, color: string): Record<string, string> => Object.fromEntries([[prefix + "-color", color], [prefix + "-source", "direct"], ...["alpha", "lum-mod", "shade"].map(key => [prefix + "-" + key, "1"]), ...["lum-off", "tint"].map(key => [prefix + "-" + key, "0"])]);
    formatButton("设置单元格填充", () => ({ "--office-cell-fill-kind": "solid", ...directColor("--office-cell-fill", fillColor.value), "background-color": fillColor.value }));
    formatButton("清除单元格填充", () => ({ "--office-cell-fill-kind": "none", "background-color": "transparent" }));
    formatButton("设置单元格边框", () => { const width = Number(borderWidth.value); if (!borderWidth.value.trim() || !Number.isFinite(width) || width < 0)
        throw new Error("请输入非负边框宽度。"); return Object.assign({}, ...["top", "right", "bottom", "left"].map(side => ({ ["border-" + side]: width + "pt " + borderStyle.value + " " + borderColor.value, ["--office-cell-border-" + side + "-kind"]: borderStyle.value === "none" ? "none" : "solid", ["--office-cell-border-" + side + "-width"]: width + "pt", ["--office-cell-border-" + side + "-dash"]: borderStyle.value === "dotted" ? "dot" : borderStyle.value === "dashed" ? "dash" : "solid", ...directColor("--office-cell-border-" + side, borderColor.value) }))); });
    formatButton("设置单元格水平对齐", () => ({ "text-align": horizontal.value }), true);
    formatButton("设置单元格垂直对齐", () => ({ "vertical-align": vertical.value }));
    return { refresh() { pane.hidden = !host.target(); }, dispose() { controls.abort(); pane.remove(); } };
}
