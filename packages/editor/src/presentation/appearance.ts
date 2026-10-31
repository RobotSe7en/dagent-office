import type { DocumentEditingSession, DocumentOperation } from "@dagent-office/core";
export interface AppearanceHost {
    readonly session: DocumentEditingSession;
    selected(): readonly string[];
    element(nodeId: string): HTMLElement | undefined;
    current(): boolean;
    perform(operations: DocumentOperation[]): unknown;
    attempt(action: () => unknown): void;
}
export interface AppearanceControls {
    refresh(): void;
    dispose(): void;
}
function hex(document: Document, value: string, fallback: string): string {
    try {
        const canvas = document.createElement("canvas");
        canvas.width = canvas.height = 1;
        const context = canvas.getContext("2d")!;
        context.fillStyle = fallback;
        context.fillStyle = value;
        context.fillRect(0, 0, 1, 1);
        return "#" + [...context.getImageData(0, 0, 1, 1).data].slice(0, 3).map(value => value.toString(16).padStart(2, "0")).join("");
    }
    catch {
        return fallback;
    }
}
/** Direct appearance declarations use the same CSS source and transaction boundary. */
export function createAppearanceControls(document: Document, parent: HTMLElement, host: AppearanceHost): AppearanceControls {
    const controls = new AbortController(), panel = document.createElement("fieldset"), legend = document.createElement("legend");
    legend.textContent = "对象外观";
    panel.append(legend);
    parent.append(panel);
    const field = (label: string, type: string, value: string): HTMLInputElement => { const labelNode = document.createElement("label"), input = document.createElement("input"); input.type = type; input.value = value; input.setAttribute("aria-label", label); labelNode.append(label, input); panel.append(labelNode); return input; };
    const select = (label: string, values: readonly (readonly [
        string,
        string
    ])[]): HTMLSelectElement => {
        const labelNode = document.createElement("label"), input = document.createElement("select");
        input.setAttribute("aria-label", label);
        for (const [value, text] of values) {
            const option = document.createElement("option");
            option.value = value;
            option.textContent = text;
            input.append(option);
        }
        labelNode.append(label, input);
        panel.append(labelNode);
        return input;
    };
    const kind = select("填充", [["none", "无填充"], ["solid", "实色"], ["linear", "线性渐变"]]), first = field("填充颜色", "color", "#4472c4"), second = field("渐变结束颜色", "color", "#ffffff"), angle = field("渐变角度", "number", "0"), opacity = field("不透明度（%）", "number", "100"), line = field("显示轮廓", "checkbox", ""), lineColor = field("轮廓颜色", "color", "#44546a"), width = field("轮廓宽度（pt）", "number", "1"), dash = select("轮廓线型", [["solid", "实线"], ["dash", "虚线"], ["dot", "点线"], ["dashDot", "点划线"], ["lgDash", "长虚线"]]), shadow = field("基础外阴影", "checkbox", "");
    angle.step = "any";
    opacity.min = "0";
    opacity.max = "100";
    opacity.step = "any";
    width.min = "0";
    width.step = "any";
    const apply = (declarations: readonly (readonly [
        string,
        string
    ])[]): unknown => host.perform(host.selected().flatMap(nodeId => declarations.map(([property, value]) => ({ op: "setStyle" as const, nodeId, property, value }))));
    const styleFor = (id: string): CSSStyleDeclaration => { const node = host.element(id)!; return node.ownerDocument.defaultView!.getComputedStyle(node); };
    const editEach = (declarations: (style: CSSStyleDeclaration, element: HTMLElement) => readonly (readonly [
        string,
        string
    ])[]): unknown => host.perform(host.selected().flatMap(nodeId => declarations(styleFor(nodeId), host.element(nodeId)!).map(([property, value]) => ({ op: "setStyle" as const, nodeId, property, value }))));
    const property = (style: CSSStyleDeclaration, name: string, fallback: string): string => style.getPropertyValue(name).trim() || fallback;
    const paintAlpha = (node: HTMLElement, prefix: string, style: CSSStyleDeclaration): string => {
        const svg = node.querySelector(":scope>svg"), stop = /^--office-fill-stop-(\d+)$/.exec(prefix), attribute = prefix === "--office-line" ? "stroke" : "fill";
        const path = [...svg?.querySelectorAll(":scope>path") ?? []].find(path => path.getAttribute(attribute) !== "none"), raw = stop ? svg?.querySelectorAll("linearGradient stop")[Number(stop[1])]?.getAttribute("stop-color") : path?.getAttribute(attribute);
        const alpha = raw?.match(/^rgba\([^,]+,[^,]+,[^,]+,\s*([\d.eE+-]+)\s*\)$/)?.[1];
        return alpha ?? property(style, prefix + "-alpha", "1");
    };
    // Core color writes reset native transforms; retain the visible opacity explicitly.
    const directPaint = (prefix: string, color: string, alpha: string): [
        string,
        string
    ][] => [[prefix + "-color", color], [prefix + "-alpha", alpha]];
    const fill = (): unknown => {
        const alpha = String(Number(opacity.value) / 100), declarations: [
            string,
            string
        ][] = [["--office-fill-kind", kind.value]];
        if (kind.value === "solid")
            declarations.push(...directPaint("--office-fill", first.value, alpha));
        if (kind.value === "linear")
            declarations.push(["--office-fill-stop-count", "2"], ["--office-fill-angle", angle.value], ["--office-fill-stop-0-position", "0"], ["--office-fill-stop-1-position", "1"], ...directPaint("--office-fill-stop-0", first.value, alpha), ...directPaint("--office-fill-stop-1", second.value, alpha));
        if (!Number.isFinite(Number(alpha)) || Number(alpha) < 0 || Number(alpha) > 1 || !Number.isFinite(Number(angle.value)))
            throw new Error("请输入有效的不透明度和渐变角度。");
        return apply(declarations);
    };
    kind.addEventListener("change", () => host.attempt(fill), { signal: controls.signal });
    first.addEventListener("change", () => host.attempt(() => editEach((style, node) => { const prefix = property(style, "--office-fill-kind", kind.value) === "linear" ? "--office-fill-stop-0" : "--office-fill"; return directPaint(prefix, first.value, paintAlpha(node, prefix, style)); })), { signal: controls.signal });
    second.addEventListener("change", () => host.attempt(() => editEach((style, node) => { const prefix = "--office-fill-stop-" + (Number(property(style, "--office-fill-stop-count", "2")) - 1); return directPaint(prefix, second.value, paintAlpha(node, prefix, style)); })), { signal: controls.signal });
    angle.addEventListener("change", () => host.attempt(() => {
        if (!angle.value.trim() || !Number.isFinite(Number(angle.value)))
            throw new Error("请输入有效渐变角度。");
        return apply([["--office-fill-angle", angle.value]]);
    }), { signal: controls.signal });
    opacity.addEventListener("change", () => host.attempt(() => {
        const alpha = Number(opacity.value) / 100;
        if (!opacity.value.trim() || !Number.isFinite(alpha) || alpha < 0 || alpha > 1)
            throw new Error("请输入 0 至 100 的不透明度。");
        return editEach(style => { const prefixes = property(style, "--office-fill-kind", kind.value) === "linear" ? Array.from({ length: Math.min(1024, Number(property(style, "--office-fill-stop-count", "2"))) }, (_, i) => "--office-fill-stop-" + i) : ["--office-fill"], declarations: [
            string,
            string
        ][] = prefixes.map(prefix => [prefix + "-alpha", String(alpha)]); for (const prefix of prefixes)
            if (property(style, prefix + "-source", "") === "reference") {
                const reference = property(style, prefix + "-reference-prefix", "");
                if (reference === "--office-fill-reference" && !declarations.some(([name]) => name === reference + "-alpha"))
                    declarations.push([reference + "-alpha", "1"]);
            } return declarations; });
    }), { signal: controls.signal });
    line.addEventListener("change", () => host.attempt(() => apply([["--office-line-kind", line.checked ? "solid" : "none"]])), { signal: controls.signal });
    lineColor.addEventListener("change", () => host.attempt(() => editEach((style, node) => [...directPaint("--office-line", lineColor.value, paintAlpha(node, "--office-line", style)), ["--office-line-kind", "solid"]])), { signal: controls.signal });
    width.addEventListener("change", () => host.attempt(() => {
        if (!width.value.trim() || !Number.isFinite(Number(width.value)) || Number(width.value) < 0)
            throw new Error("请输入非负的轮廓宽度。");
        return apply([["--office-line-width", width.value + "pt"]]);
    }), { signal: controls.signal });
    dash.addEventListener("change", () => host.attempt(() => apply([["--office-line-dash", dash.value]])), { signal: controls.signal });
    shadow.addEventListener("change", () => host.attempt(() => editEach(style => shadow.checked ? (property(style, "--office-shadow-distance", "") ? [["--office-shadow-kind", "outer"]] : [["--office-shadow-kind", "outer"], ["--office-shadow-color", "#000000"], ["--office-shadow-alpha", "0.3"], ["--office-shadow-blur", "4pt"], ["--office-shadow-distance", "3pt"], ["--office-shadow-direction", "45"]]) : [["--office-shadow-kind", "none"]])), { signal: controls.signal });
    return { refresh() {
            const ids = host.selected(), objects = host.session.snapshot().metadata.presentation!.objects, object = ids.length === 1 ? objects.find(object => object.nodeId === ids[0]) : undefined;
            panel.disabled = !host.current() || !ids.length || ids.some(id => !["text", "geometry", "connector"].includes(objects.find(object => object.nodeId === id)?.kind ?? ""));
            const element = object ? host.element(object.nodeId) : undefined;
            if (!element || !host.current())
                return;
            const style = element.ownerDocument.defaultView!.getComputedStyle(element), value = (name: string, fallback: string): string => style.getPropertyValue(name).trim() || fallback;
            kind.value = value("--office-fill-kind", object!.kind === "geometry" ? "solid" : "none");
            const svg = element.querySelector(":scope>svg"), path = svg?.querySelector(":scope>path"), stops = svg?.querySelectorAll("linearGradient stop"), last = Number(value("--office-fill-stop-count", "2")) - 1;
            first.value = hex(document, (kind.value === "linear" ? stops?.[0]?.getAttribute("stop-color") : path?.getAttribute("fill")) ?? value(kind.value === "linear" ? "--office-fill-stop-0-color" : "--office-fill-color", "#4472c4"), "#4472c4");
            second.value = hex(document, stops?.[last]?.getAttribute("stop-color") ?? value("--office-fill-stop-" + last + "-color", "#ffffff"), "#ffffff");
            angle.value = value("--office-fill-angle", "0");
            opacity.value = String(Math.round(Number(paintAlpha(element, kind.value === "linear" ? "--office-fill-stop-0" : "--office-fill", style)) * 1e10) / 1e8);
            line.checked = value("--office-line-kind", object!.kind === "connector" ? "solid" : "none") !== "none";
            lineColor.value = hex(document, path?.getAttribute("stroke") ?? value("--office-line-color", "#44546a"), "#44546a");
            width.value = String(parseFloat(value("--office-line-width", "1")));
            dash.value = value("--office-line-dash", "solid");
            shadow.checked = value("--office-shadow-kind", "none") === "outer";
            second.disabled = angle.disabled = kind.value !== "linear";
            first.disabled = opacity.disabled = kind.value === "none";
        }, dispose() { controls.abort(); panel.remove(); } };
}
