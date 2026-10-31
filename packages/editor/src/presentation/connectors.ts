import { evaluatePresentationPreset } from "@dagent-office/core";
import type { DocumentEditingSession, DocumentOperation, DocumentNode } from "@dagent-office/core";
export interface ConnectorHost {
    readonly session: DocumentEditingSession;
    node(id: string): Pick<DocumentNode, "parentNodeId" | "text"> | undefined;
    selected(): readonly string[];
    current(): boolean;
    perform(operations: DocumentOperation[]): unknown;
    attempt(action: () => unknown): void;
}
export interface ConnectorControls {
    refresh(): void;
    dispose(): void;
}
export function createConnectorControls(document: Document, parent: HTMLElement, host: ConnectorHost): ConnectorControls {
    const controls = new AbortController(), panel = document.createElement("fieldset"), legend = document.createElement("legend");
    legend.textContent = "连接线端点";
    panel.append(legend);
    parent.append(panel);
    const select = (label: string): HTMLSelectElement => { const wrap = document.createElement("label"), input = document.createElement("select"); input.setAttribute("aria-label", label); wrap.append(label, input); panel.append(wrap); return input; };
    const routing = select("连接线类型");
    for (const [value, label] of [["straight", "直线"], ["elbow", "折线"]]) {
        const option = document.createElement("option");
        option.value = value!;
        option.textContent = label!;
        routing.append(option);
    }
    const endpoints = (["start", "end"] as const).map(end => {
        const input = select(end === "start" ? "起点对象" : "终点对象"), wrap = document.createElement("label"), site = document.createElement("input");
        site.type = "number";
        site.min = "0";
        site.step = "1";
        site.setAttribute("aria-label", end === "start" ? "起点连接位置" : "终点连接位置");
        wrap.append(end === "start" ? "起点连接位置" : "终点连接位置", site);
        panel.append(wrap);
        const apply = (): unknown => host.perform([{ op: "setPresentationConnector", nodeId: host.selected()[0]!, changes: { [end]: input.value ? { nodeId: input.value, site: Number(site.value) } : null } }]);
        input.addEventListener("change", () => host.attempt(() => { site.value = "0"; return apply(); }), { signal: controls.signal });
        site.addEventListener("change", () => host.attempt(apply), { signal: controls.signal });
        return { end, input, site };
    });
    routing.addEventListener("change", () => host.attempt(() => host.perform([{ op: "setPresentationConnector", nodeId: host.selected()[0]!, changes: { routing: routing.value as "straight" | "elbow" } }])), { signal: controls.signal });
    const surface = (id: string): string | undefined => { const p = host.session.snapshot().metadata.presentation!, roots = new Set([...p.slides, ...p.masters, ...p.layouts, ...p.notesMasters, ...p.stories].map(root => root.rootNodeId)); for (let current: string | null = id; current !== null; current = host.node(current)?.parentNodeId ?? null)
        if (roots.has(current))
            return current; return undefined; };
    const siteCounts = new Map<string, number>();
    const sites = (id: string): number => {
        const object = host.session.snapshot().metadata.presentation!.objects.find(object => object.nodeId === id)!, preset = object.kind === "geometry" ? object.preset : "rect";
        if (!siteCounts.has(preset))
            siteCounts.set(preset, evaluatePresentationPreset(preset, 100, 100).connections.length);
        return siteCounts.get(preset)!;
    };
    return { refresh() {
            const ids = host.selected(), p = host.session.snapshot().metadata.presentation!, object = ids.length === 1 ? p.objects.find(object => object.nodeId === ids[0]) : undefined;
            panel.hidden = object?.kind !== "connector";
            panel.disabled = !host.current();
            if (object?.kind !== "connector")
                return;
            routing.value = object.routing;
            const owner = surface(object.nodeId), targets = p.objects.filter(target => target.nodeId !== object.nodeId && target.kind !== "connector" && surface(target.nodeId) === owner);
            for (const { end, input, site } of endpoints) {
                input.replaceChildren();
                const free = document.createElement("option");
                free.value = "";
                free.textContent = "自由端点";
                input.append(free);
                for (const [index, target] of targets.entries()) {
                    const option = document.createElement("option");
                    option.value = target.nodeId;
                    option.textContent = target.name ?? `${index + 1}. ${host.node(target.nodeId)!.text || "对象"}`;
                    option.disabled = sites(target.nodeId) === 0;
                    input.append(option);
                }
                const id = end === "start" ? object.startNodeId : object.endNodeId;
                input.value = id ?? "";
                site.value = String((end === "start" ? object.startSite : object.endSite) ?? 0);
                site.disabled = !id;
                if (id)
                    site.max = String(Math.max(0, sites(id) - 1));
            }
        }, dispose() { controls.abort(); panel.remove(); } };
}
