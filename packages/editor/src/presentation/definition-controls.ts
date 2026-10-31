import type { CandidateChange, DocumentEditingSession, DocumentOperation } from "@dagent-office/core";
export interface DefinitionControlsHost {
    readonly session: DocumentEditingSession;
    activeSlide(): string;
    surface(): string | undefined;
    object(): string | undefined;
    selectSurface(id?: string): void;
    attempt(action: () => void): void;
}
export interface DefinitionControls {
    refresh(): void;
    dispose(): void;
}
/** Global writes stay candidates until their affected pages have been shown. */
export function createDefinitionControls(document: Document, parent: HTMLElement, host: DefinitionControlsHost): DefinitionControls {
    const controls = new AbortController(), pane = document.createElement("fieldset"), legend = document.createElement("legend"), impact = document.createElement("p");
    legend.textContent = "主题、母版与布局";
    pane.append(legend);
    parent.append(pane);
    let candidate: CandidateChange | undefined;
    const button = (label: string, action: () => void): HTMLButtonElement => { const node = document.createElement("button"); node.type = "button"; node.textContent = label; node.addEventListener("click", () => host.attempt(action), { signal: controls.signal }); pane.append(node); return node; };
    const select = (label: string): HTMLSelectElement => { const wrap = document.createElement("label"), node = document.createElement("select"); node.setAttribute("aria-label", label); wrap.append(label, node); pane.append(wrap); return node; };
    const field = (label: string, type = "text"): HTMLInputElement => { const wrap = document.createElement("label"), node = document.createElement("input"); node.type = type; node.setAttribute("aria-label", label); wrap.append(label, node); pane.append(wrap); return node; };
    const discard = (): void => { if (candidate) {
        host.session.discard(candidate);
        candidate = undefined;
    } impact.textContent = ""; apply.disabled = cancel.disabled = true; };
    const stage = (operations: DocumentOperation[]): void => {
        discard();
        const { documentId, revision } = host.session.snapshot().manifest;
        candidate = host.session.prepare({ protocolVersion: 1, documentId, baseRevision: revision, operations });
        const slides = host.session.snapshot().metadata.presentation!.slides, ids = candidate.affectedSlideIds ?? [];
        const before = host.session.snapshot(), after = candidate.preview, notes = new Set([...before.metadata.presentation!.stories, ...after.metadata.presentation!.stories].filter(s => s.kind === "notes").map(s => s.storyId));
        const noteSignature = (snapshot: typeof before, model: Pick<typeof host.session, "getNode">, id: string): string => { const p = snapshot.metadata.presentation!, story = p.stories.find(s => s.storyId === id); if (!story || story.kind !== "notes")
            return ""; const master = p.notesMasters.find(m => m.notesMasterId === story.notesMasterId); return JSON.stringify([story, model.getNode(story.rootNodeId)?.html, master, master ? model.getNode(master.rootNodeId)?.html : undefined, p.themes.find(t => t.themeId === master?.themeId), snapshot.styles]); };
        const noteCount = [...notes].filter(id => noteSignature(before, host.session, id) !== noteSignature(after, candidate!.model, id)).length;
        impact.textContent = "待应用：影响 " + ids.length + " 页画布" + (ids.length ? "（" + ids.map(id => slides.find(s => s.slideId === id)?.name ?? id).join("、") + "）" : "") + (noteCount ? "；另需复核 " + noteCount + " 页备注" : "") + "。";
        apply.disabled = cancel.disabled = false;
    };
    const surface = select("定义画布");
    button("打开定义", () => host.selectSurface(surface.value));
    button("返回幻灯片", () => host.selectSurface());
    const layouts = select("当前页布局");
    button("预览布局切换", () => stage([{ op: "setPresentationSlideLayout", slideId: host.activeSlide(), layoutId: layouts.value }]));
    const themes = select("主题定义"), colorRole = select("主题颜色角色"), fontRole = select("主题字体角色");
    for (const [value, label] of [["dk1", "深色 1"], ["dk2", "深色 2"], ["lt1", "浅色 1"], ["lt2", "浅色 2"], ...Array.from({ length: 6 }, (_, i) => ["accent" + (i + 1), "强调色 " + (i + 1)]), ["hlink", "超链接"], ["folHlink", "已访问超链接"]]) {
        const option = document.createElement("option");
        option.value = "--office-theme-" + value;
        option.textContent = label!;
        colorRole.append(option);
    }
    colorRole.value = "--office-theme-accent1";
    for (const [group, name] of [["major", "标题"], ["minor", "正文"]])
        for (const [role, label] of [["ascii", "西文"], ["east-asia", "东亚文字"], ["complex-script", "复杂文字"]]) {
            const option = document.createElement("option");
            option.value = "--office-" + group + "-" + role;
            option.textContent = name + " · " + label;
            fontRole.append(option);
        }
    fontRole.value = "--office-minor-east-asia";
    const color = field("主题颜色值", "color"), font = field("主题字体名称");
    button("预览主题颜色", () => stage([{ op: "setPresentationTheme", themeId: themes.value, properties: { [colorRole.value]: color.value } }]));
    button("预览主题字体", () => stage([{ op: "setPresentationTheme", themeId: themes.value, properties: { [fontRole.value]: JSON.stringify(font.value) } }]));
    const themeReplacement = select("删除后替代主题");
    button("预览删除主题", () => stage([{ op: "removePresentationDefinition", kind: "theme", definitionId: themes.value, ...(themeReplacement.value ? { replacementId: themeReplacement.value } : {}) }]));
    button("复制主题", () => stage([{ op: "insertPresentationDefinition", kind: "theme", copyFromId: themes.value }]));
    button("新建母版", () => stage([{ op: "insertPresentationDefinition", kind: "master", themeId: themes.value }]));
    button("新建备注母版", () => stage([{ op: "insertPresentationDefinition", kind: "notes-master", themeId: themes.value }]));
    const masters = select("布局所属母版"), name = field("布局名称");
    button("新建布局", () => stage([{ op: "insertPresentationDefinition", kind: "layout", masterId: masters.value, name: name.value }]));
    const chosen = (): {
        kind: "master" | "layout" | "notes-master";
        definitionId: string;
    } => { const p = host.session.snapshot().metadata.presentation!, id = surface.value, master = p.masters.find(m => m.rootNodeId === id), layout = p.layouts.find(l => l.rootNodeId === id), notes = p.notesMasters.find(n => n.rootNodeId === id); return master ? { kind: "master", definitionId: master.masterId } : layout ? { kind: "layout", definitionId: layout.layoutId } : notes ? { kind: "notes-master", definitionId: notes.notesMasterId } : (() => { throw new Error("请选择定义。"); })(); };
    button("复制定义", () => { const d = chosen(); stage([{ op: "insertPresentationDefinition", kind: d.kind, copyFromId: d.definitionId }]); });
    button("预览定义引用", () => { const d = chosen(); stage([{ op: "updatePresentationDefinition", ...d, changes: d.kind === "layout" ? { masterId: masters.value } : { themeId: themes.value } }]); });
    button("预览布局名称", () => { const d = chosen(); if (d.kind !== "layout")
        throw new Error("请选择布局。"); stage([{ op: "updatePresentationDefinition", ...d, changes: { name: name.value } }]); });
    const replacement = select("删除后替代定义");
    button("预览删除定义", () => { const d = chosen(); stage([{ op: "removePresentationDefinition", ...d, ...(replacement.value ? { replacementId: replacement.value } : {}) }]); });
    const notes = select("当前页备注母版");
    const assignNotes = button("预览备注母版引用", () => stage([{ op: "setPresentationNotesMaster", slideId: host.activeSlide(), notesMasterId: notes.value || null }]));
    const background = field("定义背景颜色", "color");
    button("预览定义背景", () => stage([{ op: "setStyle", nodeId: surface.value, property: "background-color", value: background.value }]));
    const placeholder = select("占位符角色");
    for (const [value, label] of [["", "普通对象"], ["title", "标题"], ["ctrTitle", "居中标题"], ["subTitle", "副标题"], ["body", "正文"], ["ftr", "页脚"], ["dt", "日期"], ["sldNum", "页码"], ["obj", "内容"]]) {
        const option = document.createElement("option");
        option.value = value!;
        option.textContent = label!;
        placeholder.append(option);
    }
    const index = field("占位符索引", "number");
    index.value = "0";
    index.min = "0";
    button("预览占位符角色", () => { const nodeId = host.object(); if (!nodeId)
        throw new Error("请选择一个对象。"); stage([{ op: "setPresentationPlaceholder", nodeId, placeholder: placeholder.value ? { type: placeholder.value as "title", index: Number(index.value) } : null }]); });
    button("恢复占位符默认值", () => { const nodeId = host.object(); if (!nodeId)
        throw new Error("请选择一个占位符对象。"); stage([{ op: "resetPresentationOverrides", nodeId, geometry: true, appearance: true, text: true }]); });
    pane.append(impact);
    const apply = button("应用定义修改", () => { if (!candidate)
        return; const current = candidate; candidate = undefined; impact.textContent = ""; try {
        host.session.commit(current);
    }
    catch (error) {
        host.session.discard(current);
        throw error;
    } refresh(); }), cancel = button("放弃定义修改", discard);
    const update = (node: HTMLSelectElement, values: readonly {
        value: string;
        label: string;
    }[]): void => { const previous = node.value; node.replaceChildren(...values.map(value => { const option = document.createElement("option"); option.value = value.value; option.textContent = value.label; return option; })); if (values.some(value => value.value === previous))
        node.value = previous; };
    function refresh(): void {
        const p = host.session.snapshot().metadata.presentation!;
        if (candidate && candidate.baseRevision !== host.session.snapshot().manifest.revision) {
            candidate = undefined;
            impact.textContent = "定义候选已过期，请重新预览。";
        }
        apply.disabled = cancel.disabled = !candidate;
        update(surface, [...p.masters.map(m => ({ value: m.rootNodeId, label: "母版 · " + m.masterId })), ...p.layouts.map(l => ({ value: l.rootNodeId, label: "布局 · " + (l.name ?? l.layoutId) })), ...p.notesMasters.map(n => ({ value: n.rootNodeId, label: "备注母版 · " + n.notesMasterId }))]);
        if (host.surface())
            surface.value = host.surface()!;
        update(layouts, p.layouts.map(l => ({ value: l.layoutId, label: l.name ?? l.layoutId })));
        layouts.value = (p.slides.find(s => s.slideId === host.activeSlide()) ?? p.slides[0]!).layoutId;
        update(themes, p.themes.map((t, i) => ({ value: t.themeId, label: "主题 " + (i + 1) })));
        update(themeReplacement, [{ value: "", label: "不替换（仅未被引用时）" }, ...p.themes.filter(t => t.themeId !== themes.value).map(t => ({ value: t.themeId, label: "主题 · " + (p.themes.indexOf(t) + 1) }))]);
        const story = p.stories.find(s => s.kind === "notes" && s.slideId === host.activeSlide());
        update(notes, [{ value: "", label: "无备注母版" }, ...p.notesMasters.map((n, i) => ({ value: n.notesMasterId, label: "备注母版 " + (i + 1) }))]);
        notes.value = story?.kind === "notes" ? story.notesMasterId ?? "" : "";
        notes.disabled = assignNotes.disabled = !story;
        update(masters, p.masters.map(m => ({ value: m.masterId, label: m.masterId })));
        const d = chosen(), values = d.kind === "layout" ? p.layouts.map(l => ({ value: l.layoutId, label: l.name ?? l.layoutId })) : d.kind === "master" ? p.masters.map(m => ({ value: m.masterId, label: m.masterId })) : p.notesMasters.map(n => ({ value: n.notesMasterId, label: n.notesMasterId }));
        update(replacement, [{ value: "", label: "不替换（仅未被引用时）" }, ...values.filter(v => v.value !== d.definitionId)]);
    }
    themes.addEventListener("change", refresh, { signal: controls.signal });
    surface.addEventListener("change", refresh, { signal: controls.signal });
    const unsubscribe = host.session.subscribe(event => { if (event.type === "documentChanged")
        refresh(); });
    refresh();
    return { refresh, dispose() { unsubscribe(); try {
            discard();
        }
        catch { /* Disposed sessions have already revoked candidates. */ }
        finally {
            controls.abort();
            pane.remove();
        } } };
}
