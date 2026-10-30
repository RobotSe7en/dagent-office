import { DocumentError, openDocument } from "@dagent-office/core";
import { resolvePreviewFonts } from "./fonts.js";
import type { FontResolution } from "./fonts.js";
import type { DocumentModel, WordMetadata } from "@dagent-office/core";

export interface DocumentPreview {
  readonly element: HTMLIFrameElement;
  readonly ready: Promise<void>;
  fonts(): readonly FontResolution[];
  dispose(): void;
}
export interface DocumentPreviewOptions {
  readonly title?: string;
}

function packagePath(reference: string, entrypoint: string): string {
  const parts = entrypoint.split("/").slice(0, -1);
  for (const token of reference.split("/")) {
    const value = decodeURIComponent(token);
    if (value === "..") parts.pop();
    else if (value !== ".") parts.push(value);
  }
  return parts.join("/");
}
function numeral(value: number, format: string): string {
  if (format === "none") return "";
  if (format === "upperLetter" || format === "lowerLetter") {
    let result = "";
    for (let n = value; n > 0; n = Math.floor((n - 1) / 26))
      result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
    return format === "lowerLetter" ? result.toLowerCase() : result;
  }
  if (format === "upperRoman" || format === "lowerRoman") {
    if (value > 1_000_000)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "A derived Roman numbering label exceeds the preview budget.",
      );
    let result = "",
      n = value;
    for (const [amount, letters] of [
      [1000, "M"],
      [900, "CM"],
      [500, "D"],
      [400, "CD"],
      [100, "C"],
      [90, "XC"],
      [50, "L"],
      [40, "XL"],
      [10, "X"],
      [9, "IX"],
      [5, "V"],
      [4, "IV"],
      [1, "I"],
    ] as const) {
      while (n >= amount) {
        result += letters;
        n -= amount;
      }
    }
    return format === "lowerRoman" ? result.toLowerCase() : result;
  }
  return String(value);
}
function decorate(
  document: Document,
  word: Readonly<WordMetadata> | undefined,
): void {
  if (!word) return;
  const byId = new Map(
    [...document.querySelectorAll<HTMLElement>("[data-node-id]")].map(
      (node) => [node.dataset.nodeId!, node],
    ),
  );
  for (const run of document.querySelectorAll<HTMLElement>("[data-word-run]")) {
    const computed = document.defaultView!.getComputedStyle(run);
    const threshold = parseFloat(
      computed.getPropertyValue("--word-kerning-threshold"),
    );
    if (Number.isFinite(threshold))
      run.style.fontKerning =
        (parseFloat(computed.fontSize) * 72) / 96 >= threshold
          ? "normal"
          : "none";
  }
  for (const paragraph of document.querySelectorAll<HTMLElement>(
    "p,h1,h2,h3,h4,h5,h6",
  )) {
    const computed = document.defaultView!.getComputedStyle(paragraph);
    paragraph.style.hyphens =
      computed.getPropertyValue("--word-suppress-auto-hyphens").trim() === "1"
        ? "manual"
        : word.automaticHyphenation
          ? "auto"
          : "manual";
  }
  for (const paragraph of document.querySelectorAll<HTMLElement>(
    '[data-word-line-rule="auto"]',
  )) {
    const typed = paragraph as HTMLElement & {
      computedStyleMap?: () => {
        get: (name: string) => { unit?: string; value?: number } | undefined;
      };
    };
    const line = typed.computedStyleMap?.().get("line-height") as
      | { unit?: string; value?: number }
      | undefined;
    const multiplier =
      line?.unit === "number"
        ? (line.value ?? 0)
        : Number(paragraph.style.lineHeight);
    if (!Number.isFinite(multiplier) || multiplier <= 0) continue;
    const computed = document.defaultView!.getComputedStyle(paragraph);
    const probe = document.createElement("span");
    probe.style.cssText =
      "position:absolute;visibility:hidden;display:block;line-height:normal;white-space:pre";
    probe.style.fontFamily = computed.fontFamily;
    probe.style.fontSize = computed.fontSize;
    probe.style.fontWeight = computed.fontWeight;
    probe.style.fontStyle = computed.fontStyle;
    probe.textContent = "汉M";
    document.body.append(probe);
    const height = probe.getBoundingClientRect().height;
    probe.remove();
    paragraph.style.lineHeight = height * multiplier + "px";
  }
  for (const table of word.tables) {
    const element = byId.get(table.nodeId);
    if (!element) continue;
    const columns = document.createElement("colgroup");
    for (const width of table.grid) {
      const column = document.createElement("col");
      column.style.width = width + "pt";
      columns.append(column);
    }
    element.prepend(columns);
  }
  for (const section of document.querySelectorAll<HTMLElement>(
    '[data-office-story="main"] > section',
  ))
    section.classList.add("sdk-section");
  for (const story of word.stories) {
    const root = byId.get(story.rootNodeId);
    if (!root) continue;
    if (story.kind !== "main") {
      const heading = document.createElement("h2");
      heading.className = "sdk-story-label";
      heading.textContent =
        story.kind === "header"
          ? "页眉"
          : story.kind === "footer"
            ? "页脚"
            : "批注";
      root.before(heading);
      root.classList.add("sdk-auxiliary");
    }
    const counters = new Map<string, number[]>();
    for (const paragraph of root.querySelectorAll<HTMLElement>(
      "[data-word-numbering]",
    )) {
      const definition = word.numbering.find(
        (item) => item.numberingId === paragraph.dataset.wordNumbering,
      );
      const level = Number(paragraph.dataset.wordLevel);
      const rule = definition?.levels.find((item) => item.level === level);
      if (!definition || !rule) continue;
      const values = counters.get(definition.numberingId) ?? [];
      values[level] = (values[level] ?? rule.start - 1) + 1;
      for (const deeper of definition.levels)
        if (
          deeper.level > level &&
          deeper.restart !== 0 &&
          (deeper.restart === undefined || deeper.restart - 1 >= level)
        )
          values[deeper.level] = deeper.start - 1;
      counters.set(definition.numberingId, values);
      const marker = document.createElement("span");
      marker.className = "sdk-number-marker " + rule.cssClass;
      marker.setAttribute("aria-hidden", "true");
      marker.textContent = rule.text.replace(
        /%([1-9])/g,
        (_match, token: string) => {
          const index = Number(token) - 1;
          const format = definition.levels.find((item) => item.level === index);
          return numeral(
            values[index] ?? format?.start ?? 1,
            format?.format ?? "decimal",
          );
        },
      );
      if (rule.suffix === "space") marker.textContent += " ";
      const computed = document.defaultView!.getComputedStyle(paragraph);
      const hanging = Math.abs(parseFloat(computed.textIndent) || 24);
      paragraph.style.position = "relative";
      paragraph.style.textIndent = "0";
      marker.style.cssText +=
        ";position:absolute;left:-" +
        hanging +
        "px;width:" +
        hanging +
        "px;margin:0;padding:0;text-indent:0;white-space:pre";
      paragraph.prepend(marker);
    }
  }
  for (const field of word.fields) {
    for (const nested of field.instructionFields ?? []) {
      const definition = word.fields.find(
        (item) => item.nodeId === nested.fieldNodeId,
      );
      for (const id of definition?.resultNodeIds ?? []) {
        const node = byId.get(id);
        if (node) {
          node.hidden = true;
          node.style.display = "none";
        }
      }
    }
  }
  for (const definition of word.tabs) {
    const paragraph = byId.get(definition.paragraphNodeId);
    if (!paragraph) continue;
    // Word tab stops are relative to the text margin, before paragraph indent.
    const indent = parseFloat(
      document.defaultView!.getComputedStyle(paragraph).marginLeft,
    ) || 0;
    for (const tab of paragraph.querySelectorAll<HTMLElement>(
      "[data-word-tab]",
    )) {
      if (tab.closest("[hidden]")) continue;
      const start = paragraph.getBoundingClientRect();
      const box = tab.getBoundingClientRect();
      const cursor = ((box.left - start.left + indent) * 72) / 96;
      const stop = [...definition.stops]
        .filter((item) => item.alignment !== "clear" && item.position > cursor)
        .sort((a, b) => a.position - b.position)[0];
      const target =
        stop?.position ??
        (Math.floor(cursor / definition.defaultInterval) + 1) *
          definition.defaultInterval;
      let followingWidth = 0,
        decimalWidth = 0;
      const range = document.createRange();
      range.setStartAfter(tab);
      range.setEnd(paragraph, paragraph.childNodes.length);
      const following = range.cloneContents().textContent?.split("\t")[0] ?? "";
      const canvas = document.createElement("canvas");
      const context = canvas.getContext("2d")!;
      context.font = document.defaultView!.getComputedStyle(tab).font;
      followingWidth = (context.measureText(following).width * 72) / 96;
      decimalWidth =
        (context.measureText(following.split(/[.,]/)[0] ?? "").width * 72) / 96;
      const adjustment =
        stop?.alignment === "right"
          ? followingWidth
          : stop?.alignment === "center"
            ? followingWidth / 2
            : stop?.alignment === "decimal"
              ? decimalWidth
              : 0;
      const width = Math.max(0, target - cursor - adjustment);
      tab.textContent = "";
      tab.style.display = "inline-block";
      tab.style.width = width + "pt";
      tab.style.whiteSpace = "nowrap";
      tab.style.overflow = "hidden";
      tab.style.verticalAlign = "bottom";
      const leader = stop?.leader;
      if (leader && leader !== "none")
        tab.textContent = (
          leader === "dot"
            ? "."
            : leader === "hyphen"
              ? "-"
              : leader === "middleDot"
                ? "·"
                : "_"
        ).repeat(Math.min(2000, Math.ceil(width)));
    }
  }
}

/** Mount a validated, scriptless read-only view. The iframe DOM is never serialized. */
export function mountDocumentPreview(
  container: HTMLElement,
  input: DocumentModel,
  options: DocumentPreviewOptions = {},
): DocumentPreview {
  const model = openDocument(input.serialize());
  const snapshot = model.snapshot();
  const files = model.serialize();
  const parsed = new DOMParser().parseFromString(snapshot.html, "text/html");
  for (const link of parsed.querySelectorAll("link")) link.remove();
  const policy = parsed.createElement("meta");
  policy.httpEquiv = "Content-Security-Policy";
  policy.content =
    "default-src 'none'; img-src blob:; style-src 'unsafe-inline'; font-src 'none'; base-uri 'none'; form-action 'none'";
  parsed.head.prepend(policy);
  const style = parsed.createElement("style");
  // outerHTML leaves raw style text intact. Escape the slash in HTML end tags
  // using CSS syntax, preserving the string's value and existing escapes.
  style.textContent = snapshot.styles.map((item) => item.css).join("\n")
    .replace(/<\/style/gi, (terminator) => "<\\" + terminator.slice(1));
  parsed.head.append(style);
  const frameStyle = parsed.createElement("style");
  const setup = snapshot.metadata.pageSetup;
  frameStyle.textContent = `html{background:#eceef1}body{margin:0;padding:24px;font-family:'Noto Sans CJK SC',serif;font-size:12pt}article{box-sizing:border-box;margin:0 auto;width:${setup.width}pt;background:white;padding:${setup.margins.top}pt ${setup.margins.right}pt ${setup.margins.bottom}pt ${setup.margins.left}pt}p,h1,h2,h3,h4,h5,h6{margin:0;font-weight:inherit;line-height:normal}td,th{padding:0;font-weight:inherit;text-align:inherit}p{min-height:1em}.sdk-section{box-sizing:border-box}.sdk-story-label{font:600 12px sans-serif!important;color:#58606b!important;margin:24px 0 8px!important;padding:12px 0 0!important;border-top:1px solid #ccc!important}.sdk-auxiliary{border-left:2px solid #ddd;padding-left:12px}[data-word-field]{background:#f5f4ee}ins{color:#006400;text-decoration:underline}del{color:#a00000;text-decoration:line-through}[data-word-break="page"]{display:block;border-top:1px dashed #bbb;margin:18px 0}.sdk-number-marker{pointer-events:none}`;
  parsed.head.insertBefore(frameStyle, style);
  const byId = new Map(
    [...parsed.querySelectorAll<HTMLElement>("[data-node-id]")].map(
      (node) => [node.dataset.nodeId!, node],
    ),
  );
  let maximumWidth = setup.width;
  const sectionNodes: HTMLElement[] = [];
  for (const section of snapshot.metadata.sections) {
    const effectiveSetup = section.pageSetup ?? setup;
    const node = byId.get(section.rootNodeId);
    if (!node) continue;
    maximumWidth = Math.max(maximumWidth, effectiveSetup.width);
    const { width, margins } = effectiveSetup;
    node.classList.add("sdk-section");
    node.style.boxSizing = "border-box";
    node.style.width = width + "pt";
    node.style.margin = "0 auto";
    node.style.padding = `${margins.top}pt ${margins.right}pt ${margins.bottom}pt ${margins.left}pt`;
    sectionNodes.push(node);
  }
  const article = parsed.querySelector("article");
  if (article && sectionNodes.length && !sectionNodes.includes(article)) {
    article.style.width = maximumWidth + "pt";
    article.style.padding = "0";
    const auxiliary = parsed.createElement("div");
    auxiliary.style.padding = `0 ${setup.margins.right}pt ${setup.margins.bottom}pt ${setup.margins.left}pt`;
    for (const node of [...article.children])
      if (
        !sectionNodes.some((section) => node === section || node.contains(section))
      )
        auxiliary.append(node);
    if (auxiliary.children.length) article.append(auxiliary);
  }
  const urls: string[] = [];
  for (const image of parsed.querySelectorAll("img")) {
    const path = packagePath(
      image.getAttribute("src")!,
      snapshot.manifest.entrypoint,
    );
    const asset = snapshot.manifest.assets.find((item) => item.path === path)!;
    const bytes = files.get(path)!;
    const buffer = new ArrayBuffer(bytes.length);
    new Uint8Array(buffer).set(bytes);
    const url = URL.createObjectURL(
      new Blob([buffer], { type: asset.mimeType }),
    );
    urls.push(url);
    image.src = url;
  }
  const iframe = document.createElement("iframe");
  iframe.title = options.title ?? "文档只读预览";
  iframe.setAttribute("sandbox", "allow-same-origin");
  iframe.style.cssText = "width:100%;height:480px;border:0;display:block";
  let fontReport: readonly FontResolution[] = [];
  let disposed = false;
  let observer: ResizeObserver | undefined;
  let settle: () => void = () => {};
  let rejectReady: (reason: unknown) => void = () => {};
  const ready = new Promise<void>((resolve, reject) => {
    settle = resolve;
    rejectReady = reject;
  });
  const load = (): void => {
    const frameDocument = iframe.contentDocument;
    if (disposed || !frameDocument) return;
    void frameDocument.fonts.ready
      .then(() => {
        if (disposed) return;
        const sourceFonts = resolvePreviewFonts(frameDocument);
        decorate(
          frameDocument,
          snapshot.metadata.word as WordMetadata | undefined,
        );
        fontReport = Object.freeze([
          ...sourceFonts,
          ...resolvePreviewFonts(frameDocument, ".sdk-number-marker"),
        ]);
        const substitutions = fontReport.filter((item) => item.substituted);
        const missing = [
          ...new Set(fontReport.flatMap((item) => [...item.missingCharacters])),
        ];
        if (
          substitutions.length ||
          missing.length ||
          fontReport.some((item) => !item.glyphAuditComplete)
        ) {
          const report = frameDocument.createElement("p");
          report.setAttribute("role", "note");
          report.style.cssText =
            "font:12px sans-serif;color:#694800;margin:16px auto;max-width:600px";
          report.textContent =
            (substitutions.length
              ? "字体替代：" +
                substitutions
                  .map(
                    (item) =>
                      item.requestedFamily + " → " + item.selectedFamily,
                  )
                  .join("；") +
                "。"
              : "") +
            (missing.length
              ? "缺少字形：" +
                missing
                  .map(
                    (character) =>
                      "U+" +
                      character.codePointAt(0)!.toString(16).toUpperCase() +
                      "（" +
                      character +
                      "）",
                  )
                  .join("、") +
                "。"
              : "") +
            (fontReport.some((item) => !item.glyphAuditComplete)
              ? "部分字形未核验。"
              : "") +
            "断行和版式可能与原字体不同。";
          frameDocument.body.append(report);
        }
        const resize = (): void => {
          iframe.style.height =
            Math.max(200, frameDocument.documentElement.scrollHeight) + "px";
        };
        for (const image of frameDocument.images)
          image.addEventListener("load", resize, { once: true });
        resize();
        observer = new ResizeObserver(resize);
        observer.observe(frameDocument.body);
        settle();
      })
      .catch(rejectReady);
  };
  iframe.addEventListener("load", load, { once: true });
  iframe.srcdoc = "<!doctype html>" + parsed.documentElement.outerHTML;
  container.append(iframe);
  return {
    element: iframe,
    ready,
    fonts: () => fontReport,
    dispose: () => {
      if (disposed) return;
      disposed = true;
      observer?.disconnect();
      iframe.removeEventListener("load", load);
      iframe.remove();
      iframe.srcdoc = "";
      for (const url of urls) URL.revokeObjectURL(url);
      settle();
    },
  };
}
