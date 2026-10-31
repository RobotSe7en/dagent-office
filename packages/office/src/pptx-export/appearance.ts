import type { Css } from "../export/css.js";
import { attributes, empty } from "../export/xml.js";
import { emu, integer, rgba, points, unsupported } from "./values.js";
import type { NativeSurface, PresentationExportContext } from "./context.js";
import type { SourceElement } from "../export/source.js";
const roles = new Set(["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink", "bg1", "bg2", "tx1", "tx2"]);
const amount = (css: Css, name: string, fallback: number): number => {
    const raw = css[name];
    if (raw === undefined)
        return fallback;
    const value = Number(raw);
    if (!Number.isFinite(value))
        unsupported("A native paint constant is invalid.");
    return value;
};
export function colorXml(css: Css, prefix: string, raw: Css = {}, placeholder = false, fallback = "#000000"): string {
    const value = css[prefix + "-color"] ?? fallback, authored = raw[prefix + "-color"] ?? value, reference = /^var\(\s*--office-theme-([a-zA-Z\d]+)(?:\s*[,)]|\s*$)/.exec(authored)?.[1], isPlaceholder = placeholder && css[prefix + "-source"] === "reference";
    const parsed = isPlaceholder || reference && roles.has(reference) ? { rgb: "000000", alpha: 1 } : rgba(value, css.color ?? "#000000"), kind = isPlaceholder || reference && roles.has(reference) ? "schemeClr" : "srgbClr", val = isPlaceholder ? "phClr" : reference && roles.has(reference) ? reference : parsed.rgb;
    let result = "<a:" + kind + attributes({ val }) + ">";
    for (const [name, key, defaultValue] of [["lumMod", "lum-mod", 1], ["lumOff", "lum-off", 0], ["shade", "shade", 1], ["tint", "tint", 0], ["alpha", "alpha", 1]] as const) {
        const value = amount(css, prefix + "-" + key, defaultValue) * (name === "alpha" ? parsed.alpha : 1);
        if (value < 0 || value > 1)
            unsupported("A color transform is outside the PPTX profile.");
        if (value !== defaultValue)
            result += empty("a:" + name, { val: integer(value, 100000, 0, 100000) });
    }
    return result + "</a:" + kind + ">";
}
export function fillXml(css: Css, prefix = "--office-fill", raw: Css = {}, placeholder = false, fallback = "none"): string {
    const kind = css[prefix + "-kind"], paint = css[prefix + "-color"];
    if (kind === "none" || !kind && !paint && fallback === "none")
        return "<a:noFill/>";
    if (kind !== "linear")
        return "<a:solidFill>" + colorXml(css, prefix, raw, placeholder, fallback === "none" ? "#000000" : fallback) + "</a:solidFill>";
    const count = amount(css, prefix + "-stop-count", 0);
    if (!Number.isSafeInteger(count) || count < 2 || count > 1024)
        unsupported("A gradient stop list is outside its native range.");
    let result = "<a:gradFill" + attributes({ rotWithShape: css[prefix + "-rotate-with-shape"] === "false" ? 0 : 1 }) + "><a:gsLst>";
    for (let index = 0; index < count; index++) {
        const stop = prefix + "-stop-" + index;
        result += "<a:gs" + attributes({ pos: integer(amount(css, stop + "-position", 0), 100000, 0, 100000) }) + ">" + colorXml(css, stop, raw, placeholder) + "</a:gs>";
    }
    const angle = amount(css, prefix + "-angle", 0);
    return result + "</a:gsLst>" + empty("a:lin", { ang: integer(((angle % 360) + 360) % 360, 60000, 0, 21600000), scaled: css[prefix + "-scaled"] === "false" ? 0 : 1 }) + "</a:gradFill>";
}
export function lineXml(css: Css, prefix = "--office-line", raw: Css = {}, placeholder = false, fallback = "none", tag = "a:ln"): string {
    const width = points(css[prefix + "-width"], css, 1), kind = css[prefix + "-kind"], color = css[prefix + "-color"];
    let result = "<" + tag + attributes({ w: emu(width, 0, 2147483647), cap: css[prefix + "-cap"] ?? "flat", cmpd: css[prefix + "-cmpd"] ?? "sng", algn: css[prefix + "-algn"] ?? "ctr" }) + ">";
    if (kind === "none" || !kind && !color && fallback === "none")
        result += "<a:noFill/>";
    else
        result += "<a:solidFill>" + colorXml(css, prefix, raw, placeholder, fallback === "none" ? "#000000" : fallback) + "</a:solidFill>";
    if (css[prefix + "-dash"])
        result += empty("a:prstDash", { val: css[prefix + "-dash"] });
    const join = css[prefix + "-join"];
    if (join === "miter")
        result += empty("a:miter", { lim: integer(amount(css, prefix + "-miter-limit", 8), 100000, 0, 2147483647) });
    else if (join === "round" || join === "bevel")
        result += empty("a:" + join);
    for (const [side, name] of [["start", "headEnd"], ["end", "tailEnd"]] as const)
        if (css[prefix + "-" + side + "-arrow"])
            result += empty("a:" + name, { type: css[prefix + "-" + side + "-arrow"], w: css[prefix + "-" + side + "-w"] ?? "med", len: css[prefix + "-" + side + "-len"] ?? "med" });
    return result + "</" + tag + ">";
}
export function effectXml(css: Css, raw: Css = {}, placeholder = false): string {
    if (css["--office-shadow-kind"] !== "outer")
        return "<a:effectLst/>";
    const distance = points(css["--office-shadow-distance"], css, 0), direction = amount(css, "--office-shadow-direction", 0);
    return "<a:effectLst><a:outerShdw" + attributes({ blurRad: emu(points(css["--office-shadow-blur"], css, 0), 0), dist: emu(distance, 0), dir: integer(((direction % 360) + 360) % 360, 60000, 0, 21600000), algn: css["--office-shadow-align"] ?? "b", rotWithShape: css["--office-shadow-rotate-with-shape"] === "false" ? 0 : 1 }) + ">" + colorXml(css, "--office-shadow", raw, placeholder) + "</a:outerShdw></a:effectLst>";
}
function refPrefix(css: Css, prefix: string): string | undefined {
    const result = css[prefix + "-source"] === "reference" ? css[prefix + "-reference-prefix"] : undefined;
    if (result !== undefined && !["--office-fill-reference", "--office-line-reference", "--office-background-reference", "--office-effect-reference"].includes(result))
        unsupported("A native style color reference is invalid.");
    return result;
}
/** Keep the two native color stages rather than flattening theme references:
 * A lives in fmtScheme and B is the object's style-reference color. */
export function objectAppearance(context: PresentationExportContext, node: SourceElement, surface: NativeSurface, shapeKind: "geometry" | "text" | "connector" | "group", preset?: string): {
    properties: string;
    style: string;
} {
    let css = context.styles.get(node, surface), raw = context.styles.raw(node, surface);
    const own = context.styles.own(node, surface), refs: {
        kind: "fill" | "line" | "effect";
        index: number;
        color: string;
    }[] = [], has = (prefix: string): boolean => Object.keys(own).some(name => name.startsWith(prefix + "-")), placeholder = context.objects.get(node.attrs["data-node-id"]!)?.placeholder;
    if ((shapeKind === "geometry" || shapeKind === "text") && preset !== "line" && !css["--office-fill-kind"] && !css["--office-fill-color"] && css["background-color"] && rgba(css["background-color"]!, css.color).alpha > 0) {
        css = { ...css, "--office-fill-kind": "solid", "--office-fill-color": css["background-color"]!, "--office-fill-source": "direct" };
        raw = { ...raw, "--office-fill-color": raw["background-color"] ?? css["background-color"]! };
    }
    let properties = "";
    const fillReference = (css["--office-fill-kind"] || css["--office-fill-color"] ? refPrefix(css, "--office-fill") : undefined) ?? Array.from({ length: Math.min(1024, amount(css, "--office-fill-stop-count", 0)) }, (_, i) => refPrefix(css, "--office-fill-stop-" + i)).find(Boolean);
    if (!placeholder || has("--office-fill")) {
        if (fillReference) {
            const style = fillXml(css, "--office-fill", raw, true), index = context.format(surface.themeId, "fill", style);
            refs.push({ kind: "fill", index, color: colorXml(css, fillReference, raw) });
        }
        else
            properties += fillXml(css, "--office-fill", raw, false, shapeKind === "geometry" && preset !== "line" ? "#4472c4" : "none");
    }
    const lineReference = css["--office-line-kind"] || css["--office-line-color"] ? refPrefix(css, "--office-line") : undefined;
    if (!placeholder || has("--office-line")) {
        if (lineReference) {
            const index = context.format(surface.themeId, "line", lineXml(css, "--office-line", raw, true));
            refs.push({ kind: "line", index, color: colorXml(css, lineReference, raw) });
        }
        else
            properties += lineXml(css, "--office-line", raw, false, shapeKind === "connector" || preset === "line" ? "#000000" : "none");
    }
    const effectReference = refPrefix(css, "--office-shadow");
    if (!placeholder || has("--office-shadow")) {
        if (effectReference) {
            const index = context.format(surface.themeId, "effect", "<a:effectStyle>" + effectXml(css, raw, true) + "</a:effectStyle>");
            refs.push({ kind: "effect", index, color: colorXml(css, effectReference, raw) });
        }
        else
            properties += effectXml(css, raw);
    }
    const defaultWhite = shapeKind === "geometry" && preset !== "line" && !["linear", "none"].includes(css["--office-fill-kind"] ?? "") && !css["--office-text-color"];
    if (!refs.length && !defaultWhite)
        return { properties, style: "" };
    let style = "<p:style>";
    for (const kind of ["line", "fill", "effect"] as const) {
        const ref = refs.find(ref => ref.kind === kind), name = kind === "line" ? "lnRef" : kind === "fill" ? "fillRef" : "effectRef";
        style += "<a:" + name + attributes({ idx: ref?.index ?? 0 }) + ">" + (ref?.color ?? '<a:srgbClr val="000000"/>') + "</a:" + name + ">";
    }
    const textCss = css["--office-text-color"] ? css : { ...css, "--office-text-color": css.color ?? "#000000" }, textRaw = raw["--office-text-color"] ? raw : { ...raw, "--office-text-color": raw.color ?? css.color ?? "#000000" };
    style += '<a:fontRef idx="none">' + colorXml(textCss, "--office-text", textRaw) + '</a:fontRef></p:style>';
    return { properties, style };
}
