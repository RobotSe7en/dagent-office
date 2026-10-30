import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WordMetadata } from "@dagent-office/core";
import type { Css, ExportCss } from "./css.js";
import { cellMargins, cellProperties, paragraphKeys, paragraphProperties, pick, runKeys, runProperties, tableProperties } from "./properties.js";
import { attributes, declaration, empty, escapeXml, namespaces, value } from "./xml.js";
import { contentType } from "./opc.js";
import type { OutputOpc } from "./opc.js";

type Style = DeepReadonly<WordMetadata["styles"][number]>;
export interface HtmlNumbering { readonly id: string; readonly start: number; readonly level: number; readonly format: string; readonly text: string }
export class ExportStyles {
  readonly defaults: Css;
  readonly paragraphDefaults: Css;
  readonly ids: ReadonlyMap<string, string>;
  readonly numberingIds: Map<string, string> = new Map<string, string>();
  private readonly definitions: Map<string, Style>;
  private readonly native = new Map<string, { p: Css; r: Css }>();
  constructor(readonly css: ExportCss, readonly word: DeepReadonly<WordMetadata> | undefined, private readonly locale: string | undefined, private readonly maxDepth: number, main: Css) {
    this.definitions = new Map(word?.styles.map((style) => [style.styleId, style]) ?? []);
    this.ids = new Map([...this.definitions.keys()].map((id) => [id, id]));
    const special = runKeys.filter((key) => key === "background-color" || key.startsWith("text-decoration"));
    this.defaults = { ...pick(main, runKeys.filter((key) => !special.includes(key) && (word || key !== "--word-run-shading"))), ...(word?.defaultRunCssClass ? pick(css.classStyle(word.defaultRunCssClass, main), special) : pick(css.emptyStyle(), special)) };
    // Paragraph source nodes carry resolved defaults; a native default of zero
    // spacing is explicit so producers do not substitute their UI defaults.
    this.paragraphDefaults = { "margin-top": "0pt", "margin-bottom": "0pt" };
    word?.numbering.forEach((definition, index) => this.numberingIds.set(definition.numberingId, String(index + 1)));
  }
  style(id: string | undefined): Style | undefined { return id ? this.definitions.get(id) : undefined; }
  paragraphContext(id: string | undefined): Css { return { ...this.css.emptyStyle(), ...(id && this.style(id) ? this.nativeValues(this.style(id)!).p : {}) }; }
  runContext(paragraph: string | undefined, character?: string): Css {
    let result = this.defaults;
    for (const id of [paragraph, character]) {
      const definition = this.style(id); if (!definition) continue;
      result = { ...result, ...this.nativeValues(definition).r };
    }
    return result;
  }
  private values(style: Style): Css {
    const result = this.css.classStyle(style.cssClass, this.defaults);
    return style.runCssClass ? { ...result, ...pick(this.css.classStyle(style.runCssClass, result), ["background-color", "text-decoration-line", "text-decoration-style", "text-decoration-color"]) } : result;
  }
  private own(style: Style, kind: "p" | "r", css: Css, parent: Css): Css {
    const original = kind === "p" ? style.paragraphProperties : style.runProperties;
    const allowed = kind === "p" ? paragraphKeys : runKeys;
    const declared = new Set([...this.css.classKeys(style.cssClass), ...(style.runCssClass && kind === "r" ? this.css.classKeys(style.runCssClass) : [])]);
    const owns = (key: string): boolean => original === undefined || original.some((name) => name === key || key.startsWith(name + "-") && ["border", "margin", "padding"].some((prefix) => name.startsWith(prefix)));
    return pick(css, allowed.filter((key) => owns(key) || declared.has(key) && css[key] !== parent[key] && !(kind === "r" && style.kind !== "character" && key === "background-color" && !style.runCssClass)));
  }
  private nativeValues(style: Style, depth = 0): { p: Css; r: Css } {
    const known = this.native.get(style.styleId); if (known) return known;
    if (depth > this.maxDepth) throw new DocumentError("ResourceLimitExceeded", "Native style inheritance exceeds its depth budget.");
    const css = this.values(style), parentStyle = style.basedOn ? this.style(style.basedOn) : undefined;
    const parent = parentStyle ? this.nativeValues(parentStyle, depth + 1) : { p: {}, r: this.defaults };
    const parentCss = parentStyle ? this.values(parentStyle) : this.defaults;
    const result = { p: { ...parent.p, ...this.own(style, "p", css, parentCss) }, r: { ...parent.r, ...this.own(style, "r", css, parentCss) } };
    this.native.set(style.styleId, result); return result;
  }
  write(opc: OutputOpc): void {
    const out = opc.part("word/styles.xml", contentType("styles"));
    out.add(declaration + "<w:styles" + namespaces + ">");
    out.add("<w:docDefaults><w:rPrDefault><w:rPr>" + runProperties(pick(this.defaults, runKeys.filter((key) => key !== "vertical-align" || this.defaults[key] !== "baseline"))) + (this.locale ? empty("w:lang", { "w:val": this.locale }) : "") + "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>" + paragraphProperties(this.paragraphDefaults) + "</w:pPr></w:pPrDefault></w:docDefaults>");
    for (const style of this.definitions.values()) {
      const css = this.values(style), parent = style.basedOn ? this.values(this.definitions.get(style.basedOn)!) : this.defaults;
      out.add("<w:style" + attributes({ "w:type": style.kind, "w:styleId": this.ids.get(style.styleId)!, ...(style.kind === "paragraph" && style.styleId === "Normal" ? { "w:default": 1 } : {}) }) + ">" + value("name", style.styleId));
      for (const [key, name] of [["basedOn", "basedOn"], ["linkedStyleId", "link"], ["nextStyleId", "next"]] as const) if (style[key]) out.add(value(name, this.ids.get(style[key])));
      out.add("<w:pPr>" + paragraphProperties({ ...this.own(style, "p", css, parent), "font-size": css["font-size"]! }) + "</w:pPr>");
      out.add("<w:rPr>" + runProperties(this.own(style, "r", css, parent), parent, true) + "</w:rPr>");
      if (style.kind === "table") {
        out.add("<w:tblPr>" + (style.rowBandSize ? value("tblStyleRowBandSize", style.rowBandSize) : "") + (style.columnBandSize ? value("tblStyleColBandSize", style.columnBandSize) : "") + tableProperties(css) + cellMargins(css, "tblCellMar") + (style.tableLookMask !== undefined ? empty("w:tblLook", { "w:val": style.tableLookMask.toString(16).padStart(4, "0") }) : "") + "</w:tblPr>");
        out.add("<w:tcPr>" + cellProperties(pick(css, Object.keys(css).filter((key) => key.startsWith("padding-")))) + "</w:tcPr>");
        for (const condition of style.conditions ?? []) {
          const conditional = pick(this.css.classStyle(condition.cssClass, this.defaults), this.css.classKeys(condition.cssClass));
          out.add('<w:tblStylePr w:type="' + escapeXml(condition.region) + '"><w:pPr>' + paragraphProperties(pick(conditional, paragraphKeys)) + "</w:pPr><w:rPr>" + runProperties(pick(conditional, runKeys)) + "</w:rPr><w:tblPr>" + tableProperties(conditional) + "</w:tblPr><w:tcPr>" + cellProperties(conditional) + "</w:tcPr></w:tblStylePr>");
        }
      }
      out.add("</w:style>");
    }
    out.add("</w:styles>"); out.finish(); opc.relate("word/document.xml", "styles", "styles.xml");
  }
  writeNumbering(opc: OutputOpc, html: readonly HtmlNumbering[] = []): void {
    if (!this.word?.numbering.length && !html.length) return;
    const out = opc.part("word/numbering.xml", contentType("numbering")); out.add(declaration + "<w:numbering" + namespaces + ">");
    for (const definition of this.word?.numbering ?? []) {
      const id = this.numberingIds.get(definition.numberingId)!;
      out.add('<w:abstractNum w:abstractNumId="' + id + '">' + value("multiLevelType", "multilevel"));
      for (const level of definition.levels) {
        const css = pick(this.css.classStyle(level.cssClass, this.defaults), this.css.classKeys(level.cssClass));
        out.add('<w:lvl w:ilvl="' + level.level + '">' + value("start", level.start) + value("numFmt", level.format) + (level.restart !== undefined ? value("lvlRestart", level.restart) : "") + value("suff", level.suffix) + value("lvlText", level.text) + '<w:pPr>' + paragraphProperties(pick(css, paragraphKeys.filter((key) => !["background-color"].includes(key)))) + '</w:pPr><w:rPr>' + runProperties(pick(css, runKeys)) + "</w:rPr></w:lvl>");
      }
      out.add("</w:abstractNum>");
    }
    for (const item of html) out.add('<w:abstractNum w:abstractNumId="' + item.id + '"><w:lvl w:ilvl="' + item.level + '">' + value("start", item.start) + value("numFmt", item.format) + value("lvlText", item.text) + value("suff", "tab") + '<w:pPr><w:tabs><w:tab w:val="left" w:pos="' + ((item.level + 1) * 720) + '"/></w:tabs><w:ind w:left="' + ((item.level + 1) * 720) + '" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>');
    for (const definition of this.word?.numbering ?? []) { const id = this.numberingIds.get(definition.numberingId)!; out.add('<w:num w:numId="' + id + '">' + value("abstractNumId", id) + "</w:num>"); }
    for (const item of html) out.add('<w:num w:numId="' + item.id + '">' + value("abstractNumId", item.id) + "</w:num>");
    out.add("</w:numbering>"); out.finish(); opc.relate("word/document.xml", "numbering", "numbering.xml");
  }
}
