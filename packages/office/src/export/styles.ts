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
  constructor(readonly css: ExportCss, readonly word: DeepReadonly<WordMetadata> | undefined, private readonly locale: string | undefined) {
    this.definitions = new Map(word?.styles.map((style) => [style.styleId, style]) ?? []);
    this.ids = new Map([...this.definitions.keys()].map((id) => [id, id]));
    const main = css.classStyle("word-main");
    this.defaults = { ...pick(main, runKeys), ...pick(css.classStyle(word?.defaultRunCssClass ?? "word-default-run", main), runKeys.filter((key) => key === "background-color" || key.startsWith("text-decoration"))) };
    // Paragraph source nodes carry resolved defaults; a native default of zero
    // spacing is explicit so producers do not substitute their UI defaults.
    this.paragraphDefaults = { "margin-top": "0pt", "margin-bottom": "0pt" };
    word?.numbering.forEach((definition, index) => this.numberingIds.set(definition.numberingId, String(index + 1)));
  }
  style(id: string | undefined): Style | undefined { return id ? this.definitions.get(id) : undefined; }
  paragraphContext(id: string | undefined): Css { const style = this.style(id); return this.css.classStyle(style?.cssClass ?? "word-native-no-style"); }
  runContext(paragraph: string | undefined, character?: string): Css {
    let result = this.defaults;
    for (const id of [paragraph, character]) {
      const definition = this.style(id); if (!definition) continue;
      result = { ...result, ...pick(this.css.classStyle(definition.cssClass, result), runKeys) };
      if (definition.runCssClass) result = { ...result, ...pick(this.css.classStyle(definition.runCssClass, result), runKeys.filter((key) => key === "background-color" || key.startsWith("text-decoration"))) };
    }
    return result;
  }
  private values(style: Style): Css {
    const result = this.css.classStyle(style.cssClass, this.defaults);
    return style.runCssClass ? { ...result, ...pick(this.css.classStyle(style.runCssClass, result), ["background-color", "text-decoration-line", "text-decoration-style", "text-decoration-color"]) } : result;
  }
  write(opc: OutputOpc): void {
    const out = opc.part("word/styles.xml", contentType("styles"));
    out.add(declaration + "<w:styles" + namespaces + ">");
    out.add("<w:docDefaults><w:rPrDefault><w:rPr>" + runProperties(pick(this.defaults, runKeys.filter((key) => key !== "vertical-align" || this.defaults[key] !== "baseline"))) + (this.locale ? empty("w:lang", { "w:val": this.locale }) : "") + "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>" + paragraphProperties(this.paragraphDefaults) + "</w:pPr></w:pPrDefault></w:docDefaults>");
    for (const style of this.definitions.values()) {
      const css = this.values(style), parent = style.basedOn ? this.values(this.definitions.get(style.basedOn)!) : {};
      out.add("<w:style" + attributes({ "w:type": style.kind, "w:styleId": this.ids.get(style.styleId)!, ...(style.kind === "paragraph" && style.styleId === "Normal" ? { "w:default": 1 } : {}) }) + ">" + value("name", style.styleId));
      for (const [key, name] of [["basedOn", "basedOn"], ["linkedStyleId", "link"], ["nextStyleId", "next"]] as const) if (style[key]) out.add(value(name, this.ids.get(style[key])));
      out.add("<w:pPr>" + paragraphProperties(pick(css, style.paragraphProperties ?? paragraphKeys)) + "</w:pPr>");
      out.add("<w:rPr>" + runProperties(pick(css, style.runProperties ?? runKeys), parent, true) + "</w:rPr>");
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
