import { C, check, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";

/** Office infers crossesAt from the perpendicular axis (MS-OI29500
 * §2.1.1447). The selected basic profile has one category/value axis pair. */
export function chartAxisCrossing(source: XmlNode): Css {
  const choices = elements(source, undefined, C).filter(node => node.name === "crosses" || node.name === "crossesAt");
  if (choices.length > 1) fail(source, "A chart axis has multiple crossing choices.");
  const choice = choices[0];
  if (!choice) return {};
  if ([...choice.attrs.keys()].some(key => key !== " val")) unsupported(choice, "A chart axis crossing has unknown attributes.");
  check(choice, [], C);
  if (choice.content.some(item => typeof item === "string" && item.length)) fail(choice, "A chart axis crossing has unexpected text.");
  const raw = choice.attrs.get(" val")?.replace(/^[\x20\t\r\n]+|[\x20\t\r\n]+$/g, "");
  if (choice.name === "crosses") {
    if (!raw || !["autoZero", "min", "max"].includes(raw)) fail(choice, "A chart axis crossing has an invalid enum value.");
    return { "--office-axis-crosses": raw };
  }
  if (!raw || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(raw) || !Number.isFinite(Number(raw))) fail(choice, "A chart axis crossing requires a finite native number.");
  if (source.name === "valAx" && !Number.isInteger(Number(raw))) unsupported(choice, "Office category crossing values must be integral category numbers.");
  return { "--office-axis-crosses-at": String(Number(raw)) };
}
