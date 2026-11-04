import { z } from "zod";

const string = z.string().max(32767).optional();
/** Document provenance supplied by the file, never host identity or access
 * control. Dates retain native lexical precision and timezone as source text. */
export const DocumentPropertiesSchema = z.strictObject({
  creator: string, subject: string, description: string, keywords: string,
  lastModifiedBy: string, revision: string, category: string, contentStatus: string,
  identifier: string, version: string, contentType: string, created: string, modified: string, lastPrinted: string,
  template: string, manager: string, company: string, application: string, applicationVersion: string,
  totalEditingMinutes: z.number().int().min(0).max(0x7fff_ffff).optional(),
});
