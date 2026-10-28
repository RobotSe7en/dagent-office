import { z } from "zod";

export const FORMAT_VERSION = 1;
export const PROTOCOL_VERSION = 1;

export const IdentifierSchema = z.string().min(1);
export const RevisionSchema = z.number().int().min(0).max(Number.MAX_SAFE_INTEGER);

// Canonical POSIX package paths, not filesystem paths or resource permissions.
export const PackagePathSchema = z.string().regex(
  /^(?!\.{1,2}(?:\/|$))(?![\s\S]*\/\.{1,2}(?:\/|$))[^/\\:\u0000-\u001f\u007f?#]+(?:\/[^/\\:\u0000-\u001f\u007f?#]+)*(?![\s\S])/,
);

export const DocumentKindSchema = z.enum(["document", "presentation", "workbook"]);
export type DocumentKind = z.infer<typeof DocumentKindSchema>;

export const RectangleSchema = z.strictObject({
  x: z.number().min(0),
  y: z.number().min(0),
  width: z.number().positive(),
  height: z.number().positive(),
});
export type Rectangle = z.infer<typeof RectangleSchema>;
