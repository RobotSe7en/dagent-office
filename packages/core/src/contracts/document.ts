import { z } from "zod";
import { DocumentKindSchema, FORMAT_VERSION, IdentifierSchema, PackagePathSchema, RevisionSchema } from "./primitives.js";

export const AssetDescriptorSchema = z.strictObject({
  path: PackagePathSchema,
  mimeType: z.string().min(1),
  sha256: z.string().length(64).regex(/^[a-f0-9]{64}$/),
});
export type AssetDescriptor = z.infer<typeof AssetDescriptorSchema>;

export const DocumentManifestSchema = z.strictObject({
  format: z.literal("dagent-office"),
  formatVersion: z.literal(FORMAT_VERSION),
  documentId: IdentifierSchema,
  kind: DocumentKindSchema,
  revision: RevisionSchema,
  entrypoint: PackagePathSchema,
  styles: z.array(PackagePathSchema),
  metadata: PackagePathSchema,
  annotations: PackagePathSchema,
  assets: z.array(AssetDescriptorSchema),
});
export type DocumentManifest = z.infer<typeof DocumentManifestSchema>;

export const PageSetupSchema = z.strictObject({
  unit: z.literal("pt"),
  width: z.number().positive(),
  height: z.number().positive(),
  margins: z.strictObject({
    top: z.number().min(0),
    right: z.number().min(0),
    bottom: z.number().min(0),
    left: z.number().min(0),
  }),
});
export type PageSetup = z.infer<typeof PageSetupSchema>;

export const DocumentSectionSchema = z.strictObject({
  sectionId: IdentifierSchema,
  rootNodeId: IdentifierSchema,
});
export type DocumentSection = z.infer<typeof DocumentSectionSchema>;

// Word metadata currently defined by the document specification.
// Numbering and header/footer mappings will be added explicitly during DOCX work.
export const DocumentMetadataSchema = z.strictObject({
  title: z.string().optional(),
  locale: IdentifierSchema.optional(),
  pageSetup: PageSetupSchema,
  sections: z.array(DocumentSectionSchema).min(1),
});
export type DocumentMetadata = z.infer<typeof DocumentMetadataSchema>;
