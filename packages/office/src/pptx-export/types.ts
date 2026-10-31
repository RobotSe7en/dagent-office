import type { DocumentZipOptions } from "../document-zip.js";
export interface PptxExportOptions extends DocumentZipOptions {
    readonly commentFormat?: "preserve" | "modern" | "classic";
}
export interface PptxAnnotationLoss {
    readonly annotationId: string;
    readonly kind: "anchor" | "thread" | "status" | "format" | "metadata";
    readonly message: string;
}
export interface UnexportedPptxAnnotation {
    readonly annotationId: string;
    readonly reason: "missing_surface" | "unsupported_anchor";
}
export interface PptxExportReport {
    readonly documentId: string;
    readonly revision: number;
    readonly exportedAnnotationIds: readonly string[];
    readonly unexportedAnnotations: readonly UnexportedPptxAnnotation[];
    readonly annotationLosses: readonly PptxAnnotationLoss[];
}
export interface PptxExportResult {
    readonly bytes: Uint8Array;
    readonly report: PptxExportReport;
}
