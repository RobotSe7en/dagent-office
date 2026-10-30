import { parseContract } from "../schema-registry.js";
import type { DocumentPackage, DocumentReadLimits, DocumentSnapshot } from "../runtime.js";
import { normalizeCss } from "./css.js";
import { parseHtml } from "./html.js";
import { validateAssets } from "./assets.js";
import { validateWord } from "./word.js";
import { readManifest } from "./manifest.js";
import { encode, readJson, readText } from "./package.js";
import { checkLimit, freeze, invalid, unsupported } from "./validation.js";

/** Private bytes are never exposed; unchanged assets may be shared across states. */
export interface DocumentState {
  readonly files: DocumentPackage;
  readonly snapshot: DocumentSnapshot;
  readonly html: ReturnType<typeof parseHtml>;
}

export function buildDocumentState(input: DocumentPackage, limits: DocumentReadLimits): DocumentState {
  const files = new Map(input);
  const manifest = readManifest(files);
  if (manifest.kind !== "document") unsupported("Only document packages can be opened by this reader.");

  validateAssets(manifest.assets, files);
  const metadata = parseContract("document-metadata", readJson(files, manifest.metadata));
  const annotations = parseContract("annotations", readJson(files, manifest.annotations));
  const cssBudget = { nodes: 0 };
  const styles = manifest.styles.map((path) => ({ path, css: normalizeCss(readText(files, path), "stylesheet", limits, cssBudget) }));
  const html = parseHtml(readText(files, manifest.entrypoint), manifest, files, limits, cssBudget);

  const sectionIds = new Set<string>();
  const sectionRoots = new Set<string>();
  for (const section of metadata.sections) {
    if (sectionIds.has(section.sectionId) || sectionRoots.has(section.rootNodeId) || !html.isContainer(section.rootNodeId)) {
      invalid("Document sections require unique identities and existing root containers.");
    }
    sectionIds.add(section.sectionId);
    sectionRoots.add(section.rootNodeId);
  }
  const annotationIds = new Set<string>();
  for (const annotation of annotations) {
    if (annotationIds.has(annotation.annotationId) || annotation.documentId !== manifest.documentId
      || annotation.baseRevision > manifest.revision) {
      invalid("Annotations require unique identities and must belong to this document and its history.");
    }
    annotationIds.add(annotation.annotationId);
    if (annotation.status !== "needs_reanchor") {
      const anchor = annotation.anchor;
      const nodeIds = anchor.kind === "text" || anchor.kind === "node" ? [anchor.nodeId]
        : anchor.kind === "text-set" ? anchor.ranges.map((range) => range.nodeId) : anchor.nodeIds;
      if (nodeIds.some((nodeId) => !html.hasNode(nodeId))) {
        invalid("An active or resolved annotation refers to a missing document node.");
      }
    }
  }

  validateWord(metadata, annotations, html);
  const snapshot: DocumentSnapshot = freeze({ manifest, html: html.html, styles, metadata, annotations });
  files.set("manifest.json", encode(JSON.stringify(manifest, null, 2) + "\n"));
  files.set(manifest.entrypoint, encode(snapshot.html));
  files.set(manifest.metadata, encode(JSON.stringify(metadata, null, 2) + "\n"));
  files.set(manifest.annotations, encode(JSON.stringify(annotations, null, 2) + "\n"));
  for (const style of styles) files.set(style.path, encode(style.css));
  let total = 0;
  for (const bytes of files.values()) {
    checkLimit(bytes.byteLength, limits.maxFileBytes);
    total += bytes.byteLength;
  }
  checkLimit(total, limits.maxPackageBytes);

  return { files, snapshot, html };
}
