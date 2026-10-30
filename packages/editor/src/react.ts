import { createElement, useEffect, useRef } from "react";
import type { ReactElement } from "react";
import type { DocumentSession } from "@dagent-office/core";
import { mountDocumentWorkbench } from "./workbench.js";
import type { DocumentWorkbenchOptions, DocumentWorkbenchView } from "./workbench.js";

export interface DocumentWorkbenchProps extends DocumentWorkbenchOptions {
  readonly session: DocumentSession;
  readonly className?: string;
  readonly onViewChange?: (view: DocumentWorkbenchView | undefined) => void;
}

/** Optional React entry; ordinary JavaScript consumers need no React installation. */
export function DocumentWorkbench(props: DocumentWorkbenchProps): ReactElement {
  const container = useRef<HTMLDivElement>(null), latest = useRef(props);
  latest.current = props;
  const hasSubmission = Boolean(props.onSubmitAnnotations);
  useEffect(() => {
    if (!container.current) return;
    const view = mountDocumentWorkbench(container.current, props.session, {
      ...(props.editable !== undefined ? { editable: props.editable } : {}),
      ...(props.title ? { title: props.title } : {}),
      ...(props.supportedOperations ? { supportedOperations: props.supportedOperations } : {}),
      ...(props.contextLimits ? { contextLimits: props.contextLimits } : {}),
      ...(hasSubmission ? { onSubmitAnnotations: (request) => latest.current.onSubmitAnnotations?.(request) } : {}),
      onSelectionChange: (selection) => latest.current.onSelectionChange?.(selection),
      onError: (error) => latest.current.onError?.(error),
      onEvent: (event) => latest.current.onEvent?.(event),
    });
    latest.current.onViewChange?.(view);
    let active = true;
    void view.ready.catch((error: unknown) => { if (active) latest.current.onError?.(error instanceof Error ? error : new Error("Preview failed.")); });
    return () => { active = false; view.dispose(); latest.current.onViewChange?.(undefined); };
  }, [props.session, props.title, props.supportedOperations, props.contextLimits, props.editable, hasSubmission]);
  return createElement("div", { ref: container, className: props.className });
}
