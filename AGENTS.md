# AGENTS.md

## Scope

`dagent-office` is an independent TypeScript SDK for HTML-based Office editing and human/AI collaboration. Read [README.md](README.md) and the task-relevant design documents before implementation. The core implements structural contracts, JSON Schemas, Word semantics, a read-only HTML/CSS document package model and resource validation/collection; office implements bounded package ZIP adapters and DOCX import. Editor provides a scriptless read-only preview. Three-format model design, declared DOCX import acceptance and R1-06 core transactions/history are complete; R1-07 precise anchor and annotation mapping is next. Editing sessions support isolated candidates, content/annotation operations and bounded snapshot history. Request/proposal execution, rich-text editing, automatic pagination and Office/PDF export are not implemented yet.

## Working with the user

- Complete authorized work and meaningful verification autonomously. Existing user decisions persist; do not add approval steps for routine implementation.
- Do not use sub-agents unless the user explicitly requests them for the current task.
- Keep changes focused. Never overwrite user changes or commit to `main` without explicit instruction. Do not publish packages or push unless requested.

## Current implementation order

- Design the complete V1 model for Word, PPT and Excel first (R1-05A); implement Word in the current stage.
- Complete R1-10 DOCX reading, style mapping, read-only preview and import acceptance before R1-06 or subsequent editing work. Keep existing task IDs.
- Import preserves fields and cached results, classic comments, and text insertion/deletion revisions. Unsupported content fails explicitly; no default partial import. Export remains R1-13.
- See [data model](docs/data-model.md) and [iteration 1](docs/iteration-1.md) for detailed contracts and gates.

## Architecture and design principles

- Follow established software architecture and design principles: modularity, high cohesion, low coupling, separation of concerns, information hiding and dependency inversion.
- Clearly define module responsibilities, interface contracts, dependency directions and abstraction boundaries to support extensibility, replaceability, maintainability and architectural evolution.
- Prefer mature, simple and general-purpose designs. Avoid excessive abstraction, unnecessary indirection and overengineering; use the minimum necessary complexity to support stable, long-term evolution.

## Boundaries

- Keep SDK code independent of Enterprise and `dagent-ai`. Hosts own identity, permissions, credentials, authoritative storage, model execution and deployment.
- Accept authorized bytes and opaque references through host interfaces. Document IDs, node IDs, HTML attributes and cache keys never grant access.
- Do not read Enterprise databases or knowledge tables. A host supplies only context authorized for its current principal.
- Keep per-session document and model state isolated. Never share private document content or model contexts between host cache namespaces.
- Do not copy or translate ibisOffice source into this project. Implement from the document specification and public format standards; record licenses for reused dependencies.
- Implement DOCX/PPTX/XLSX OOXML reading, writing and style mapping in this project. Do not use docx, PptxGenJS or ExcelJS for Office conversion. Mature ZIP, XML and hashing primitives may be reused behind adapters.

## Document and editing invariants

- The persisted source is a versioned HTML document package. Runtime trees and editor DOMs must serialize to that same source without introducing competing authoritative copies.
- Put human and AI edits through the same typed transaction boundary. Validate the complete batch before committing one new revision and one undo step.
- Bind model proposals and anchors to a document revision. Reject stale proposals in V1; do not guess a semantic merge.
- Treat generated HTML as data. Reject executable scripts, handlers and unauthorized resource references.
- Keep import, editing, export and model orchestration separate. Use adapters for format libraries; never invent a second agent runtime.
- Preserve supported package readers once published; reject unknown versions without modifying their bytes. Do not rewrite immutable revisions in place.

- Chinese typography must preserve Latin/east-Asian font roles, source font names, UTF-16 text and variation sequences. Report preview substitutions and missing glyphs; pin and record fonts for visual fixtures. Do not silently download fonts. Host-authorized local font installation is separate from SDK runtime behavior; record source, version, license and hashes.

## Verification

Use the narrowest meaningful checks. Changes to editing need transaction, undo and conflict tests; Office changes need import, editable export and reopen checks; host changes need ownership and concurrent-save coverage. Layout changes need relevant visual fixtures with fixed fonts and viewport.

For design-only work, check Markdown links, JSON examples, example consistency and `git diff --check`. Do not claim planned capabilities as implemented behavior.
