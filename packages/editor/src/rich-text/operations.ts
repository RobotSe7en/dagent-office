import { blockSelector, caretAt, inputStep, offsetAt, sourceText } from "./dom.js";
import type { IgnoreNode } from "./dom.js";
/** Shared native DOM operations; the source transaction remains the commit boundary. */
export function deleteNativeRange(range: Range): void {
    const blockAt = (node: Node): Element | null => (node.nodeType === 1 ? node as Element : node.parentElement)?.closest(blockSelector) ?? null;
    const first = blockAt(range.startContainer), last = blockAt(range.endContainer), join = first && last && first !== last && first.parentElement === last.parentElement && !first.matches("td,th,li") && !last.matches("td,th,li");
    range.deleteContents();
    if (join) {
        const offset = first.childNodes.length;
        first.append(...last.childNodes);
        last.remove();
        range.setStart(first, offset);
        range.collapse(true);
    }
}
export function nativeDeletionRange(root: HTMLElement, range: Range, kind: string, ignored: IgnoreNode): Range {
    const frame = root.ownerDocument, step = inputStep(root, range, ignored);
    if (step.from !== step.to)
        return range;
    const blockAt = (node: Node): HTMLElement | null => (node.nodeType === 1 ? node as Element : node.parentElement)?.closest<HTMLElement>(blockSelector) ?? null;
    const block = blockAt(range.startContainer);
    if (!block || block !== blockAt(range.endContainer))
        return range;
    const text = sourceText(block, ignored), offset = offsetAt(block, range.startContainer, range.startOffset, ignored), backward = kind.endsWith("Backward");
    if (backward && offset === 0 || !backward && offset === text.length) {
        let sibling = backward ? block.previousElementSibling : block.nextElementSibling;
        while (sibling && ignored(sibling))
            sibling = backward ? sibling.previousElementSibling : sibling.nextElementSibling;
        if (!block.matches("td,th,li") && sibling?.matches("p,h1,h2,h3,h4,h5,h6")) {
            const joined = frame.createRange(), first = backward ? sibling : block, last = backward ? block : sibling;
            joined.setStart(...caretAt(first, sourceText(first, ignored).length, ignored));
            joined.setEnd(...caretAt(last, 0, ignored, "forward"));
            return joined;
        }
        return range;
    }
    let from = offset, to = offset;
    if (kind.includes("Word")) {
        if (backward)
            from -= /((?:\S+\s*)|\s+)$/.exec(text.slice(0, offset))?.[0].length ?? 0;
        else
            to += /^(?:\s*\S+|\s+)/.exec(text.slice(offset))?.[0].length ?? 0;
    }
    else {
        const segmenter = typeof Intl.Segmenter === "function" ? new Intl.Segmenter(undefined, { granularity: "grapheme" }) : undefined;
        if (segmenter) {
            const part = segmenter.segment(text).containing(backward ? Math.max(0, offset - 1) : offset);
            if (part) {
                if (backward)
                    from = part.index;
                else
                    to = part.index + part.segment.length;
            }
        }
        else {
            const boundaries = [0];
            let position = 0, previous = "", regional = 0;
            for (const character of text) {
                const code = character.codePointAt(0)!, isRegional = code >= 0x1f1e6 && code <= 0x1f1ff;
                if (position && !/[\p{Mark}\uFE00-\uFE0F\u{E0100}-\u{E01EF}\u200D]/u.test(character) && previous !== "\u200d" && !(code >= 0x1f3fb && code <= 0x1f3ff) && !(isRegional && regional % 2 === 1))
                    boundaries.push(position);
                regional = isRegional ? regional + 1 : 0;
                previous = character;
                position += character.length;
            }
            boundaries.push(text.length);
            if (backward)
                from = boundaries.filter(value => value < offset).at(-1) ?? 0;
            else
                to = boundaries.find(value => value > offset) ?? text.length;
        }
    }
    const expanded = frame.createRange();
    expanded.setStart(...caretAt(block, from, ignored));
    expanded.setEnd(...caretAt(block, to, ignored, "forward"));
    return expanded;
}
export function styleNativeRange(root: HTMLElement, range: Range, properties: ReadonlyMap<string, string>, ignored: IgnoreNode): void {
    const frame = root.ownerDocument, walker = frame.createTreeWalker(root, 4), ranges: Range[] = [];
    let node: Node | null;
    while ((node = walker.nextNode())) {
        if (ignored(node.parentNode!) || !range.intersectsNode(node))
            continue;
        const part = frame.createRange();
        part.selectNodeContents(node);
        if (range.compareBoundaryPoints(0, part) > 0)
            part.setStart(range.startContainer, range.startOffset);
        if (range.compareBoundaryPoints(2, part) < 0)
            part.setEnd(range.endContainer, range.endOffset);
        if (!part.collapsed)
            ranges.push(part);
    }
    for (const part of ranges.reverse()) {
        const span = frame.createElement("span");
        for (const [property, value] of properties)
            span.style.setProperty(property, value);
        part.surroundContents(span);
    }
}
