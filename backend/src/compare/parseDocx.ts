/**
 * DOCX -> ordered paragraph/run model + shared read-side OOXML helpers.
 *
 * Loads a `.docx`, reads `word/document.xml`, and parses it with
 * `fast-xml-parser` in `preserveOrder` mode to produce an ordered list of body
 * paragraphs -- each carrying its text-bearing runs and their preserved
 * `<w:rPr>` -- while keeping the full parsed preserve-order tree and the loaded
 * archive available for later tracked-changes emission.
 *
 * This module is the OWNER of the shared OOXML node-model type (`XNode`) and
 * the read-side primitives (`createParser`, `elName`, `isTextNode`,
 * `elChildren`, `setChildren`, `elAttrs`, `getTextContent`, `cloneNode`,
 * `findBody`). They are exported here for reuse by `normalize.ts` and
 * `emitTrackedChanges.ts`. They mirror -- but deliberately do NOT import -- the
 * private helpers in `../lib/docxTrackedChanges` (those are not exported). The
 * only public symbol imported from that module is `extractDocxBodyText`, used
 * to compute a whole-document plaintext sanity value.
 *
 * Parsing is deterministic: original order and `<w:rPr>` are preserved exactly;
 * nothing is reordered, sorted, or mutated during parse, and the input bytes
 * are never modified. Accept-all normalization of pre-existing tracked changes
 * (unwrapping `<w:ins>`, dropping `<w:del>`) is intentionally NOT performed here
 * -- it is the single, separately-testable responsibility of `normalize.ts`.
 * Accordingly, `parseDocx` reads each paragraph's DIRECT `<w:r>` runs as-is.
 */

import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { extractDocxBodyText } from "../lib/docxTrackedChanges";

// ---------------------------------------------------------------------------
// XML parser configuration
// ---------------------------------------------------------------------------
//
// These flags MIRROR the parser used by the existing tracked-changes engine
// (`../lib/docxTrackedChanges`) exactly, so documents parsed here and XML
// emitted later stay compatible with MS Word / Google Docs:
//   - preserveOrder: true       -> ordered node model (children are arrays)
//   - trimValues: false         -> whitespace is preserved (diff fidelity)
//   - attributeNamePrefix "@_"  -> attributes live under the ":@" key as "@_*"
//   - ignoreAttributes: false   -> attributes are retained, not dropped
//   - parseAttributeValue: false -> attribute values stay strings (no coercion)
//   - processEntities: true     -> entities round-trip on parse/emit
// Do not change any flag: the emitter relies on this exact shape.

/** Key under which `fast-xml-parser` stores an element's attribute map. */
export const ATTR_KEY = ":@";
/** Key under which `fast-xml-parser` stores a bare text value. */
export const TEXT_KEY = "#text";

/**
 * Build an `XMLParser` configured for OOXML preserve-order parsing. A fresh
 * instance is returned per call so callers never share parser state.
 */
export function createParser(): XMLParser {
    return new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: "@_",
        preserveOrder: true,
        trimValues: false,
        parseAttributeValue: false,
        processEntities: true,
    });
}

// ---------------------------------------------------------------------------
// Preserve-order node model + read-side helpers
// ---------------------------------------------------------------------------
//
// In `preserveOrder` mode every element is an object whose single own element
// key is the tag name; that key maps to an array of child nodes. Attributes
// (if any) live under the ":@" key, and text nodes are bare objects of the
// form `{ "#text": "..." }`. These helpers are reimplemented here (NOT imported
// from `../lib/docxTrackedChanges`, whose copies are private) so this module
// can OWN them for the rest of the compare engine.

/** A single preserve-order XML node (one tag key + optional ":@"/"#text"). */
export type XNode = Record<string, unknown>;

/**
 * Return the element's tag name -- the first own key that is not the attribute
 * (":@") or text ("#text") key -- or `null` for bare text nodes / non-objects.
 */
export function elName(n: unknown): string | null {
    if (!n || typeof n !== "object") return null;
    for (const k of Object.keys(n as XNode)) {
        if (k === ATTR_KEY || k === TEXT_KEY) continue;
        return k;
    }
    return null;
}

/** Type guard: a bare text node `{ "#text": string }` carrying no element tag. */
export function isTextNode(n: unknown): n is { "#text": string } {
    if (!n || typeof n !== "object") return false;
    const obj = n as XNode;
    return TEXT_KEY in obj && elName(n) === null;
}

/** Return the child-node array stored under the element's tag key (or `[]`). */
export function elChildren(n: unknown): XNode[] {
    const name = elName(n);
    if (!name) return [];
    const v = (n as XNode)[name];
    return Array.isArray(v) ? (v as XNode[]) : [];
}

/** Replace the element's child array in place. No-op for non-element nodes. */
export function setChildren(n: XNode, children: XNode[]): void {
    const name = elName(n);
    if (!name) return;
    n[name] = children;
}

/** Return the element's attribute map (object under ":@"), or `{}` if none. */
export function elAttrs(n: unknown): Record<string, string> {
    if (!n || typeof n !== "object") return {};
    const a = (n as XNode)[ATTR_KEY];
    return (a as Record<string, string>) ?? {};
}

/**
 * Concatenate the `#text` children of an element (e.g. a `<w:t>`). Values are
 * coerced with `String(...)` so any numeric-looking text the parser may have
 * surfaced as a number is normalized back to a string.
 */
export function getTextContent(wtEl: XNode): string {
    const kids = elChildren(wtEl);
    let out = "";
    for (const k of kids) {
        if (isTextNode(k)) out += String(k[TEXT_KEY] ?? "");
    }
    return out;
}

/**
 * Deep-clone a node via a JSON round-trip. Preserve-order nodes are plain JSON
 * (objects, arrays, strings), so this is a safe and deterministic clone.
 */
export function cloneNode<T>(n: T): T {
    return JSON.parse(JSON.stringify(n)) as T;
}

/**
 * Locate the `<w:body>` children array by walking the top-level nodes ->
 * `w:document` -> `w:body`. Returns the body's child array, or `null` when the
 * document has no body element.
 */
export function findBody(doc: XNode[]): XNode[] | null {
    for (const top of doc) {
        if (elName(top) === "w:document") {
            for (const c of elChildren(top)) {
                if (elName(c) === "w:body") return elChildren(c);
            }
        }
    }
    return null;
}

// ---------------------------------------------------------------------------
// Parsed paragraph / run model (this module's primary contract)
// ---------------------------------------------------------------------------

/**
 * A single run. One `ParsedRun` is produced per DIRECT `<w:r>` child of a
 * paragraph; its `text` may be empty for runs that carry only non-text content
 * (e.g. `<w:br>` / `<w:tab>` / `<w:sym>`), which keeps the model faithful to
 * the document as-is rather than silently dropping such runs.
 */
export interface ParsedRun {
    /**
     * The run's `<w:rPr>` element node, or `null` when the run has no
     * properties. Stored as a REFERENCE into the parsed tree (not a clone):
     * the emitter is expected to `cloneNode(rPr)` before attaching it to an
     * emitted run, so sharing the reference here is safe and avoids copying.
     */
    rPr: XNode | null;
    /** Concatenated visible text of the run's `<w:t>` children. */
    text: string;
}

/**
 * A single body paragraph (`<w:p>`), its ordered text-bearing runs, and its
 * concatenated plaintext.
 */
export interface ParsedParagraph {
    /** Ordered runs (one entry per direct `<w:r>` child), with preserved `<w:rPr>`. */
    runs: ParsedRun[];
    /** Concatenation of the run texts = the paragraph plaintext (NO trailing newline). */
    text: string;
    /** The underlying `<w:p>` element node, retained for tracked-changes emission. */
    node: XNode;
}

/**
 * The fully parsed document: ordered paragraphs plus the raw tree and loaded
 * archive needed to rebuild and rezip a valid `.docx` downstream.
 */
export interface ParsedDocx {
    /** Body paragraphs in document order (including those nested in tables / `w:sdt`). */
    paragraphs: ParsedParagraph[];
    /** The full parsed preserve-order tree (consumed by `emitTrackedChanges.ts`). */
    tree: XNode[];
    /** The loaded archive, so the emitter can rezip the SAME archive (preserving styles/rels/etc.). */
    zip: JSZip;
    /** Result of `extractDocxBodyText(bytes)` -- whole-document sanity plaintext. */
    bodyPlainText: string;
}

// ---------------------------------------------------------------------------
// Zip entry resolution
// ---------------------------------------------------------------------------
//
// Some older Windows/Word archives store entries with backslash separators
// (e.g. `word\document.xml`) even though the zip spec requires forward slashes.
// JSZip looks up entries by exact string, so we accept the canonical
// forward-slash form and transparently fall back to the backslash variant.
// Mirrors the read-side fallback in the existing engine.

function getZipEntry(zip: JSZip, pathSlash: string): JSZip.JSZipObject | null {
    const direct = zip.file(pathSlash);
    if (direct) return direct;
    return zip.file(pathSlash.replace(/\//g, "\\"));
}

// ---------------------------------------------------------------------------
// Paragraph construction
// ---------------------------------------------------------------------------

/**
 * Build a `ParsedParagraph` from a `<w:p>` node by reading its DIRECT `<w:r>`
 * children only. `<w:ins>` / `<w:del>` wrappers are intentionally NOT descended
 * into here -- accept-all normalization is the job of `normalize.ts`, which
 * keeps that transformation a single, independently-testable step. For each
 * run, the `<w:rPr>` is captured (by reference) and the text of its `<w:t>`
 * children is concatenated; non-text run children (`<w:tab>` / `<w:br>` /
 * `<w:sym>`) are ignored for the text stream, mirroring the reference engine's
 * `flattenParagraph`, which counts only `<w:t>`.
 */
function buildParagraph(pNode: XNode): ParsedParagraph {
    const runs: ParsedRun[] = [];
    let text = "";
    for (const child of elChildren(pNode)) {
        if (elName(child) !== "w:r") continue;
        let rPr: XNode | null = null;
        let runText = "";
        for (const rk of elChildren(child)) {
            const name = elName(rk);
            if (name === "w:rPr") {
                rPr = rk;
            } else if (name === "w:t") {
                runText += getTextContent(rk);
            }
            // Other run children (w:tab, w:br, w:sym, ...) are not part of the
            // text stream; the underlying w:p node is retained for emission.
        }
        runs.push({ rPr, text: runText });
        text += runText;
    }
    return { runs, text, node: pNode };
}

// ---------------------------------------------------------------------------
// Resource guardrails (deterministic fail-fast limits)
// ---------------------------------------------------------------------------
//
// These fixed bounds cap the uncompressed main-part size and the paragraph/run
// counts a single parse will accept; exceeding any bound throws a controlled
// error.
// (Thresholds and rationale: see docs/decisions/document-compare-decision-log.md.)

/** Max uncompressed size (bytes) of `word/document.xml` accepted for parsing. */
const MAX_DOCUMENT_XML_BYTES = 100 * 1024 * 1024;

/** Max body paragraphs accepted from a single document. */
const MAX_PARAGRAPHS = 50_000;

/** Max total text runs accepted across all body paragraphs. */
const MAX_RUNS = 500_000;

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * Parse a `.docx` buffer into an ordered paragraph/run model.
 *
 * Walks the document body in render order, descending into tables
 * (`w:tbl`/`w:tr`/`w:tc`) and structured-document-tag containers
 * (`w:sdt`/`w:sdtContent`) to collect every `<w:p>` -- matching the reference
 * engine's paragraph collection.
 *
 * @param bytes - the raw `.docx` archive bytes (never mutated)
 * @returns the parsed paragraphs, the raw preserve-order tree, the loaded
 *          archive, and a whole-document plaintext sanity value
 * @throws Error if `word/document.xml` or `<w:body>` is missing (a valid
 *         `.docx` input is required; a parse failure here is fatal)
 */
export async function parseDocx(bytes: Buffer): Promise<ParsedDocx> {
    const zip = await JSZip.loadAsync(bytes);

    const docXmlFile = getZipEntry(zip, "word/document.xml");
    if (!docXmlFile) throw new Error("document.xml missing from docx");
    const xml = await docXmlFile.async("string");

    // Guardrail: bound the uncompressed main-part size before parsing it.
    const xmlBytes = Buffer.byteLength(xml, "utf8");
    if (xmlBytes > MAX_DOCUMENT_XML_BYTES) {
        throw new Error(
            `compare: word/document.xml size ${xmlBytes} bytes exceeds limit ` +
                `(max=${MAX_DOCUMENT_XML_BYTES})`,
        );
    }

    const tree = createParser().parse(xml) as XNode[];

    const body = findBody(tree);
    if (!body) throw new Error("w:body missing from document.xml");

    const paragraphs: ParsedParagraph[] = [];
    const collectParagraphs = (nodes: XNode[]): void => {
        for (const n of nodes) {
            const name = elName(n);
            if (!name) continue;
            if (name === "w:p") {
                paragraphs.push(buildParagraph(n));
            } else if (
                name === "w:tbl" ||
                name === "w:tr" ||
                name === "w:tc" ||
                name === "w:sdt" ||
                name === "w:sdtContent"
            ) {
                collectParagraphs(elChildren(n));
            }
        }
    };
    collectParagraphs(body);

    // Guardrail: bound paragraph and run counts before the document flows into
    // the O(n*m) alignment.
    if (paragraphs.length > MAX_PARAGRAPHS) {
        throw new Error(
            `compare: paragraph count ${paragraphs.length} exceeds limit ` +
                `(max=${MAX_PARAGRAPHS})`,
        );
    }
    let totalRuns = 0;
    for (const paragraph of paragraphs) totalRuns += paragraph.runs.length;
    if (totalRuns > MAX_RUNS) {
        throw new Error(
            `compare: run count ${totalRuns} exceeds limit (max=${MAX_RUNS})`,
        );
    }

    // Whole-document plaintext via the single public helper we reuse. Kept for
    // sanity/debugging; the alignment path operates on `paragraph.text`.
    const bodyPlainText = await extractDocxBodyText(bytes);

    return { paragraphs, tree, zip, bodyPlainText };
}
