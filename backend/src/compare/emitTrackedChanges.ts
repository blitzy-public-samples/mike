/**
 * Native Word tracked-changes emitter for the Document Compare / Redline engine.
 *
 * Given a parsed BASE `.docx`, the ordered paragraph alignment, and the optional
 * per-paragraph word-diff segments for modified pairs, this module merges the
 * two documents into a single `word/document.xml` carrying NATIVE Word tracked
 * changes and rezips a valid `.docx` {@link Buffer}:
 *
 *   - Insertions are wrapped in `<w:ins w:id w:author w:date>`.
 *   - Deletions are wrapped in `<w:del w:id w:author w:date>` with the run text
 *     emitted as `<w:delText>` (rather than `<w:t>`).
 *   - A deleted paragraph mark is represented by
 *     `<w:pPr><w:rPr><w:del .../></w:rPr></w:pPr>`, and an inserted paragraph
 *     mark by the matching `<w:ins .../>` inside `<w:pPr><w:rPr>`.
 *
 * Every emitted `<w:ins>` / `<w:del>` receives a monotonic, unique `w:id` that
 * starts above the largest pre-existing tracked-change id, each carries the
 * caller-supplied `w:author` / `w:date`, and each emitted run carries the source
 * `<w:rPr>` so formatting round-trips. The output is byte-deterministic: identical
 * inputs (including the same `date`) always produce byte-identical bytes, which is
 * a hard requirement of the engine (AAP 0.8.1). Determinism is achieved by taking
 * `author` / `date` from the caller (never the clock) and by pinning every zip
 * entry's timestamp and the DEFLATE level so the DOS timestamps embedded by JSZip
 * cannot drift between runs.
 *
 * Isolation (AAP 0.8.1): the OOXML emission internals are REIMPLEMENTED here and
 * this module does NOT import the private helpers of `../lib/docxTrackedChanges`;
 * it reuses only the READ-side node-model helpers exported by `./parseDocx`
 * (keeping the parse/emit node model identical) plus the `jszip` and
 * `fast-xml-parser` libraries already present in the backend. It reaches into no
 * route, persistence, assistant/chat, tabular-review, or workflow code path.
 *
 * V1 scope (AAP 0.7.2): only `<w:ins>` / `<w:del>` / `<w:delText>` and the
 * paragraph-mark ins/del are emitted. Formatting-only tracking
 * (`<w:rPrChange>` / `<w:pPrChange>`) and moves (`moveFrom` / `moveTo`) are NOT
 * emitted (a move renders as a delete + an insert); comments, footnotes,
 * headers/footers, and images are passed through unchanged.
 *
 * @module compare/emitTrackedChanges
 */

import JSZip from "jszip";
import { XMLBuilder } from "fast-xml-parser";

import {
  ATTR_KEY,
  TEXT_KEY,
  type XNode,
  elName,
  elChildren,
  setChildren,
  elAttrs,
  cloneNode,
  findBody,
  type ParsedDocx,
  type ParsedParagraph,
} from "./parseDocx";
import type { AlignOp } from "./alignParagraphs";
import type { DiffSegment } from "./diffJson";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** The WordprocessingML main namespace, ensured on the `<w:document>` root. */
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Canonical OOXML main-document part path (forward-slash form). */
const DOCUMENT_XML_PATH = "word/document.xml";

/**
 * A shared, mutable monotonic counter for tracked-change `w:id` values. Passed
 * by reference so every emit helper draws the next unique id from the same
 * sequence, guaranteeing ids are unique AND monotonic across a single emit.
 */
interface IdCounter {
  /** The next `w:id` to assign; incremented on each use. */
  next: number;
}

// ---------------------------------------------------------------------------
// Write-side node-model helpers (owned locally; mirror docxTrackedChanges.ts)
// ---------------------------------------------------------------------------
//
// The preserve-order node model is shared with `./parseDocx` (element =
// `{ [tag]: XNode[], ":@"?: { "@_attr": value } }`; text = `{ "#text": string }`).
// The READ-side helpers (elName / elChildren / setChildren / elAttrs / cloneNode
// / findBody) are imported from `./parseDocx`; the WRITE-side constructors below
// are this module's own, reimplemented to avoid importing the private internals
// of `../lib/docxTrackedChanges`.

/**
 * Build an `XMLBuilder` configured to round-trip the OOXML preserve-order model.
 * The flags EXACTLY mirror the parser used by `./parseDocx` (and the reference
 * engine) so the emitted XML stays compatible with MS Word / Google Docs. A
 * fresh instance is returned per call so callers never share builder state.
 */
function createBuilder(): XMLBuilder {
  return new XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    preserveOrder: true,
    suppressEmptyNode: false,
    processEntities: true,
  });
}

/**
 * Construct a preserve-order element node.
 *
 * @param name - the element tag (e.g. `"w:r"`, `"w:ins"`)
 * @param children - the ordered child nodes (defaults to empty)
 * @param attrs - optional attribute map with UNPREFIXED keys (e.g.
 *   `{ "w:id": "5" }`); each key is stored prefixed with `@_` under the `":@"`
 *   attribute key, matching how `fast-xml-parser` represents attributes
 * @returns the constructed element node
 */
function makeEl(
  name: string,
  children: XNode[] = [],
  attrs?: Record<string, string>,
): XNode {
  const el: XNode = { [name]: children };
  if (attrs) {
    const attrObj: Record<string, string> = {};
    for (const [k, v] of Object.entries(attrs)) {
      attrObj[`@_${k}`] = v;
    }
    el[ATTR_KEY] = attrObj;
  }
  return el;
}

/** Construct a bare text node `{ "#text": s }`. */
function makeText(s: string): XNode {
  return { [TEXT_KEY]: s };
}

/**
 * Build a `<w:r>` run element wrapping `text`.
 *
 * The run's `<w:rPr>` (if provided) is CLONED and placed first so the source
 * formatting round-trips without aliasing the parsed tree. The text is split on
 * `"\n"`: each non-empty segment becomes a `<w:t>` (or `<w:delText>` when
 * `asDelText`) carrying `xml:space="preserve"`, and each newline boundary emits
 * a `<w:br/>` soft line break — so a multi-line replacement never surfaces a
 * literal `"\n"` as visible text. Mirrors `docxTrackedChanges.buildRun`.
 *
 * @param text - the run's visible text (may contain `"\n"`)
 * @param rPr - the run properties to carry (cloned), or `null`
 * @param asDelText - `true` to emit `<w:delText>` (for deletions), `false` for `<w:t>`
 * @returns the constructed `<w:r>` element
 */
function buildRun(text: string, rPr: XNode | null, asDelText: boolean): XNode {
  const tagName = asDelText ? "w:delText" : "w:t";
  const children: XNode[] = [];
  if (rPr) {
    // Clone so the emitted run never aliases a node in the parsed tree.
    children.push(cloneNode(rPr));
  }
  const segments = text.split("\n");
  for (let i = 0; i < segments.length; i++) {
    if (i > 0) {
      children.push(makeEl("w:br", []));
    }
    const seg = segments[i];
    if (seg.length > 0) {
      children.push(makeEl(tagName, [makeText(seg)], { "xml:space": "preserve" }));
    }
  }
  return makeEl("w:r", children);
}

/**
 * Allocate the next tracked-change attribute set `{ w:id, w:author, w:date }`,
 * advancing the shared id counter. `w:date` is used VERBATIM (never the clock)
 * so the output stays deterministic.
 */
function mkChangeAttrs(
  ids: IdCounter,
  author: string,
  date: string,
): Record<string, string> {
  return {
    "w:id": String(ids.next++),
    "w:author": author,
    "w:date": date,
  };
}

/**
 * Walk a preserve-order tree and return the largest `w:id` carried by any
 * `<w:ins>` / `<w:del>` element (0 when there are none). New tracked changes
 * start their numbering at `maxTrackedId(tree) + 1` so ids never collide with
 * pre-existing ones. Mirrors `docxTrackedChanges.maxTrackedId`.
 */
function maxTrackedId(doc: XNode[]): number {
  let max = 0;
  const visit = (n: unknown): void => {
    const name = elName(n);
    if (!name) return;
    if (name === "w:ins" || name === "w:del") {
      const raw = elAttrs(n)["@_w:id"];
      if (raw != null) {
        const v = parseInt(String(raw), 10);
        if (Number.isFinite(v) && v > max) max = v;
      }
    }
    for (const c of elChildren(n as XNode)) visit(c);
  };
  for (const top of doc) visit(top);
  return max;
}

/**
 * Prepend the standard XML declaration when the built string lacks one, so the
 * emitted `word/document.xml` always opens with a proper prolog. Mirrors
 * `docxTrackedChanges.ensureXmlDeclaration`.
 */
function ensureXmlDeclaration(xml: string): string {
  if (xml.startsWith("<?xml")) return xml;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
}

/**
 * Ensure the `<w:document>` root carries the `w` namespace declaration
 * (`xmlns:w`). Standalone `<w:ins>` / `<w:del>` elements only render correctly
 * when the namespace is in scope; a valid `.docx` already declares it, but this
 * guards inputs that somehow omit it. Mirrors the reference engine's namespace
 * ensure (`W_NS_ATTRS`).
 */
function ensureWNamespace(tree: XNode[]): void {
  for (const top of tree) {
    if (elName(top) !== "w:document") continue;
    const attrs = (top[ATTR_KEY] as Record<string, string> | undefined) ?? {};
    if (attrs["@_xmlns:w"] === undefined) {
      attrs["@_xmlns:w"] = W_NS;
      top[ATTR_KEY] = attrs;
    }
  }
}

// ---------------------------------------------------------------------------
// Emit helpers: <w:ins> / <w:del> wrappers and paragraph-mark tracking
// ---------------------------------------------------------------------------

/**
 * Wrap one run (or several runs) in a tracked-insertion `<w:ins>` element with a
 * fresh `w:id` and the caller's `w:author` / `w:date`. Accepting an array lets a
 * whole inserted paragraph's runs share a single `<w:ins>` wrapper, while a
 * single word-diff insertion segment passes just its one run.
 */
function emitIns(
  run: XNode | XNode[],
  ids: IdCounter,
  author: string,
  date: string,
): XNode {
  const runs = Array.isArray(run) ? run : [run];
  return makeEl("w:ins", runs, mkChangeAttrs(ids, author, date));
}

/**
 * Wrap one or more `<w:delText>`-bearing runs in a tracked-deletion `<w:del>`
 * element with a fresh `w:id` and the caller's `w:author` / `w:date`.
 */
function emitDel(
  runs: XNode[],
  ids: IdCounter,
  author: string,
  date: string,
): XNode {
  return makeEl("w:del", runs, mkChangeAttrs(ids, author, date));
}

/**
 * Mark a paragraph's end-of-paragraph glyph (the paragraph mark) as inserted or
 * deleted by ensuring the paragraph's `<w:pPr>` contains a `<w:rPr>` whose FIRST
 * child is a `<w:ins/>` (kind `"ins"`) or `<w:del/>` (kind `"del"`) carrying a
 * fresh `w:id` / `w:author` / `w:date`.
 *
 * When `w:rPr` is a child of `w:pPr` it represents the paragraph mark's run
 * properties, and per ISO-29500 (CT_ParaRPr) its `<w:ins>` / `<w:del>` tracking
 * element comes FIRST — hence the change is prepended to any existing `<w:rPr>`
 * children. A missing `<w:pPr>` / `<w:rPr>` is created as needed. Mirrors the
 * OOXML pattern confirmed in AAP 0.2.2 (Microsoft Learn / ISO-29500):
 * a deleted paragraph mark = `<w:pPr><w:rPr><w:del .../></w:rPr></w:pPr>`.
 *
 * @param pPr - the paragraph's existing `<w:pPr>` node, or `null` to create one
 * @param kind - `"ins"` for an inserted paragraph mark, `"del"` for deleted
 * @returns the `<w:pPr>` node to use as the paragraph's first child
 */
function ensureParaMarkChange(
  pPr: XNode | null,
  kind: "ins" | "del",
  ids: IdCounter,
  author: string,
  date: string,
): XNode {
  const pPrNode = pPr ?? makeEl("w:pPr", []);
  const change = makeEl(
    kind === "ins" ? "w:ins" : "w:del",
    [],
    mkChangeAttrs(ids, author, date),
  );

  const pPrKids = elChildren(pPrNode);
  let rPr: XNode | null = null;
  for (const k of pPrKids) {
    if (elName(k) === "w:rPr") {
      rPr = k;
      break;
    }
  }

  if (rPr) {
    // Prepend the tracking element: inside a paragraph-mark rPr it must be first.
    setChildren(rPr, [change, ...elChildren(rPr)]);
  } else {
    // No paragraph-mark rPr yet: create one carrying only the tracking element.
    rPr = makeEl("w:rPr", [change]);
    setChildren(pPrNode, [...pPrKids, rPr]);
  }
  return pPrNode;
}

// ---------------------------------------------------------------------------
// Body paragraph collection (mirror parseDocx traversal order)
// ---------------------------------------------------------------------------

/**
 * A located paragraph in the cloned tree: the `<w:p>` node together with the
 * child array that directly contains it, so a sibling paragraph can be spliced
 * in next to it during insertion.
 */
interface ParaSlot {
  /** The `<w:p>` element node in the cloned tree. */
  node: XNode;
  /** The child array (body or a table cell / sdt content) that contains `node`. */
  parent: XNode[];
}

/**
 * Collect every body paragraph (`<w:p>`) in EXACTLY the same document order as
 * `parseDocx` — descending into tables (`w:tbl`/`w:tr`/`w:tc`) and
 * structured-document-tag containers (`w:sdt`/`w:sdtContent`) — recording each
 * paragraph's containing child array. Because the traversal matches `parseDocx`,
 * `collectParaSlots(...)[i]` corresponds to `base.paragraphs[i]` and hence to an
 * alignment op whose `baseIndex` is `i`.
 */
function collectParaSlots(bodyChildren: XNode[]): ParaSlot[] {
  const slots: ParaSlot[] = [];
  const walk = (nodes: XNode[]): void => {
    for (const n of nodes) {
      const name = elName(n);
      if (!name) continue;
      if (name === "w:p") {
        slots.push({ node: n, parent: nodes });
      } else if (
        name === "w:tbl" ||
        name === "w:tr" ||
        name === "w:tc" ||
        name === "w:sdt" ||
        name === "w:sdtContent"
      ) {
        walk(elChildren(n));
      }
    }
  };
  walk(bodyChildren);
  return slots;
}

/**
 * Return the paragraph's `<w:pPr>` node when it is present as the FIRST child of
 * the `<w:p>` (its canonical position), else `null`.
 */
function firstPPr(paraChildren: XNode[]): XNode | null {
  if (paraChildren.length > 0 && elName(paraChildren[0]) === "w:pPr") {
    return paraChildren[0];
  }
  return null;
}

/**
 * Return the source `<w:rPr>` of the run that covers character `offset` within a
 * paragraph's concatenated run text, so a rebuilt word-diff segment carries the
 * formatting of the exact base/revised run it originated from. Empty runs are
 * skipped; an `offset` at or past the end falls back to the last run's `<w:rPr>`.
 *
 * A segment that happens to straddle two differently-formatted runs takes the
 * `<w:rPr>` at its START — a documented V1 fidelity boundary, since the word-diff
 * operates on plaintext and does not sub-split at run boundaries.
 */
function rPrAtOffset(runs: ParsedParagraph["runs"], offset: number): XNode | null {
  let acc = 0;
  for (const run of runs) {
    const end = acc + run.text.length;
    if (offset < end) return run.rPr;
    acc = end;
  }
  return runs.length > 0 ? runs[runs.length - 1].rPr : null;
}

// ---------------------------------------------------------------------------
// Paragraph transforms: modified / deleted / inserted
// ---------------------------------------------------------------------------

/**
 * Rewrite a matched-but-modified paragraph IN PLACE from its word-diff segments.
 *
 * The paragraph's own `<w:pPr>` is preserved; its run content is rebuilt so that
 * each segment becomes: an untracked run (`equal`), a `<w:del>` of `<w:delText>`
 * (`del`), or a `<w:ins>` of `<w:t>` (`ins`). Each segment carries the `<w:rPr>`
 * of the source run it came from — the BASE run at the running base-text offset
 * for `equal` / `del` text, and the REVISED run at the running revised-text
 * offset for `ins` text — so per-run formatting round-trips faithfully (see
 * {@link rPrAtOffset} for the straddling-run fidelity boundary).
 *
 * @param pNode - the cloned base `<w:p>` node to transform in place
 * @param segments - the ordered word-diff segments for this pair
 * @param basePara - the base paragraph (source of `equal` / `del` run rPr)
 * @param revisedPara - the revised paragraph (source of `ins` run rPr)
 */
function rewriteParagraphToModified(
  pNode: XNode,
  segments: DiffSegment[],
  basePara: ParsedParagraph,
  revisedPara: ParsedParagraph,
  ids: IdCounter,
  author: string,
  date: string,
): void {
  const pPr = firstPPr(elChildren(pNode));

  const newKids: XNode[] = [];
  if (pPr) newKids.push(pPr);

  // Character cursors into the base and revised plaintext run streams, advanced
  // as segments are consumed so each carries the rPr of its originating run.
  let baseCursor = 0;
  let revisedCursor = 0;

  for (const seg of segments) {
    if (seg.type === "equal") {
      newKids.push(buildRun(seg.text, rPrAtOffset(basePara.runs, baseCursor), false));
      baseCursor += seg.text.length;
      revisedCursor += seg.text.length;
    } else if (seg.type === "del") {
      newKids.push(
        emitDel([buildRun(seg.text, rPrAtOffset(basePara.runs, baseCursor), true)], ids, author, date),
      );
      baseCursor += seg.text.length;
    } else {
      // "ins"
      newKids.push(
        emitIns(buildRun(seg.text, rPrAtOffset(revisedPara.runs, revisedCursor), false), ids, author, date),
      );
      revisedCursor += seg.text.length;
    }
  }

  setChildren(pNode, newKids);
}

/**
 * Rewrite a whole unmatched BASE paragraph IN PLACE as a tracked deletion.
 *
 * The paragraph's `<w:pPr>` is preserved and its paragraph mark is marked deleted
 * (`<w:pPr><w:rPr><w:del/></w:rPr></w:pPr>`); every base run is re-emitted as
 * `<w:delText>` — each carrying its own source `<w:rPr>` — wrapped in a single
 * `<w:del>`. An empty base paragraph yields just the deleted paragraph mark.
 *
 * @param pNode - the cloned base `<w:p>` node to transform in place
 * @param basePara - the base paragraph whose runs are converted to deletions
 */
function rewriteParagraphToDeletion(
  pNode: XNode,
  basePara: ParsedParagraph,
  ids: IdCounter,
  author: string,
  date: string,
): void {
  const pPr = ensureParaMarkChange(
    firstPPr(elChildren(pNode)),
    "del",
    ids,
    author,
    date,
  );

  const delRuns: XNode[] = [];
  for (const run of basePara.runs) {
    if (run.text.length > 0) {
      delRuns.push(buildRun(run.text, run.rPr, true));
    }
  }

  const newKids: XNode[] = [pPr];
  if (delRuns.length > 0) {
    newKids.push(emitDel(delRuns, ids, author, date));
  }
  setChildren(pNode, newKids);
}

/**
 * Build a NEW `<w:p>` node for an unmatched REVISED paragraph, emitted as a
 * tracked insertion.
 *
 * The revised paragraph's `<w:pPr>` (if any) is cloned and its paragraph mark is
 * marked inserted (`<w:pPr><w:rPr><w:ins/></w:rPr></w:pPr>`); every revised run
 * is re-emitted as `<w:t>` — each carrying its own source `<w:rPr>` — wrapped in
 * a single `<w:ins>`. An empty revised paragraph yields just the inserted
 * paragraph mark. The revised parse tree is never mutated (its `<w:pPr>` and
 * `<w:rPr>` are cloned before use).
 *
 * @param revisedPara - the revised paragraph to render as an insertion
 * @returns the freshly constructed `<w:p>` node
 */
function buildInsertedParagraph(
  revisedPara: ParsedParagraph,
  ids: IdCounter,
  author: string,
  date: string,
): XNode {
  const srcPPr = firstPPr(elChildren(revisedPara.node));
  const pPr = ensureParaMarkChange(
    srcPPr ? cloneNode(srcPPr) : null,
    "ins",
    ids,
    author,
    date,
  );

  const insRuns: XNode[] = [];
  for (const run of revisedPara.runs) {
    if (run.text.length > 0) {
      // buildRun clones run.rPr, so the revised parse tree is not aliased.
      insRuns.push(buildRun(run.text, run.rPr, false));
    }
  }

  const children: XNode[] = [pPr];
  if (insRuns.length > 0) {
    children.push(emitIns(insRuns, ids, author, date));
  }
  return makeEl("w:p", children);
}


// ---------------------------------------------------------------------------
// Deterministic rezip
// ---------------------------------------------------------------------------

/**
 * Overwrite the `word/document.xml` entry with `xml`, pinning its timestamp to
 * `date`. Handles the rare archives that store the entry under a backslash path
 * (`word\document.xml`) by writing back to that same path, so we never emit two
 * `document.xml` variants side by side. Passing an explicit `date` prevents
 * JSZip from stamping the freshly-written entry with `new Date()`.
 */
function writeDocumentXml(zip: JSZip, xml: string, date: Date): void {
  const backslash = DOCUMENT_XML_PATH.replace(/\//g, "\\");
  if (!zip.file(DOCUMENT_XML_PATH) && zip.file(backslash)) {
    zip.file(backslash, xml, { date });
    return;
  }
  zip.file(DOCUMENT_XML_PATH, xml, { date });
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Inputs to {@link emitTrackedChanges}.
 *
 * The alignment (`align`) is walked in document order; `segmentsByPair` supplies
 * the word-diff for matched-but-modified pairs, keyed by BASE paragraph index
 * (i.e. the alignment op's `baseIndex`). `author` and `date` are written verbatim
 * as the `w:author` / `w:date` of every emitted change, and `date` also pins the
 * output zip's timestamps — so identical inputs yield byte-identical output.
 */
export interface EmitInput {
  /** The parsed BASE document; a CLONE of its `tree` is the merge target. */
  base: ParsedDocx;
  /** The ordered paragraph alignment (`equal` / `del` / `ins`). */
  align: AlignOp[];
  /**
   * Word-diff segments for matched-but-MODIFIED pairs, keyed by BASE paragraph
   * index. An `equal` op WITHOUT an entry here is an identical paragraph (kept
   * as-is); an `equal` op WITH an entry is rebuilt from its segments.
   */
  segmentsByPair: Map<number, DiffSegment[]>;
  /** The parsed REVISED document; source of inserted paragraphs and their rPr. */
  revised: ParsedDocx;
  /** The `w:author` written on every emitted tracked change. */
  author: string;
  /**
   * The `w:date` (an ISO-8601 string) written verbatim on every emitted change,
   * and parsed via `new Date(date)` to pin the output zip's entry timestamps.
   * Supplying it (never the clock) is what makes the output deterministic.
   */
  date: string;
}

/**
 * Merge the BASE and REVISED documents into a single `word/document.xml` carrying
 * native Word tracked changes, and rezip a valid, byte-deterministic `.docx`.
 *
 * The base tree is cloned (never mutated) and the merge is a single in-order pass
 * over `align`:
 *
 *   - `equal` with a `segmentsByPair` entry  -> the base paragraph is rewritten
 *     in place from its word-diff segments (untracked / `<w:del>` / `<w:ins>`);
 *   - `equal` without an entry               -> the base paragraph is kept as-is;
 *   - `del`                                   -> the base paragraph is rewritten
 *     in place as a tracked deletion (runs -> `<w:del>`/`<w:delText>`, paragraph
 *     mark deleted);
 *   - `ins`                                   -> a new paragraph built from the
 *     revised paragraph is spliced in as a tracked insertion (runs -> `<w:ins>`,
 *     paragraph mark inserted), positioned immediately after the preceding
 *     emitted paragraph (or at the body start for a leading insertion).
 *
 * Non-paragraph body content (tables' surrounding structure, `<w:sectPr>`, etc.)
 * is preserved because paragraphs are transformed in place and insertions are
 * spliced next to their neighbours rather than rebuilding the body from scratch.
 *
 * @param input - see {@link EmitInput}
 * @returns the merged redline `.docx` as a Node {@link Buffer}
 * @throws Error if the base document has no `<w:body>`
 */
export async function emitTrackedChanges(input: EmitInput): Promise<Buffer> {
  const { base, revised, align, segmentsByPair, author, date } = input;

  // 1. Clone the base tree so the original parse is never mutated, and locate
  //    the body child array (also the target for leading insertions).
  const tree = cloneNode(base.tree);
  const bodyChildren = findBody(tree);
  if (!bodyChildren) {
    throw new Error("w:body missing from base document.xml");
  }

  // 2. Re-collect the cloned tree's paragraph nodes in parseDocx traversal order,
  //    so slot i lines up with base.paragraphs[i] / an op whose baseIndex is i.
  const baseSlots = collectParaSlots(bodyChildren);

  // 3. Start new tracked-change ids above the largest pre-existing one.
  const ids: IdCounter = { next: maxTrackedId(tree) + 1 };

  // 4. Single ordered pass over the alignment. `anchor` tracks the node after
  //    which the next inserted paragraph is spliced (chaining consecutive
  //    insertions); `null` means "insert at the body start".
  let anchor: { parent: XNode[]; node: XNode } | null = null;

  for (const op of align) {
    if (op.type === "equal") {
      const i = op.baseIndex as number;
      const j = op.revisedIndex as number;
      const slot = baseSlots[i];
      if (!slot) continue;
      const segments = segmentsByPair.get(i);
      if (segments) {
        // Matched-but-modified: rebuild runs from the word-diff segments.
        rewriteParagraphToModified(
          slot.node,
          segments,
          base.paragraphs[i],
          revised.paragraphs[j],
          ids,
          author,
          date,
        );
      }
      // Otherwise identical text: keep the cloned base paragraph unchanged.
      anchor = { parent: slot.parent, node: slot.node };
    } else if (op.type === "del") {
      const i = op.baseIndex as number;
      const slot = baseSlots[i];
      if (!slot) continue;
      rewriteParagraphToDeletion(slot.node, base.paragraphs[i], ids, author, date);
      anchor = { parent: slot.parent, node: slot.node };
    } else {
      // "ins": splice a freshly built inserted paragraph next to its neighbour.
      const j = op.revisedIndex as number;
      const revisedPara = revised.paragraphs[j];
      if (!revisedPara) continue;
      const newPara = buildInsertedParagraph(revisedPara, ids, author, date);
      if (anchor) {
        const idx = anchor.parent.indexOf(anchor.node);
        if (idx >= 0) {
          anchor.parent.splice(idx + 1, 0, newPara);
        } else {
          anchor.parent.push(newPara);
        }
        anchor = { parent: anchor.parent, node: newPara };
      } else {
        // Leading insertion: place before the first existing body child.
        bodyChildren.unshift(newPara);
        anchor = { parent: bodyChildren, node: newPara };
      }
    }
  }

  // 5. Ensure the document root declares the WordprocessingML namespace.
  ensureWNamespace(tree);

  // Serialize the merged tree back to XML.
  const xml = ensureXmlDeclaration(createBuilder().build(tree));

  // Deterministic rezip: reuse the base archive (preserving styles/rels/etc.),
  // then normalize entry timestamps + DEFLATE level so the DOS timestamps JSZip
  // embeds cannot drift between runs (see AAP 0.8.1 / decision log).
  const zip = base.zip;
  const fixedDate = new Date(date);
  // Write the merged document.xml FIRST: adding it makes JSZip materialize any
  // missing parent-folder entry (e.g. `word/`) as a side effect, so that folder
  // must exist before we normalize timestamps below.
  writeDocumentXml(zip, xml, fixedDate);
  // Normalize EVERY entry's date -- INCLUDING directory entries -- for
  // deterministic output. A prior `if (!file.dir)` guard skipped directories,
  // leaving the implicit `word/` folder (materialized by the add above with
  // `new Date()`) stamped with the wall clock; that made the redline bytes drift
  // between otherwise-identical runs, violating the byte-determinism mandate.
  zip.forEach((_path, file) => {
    file.date = fixedDate;
  });

  const out = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    // Pin the DEFLATE level so the compressed stream is stable across runs.
    compressionOptions: { level: 6 },
  });
  return out as Buffer;
}

