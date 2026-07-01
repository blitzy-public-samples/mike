/**
 * Accept-all normalization of a parsed `.docx` to clean "accepted-view" text.
 *
 * Before two documents can be diffed as a clean two-way comparison, any
 * pre-existing tracked changes carried by an input `.docx` must be resolved:
 * insertions (`<w:ins>`) are ACCEPTED (their inner runs become normal runs) and
 * deletions (`<w:del>`) are DISCARDED (their `<w:delText>` content is excluded).
 * The result is the plaintext a reader sees when Word/Google Docs display the
 * document in "all changes accepted" view.
 *
 * This transformation operates purely on the ordered paragraph/run model
 * produced by `./parseDocx` (which deliberately leaves `<w:ins>` / `<w:del>`
 * un-flattened); this module is where that un-flattening happens for the
 * compare engine.
 *
 * The functions here are pure: they read the parsed model, never mutate the
 * input `ParsedDocx` or its nodes, perform no I/O, and use no randomness or
 * clock, so identical inputs always yield identical output. Preserved `<w:rPr>`
 * nodes are carried by reference (the emitter clones them before use).
 *
 * @module compare/normalize
 */

import {
  type ParsedDocx,
  type ParsedRun,
  type XNode,
  elName,
  elChildren,
  getTextContent,
} from "./parseDocx";

// ---------------------------------------------------------------------------
// Public contract
// ---------------------------------------------------------------------------

/**
 * A single body paragraph after accept-all normalization: its accepted-view
 * runs (in document order, with `<w:rPr>` preserved) and their concatenated
 * plaintext. Produced 1:1 with `ParsedDocx.paragraphs`.
 */
export interface NormalizedParagraph {
  /**
   * Accepted-view runs in document order. Each entry corresponds to a direct
   * `<w:r>` of the paragraph or to a `<w:r>` unwrapped from an accepted
   * `<w:ins>`; runs inside a discarded `<w:del>` are absent. Each run's `rPr`
   * is the same `<w:rPr>` node reference carried by the parsed model.
   */
  runs: ParsedRun[];
  /**
   * Accepted-view plaintext of the paragraph: the concatenation of the run
   * texts, with NO trailing newline (mirrors `ParsedParagraph.text`).
   */
  text: string;
}

// ---------------------------------------------------------------------------
// Run reading (accepted-view text of a single <w:r>)
// ---------------------------------------------------------------------------

/**
 * Read a single `<w:r>` element into a {@link ParsedRun}. The run's `<w:rPr>`
 * (if present) is captured by reference and the text of its `<w:t>` children is
 * concatenated. Non-text run children (`<w:tab>` / `<w:br>` / `<w:sym>`) and
 * any `<w:delText>` are excluded from the text stream, mirroring how
 * `parseDocx` reads a run (`<w:t>` only) so run/text offsets stay consistent
 * across the engine.
 */
function readRun(rNode: XNode): ParsedRun {
  let rPr: XNode | null = null;
  let text = "";
  for (const child of elChildren(rNode)) {
    const name = elName(child);
    if (name === "w:rPr") {
      // Run properties are carried by reference, as in the parsed model.
      rPr = child;
    } else if (name === "w:t") {
      text += getTextContent(child);
    }
    // Other run children (including <w:delText>) contribute no accepted-view
    // text and are intentionally skipped.
  }
  return { rPr, text };
}

// ---------------------------------------------------------------------------
// Paragraph normalization (accept-all over the parsed model)
// ---------------------------------------------------------------------------

/**
 * Accept-all normalize every body paragraph of a parsed document for clean-text
 * diffing.
 *
 * For each paragraph the raw `<w:p>` children (`paragraph.node`) are walked in
 * document order and reduced to accepted-view runs:
 *
 * - a direct `<w:r>` run is kept as-is (its `<w:rPr>` and `<w:t>` text);
 * - a `<w:ins>` wrapper is UNWRAPPED — its inner `<w:r>` children are included
 *   as if they were normal runs (an accepted insertion);
 * - a `<w:del>` wrapper is DROPPED entirely — its `<w:delText>` content is
 *   excluded from the accepted view (an accepted deletion);
 * - any other element (paragraph properties, bookmarks, `<w:sdt>`, …)
 *   contributes nothing to the text stream and is skipped.
 *
 * The recursion depth into `<w:ins>` matches the reference engine's pragmatic
 * depth: only the wrapper's direct `<w:r>` children are unwrapped. Output is
 * produced 1:1 with `doc.paragraphs`, preserving document order, and the input
 * `ParsedDocx` and its nodes are never mutated.
 *
 * @param doc - the parsed document (read only; never mutated)
 * @returns one {@link NormalizedParagraph} per body paragraph, in order
 */
export function normalizeParagraphs(doc: ParsedDocx): NormalizedParagraph[] {
  const normalized: NormalizedParagraph[] = [];

  for (const paragraph of doc.paragraphs) {
    const runs: ParsedRun[] = [];
    let text = "";

    for (const child of elChildren(paragraph.node)) {
      const name = elName(child);

      if (name === "w:r") {
        // Direct run: keep it as an accepted-view run.
        const run = readRun(child);
        runs.push(run);
        text += run.text;
      } else if (name === "w:ins") {
        // Accepted insertion: unwrap the wrapper and include its inner
        // <w:r> children as if they were normal runs.
        for (const inner of elChildren(child)) {
          if (elName(inner) === "w:r") {
            const run = readRun(inner);
            runs.push(run);
            text += run.text;
          }
        }
      }
      // "w:del": discarded entirely (its <w:delText> is excluded from the
      // accepted view). All other elements contribute no text and are
      // skipped here.
    }

    normalized.push({ runs, text });
  }

  return normalized;
}

/**
 * Convenience: the whole-document accepted-view plaintext, with paragraphs
 * joined by a single `"\n"` and no trailing newline. Equivalent to
 * `normalizeParagraphs(doc).map((p) => p.text).join("\n")`. Useful as a
 * deterministic sanity value for the orchestrator.
 *
 * @param doc - the parsed document (read only; never mutated)
 * @returns the accepted-view plaintext of the whole body
 */
export function normalizedText(doc: ParsedDocx): string {
  return normalizeParagraphs(doc)
    .map((paragraph) => paragraph.text)
    .join("\n");
}
