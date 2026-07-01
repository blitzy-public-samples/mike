/**
 * Deterministic Document Compare / Redline engine — public orchestrator.
 *
 * {@link runComparison} is the SINGLE public entrypoint of the compare engine
 * and the only symbol `backend/src/routes/comparisons.ts` imports from this
 * module (alongside the re-exported structured-diff contract types). Given two
 * `.docx` buffers (a BASE and a REVISED version) it deterministically computes
 * their differences and produces two artifacts:
 *
 *   1. `redlineBytes` — a valid `.docx` {@link Buffer} carrying NATIVE Word
 *      tracked changes (`<w:ins>` / `<w:del>` with the caller's `w:author` /
 *      `w:date`) that opens cleanly in MS Word and Google Docs; and
 *   2. `diff` — the ordered-hunks {@link DiffJson} that powers the in-app inline
 *      and side-by-side redline renderers and is persisted to R2 as `diff.json`.
 *
 * This file is PURE ORCHESTRATION: every non-trivial step lives in a focused,
 * independently-testable sibling module and is invoked here in a fixed order.
 * That keeps the engine within its mandated 9-file layout and satisfies the
 * strict-isolation rule (AAP §0.8.1) — nothing here reaches into the assistant /
 * chat, document-processing, tabular-review, or workflow code paths, and the
 * only imports are the sibling engine modules.
 *
 * Determinism (AAP §0.8.1 — non-negotiable): there is NO AI in the diff path and
 * NO source of nondeterminism in this file — no clock read (`Date.now()` /
 * `new Date()`), no randomness, and no iteration-order-dependent container.
 * Every step it calls is itself deterministic (stable parse, deterministic
 * accept-all normalization, LCS alignment with a fixed tie-break, `fast-diff`
 * word diff, and an emitter whose only time source is the caller-supplied
 * `opts.date`). Consequently `runComparison(b, r, opts)` invoked twice with the
 * same buffers and the same `opts` returns byte-identical `redlineBytes` and a
 * deep-equal `diff`. `opts.date` (threaded verbatim into the emitter, which also
 * uses it to pin the output zip's entry timestamps) is the sole time source and
 * is what makes the redline byte-identical across runs.
 *
 * @module compare
 */

import { parseDocx } from "./parseDocx";
import { normalizeParagraphs } from "./normalize";
import { alignParagraphs } from "./alignParagraphs";
import { wordDiff } from "./wordDiff";
import { emitTrackedChanges } from "./emitTrackedChanges";
import { buildDiffJson } from "./diffJson";
import { validate } from "./validate";

// Structured-diff contract types. `DiffSegment` and `DiffJson` and
// `DiffHunkType` are consumed locally below; all four public contract types are
// re-exported (see below) so the route can import them from the engine root.
import type {
  DiffSegment,
  DiffJson,
  DiffHunk,
  DiffHunkType,
  DiffRange,
} from "./diffJson";

// Re-export the structured-diff contract from the engine root so the comparisons
// route (and any other consumer) can import the value and its types together:
//   import { runComparison, DiffJson, DiffHunk } from "../compare";
// `DiffSegment` is an internal builder detail and is intentionally NOT re-exported.
export type { DiffJson, DiffHunk, DiffHunkType, DiffRange };

// ---------------------------------------------------------------------------
// Diff-JSON segment stream: paragraph-mark separator convention
// ---------------------------------------------------------------------------
//
// The document-order segment stream fed to `buildDiffJson` must reconstruct the
// two column plaintexts EXACTLY, because `buildDiffJson` assigns each hunk's
// per-column character offsets from the running length of that stream and the
// frontend indexes the base/revised columns with those offsets. The base and
// revised plaintexts are the paragraph texts joined by a single "\n" — the same
// join `extractDocxBodyText` (reused by `parseDocx`) and `normalizedText` use —
// so the builder's reconstruction invariant for the WHOLE document is:
//
//   concat(text of segments whose type is `equal` or `del`) === base  plaintext
//   concat(text of segments whose type is `equal` or `ins`) === revised plaintext
//
// A paragraph mark (the inter-paragraph "\n") therefore belongs to the BASE
// column only between two consecutive BASE paragraphs, and to the REVISED column
// only between two consecutive REVISED paragraphs. Because unmatched paragraphs
// (pure `del` / `ins` ops) make those boundaries differ between the two columns,
// a single untyped `{ equal, "\n" }` separator would inject a spurious "\n" into
// the wrong column and break the reconstruction invariant. The separator is thus
// TYPED per column: emitted for BASE only when this op is not the last BASE
// paragraph, and for REVISED only when it is not the last REVISED paragraph —
// `equal` when both, `del` when base-only, `ins` when revised-only, and omitted
// after the final paragraph of each column (no trailing separator). This is a
// correctness-preserving refinement of the AAP's simplified "{equal,\n} between
// paragraphs" note, adopted specifically to honor its explicit "base/revised
// offsets reconstruct exactly" requirement (AAP §0.6.2, key insights).
/** The inter-paragraph mark that matches the engine's `"\n"` paragraph join. */
const PARAGRAPH_SEPARATOR = "\n";

/**
 * Append `text` (classified as `type`) to a document-order segment stream,
 * coalescing it into the trailing segment when that segment shares the same
 * `type`, and dropping zero-length text.
 *
 * Coalescing keeps the emitted {@link DiffJson} compact (adjacent same-type
 * runs — e.g. a paragraph's text and the following paragraph mark, or a stretch
 * of unchanged paragraphs — collapse into one hunk) and mirrors the safety net
 * `wordDiff` applies internally, so both code paths behave uniformly.
 *
 * Coalescing REPLACES the trailing element with a NEW object rather than
 * mutating it in place. This matters: the word-diff segments for a modified
 * paragraph are pushed into BOTH this stream and `segmentsByPair` (which the
 * emitter consumes); pushing only the primitive `type`/`text` here and never
 * mutating an existing object guarantees the two consumers can never alias or
 * corrupt each other.
 *
 * @param segments - the target document-order stream (mutated by append)
 * @param type - the segment classification (`equal` / `del` / `ins`)
 * @param text - the verbatim text to append (ignored when empty)
 */
function pushSegment(
  segments: DiffSegment[],
  type: DiffHunkType,
  text: string,
): void {
  if (text.length === 0) {
    // Zero-length text contributes no column content and no offset movement.
    return;
  }
  const lastIndex = segments.length - 1;
  if (lastIndex >= 0 && segments[lastIndex].type === type) {
    // Coalesce by replacing the trailing element with a fresh object; never
    // mutate in place (the source object may be shared with `segmentsByPair`).
    segments[lastIndex] = { type, text: segments[lastIndex].text + text };
    return;
  }
  segments.push({ type, text });
}

// ---------------------------------------------------------------------------
// Resource guardrail (deterministic fail-fast limit on input size)
// ---------------------------------------------------------------------------

/** Max accepted size (bytes) of a single input `.docx` buffer. */
const MAX_INPUT_DOCX_BYTES = 50 * 1024 * 1024;

/**
 * Run the deterministic compare pipeline end-to-end.
 *
 * Pipeline (fixed order; every step deterministic):
 *  1. Parse both `.docx` inputs into the ordered paragraph/run model.
 *  2. Accept-all normalize EACH input, so the diff runs against clean
 *     "accepted view" text even when an input already carries tracked changes.
 *  3. Align the normalized paragraph lists (LCS) into ordered `equal`/`del`/`ins`
 *     operations.
 *  4. Walk the alignment once, producing (a) `segmentsByPair` — the word-level
 *     diff for each matched-but-MODIFIED paragraph, keyed by BASE paragraph
 *     index (consumed by the emitter), and (b) a document-order segment stream
 *     with typed paragraph-mark separators (consumed by the diff-JSON builder).
 *  5. Emit the redline `.docx` (native tracked changes) from the base tree,
 *     the alignment, and the per-pair word diffs.
 *  6. Build the structured {@link DiffJson} from the document-order stream.
 *  7. Validate the emitted redline (fatal reparse check) BEFORE returning.
 *  8. Return the redline bytes and the structured diff.
 *
 * @param baseBytes - the BASE `.docx` archive bytes (never mutated)
 * @param revisedBytes - the REVISED `.docx` archive bytes (never mutated)
 * @param opts.author - the `w:author` written on every emitted tracked change
 * @param opts.date - the `w:date` (ISO-8601 string) written verbatim on every
 *   emitted change AND used to pin the output zip's timestamps; supplying it
 *   (never the clock) is what makes repeated runs byte-identical
 * @returns the redline `.docx` {@link Buffer} and the structured {@link DiffJson}
 * @throws Error if either input is not a valid `.docx` (missing
 *   `word/document.xml` / `<w:body>`), if the emitted redline fails the fatal
 *   reparse check, or if `opts.date` is empty
 */
export async function runComparison(
  baseBytes: Buffer,
  revisedBytes: Buffer,
  opts: { author: string; date: string },
): Promise<{ redlineBytes: Buffer; diff: DiffJson }> {
  // Reject an empty `date` at this public boundary: the emitter derives the
  // output zip's entry timestamps from it (via `new Date(date)`), so the caller
  // must supply it. `author` may legitimately be empty.
  // (Determinism rationale: see docs/decisions/document-compare-decision-log.md.)
  if (opts.date.length === 0) {
    throw new Error(
      "runComparison: opts.date must be a non-empty ISO-8601 date string",
    );
  }

  // Guardrail: reject oversized inputs before any parsing/allocation.
  if (
    baseBytes.length > MAX_INPUT_DOCX_BYTES ||
    revisedBytes.length > MAX_INPUT_DOCX_BYTES
  ) {
    throw new Error(
      `compare: input .docx size exceeds limit (base=${baseBytes.length}, ` +
        `revised=${revisedBytes.length}, max=${MAX_INPUT_DOCX_BYTES} bytes)`,
    );
  }

  // Step 1 — Parse both inputs into the ordered paragraph/run model. `parseDocx`
  // throws on a non-`.docx` / structurally invalid input, which propagates.
  const base = await parseDocx(baseBytes);
  const revised = await parseDocx(revisedBytes);

  // Step 2 — Accept-all normalization on EACH input: pre-existing `<w:ins>` are
  // accepted and `<w:del>` are dropped, so the alignment/diff operate on the
  // clean text a reader sees in "all changes accepted" view (AAP §0.1.1 step 2).
  const baseParas = normalizeParagraphs(base);
  const revisedParas = normalizeParagraphs(revised);

  // Step 3 — Deterministic paragraph-level alignment (LCS, document order).
  const align = alignParagraphs(baseParas, revisedParas);

  // Step 4 — Single ordered walk of the alignment, building both downstream
  // structures at once so the word diff for a modified pair is computed exactly
  // once and shared:
  //   * `segmentsByPair` — keyed by BASE paragraph index; only matched-but-
  //     MODIFIED pairs get an entry (identical pairs pass through untouched).
  //     Consumed by the emitter to rebuild those paragraphs as tracked changes.
  //   * `documentOrderSegments` — the linear typed-segment stream (with typed
  //     paragraph-mark separators) consumed by `buildDiffJson`.
  const segmentsByPair = new Map<number, DiffSegment[]>();
  const documentOrderSegments: DiffSegment[] = [];

  // The last paragraph index in each column: a paragraph mark is emitted for a
  // column only when the current op is NOT that column's final paragraph (so no
  // trailing "\n" is produced). Empty columns yield -1, which no index matches.
  const lastBaseIndex = baseParas.length - 1;
  const lastRevisedIndex = revisedParas.length - 1;

  for (const op of align) {
    // 4a — Paragraph content for this op.
    if (op.type === "equal") {
      // Matched pair: both indices are non-null for an `equal` op.
      const i = op.baseIndex as number;
      const j = op.revisedIndex as number;
      const baseText = baseParas[i].text;
      const revisedText = revisedParas[j].text;

      if (baseText === revisedText) {
        // Identical paragraph: it passes through as unchanged text; no entry in
        // `segmentsByPair` (the emitter keeps the base paragraph as-is).
        pushSegment(documentOrderSegments, "equal", baseText);
      } else {
        // Modified paragraph: compute the word-level diff ONCE and reuse it for
        // BOTH outputs — stored under the BASE index for the emitter, and
        // expanded into the document-order stream for the diff JSON.
        const segments = wordDiff(baseText, revisedText);
        segmentsByPair.set(i, segments);
        for (const seg of segments) {
          pushSegment(documentOrderSegments, seg.type, seg.text);
        }
      }
    } else if (op.type === "del") {
      // Unmatched BASE paragraph (a deletion): base-only content.
      const i = op.baseIndex as number;
      pushSegment(documentOrderSegments, "del", baseParas[i].text);
    } else {
      // op.type === "ins" — unmatched REVISED paragraph (an insertion):
      // revised-only content.
      const j = op.revisedIndex as number;
      pushSegment(documentOrderSegments, "ins", revisedParas[j].text);
    }

    // 4b — Typed paragraph-mark separator BETWEEN paragraphs (see the convention
    // comment above `PARAGRAPH_SEPARATOR`). The mark belongs to a column only
    // when this op is not that column's last paragraph.
    const baseHasMark = op.baseIndex !== null && op.baseIndex !== lastBaseIndex;
    const revisedHasMark =
      op.revisedIndex !== null && op.revisedIndex !== lastRevisedIndex;

    if (baseHasMark && revisedHasMark) {
      // Boundary exists in BOTH columns: the mark is unchanged.
      pushSegment(documentOrderSegments, "equal", PARAGRAPH_SEPARATOR);
    } else if (baseHasMark) {
      // Boundary exists only in BASE (e.g. before/after a pure deletion): the
      // mark is a base-only deletion.
      pushSegment(documentOrderSegments, "del", PARAGRAPH_SEPARATOR);
    } else if (revisedHasMark) {
      // Boundary exists only in REVISED (e.g. before/after a pure insertion):
      // the mark is a revised-only insertion.
      pushSegment(documentOrderSegments, "ins", PARAGRAPH_SEPARATOR);
    }
    // else: this op is the last paragraph of its column(s) — no trailing mark.
  }

  // Step 5 — Emit the redline `.docx` with native Word tracked changes. The
  // emitter clones the base tree (never mutating our parsed model), splices in
  // insertions from the revised document, rewrites deletions and modified
  // paragraphs, and rezips deterministically using `opts.date`.
  const redlineBytes = await emitTrackedChanges({
    base,
    align,
    segmentsByPair,
    revised,
    author: opts.author,
    date: opts.date,
  });

  // Step 6 — Build the structured diff JSON (ordered hunks with per-column
  // character ranges) from the document-order segment stream.
  const diff = buildDiffJson(documentOrderSegments);

  // Step 7 — Validate BEFORE returning (AAP §0.6.2: validation precedes
  // persistence — here, before handing bytes back to the route). The reparse
  // check is fatal and throws on a malformed emission; the LibreOffice render
  // probe inside `validate` is best-effort and never fails the comparison, and
  // it does not affect the returned bytes or diff.
  await validate(redlineBytes);

  // Step 8 — Return the two artifacts.
  return { redlineBytes, diff };
}
