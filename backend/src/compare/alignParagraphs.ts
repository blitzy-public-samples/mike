/**
 * Deterministic paragraph-level alignment for the Document Compare / Redline
 * engine.
 *
 * Given the accept-all-normalized BASE and REVISED paragraph lists (produced by
 * `./normalize`), this module computes an ordered alignment that classifies
 * every paragraph as exactly one of:
 *
 *   - `equal` -- a matched pair present in BOTH documents (identical normalized
 *     text); a downstream word-level diff refines these pairs,
 *   - `del`   -- a BASE paragraph with no match in REVISED (a deletion),
 *   - `ins`   -- a REVISED paragraph with no match in BASE (an insertion).
 *
 * The alignment is a classic O(n*m) dynamic-programming Longest Common
 * Subsequence (LCS) over the two paragraph lists, where paragraph equality is
 * EXACT normalized-text equality. A SHA-1 hash of each paragraph's text (via
 * `node:crypto`) is used only as a cheap pre-filter for that comparison; every
 * hash match is confirmed by raw-text equality, so hashing is an optimization
 * that never affects correctness (a hash collision cannot assert a false match).
 *
 * Determinism (AAP 0.8.1) is the hard requirement: the function is pure, reads
 * only its inputs, performs no I/O, and uses no clock, randomness, or
 * iteration-order-dependent container, so identical inputs always yield an
 * identical `AlignOp[]`. Behavior note: on tied LCS values the backtrack favors
 * the insertion step, which -- because ops are collected backward then reversed
 * -- yields deletion-before-insertion in document order for changed/disjoint
 * paragraphs, fixing the traversal deterministically.
 *
 * The ops are returned in document (top-to-bottom) order so the downstream
 * tracked-changes emitter and the diff-JSON builder can walk the same ordered
 * sequence and reconstruct the merged document consistently.
 *
 * @module compare/alignParagraphs
 */

import { createHash } from "node:crypto";

import type { NormalizedParagraph } from "./normalize";

// ---------------------------------------------------------------------------
// Public contract
// ---------------------------------------------------------------------------

/**
 * Classifies a single alignment op by which document(s) contain the paragraph.
 *
 * - `"equal"` -- a matched pair present in both BASE and REVISED.
 * - `"del"`   -- a BASE-only paragraph (unmatched in REVISED): a deletion.
 * - `"ins"`   -- a REVISED-only paragraph (unmatched in BASE): an insertion.
 */
export type AlignOpType = "equal" | "del" | "ins";

/**
 * One ordered paragraph-alignment operation.
 *
 * Index nullability is fully determined by {@link AlignOp.type}:
 * - `equal` -> both {@link AlignOp.baseIndex} and {@link AlignOp.revisedIndex}
 *   are non-null (the matched pair).
 * - `del`   -> {@link AlignOp.baseIndex} is non-null and
 *   {@link AlignOp.revisedIndex} is `null`.
 * - `ins`   -> {@link AlignOp.revisedIndex} is non-null and
 *   {@link AlignOp.baseIndex} is `null`.
 *
 * Ops are emitted in document order, so a consumer can rebuild the merged
 * document (and each column) with a single walk of the array.
 */
export interface AlignOp {
  /** Which document(s) contain this paragraph. */
  type: AlignOpType;
  /** Index into the BASE paragraph list; `null` for pure insertions (`ins`). */
  baseIndex: number | null;
  /** Index into the REVISED paragraph list; `null` for pure deletions (`del`). */
  revisedIndex: number | null;
}

// ---------------------------------------------------------------------------
// Op constructors (enforce the per-type index-nullability invariants)
// ---------------------------------------------------------------------------
//
// Building ops through these helpers makes the null-ness rules structural
// rather than incidental: an `equal` always carries both indices, a `del`
// always nulls the revised index, and an `ins` always nulls the base index.
// This guarantees the "no NaN/undefined; equal has both, del only base, ins
// only revised" contract at every emission site.

/** Build an `equal` op for a matched pair (both indices non-null). */
function equalOp(baseIndex: number, revisedIndex: number): AlignOp {
  return { type: "equal", baseIndex, revisedIndex };
}

/** Build a `del` op for an unmatched BASE paragraph (revised index `null`). */
function delOp(baseIndex: number): AlignOp {
  return { type: "del", baseIndex, revisedIndex: null };
}

/** Build an `ins` op for an unmatched REVISED paragraph (base index `null`). */
function insOp(revisedIndex: number): AlignOp {
  return { type: "ins", baseIndex: null, revisedIndex };
}

// ---------------------------------------------------------------------------
// Paragraph equality (exact normalized-text equality, hash-accelerated)
// ---------------------------------------------------------------------------

/**
 * SHA-1 hex digest of a paragraph's normalized text, used only as a cheap
 * pre-filter for paragraph equality. It is a pure function of the input string
 * (no salt, clock, or randomness), so identical text always yields an identical
 * digest -- keeping the alignment fully deterministic.
 *
 * @param text - the paragraph's normalized plaintext
 * @returns the lowercase hexadecimal SHA-1 digest of `text`
 */
function hashParagraphText(text: string): string {
  return createHash("sha1").update(text, "utf8").digest("hex");
}

// ---------------------------------------------------------------------------
// Resource guardrails (deterministic fail-fast limits)
// ---------------------------------------------------------------------------
//
// The alignment allocates an O(n*m) LCS table; these fixed bounds are checked
// BEFORE that allocation and throw a controlled error when exceeded.
// (Thresholds and rationale: see docs/decisions/document-compare-decision-log.md.)

/** Max body paragraphs accepted per document (per side) before alignment. */
const MAX_ALIGN_PARAGRAPHS_PER_DOC = 50_000;

/**
 * Max LCS dynamic-programming cells -- `(base.length + 1) * (revised.length + 1)`
 * -- allocated for a single alignment, bounding the DP table's memory footprint.
 */
const MAX_LCS_CELLS = 25_000_000;

// ---------------------------------------------------------------------------
// Deterministic LCS alignment
// ---------------------------------------------------------------------------

/**
 * Align the normalized BASE paragraphs against the normalized REVISED
 * paragraphs, returning the ordered list of alignment operations.
 *
 * Algorithm (deterministic LCS):
 * 1. Precompute a SHA-1 digest per paragraph as a fast equality pre-filter;
 *    equality is then confirmed by exact normalized-text comparison.
 * 2. Fill the standard LCS dynamic-programming table `dp`, where `dp[i][j]` is
 *    the LCS length of `base[0..i-1]` and `revised[0..j-1]`.
 * 3. Backtrack from `(n, m)` to `(0, 0)`: a matched pair moves diagonally and
 *    emits `equal`; otherwise, when `dp[i-1][j] > dp[i][j-1]` a BASE paragraph
 *    is consumed as a `del`, else (including on tied DP values) a REVISED
 *    paragraph is consumed as an `ins`. Because ops are collected backward and
 *    reversed in step 4, sending ties to the `ins` step yields
 *    deletion-before-insertion in document order -- the required ordering for
 *    changed/disjoint paragraphs -- deterministically. Once one side is
 *    exhausted, the remaining BASE paragraphs are `del`s and the remaining
 *    REVISED paragraphs are `ins`s.
 * 4. Backtracking collects ops in reverse; the list is reversed once so it is
 *    returned in document (top-to-bottom) order.
 *
 * The result is pure and deterministic: identical inputs always produce a
 * deeply-equal `AlignOp[]`. Complexity is O(n*m) time and space, which is well
 * within budget for V1 contract-sized documents (hundreds to low-thousands of
 * paragraphs).
 *
 * @param base - normalized BASE paragraphs in document order
 * @param revised - normalized REVISED paragraphs in document order
 * @returns the ordered alignment ops (`equal` / `del` / `ins`) in document order
 *
 * @example
 * // base = ["A", "B"], revised = ["A", "N", "B"] (one inserted paragraph) ->
 * //   equal { baseIndex: 0,    revisedIndex: 0    }
 * //   ins   { baseIndex: null, revisedIndex: 1    }
 * //   equal { baseIndex: 1,    revisedIndex: 2    }
 */
export function alignParagraphs(
  base: NormalizedParagraph[],
  revised: NormalizedParagraph[],
): AlignOp[] {
  const n = base.length;
  const m = revised.length;

  // Guardrail: reject pathological paragraph counts before doing any work.
  if (n > MAX_ALIGN_PARAGRAPHS_PER_DOC || m > MAX_ALIGN_PARAGRAPHS_PER_DOC) {
    throw new Error(
      `compare: paragraph count exceeds limit (base=${n}, revised=${m}, ` +
        `max=${MAX_ALIGN_PARAGRAPHS_PER_DOC} per document)`,
    );
  }

  // Per-paragraph SHA-1 digests: a cheap, deterministic equality pre-filter.
  const baseKeys = base.map((paragraph) => hashParagraphText(paragraph.text));
  const revisedKeys = revised.map((paragraph) =>
    hashParagraphText(paragraph.text),
  );

  // Paragraph equality: require a hash match (fast reject on mismatch), then
  // confirm with exact normalized-text equality so a hash collision can never
  // assert a false match. Correctness therefore rests on text, not the hash.
  const equalAt = (i: number, j: number): boolean =>
    baseKeys[i] === revisedKeys[j] && base[i].text === revised[j].text;

  // Guardrail: bound the LCS table size before the Array.from allocation below.
  const lcsCells = (n + 1) * (m + 1);
  if (lcsCells > MAX_LCS_CELLS) {
    throw new Error(
      `compare: alignment matrix ${n + 1}x${m + 1} (${lcsCells} cells) ` +
        `exceeds limit (max=${MAX_LCS_CELLS})`,
    );
  }

  // Standard LCS DP table (prefix formulation): dp[i][j] = LCS length of
  // base[0..i-1] and revised[0..j-1]. Row 0 and column 0 model empty prefixes
  // and stay 0.
  const dp: number[][] = Array.from({ length: n + 1 }, () =>
    new Array<number>(m + 1).fill(0),
  );
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      dp[i][j] = equalAt(i - 1, j - 1)
        ? dp[i - 1][j - 1] + 1
        : Math.max(dp[i - 1][j], dp[i][j - 1]);
    }
  }

  // Backtrack from (n, m) toward (0, 0). Ops are collected in reverse document
  // order and reversed once at the end.
  const reversedOps: AlignOp[] = [];
  let i = n;
  let j = m;
  while (i > 0 && j > 0) {
    if (equalAt(i - 1, j - 1)) {
      // Matched pair: consume one paragraph from each side (move diagonally).
      reversedOps.push(equalOp(i - 1, j - 1));
      i -= 1;
      j -= 1;
    } else if (dp[i - 1][j] > dp[i][j - 1]) {
      // Deletion step. The strict `>` sends tied DP values to the insertion
      // branch below; because the backtrack runs backward and the ops are
      // reversed once at the end, consuming the insertion first on a tie makes
      // the deletion precede the insertion in final document order
      // (deletion-before-insertion for changed/disjoint paragraphs).
      reversedOps.push(delOp(i - 1));
      i -= 1;
    } else {
      // Insertion step -- also the tie case, per the strict `>` above.
      reversedOps.push(insOp(j - 1));
      j -= 1;
    }
  }
  // One side is exhausted: any remaining BASE paragraphs are deletions and any
  // remaining REVISED paragraphs are insertions.
  while (i > 0) {
    reversedOps.push(delOp(i - 1));
    i -= 1;
  }
  while (j > 0) {
    reversedOps.push(insOp(j - 1));
    j -= 1;
  }

  // Restore document (top-to-bottom) order.
  reversedOps.reverse();
  return reversedOps;
}
