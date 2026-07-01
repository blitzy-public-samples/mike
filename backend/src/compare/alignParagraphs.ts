/**
 * Deterministic paragraph-level alignment for the Document Compare / Redline
 * engine.
 *
 * Given the accept-all-normalized BASE and REVISED paragraph lists (produced by
 * `./normalize`), this module computes an ordered alignment that classifies
 * every paragraph as exactly one of:
 *
 *   - `equal` -- a matched pair present in BOTH documents. This covers two
 *     cases: an IDENTICAL pair (same normalized text, an anchor) and a MODIFIED
 *     pair (different text but the same underlying paragraph). The orchestrator
 *     passes each `equal` pair to a downstream word-level diff, which is a
 *     no-op for identical text and produces `<w:ins>`/`<w:del>` runs for a
 *     modified pair,
 *   - `del`   -- a BASE paragraph with no match in REVISED (a deletion),
 *   - `ins`   -- a REVISED paragraph with no match in BASE (an insertion).
 *
 * Alignment runs in two deterministic stages:
 *
 * 1. EXACT-ANCHOR LCS -- a classic O(n*m) dynamic-programming Longest Common
 *    Subsequence over the two paragraph lists, where paragraph equality is
 *    EXACT normalized-text equality. A SHA-1 hash of each paragraph's text (via
 *    `node:crypto`) is used only as a cheap pre-filter for that comparison;
 *    every hash match is confirmed by raw-text equality, so hashing is an
 *    optimization that never affects correctness (a hash collision cannot assert
 *    a false match). This stage pins identical paragraphs as stable anchors.
 * 2. SIMILARITY REFINEMENT -- within each maximal "change region" between two
 *    anchors (a run of `del`/`ins` ops), a SECOND LCS is run over the region's
 *    deleted and inserted paragraphs using a word-token Dice-coefficient
 *    similarity predicate (>= {@link SIMILARITY_THRESHOLD}). Similar delete+insert
 *    paragraphs are re-paired into MODIFIED `equal` ops so the downstream
 *    word-level diff runs for realistic single-paragraph edits (rather than
 *    degenerating to a whole-paragraph delete + insert). Non-similar paragraphs
 *    remain `del`/`ins`. This stage is what lets a small in-paragraph edit
 *    render as an intra-paragraph redline, per AAP 0.6.2.
 *
 * Determinism (AAP 0.8.1) is the hard requirement: the function is pure, reads
 * only its inputs, performs no I/O, and uses no clock, randomness, or
 * iteration-order-dependent container, so identical inputs always yield an
 * identical `AlignOp[]`. Both LCS stages share the same tie-break: on tied LCS
 * values the backtrack favors the insertion step, which -- because ops are
 * collected backward then reversed -- yields deletion-before-insertion in
 * document order for changed/disjoint paragraphs, fixing the traversal
 * deterministically. The Dice similarity predicate is a pure function of the
 * two paragraphs' word tokens, so the refinement stage is equally deterministic.
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
 * - `"equal"` -- a matched pair present in both BASE and REVISED (either an
 *   IDENTICAL anchor or a MODIFIED pair discovered by similarity refinement;
 *   the orchestrator word-diffs the pair to tell them apart).
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
// Similarity refinement (stage 2) parameters
// ---------------------------------------------------------------------------
//
// After the exact-anchor LCS, each change region is re-aligned by paragraph
// SIMILARITY so realistic single-paragraph edits become MODIFIED pairs (an
// intra-paragraph word diff) instead of a whole-paragraph delete + insert.
// (Threshold rationale and worked examples: see the decision log.)

/**
 * Minimum Sorensen-Dice coefficient of two paragraphs' distinct word-token sets
 * for them to be treated as the SAME (modified) paragraph during refinement.
 *
 * Dice = 2*|A ∩ B| / (|A| + |B|), range [0, 1]. A value of 0.5 means the two
 * paragraphs must share at least half their tokens (weighted by set size) to be
 * paired -- high enough to reject coincidental single-word overlaps between
 * unrelated short paragraphs, low enough to catch typical clause edits (e.g. the
 * word-insert/word-delete/mixed-edit fixtures score ~0.89 / ~0.86 / 0.60). This
 * constant is the single deterministic knob governing modified-pair detection.
 */
const SIMILARITY_THRESHOLD = 0.5;

/**
 * Max secondary-LCS cells (`|deletes| * |inserts|`) refined for a single change
 * region. Beyond this bound the region is left as-is (pure `del` + `ins`) -- a
 * deterministic graceful degradation that keeps refinement cost bounded on
 * pathologically large change regions without throwing.
 */
const MAX_REGION_PAIR_CELLS = 1_000_000;

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
 * 5. REFINEMENT: {@link refineChangeRegions} re-pairs similar delete+insert
 *    paragraphs within each change region into MODIFIED `equal` ops (a second,
 *    similarity-driven LCS), so an in-paragraph edit word-diffs rather than
 *    degenerating to a whole-paragraph delete + insert.
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
 *
 * @example
 * // base = ["the quick brown fox"], revised = ["the quick red fox"]
 * // (one modified paragraph; Dice similarity above threshold) ->
 * //   equal { baseIndex: 0, revisedIndex: 0 }   // a MODIFIED pair; word-diffed
 * //   downstream into equal "the quick " + del "brown" + ins "red" + equal " fox"
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

  // Stage 2: refine each change region so similar delete+insert paragraphs
  // become MODIFIED `equal` pairs (see refineChangeRegions). Anchors from the
  // exact LCS above are preserved untouched.
  return refineChangeRegions(reversedOps, base, revised);
}

// ---------------------------------------------------------------------------
// Stage 2: similarity-based modified-paragraph refinement
// ---------------------------------------------------------------------------

/**
 * The distinct, lowercased whitespace-delimited word tokens of a paragraph.
 *
 * Pure and deterministic: splitting and lowercasing depend only on the input
 * string. The result is a Set used solely for membership/size in the Dice
 * calculation, so its (insertion-order) iteration never affects any numeric
 * outcome.
 *
 * @param text - the paragraph's normalized plaintext
 * @returns the set of distinct lowercase tokens (empty for whitespace-only text)
 */
function tokenSet(text: string): Set<string> {
  const tokens = new Set<string>();
  for (const raw of text.split(/\s+/)) {
    if (raw.length > 0) tokens.add(raw.toLowerCase());
  }
  return tokens;
}

/**
 * Sorensen-Dice similarity of two token sets: `2*|a ∩ b| / (|a| + |b|)`.
 *
 * Returns 1 when both sets are empty (two whitespace-only paragraphs are
 * considered identical) and 0 when exactly one is empty. Pure and order-
 * independent, so it preserves alignment determinism.
 *
 * @param a - first paragraph's token set
 * @param b - second paragraph's token set
 * @returns the Dice coefficient in [0, 1]
 */
function diceSimilarity(a: Set<string>, b: Set<string>): number {
  if (a.size === 0 && b.size === 0) return 1;
  if (a.size === 0 || b.size === 0) return 0;
  // Iterate the smaller set for the intersection count (order-independent).
  const [small, large] = a.size <= b.size ? [a, b] : [b, a];
  let intersection = 0;
  for (const token of small) {
    if (large.has(token)) intersection += 1;
  }
  return (2 * intersection) / (a.size + b.size);
}

/**
 * Walk the exact-LCS ops and refine each maximal CHANGE REGION -- a run of
 * consecutive non-`equal` (`del`/`ins`) ops bounded by anchors or the ends of
 * the list. `equal` anchors from stage 1 are copied through untouched; each
 * change region is handed to {@link refineChangeRegion} for similarity pairing.
 *
 * @param ops - stage-1 ops in document order
 * @param base - normalized BASE paragraphs (for token comparison)
 * @param revised - normalized REVISED paragraphs (for token comparison)
 * @returns the refined ops in document order (modified pairs now `equal`)
 */
function refineChangeRegions(
  ops: AlignOp[],
  base: NormalizedParagraph[],
  revised: NormalizedParagraph[],
): AlignOp[] {
  const out: AlignOp[] = [];
  let k = 0;
  while (k < ops.length) {
    if (ops[k].type === "equal") {
      // Stable anchor from stage 1 -- preserve exactly.
      out.push(ops[k]);
      k += 1;
      continue;
    }
    // Gather the maximal run of non-`equal` ops (this change region).
    let end = k;
    while (end < ops.length && ops[end].type !== "equal") end += 1;
    const refined = refineChangeRegion(ops.slice(k, end), base, revised);
    for (const op of refined) out.push(op);
    k = end;
  }
  return out;
}

/**
 * Re-align one change region's deleted and inserted paragraphs by SIMILARITY,
 * re-pairing similar delete+insert paragraphs into MODIFIED `equal` ops.
 *
 * The region's `del` base indices (in base order) and `ins` revised indices (in
 * revised order) are aligned with a second LCS whose match predicate is
 * "Dice similarity >= {@link SIMILARITY_THRESHOLD}", using the SAME
 * deletion-before-insertion tie-break as the primary alignment. A matched pair
 * becomes `equalOp(baseIndex, revisedIndex)` (a modified pair the orchestrator
 * will word-diff); unmatched paragraphs stay `del` / `ins`.
 *
 * Because both index lists are monotonically increasing and the LCS preserves
 * order, the emitted base/revised indices stay monotonic within the region, so
 * global document order (anchored on both sides) is preserved.
 *
 * Regions that are pure deletions or pure insertions (nothing to pair), or that
 * exceed {@link MAX_REGION_PAIR_CELLS}, are returned unchanged.
 *
 * @param region - the run of `del`/`ins` ops for one change region
 * @param base - normalized BASE paragraphs
 * @param revised - normalized REVISED paragraphs
 * @returns the region's ops after similarity re-pairing, in document order
 */
function refineChangeRegion(
  region: AlignOp[],
  base: NormalizedParagraph[],
  revised: NormalizedParagraph[],
): AlignOp[] {
  const deletes: number[] = [];
  const inserts: number[] = [];
  for (const op of region) {
    if (op.type === "del") deletes.push(op.baseIndex as number);
    else if (op.type === "ins") inserts.push(op.revisedIndex as number);
  }

  const p = deletes.length;
  const q = inserts.length;
  // Nothing to pair (pure deletion or pure insertion region).
  if (p === 0 || q === 0) return region;
  // Guardrail: leave pathologically large regions unrefined (deterministic).
  if (p * q > MAX_REGION_PAIR_CELLS) return region;

  // Precompute token sets once per region paragraph.
  const deleteTokens = deletes.map((bi) => tokenSet(base[bi].text));
  const insertTokens = inserts.map((ri) => tokenSet(revised[ri].text));
  const similar = (x: number, y: number): boolean =>
    diceSimilarity(deleteTokens[x], insertTokens[y]) >= SIMILARITY_THRESHOLD;

  // Second LCS over the similarity predicate.
  const dp: number[][] = Array.from({ length: p + 1 }, () =>
    new Array<number>(q + 1).fill(0),
  );
  for (let x = 1; x <= p; x++) {
    for (let y = 1; y <= q; y++) {
      dp[x][y] = similar(x - 1, y - 1)
        ? dp[x - 1][y - 1] + 1
        : Math.max(dp[x - 1][y], dp[x][y - 1]);
    }
  }

  // Backtrack with the same tie-break as the primary alignment (ties -> ins,
  // yielding deletion-before-insertion in final document order).
  const reversed: AlignOp[] = [];
  let x = p;
  let y = q;
  while (x > 0 && y > 0) {
    if (similar(x - 1, y - 1)) {
      // Modified pair: emit an `equal` op carrying BOTH indices so the
      // orchestrator runs a word-level diff over the two paragraphs.
      reversed.push(equalOp(deletes[x - 1], inserts[y - 1]));
      x -= 1;
      y -= 1;
    } else if (dp[x - 1][y] > dp[x][y - 1]) {
      reversed.push(delOp(deletes[x - 1]));
      x -= 1;
    } else {
      reversed.push(insOp(inserts[y - 1]));
      y -= 1;
    }
  }
  while (x > 0) {
    reversed.push(delOp(deletes[x - 1]));
    x -= 1;
  }
  while (y > 0) {
    reversed.push(insOp(inserts[y - 1]));
    y -= 1;
  }

  reversed.reverse();
  return reversed;
}
