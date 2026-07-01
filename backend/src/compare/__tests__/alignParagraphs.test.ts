/**
 * Unit tests for `backend/src/compare/alignParagraphs.ts` -- the deterministic
 * LCS-based paragraph alignment of the document-compare engine.
 *
 * `alignParagraphs(base, revised)` walks two accept-all-normalized paragraph
 * lists and emits an ordered list of {@link AlignOp} operations that classify
 * every paragraph as exactly one of:
 *   - `equal` -- a matched pair present in BOTH documents (identical normalized
 *     text; a downstream word-level diff refines these pairs);
 *   - `del`   -- a BASE-only paragraph (unmatched in REVISED): a deletion;
 *   - `ins`   -- a REVISED-only paragraph (unmatched in BASE): an insertion.
 *
 * These tests exercise the full contract with fixture-free, plain-object inputs:
 *   - identical documents collapse to all `equal`;
 *   - a single added / deleted paragraph yields one `ins` / one `del` with the
 *     surrounding matches preserved and correctly indexed;
 *   - a changed paragraph surfaces (at the paragraph level) as a `del`+`ins`
 *     pair (word-level refinement is `wordDiff`'s job, tested separately);
 *   - the FIXED, deterministic tie-break for two fully-disjoint paragraphs
 *     (LCS length 0) produces a stable ordering;
 *   - empty / one-sided inputs, the per-type index-nullability invariants, and
 *     determinism across repeated calls.
 *
 * Determinism (AAP s0.8.1) is the hard requirement under test: `alignParagraphs`
 * is pure, so identical inputs must always yield a deeply-equal `AlignOp[]`.
 *
 * Strict isolation (AAP s0.8.1): this file imports ONLY the engine under test
 * (`../alignParagraphs`), the type-only `NormalizedParagraph` shape it consumes
 * (`../normalize`), and `vitest`. It never imports `jszip`, `fast-xml-parser`,
 * `fast-diff`, `node:crypto`, `../../lib/*`, `../../routes/*`, any fixtures, or
 * any assistant/chat/doc-processing/tabular/workflow code. No Node built-ins are
 * needed: every input is constructed as a plain object literal.
 *
 * Build note: the backend `tsconfig` compiles `src/**` (including this test) and
 * does NOT register `vitest/globals` types, so the vitest globals are imported
 * explicitly to keep `npm run build` (tsc) green under `strict`.
 */

import { describe, it, expect } from "vitest";

import { alignParagraphs } from "../alignParagraphs";
import type { AlignOp } from "../alignParagraphs";
import type { NormalizedParagraph } from "../normalize";

// ---------------------------------------------------------------------------
// Fixture-free input builder
// ---------------------------------------------------------------------------
//
// `alignParagraphs` compares paragraphs by EXACT normalized-text equality and
// reads ONLY `NormalizedParagraph.text` (the `runs` array is irrelevant to
// alignment -- it is consumed later by the emitter / word-diff). So a paragraph
// with an empty `runs: []` is a fully valid, deterministic unit-test input:
// there is no need to parse or load any `.docx` fixture here.
const p = (text: string): NormalizedParagraph => ({ runs: [], text });

// ---------------------------------------------------------------------------
// Structural invariant helper (also exercises the imported `AlignOp` type)
// ---------------------------------------------------------------------------
//
// Verifies the per-type index-nullability contract documented on `AlignOp`:
//   - `equal` -> BOTH indices are real numbers;
//   - `del`   -> `baseIndex` is a real number and `revisedIndex` is `null`;
//   - `ins`   -> `baseIndex` is `null` and `revisedIndex` is a real number.
// It also proves neither index is ever `NaN` or `undefined` (only a
// non-negative integer or `null`).
const assertIndexInvariants = (ops: readonly AlignOp[]): void => {
  for (const op of ops) {
    if (op.type === "equal") {
      expect(typeof op.baseIndex).toBe("number");
      expect(typeof op.revisedIndex).toBe("number");
    } else if (op.type === "del") {
      expect(typeof op.baseIndex).toBe("number");
      expect(op.revisedIndex).toBeNull();
    } else {
      expect(op.baseIndex).toBeNull();
      expect(typeof op.revisedIndex).toBe("number");
    }
    // Each index is either `null` or a genuine non-negative integer -- never
    // `NaN` (not an integer) and never `undefined` (neither null nor integer).
    expect(op.baseIndex === null || Number.isInteger(op.baseIndex)).toBe(true);
    expect(op.revisedIndex === null || Number.isInteger(op.revisedIndex)).toBe(
      true,
    );
  }
};

// ---------------------------------------------------------------------------
// Phase 2 -- Identical documents collapse to all `equal`
// ---------------------------------------------------------------------------

describe("alignParagraphs: identical documents", () => {
  it("maps every paragraph to an `equal` op with matching indices", () => {
    const base = [p("A"), p("B"), p("C")];
    const revised = [p("A"), p("B"), p("C")];

    // With identical text throughout, the LCS spans the whole document and each
    // paragraph pairs diagonally: no insertions, no deletions.
    expect(alignParagraphs(base, revised)).toEqual([
      { type: "equal", baseIndex: 0, revisedIndex: 0 },
      { type: "equal", baseIndex: 1, revisedIndex: 1 },
      { type: "equal", baseIndex: 2, revisedIndex: 2 },
    ]);
  });
});

// ---------------------------------------------------------------------------
// Phase 3 -- Added paragraph (revised has one extra)
// ---------------------------------------------------------------------------

describe("alignParagraphs: added paragraph", () => {
  it("emits exactly one `ins` between the two preserved matches", () => {
    const ops = alignParagraphs([p("A"), p("B")], [p("A"), p("X"), p("B")]);

    // One pure insertion, zero deletions, and the two surrounding matches.
    expect(ops.filter((o) => o.type === "ins").length).toBe(1);
    expect(ops.filter((o) => o.type === "del").length).toBe(0);
    expect(ops.filter((o) => o.type === "equal").length).toBe(2);
  });

  it("points the `ins` at the extra revised paragraph with a null base index", () => {
    const ops = alignParagraphs([p("A"), p("B")], [p("A"), p("X"), p("B")]);

    const ins = ops.find((o) => o.type === "ins")!;
    expect(ins.baseIndex).toBeNull();
    expect(ins.revisedIndex).toBe(1);
  });

  it("returns the ops in document order with correctly re-based matches", () => {
    const ops = alignParagraphs([p("A"), p("B")], [p("A"), p("X"), p("B")]);

    // Document order: A matches, then the inserted X, then B matches.
    expect(ops.map((o) => o.type)).toEqual(["equal", "ins", "equal"]);
    // A maps 0 -> 0; B maps 1 -> 2 (its revised index shifted by the insertion).
    expect(ops[0]).toEqual({ type: "equal", baseIndex: 0, revisedIndex: 0 });
    expect(ops[2]).toEqual({ type: "equal", baseIndex: 1, revisedIndex: 2 });
  });
});

// ---------------------------------------------------------------------------
// Phase 4 -- Deleted paragraph (base has one extra)
// ---------------------------------------------------------------------------

describe("alignParagraphs: deleted paragraph", () => {
  it("emits exactly one `del` between the two preserved matches", () => {
    const ops = alignParagraphs([p("A"), p("X"), p("B")], [p("A"), p("B")]);

    // One pure deletion, zero insertions, and the two surrounding matches.
    expect(ops.filter((o) => o.type === "del").length).toBe(1);
    expect(ops.filter((o) => o.type === "ins").length).toBe(0);
    expect(ops.filter((o) => o.type === "equal").length).toBe(2);
  });

  it("points the `del` at the extra base paragraph with a null revised index", () => {
    const ops = alignParagraphs([p("A"), p("X"), p("B")], [p("A"), p("B")]);

    const del = ops.find((o) => o.type === "del")!;
    expect(del.baseIndex).toBe(1);
    expect(del.revisedIndex).toBeNull();
  });

  it("returns the ops in document order with correctly re-based matches", () => {
    const ops = alignParagraphs([p("A"), p("X"), p("B")], [p("A"), p("B")]);

    // Document order: A matches, then the deleted X, then B matches.
    expect(ops.map((o) => o.type)).toEqual(["equal", "del", "equal"]);
    // A maps 0 -> 0; B maps 2 -> 1 (its base index shifted by the deletion).
    expect(ops[0]).toEqual({ type: "equal", baseIndex: 0, revisedIndex: 0 });
    expect(ops[2]).toEqual({ type: "equal", baseIndex: 2, revisedIndex: 1 });
  });
});

// ---------------------------------------------------------------------------
// Phase 5 -- Changed paragraph -> del(old) + ins(new)
// ---------------------------------------------------------------------------

describe("alignParagraphs: changed paragraph", () => {
  it("keeps the unchanged paragraph equal and splits the change into del+ins", () => {
    // B and B2 differ by exact normalized text, so the pair does not match and
    // surfaces (at the paragraph level) as a deletion of the old paragraph plus
    // an insertion of the new one.
    const ops = alignParagraphs([p("A"), p("B")], [p("A"), p("B2")]);

    // The unchanged paragraph A yields a single equal pair mapping 0 -> 0.
    expect(ops.filter((o) => o.type === "equal")).toEqual([
      { type: "equal", baseIndex: 0, revisedIndex: 0 },
    ]);

    // Exactly one del and one ins carry the changed paragraph's indices. We do
    // NOT assert their relative order: the paragraph-level del/ins pairing is an
    // internal backtrack detail, and word-level refinement of a changed pair is
    // `wordDiff`'s responsibility (covered by its own tests).
    const dels = ops.filter((o) => o.type === "del");
    const inses = ops.filter((o) => o.type === "ins");
    expect(dels.length).toBe(1);
    expect(inses.length).toBe(1);
    expect(dels[0].baseIndex).toBe(1);
    expect(dels[0].revisedIndex).toBeNull();
    expect(inses[0].baseIndex).toBeNull();
    expect(inses[0].revisedIndex).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// Phase 6 -- Fixed, deterministic tie-break for fully-disjoint paragraphs
// ---------------------------------------------------------------------------

describe("alignParagraphs: deterministic tie-break (disjoint paragraphs)", () => {
  it("resolves two disjoint single paragraphs to a stable [ins, del] order", () => {
    // A single base paragraph "A" and a single revised paragraph "B" share no
    // common subsequence (LCS length 0), so the pair is emitted as one deletion
    // of the base paragraph plus one insertion of the revised paragraph.
    const ops = alignParagraphs([p("A")], [p("B")]);

    // The ordering is FIXED and deterministic. The LCS backtrack's `>=`
    // tie-break consumes the BASE paragraph (the deletion) first while walking
    // the DP table backward from the document end; after the single reverse into
    // document order, the deletion therefore lands AFTER the insertion, yielding
    // exactly [ins, del]. Asserting this precise sequence proves the tie-break is
    // resolved deterministically (never nondeterministically, and never the
    // opposite permutation).
    expect(ops.map((o) => o.type)).toEqual(["ins", "del"]);
    expect(ops[0]).toEqual({ type: "ins", baseIndex: null, revisedIndex: 0 });
    expect(ops[1]).toEqual({ type: "del", baseIndex: 0, revisedIndex: null });
  });
});

// ---------------------------------------------------------------------------
// Phase 7 -- Edge cases, structural invariants, and determinism
// ---------------------------------------------------------------------------

describe("alignParagraphs: edge cases", () => {
  it("returns an empty op list for two empty documents", () => {
    expect(alignParagraphs([], [])).toEqual([]);
  });

  it("emits a single `del` when only the base has a paragraph", () => {
    expect(alignParagraphs([p("A")], [])).toEqual([
      { type: "del", baseIndex: 0, revisedIndex: null },
    ]);
  });

  it("emits a single `ins` when only the revised has a paragraph", () => {
    expect(alignParagraphs([], [p("A")])).toEqual([
      { type: "ins", baseIndex: null, revisedIndex: 0 },
    ]);
  });
});

describe("alignParagraphs: structural invariants and determinism", () => {
  it("honors the per-type index-nullability contract (no NaN / undefined)", () => {
    // Use a non-trivial mix (equal + ins + equal) so every op type is present.
    const ops = alignParagraphs([p("A"), p("B")], [p("A"), p("X"), p("B")]);
    assertIndexInvariants(ops);
  });

  it("produces deeply-equal output across repeated calls (pure + deterministic)", () => {
    const base = [p("A"), p("B")];
    const revised = [p("A"), p("X"), p("B")];

    // Same inputs -> same ops, every time: no clock, randomness, or
    // iteration-order dependence leaks into the alignment.
    expect(alignParagraphs(base, revised)).toEqual(
      alignParagraphs(base, revised),
    );
  });

  it("keeps base and revised indices monotonically increasing in document order", () => {
    const ops = alignParagraphs([p("A"), p("B")], [p("A"), p("X"), p("B")]);

    // Among ops that carry a base index (everything except `ins`), the base
    // index strictly increases -- the walk never revisits or reorders BASE
    // paragraphs. The `as number` cast is sound because non-`ins` ops always
    // carry a numeric baseIndex per the index-nullability invariant proven above.
    const baseSeq = ops
      .filter((o) => o.type !== "ins")
      .map((o) => o.baseIndex as number);
    for (let k = 1; k < baseSeq.length; k++) {
      expect(baseSeq[k]).toBeGreaterThan(baseSeq[k - 1]);
    }

    // Symmetrically, among ops that carry a revised index (everything except
    // `del`), the revised index strictly increases.
    const revisedSeq = ops
      .filter((o) => o.type !== "del")
      .map((o) => o.revisedIndex as number);
    for (let k = 1; k < revisedSeq.length; k++) {
      expect(revisedSeq[k]).toBeGreaterThan(revisedSeq[k - 1]);
    }
  });
});
