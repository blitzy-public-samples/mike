/**
 * Unit tests for the word-level diff of the Document Compare / Redline engine.
 *
 * Exercises {@link wordDiff} from `../wordDiff`: the mapping of the underlying
 * `fast-diff` operations onto typed {@link DiffSegment}s (`ins` / `del` /
 * `equal`), the identical / pure-insert / pure-delete edge cases, and — most
 * importantly — the RECONSTRUCTION INVARIANT that the downstream
 * `buildDiffJson` relies on to assign contiguous per-column character offsets:
 *
 *   concat(text of segments whose type is `equal` or `del`) === baseText
 *   concat(text of segments whose type is `equal` or `ins`) === revisedText
 *
 * `wordDiff` never adds, drops, or mutates a character — it only classifies
 * each character into a segment — so this invariant is the single most
 * important property to lock. Determinism is asserted too (identical inputs
 * always yield identical segment arrays, since the wrapper uses no clock,
 * randomness, or cursor hint).
 *
 * Robustness to engine latitude (AAP §0.6.2): the engine MAY coalesce adjacent
 * same-type segments and MAY snap change boundaries to whole words. Mixed-edit
 * cases therefore assert INVARIANTS and substring membership rather than exact
 * segment arrays; exact-array assertions are reserved for the identical /
 * pure-insert / pure-delete cases, whose output is unambiguous.
 *
 * Isolation (AAP §0.8.1): this suite imports ONLY the wrapper under test
 * (`../wordDiff`), the type-only segment contract (`../diffJson`), and
 * `vitest`. It deliberately does NOT import `fast-diff` itself, any
 * `../../lib/*` or `../../routes/*` module, `jszip` / `fast-xml-parser`, any
 * `.docx` fixture, or any Node built-in — `wordDiff` needs none of them, and
 * the point of the test is to lock the wrapper's contract, not the library it
 * happens to wrap.
 *
 * Build note: the explicit `vitest` import below is MANDATORY. `backend/tsconfig.json`
 * compiles `src/**` with `tsc` (the production `npm run build`) and does NOT
 * register the `vitest/globals` ambient types, so `describe` / `it` / `expect`
 * must be imported explicitly for this file to type-check. Relying on the
 * runtime `vitest.config.ts` alone (which the `tsc` build does not observe)
 * would break `npm run build`.
 */

// Explicit vitest import (MANDATORY — see the build note above; do NOT rely on
// vitest.config.ts globals, which the `tsc` build does not observe).
import { describe, it, expect } from "vitest";

import { wordDiff } from "../wordDiff";
import type { DiffSegment } from "../diffJson";

// ---------------------------------------------------------------------------
// Local reconstruction helpers.
//
// A redline consumer rebuilds each column by walking the segments once: the
// BASE column is every non-`ins` segment's text in order, and the REVISED
// column is every non-`del` segment's text in order. `textOfType` isolates the
// text classified into exactly one bucket, used for change-membership checks.
// ---------------------------------------------------------------------------

/** Reconstructs the BASE-column plaintext (`equal` + `del` text, in order). */
const baseOf = (segs: DiffSegment[]): string =>
  segs
    .filter((s) => s.type !== "ins")
    .map((s) => s.text)
    .join("");

/** Reconstructs the REVISED-column plaintext (`equal` + `ins` text, in order). */
const revOf = (segs: DiffSegment[]): string =>
  segs
    .filter((s) => s.type !== "del")
    .map((s) => s.text)
    .join("");

/** Concatenates, in document order, the text of every segment of type `t`. */
const textOfType = (segs: DiffSegment[], t: DiffSegment["type"]): string =>
  segs
    .filter((s) => s.type === t)
    .map((s) => s.text)
    .join("");

// ---------------------------------------------------------------------------
// Phase 2 — Identical and pure-edit edge cases (EXACT assertions).
//
// These inputs have an unambiguous canonical output, so the exact segment
// arrays are asserted directly (no coalescing/word-snapping latitude applies).
// ---------------------------------------------------------------------------

describe("wordDiff — identical and pure-edit edge cases (exact)", () => {
  it("returns a single equal segment for identical non-empty text", () => {
    expect(wordDiff("abc", "abc")).toEqual([{ type: "equal", text: "abc" }]);
  });

  it("returns no segments for two empty strings", () => {
    // Both inputs are empty: the sole (empty) equal segment is dropped, so the
    // result is the empty array rather than a zero-length `equal` segment.
    expect(wordDiff("", "")).toEqual([]);
  });

  it("classifies a pure insertion as all-ins (empty base, full revised)", () => {
    const s = wordDiff("", "abc");

    expect(baseOf(s)).toBe("");
    expect(revOf(s)).toBe("abc");
    // Every emitted segment is a non-empty insertion.
    expect(s.every((x) => x.type === "ins" && x.text.length > 0)).toBe(true);
    expect(textOfType(s, "ins")).toBe("abc");
  });

  it("classifies a pure deletion as all-del (full base, empty revised)", () => {
    const s = wordDiff("abc", "");

    expect(baseOf(s)).toBe("abc");
    expect(revOf(s)).toBe("");
    // Every emitted segment is a non-empty deletion.
    expect(s.every((x) => x.type === "del" && x.text.length > 0)).toBe(true);
    expect(textOfType(s, "del")).toBe("abc");
  });
});

// ---------------------------------------------------------------------------
// Phase 3 — Mixed word edits (INVARIANTS + substring membership).
//
// The engine may coalesce adjacent same-type segments and may snap change
// boundaries to whole words, so exact segment arrays are deliberately NOT
// asserted here. Instead we lock the reconstruction invariant and assert that
// the changed token lands in the correct bucket (and only there).
// ---------------------------------------------------------------------------

describe("wordDiff — mixed word edits (invariants + substring membership)", () => {
  it("replaces a single interior word (quick -> slow)", () => {
    const s = wordDiff("the quick brown fox", "the slow brown fox");

    // Reconstruction — the critical property for diff-JSON offset correctness.
    expect(baseOf(s)).toBe("the quick brown fox");
    expect(revOf(s)).toBe("the slow brown fox");

    // The changed word lands in the correct bucket...
    expect(textOfType(s, "del")).toContain("quick");
    expect(textOfType(s, "ins")).toContain("slow");

    // ...and only there (no cross-contamination between the del/ins buckets).
    expect(textOfType(s, "del")).not.toContain("slow");
    expect(textOfType(s, "ins")).not.toContain("quick");
  });

  it("replaces a numeric token (30 -> 45), leaving the surround equal", () => {
    const s = wordDiff(
      "Payment is due in 30 days.",
      "Payment is due in 45 days.",
    );

    expect(baseOf(s)).toBe("Payment is due in 30 days.");
    expect(revOf(s)).toBe("Payment is due in 45 days.");

    expect(textOfType(s, "del")).toContain("30");
    expect(textOfType(s, "ins")).toContain("45");
    expect(textOfType(s, "del")).not.toContain("45");
    expect(textOfType(s, "ins")).not.toContain("30");
  });
});

// ---------------------------------------------------------------------------
// Phase 4 — General invariants over a table of edits.
//
// For every [base, revised] pair the same three properties must hold, whatever
// segmentation the engine chooses:
//   1. reconstruction (baseOf === base, revOf === revised) — the property the
//      diff-JSON offsets depend on,
//   2. no empty segments (the wrapper drops zero-length text), and
//   3. only the valid `ins` / `del` / `equal` op mapping is emitted.
// ---------------------------------------------------------------------------

describe("wordDiff — general invariants over a table of edits", () => {
  const cases: ReadonlyArray<{ name: string; base: string; revised: string }> =
    [
      {
        name: "interior word replacement",
        base: "the quick brown fox",
        revised: "the slow brown fox",
      },
      {
        name: "numeric token replacement",
        base: "Payment is due in 30 days.",
        revised: "Payment is due in 45 days.",
      },
      {
        name: "prefix-only change",
        base: "hello world",
        revised: "goodbye world",
      },
      {
        name: "suffix-only change",
        base: "hello world",
        revised: "hello there",
      },
      {
        name: "insertion in the middle",
        base: "alpha gamma",
        revised: "alpha beta gamma",
      },
      {
        name: "deletion in the middle",
        base: "alpha beta gamma",
        revised: "alpha gamma",
      },
    ];

  for (const { name, base, revised } of cases) {
    describe(name, () => {
      // `wordDiff` is pure, so computing once at collection time is safe.
      const s = wordDiff(base, revised);

      it("reconstructs the base column from every non-ins segment", () => {
        expect(baseOf(s)).toBe(base);
      });

      it("reconstructs the revised column from every non-del segment", () => {
        expect(revOf(s)).toBe(revised);
      });

      it("emits no empty segments", () => {
        expect(s.every((x) => x.text.length > 0)).toBe(true);
      });

      it("emits only valid ins/del/equal segment types", () => {
        expect(
          s.every(
            (x) => x.type === "ins" || x.type === "del" || x.type === "equal",
          ),
        ).toBe(true);
      });
    });
  }
});

// ---------------------------------------------------------------------------
// Phase 5 — Determinism.
//
// `wordDiff` uses no clock, randomness, or cursor hint, so repeated calls on
// the same inputs must produce a deeply-equal, order-stable result.
// ---------------------------------------------------------------------------

describe("wordDiff — determinism", () => {
  it("produces deeply-equal output for repeated identical calls", () => {
    const a = wordDiff("the quick brown fox", "the slow brown fox");
    const b = wordDiff("the quick brown fox", "the slow brown fox");

    expect(a).toEqual(b);
    // Order-stable serialization also holds for a pure, cursor-free diff.
    expect(JSON.stringify(a)).toBe(JSON.stringify(b));
  });
});
