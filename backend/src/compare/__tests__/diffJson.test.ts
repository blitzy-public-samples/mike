/**
 * Unit tests for the diff-JSON contract + deterministic builder.
 *
 * Exercises {@link buildDiffJson} from `../diffJson`: per-column offset/range
 * correctness, the zero-length skip rule, the null-ness rules per hunk type,
 * and the reconstruction invariant (base column = `equal` + `del` text; revised
 * column = `equal` + `ins` text) together with per-column contiguity. These
 * properties are what let the frontend rebuild both redline columns from a
 * single walk of the hunk array, so they are the engine's central correctness
 * guarantees.
 *
 * Isolation (AAP §0.8.1): this suite imports ONLY the engine under test and
 * `vitest`. It pulls in no `jszip` / `fast-xml-parser` / `fast-diff`, no
 * `../../lib/*` or `../../routes/*` module, no `.docx` fixtures, and no Node
 * built-ins — `buildDiffJson` is dependency-free and needs none of them.
 *
 * Build note: the explicit `vitest` import below is MANDATORY. `backend/tsconfig.json`
 * compiles `src/**` with `tsc` (the production `npm run build`) and does NOT
 * register the `vitest/globals` ambient types, so `describe` / `it` / `expect`
 * must be imported explicitly for this file to type-check. Relying on the
 * runtime `globals: true` from `vitest.config.ts` alone would break the build.
 */

// Explicit vitest import (MANDATORY — see the build note above; do NOT rely on
// vitest.config.ts `globals: true`, which the `tsc` build does not observe).
import { describe, it, expect } from "vitest";

import { buildDiffJson } from "../diffJson";
import type { DiffJson, DiffHunk, DiffSegment } from "../diffJson";

// ---------------------------------------------------------------------------
// Local reconstruction helpers.
//
// A consumer rebuilds each redline column by walking the hunks once (AAP
// §0.6.2): the BASE column is every non-`ins` hunk's text in order, and the
// REVISED column is every non-`del` hunk's text in order.
// ---------------------------------------------------------------------------

/** Reconstructs the BASE-column plaintext (`equal` + `del` text, in order). */
function baseText(d: DiffJson): string {
  return d.hunks
    .filter((h) => h.type !== "ins")
    .map((h) => h.text)
    .join("");
}

/** Reconstructs the REVISED-column plaintext (`equal` + `ins` text, in order). */
function revisedText(d: DiffJson): string {
  return d.hunks
    .filter((h) => h.type !== "del")
    .map((h) => h.text)
    .join("");
}

/**
 * Asserts per-column contiguity for one column of a diff:
 * - the relevant range is non-null on every contributing hunk,
 * - the first contributing hunk starts at offset 0,
 * - each subsequent range starts exactly where the previous one ended
 *   (no gaps, no overlaps), and
 * - each range's width equals the hunk's `text.length`.
 *
 * The check is coalescing-robust: it never assumes a hunk count, only that the
 * offsets tile the column exactly.
 *
 * @param hunks  All hunks of the diff, in document order.
 * @param column Which column to check: BASE (skip `ins`) or REVISED (skip `del`).
 */
function assertColumnContiguity(
  hunks: DiffHunk[],
  column: "base" | "revised",
): void {
  const skipType: DiffHunk["type"] = column === "base" ? "ins" : "del";
  let cursor = 0;

  for (const hunk of hunks) {
    if (hunk.type === skipType) {
      continue;
    }

    const range = column === "base" ? hunk.baseRange : hunk.revisedRange;

    // A contributing hunk MUST carry the column's range. `expect` is not a TS
    // type guard, so narrow explicitly afterwards to stay strict-null-safe
    // without a `!` non-null assertion.
    expect(range).not.toBeNull();
    if (range === null) {
      throw new Error(
        `Expected a non-null ${column}Range for ${hunk.type} hunk "${hunk.text}"`,
      );
    }

    expect(range.start).toBe(cursor);
    expect(range.end - range.start).toBe(hunk.text.length);

    cursor = range.end;
  }
}

// ---------------------------------------------------------------------------
// Phase 2 — the canonical hand-trace (exact assertions).
//
// The input has no adjacent same-type segments and no zero-length segments, so
// the exact ranges are guaranteed regardless of whether the builder coalesces.
// ---------------------------------------------------------------------------

describe("buildDiffJson — canonical hand-trace", () => {
  it("assigns exact per-column ranges for a mixed equal/del/ins/equal diff", () => {
    const segments: DiffSegment[] = [
      { type: "equal", text: "Hello " },
      { type: "del", text: "cruel " },
      { type: "ins", text: "kind " },
      { type: "equal", text: "world" },
    ];

    const { hunks } = buildDiffJson(segments);

    // Hand-trace (matches diffJson.ts's own @example, lines 121-126):
    //   "Hello " len 6 -> base 0..6,   revised 0..6
    //   "cruel " len 6 -> base 6..12,  revised null   (del advances base only)
    //   "kind "  len 5 -> base null,   revised 6..11  (ins advances revised only)
    //   "world"  len 5 -> base 12..17, revised 11..16
    // The final revisedRange is 11..16: the revised cursor sits at 11 after
    // "kind " (6..11), so "world" (len 5) occupies 11..16.
    expect(hunks).toEqual([
      {
        type: "equal",
        text: "Hello ",
        baseRange: { start: 0, end: 6 },
        revisedRange: { start: 0, end: 6 },
      },
      {
        type: "del",
        text: "cruel ",
        baseRange: { start: 6, end: 12 },
        revisedRange: null,
      },
      {
        type: "ins",
        text: "kind ",
        baseRange: null,
        revisedRange: { start: 6, end: 11 },
      },
      {
        type: "equal",
        text: "world",
        baseRange: { start: 12, end: 17 },
        revisedRange: { start: 11, end: 16 },
      },
    ]);
  });
});

// ---------------------------------------------------------------------------
// Phase 3 — per-type range/null rules.
// ---------------------------------------------------------------------------

describe("buildDiffJson — per-type range and null rules", () => {
  it("emits a pure insertion with a null baseRange and a revised range", () => {
    expect(buildDiffJson([{ type: "ins", text: "abc" }]).hunks).toEqual([
      {
        type: "ins",
        text: "abc",
        baseRange: null,
        revisedRange: { start: 0, end: 3 },
      },
    ]);
  });

  it("emits a pure deletion with a null revisedRange and a base range", () => {
    expect(buildDiffJson([{ type: "del", text: "abc" }]).hunks).toEqual([
      {
        type: "del",
        text: "abc",
        baseRange: { start: 0, end: 3 },
        revisedRange: null,
      },
    ]);
  });

  it("emits a single equal spanning both columns", () => {
    expect(buildDiffJson([{ type: "equal", text: "abcd" }]).hunks).toEqual([
      {
        type: "equal",
        text: "abcd",
        baseRange: { start: 0, end: 4 },
        revisedRange: { start: 0, end: 4 },
      },
    ]);
  });

  it("returns no hunks for empty input", () => {
    expect(buildDiffJson([]).hunks).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// Phase 4 — zero-length segments are skipped.
// ---------------------------------------------------------------------------

describe("buildDiffJson — zero-length segments are skipped", () => {
  it("drops every empty segment and never emits an empty hunk", () => {
    const result = buildDiffJson([
      { type: "equal", text: "a" },
      { type: "equal", text: "" },
      { type: "del", text: "" },
      { type: "ins", text: "" },
      { type: "del", text: "b" },
    ]);

    // No hunk carries empty text.
    expect(result.hunks.every((h) => h.text.length > 0)).toBe(true);

    // "a" (equal) + "b" (del) -> base column "ab"; the revised column sees only
    // the "a" equal because the sole non-empty deletion is base-only.
    expect(baseText(result)).toBe("ab");
    expect(revisedText(result)).toBe("a");
  });
});

// ---------------------------------------------------------------------------
// Phase 5 — reconstruction invariant + per-column contiguity + null rules.
//
// These assertions hold whether or not the builder coalesces adjacent
// same-type segments, so exact hunk counts are deliberately NOT asserted for
// the adjacent-same-type case.
// ---------------------------------------------------------------------------

describe("buildDiffJson — reconstruction invariant, contiguity, and null rules", () => {
  const cases: ReadonlyArray<{ name: string; segments: DiffSegment[] }> = [
    {
      name: "no adjacent same-type segments (canonical list)",
      segments: [
        { type: "equal", text: "Hello " },
        { type: "del", text: "cruel " },
        { type: "ins", text: "kind " },
        { type: "equal", text: "world" },
      ],
    },
    {
      name: "adjacent same-type segments (coalescing-robust)",
      segments: [
        { type: "equal", text: "a" },
        { type: "equal", text: "b" },
        { type: "del", text: "c" },
        { type: "ins", text: "d" },
        { type: "ins", text: "e" },
      ],
    },
  ];

  for (const { name, segments } of cases) {
    describe(name, () => {
      // `buildDiffJson` is pure, so computing once at collection time is safe.
      const result = buildDiffJson(segments);

      it("reconstructs the base column from every non-ins segment text", () => {
        const expectedBase = segments
          .filter((s) => s.type !== "ins")
          .map((s) => s.text)
          .join("");
        expect(baseText(result)).toBe(expectedBase);
      });

      it("reconstructs the revised column from every non-del segment text", () => {
        const expectedRevised = segments
          .filter((s) => s.type !== "del")
          .map((s) => s.text)
          .join("");
        expect(revisedText(result)).toBe(expectedRevised);
      });

      it("keeps the base column contiguous (starts at 0, no gaps or overlaps)", () => {
        assertColumnContiguity(result.hunks, "base");
      });

      it("keeps the revised column contiguous (starts at 0, no gaps or overlaps)", () => {
        assertColumnContiguity(result.hunks, "revised");
      });

      it("obeys the null rules for every hunk type", () => {
        for (const hunk of result.hunks) {
          if (hunk.type === "del") {
            expect(hunk.revisedRange).toBeNull();
            expect(hunk.baseRange).not.toBeNull();
          } else if (hunk.type === "ins") {
            expect(hunk.baseRange).toBeNull();
            expect(hunk.revisedRange).not.toBeNull();
          } else {
            expect(hunk.baseRange).not.toBeNull();
            expect(hunk.revisedRange).not.toBeNull();
          }
        }
      });
    });
  }
});

// ---------------------------------------------------------------------------
// Phase 6 — determinism.
// ---------------------------------------------------------------------------

describe("buildDiffJson — determinism", () => {
  it("produces deeply-equal, byte-identical JSON for identical inputs", () => {
    const segments: DiffSegment[] = [
      { type: "equal", text: "Hello " },
      { type: "del", text: "cruel " },
      { type: "ins", text: "kind " },
      { type: "equal", text: "world" },
    ];

    const a = buildDiffJson(segments);
    const b = buildDiffJson(segments);

    // Deep equality and order-stable serialization both hold for a pure builder.
    expect(a).toEqual(b);
    expect(JSON.stringify(a)).toBe(JSON.stringify(b));
  });
});
