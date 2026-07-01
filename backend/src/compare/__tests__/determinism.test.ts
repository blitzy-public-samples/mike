/**
 * Byte-identical DETERMINISM guard for the Document Compare engine.
 *
 * This is the single most important behavioral test of the whole feature. The
 * product requirement is an in-house, deterministic "Litera Compare": there is
 * NO AI in the diff path, so identical inputs plus identical `opts` MUST always
 * yield byte-identical output (AAP §0.1.1 "Deterministic means identical inputs
 * always yield byte-identical output — there is no AI in the diff path"; §0.6.2).
 *
 * The test drives the public engine entry {@link runComparison} MULTIPLE times
 * with the SAME arguments and asserts two things every run:
 *   1. `redlineBytes` are byte-for-byte identical (`Buffer.equals` — exact), and
 *   2. the structured `diff` objects deep-equal (`toEqual`).
 *
 * It is the regression guard for the empirically-verified JSZip entry-date fix
 * in `emitTrackedChanges.ts`: JSZip stamps wall-clock time (`new Date()`) into
 * every zip entry by default, so without normalizing each entry's `date` to a
 * fixed value derived from `opts.date` — and writing `word/document.xml` with an
 * explicit `{ date }` at a pinned DEFLATE level — the emitted `.docx` bytes drift
 * from run to run even for identical inputs. If this suite ever fails, that
 * timestamp normalization is the most likely culprit.
 *
 * Strict isolation (AAP §0.8.1): this file imports ONLY from the public engine
 * entry (`../index`), Node built-ins (`node:fs`, `node:path`), and `vitest`. It
 * never imports `jszip`, `fast-xml-parser`, `fast-diff`, `../../lib/*`,
 * `../../routes/*`, `node:crypto`, or any assistant/chat/doc-processing/tabular/
 * workflow code. Byte comparison uses `Buffer.equals` (a Node built-in that is
 * an exact, length-and-content check) — no hashing dependency is required.
 *
 * Build note: the explicit `vitest` import below is MANDATORY. `backend/tsconfig.json`
 * compiles `src/**` with `tsc` (the production `npm run build`) and does NOT
 * register the `vitest/globals` ambient types, so `describe` / `it` / `expect`
 * must be imported explicitly for this file to type-check. Relying on the
 * runtime `globals`/config alone would break `npm run build`. `__dirname` is
 * available because the backend compiles to CommonJS.
 */

// Explicit vitest import (MANDATORY — see the build note above; do NOT rely on
// vitest.config.ts globals, which the `tsc` build does not observe).
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { runComparison } from "../index";

// ---------------------------------------------------------------------------
// Fixture loader
// ---------------------------------------------------------------------------
//
// The committed golden `.docx` pairs live alongside this file in `./fixtures/`.
// They are loaded synchronously via Node's `fs`; each `load(...)` call returns a
// FRESH `Buffer` (a new read of the file), which lets Phase 4 prove that
// determinism is a property of the byte CONTENT, not of a memoized object.
const fixturesDir = join(__dirname, "fixtures");
const load = (name: string): Buffer => readFileSync(join(fixturesDir, name));

// A FIXED `opts` is what makes determinism possible: a stable `date` (never
// `new Date()`) is threaded through the engine as both the `w:author` /
// `w:date` written on every tracked change AND the pinned timestamp on every
// output zip entry. Keeping `date` constant across runs is the precondition the
// byte-identity assertions below rely on.
const OPTS = { author: "Compare Bot", date: "2026-01-01T00:00:00Z" };

// The golden scenarios exercised for cross-fixture determinism (Phase 3): an
// in-place edit, a pure word insertion, a whole added paragraph, and a
// no-change (identical) document. Together they prove determinism is a property
// of the engine rather than an accident of one fixture.
const SCENARIOS = [
  "mixed-edit",
  "word-insert",
  "added-paragraph",
  "identical",
] as const;

// ---------------------------------------------------------------------------
// Phase 2 — Byte-identical redline across repeated runs (the core guard)
// ---------------------------------------------------------------------------
describe("runComparison — byte-identical output across repeated runs", () => {
  it("produces identical redlineBytes and diff across THREE runs (mixed-edit)", async () => {
    const base = load("mixed-edit.base.docx");
    const revised = load("mixed-edit.revised.docx");

    // Three runs with the SAME arguments. Transitivity of equality across the
    // adjacent pairs (r1==r2, r2==r3) implies all three are identical.
    const r1 = await runComparison(base, revised, OPTS);
    const r2 = await runComparison(base, revised, OPTS);
    const r3 = await runComparison(base, revised, OPTS);

    // Coarse length signal first (clearer failure message if sizes diverge),
    // then the authoritative exact byte-identity check via Buffer.equals.
    expect(r1.redlineBytes.length).toBe(r2.redlineBytes.length);
    expect(r1.redlineBytes.equals(r2.redlineBytes)).toBe(true);
    expect(r2.redlineBytes.equals(r3.redlineBytes)).toBe(true);

    // The structured diff must be deep-equal across runs as well.
    expect(r1.diff).toEqual(r2.diff);
    expect(r2.diff).toEqual(r3.diff);
  });
});

// ---------------------------------------------------------------------------
// Phase 3 — Determinism holds across multiple scenarios
// ---------------------------------------------------------------------------
describe("runComparison — determinism across multiple scenarios", () => {
  for (const name of SCENARIOS) {
    it(`is byte-identical and diff-stable for '${name}' across two runs`, async () => {
      const base = load(`${name}.base.docx`);
      const revised = load(`${name}.revised.docx`);

      const a = await runComparison(base, revised, OPTS);
      const b = await runComparison(base, revised, OPTS);

      expect(a.redlineBytes.equals(b.redlineBytes)).toBe(true);
      expect(a.diff).toEqual(b.diff);
    });
  }
});

// ---------------------------------------------------------------------------
// Phase 4 — Fresh buffer reads do not affect determinism
// ---------------------------------------------------------------------------
describe("runComparison — determinism depends on byte content, not object identity", () => {
  it("is byte-identical when inputs are re-read into NEW Buffer instances", async () => {
    // Each `load(...)` returns a distinct Buffer object read afresh from disk;
    // no Buffer instance is shared between the two runs. Identical bytes must
    // still yield an identical redline, proving content-determinism rather than
    // accidental memoization of a Buffer object.
    const a = await runComparison(
      load("mixed-edit.base.docx"),
      load("mixed-edit.revised.docx"),
      OPTS,
    );
    const b = await runComparison(
      load("mixed-edit.base.docx"),
      load("mixed-edit.revised.docx"),
      OPTS,
    );

    expect(a.redlineBytes.equals(b.redlineBytes)).toBe(true);
    expect(a.diff).toEqual(b.diff);
  });
});

// ---------------------------------------------------------------------------
// Phase 5 — `opts` genuinely influences the output (sanity check)
// ---------------------------------------------------------------------------
describe("runComparison — opts.date is genuinely consumed", () => {
  it("produces DIFFERENT bytes for a different opts.date", async () => {
    const base = load("mixed-edit.base.docx");
    const revised = load("mixed-edit.revised.docx");

    // Baseline with the fixed OPTS date.
    const a = await runComparison(base, revised, OPTS);

    // A DIFFERENT date changes both the emitted `w:date` attributes and the
    // normalized zip-entry timestamps, so the bytes MUST differ. This confirms
    // `opts.date` is truly threaded through the engine and not ignored — the
    // flip side of the determinism guarantee above. (Documents intent; the
    // four-year date gap is far larger than any timezone rounding.)
    const c = await runComparison(base, revised, {
      author: "Compare Bot",
      date: "2030-12-31T23:59:59Z",
    });

    expect(a.redlineBytes.equals(c.redlineBytes)).toBe(false);
  });
});
