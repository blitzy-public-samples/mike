/**
 * End-to-end integration tests for `backend/src/compare/index.ts` -- the public
 * orchestrator `runComparison(baseBytes, revisedBytes, opts)` of the
 * deterministic Document Compare / Redline engine.
 *
 * Unlike the sibling unit suites (which each exercise a single engine module in
 * isolation), this file drives the WHOLE pipeline -- parse -> accept-all
 * normalize -> paragraph align -> word diff -> emit tracked changes -> build
 * diff JSON -> validate -> return -- over the committed golden `.docx` fixture
 * PAIRS in `./fixtures/`, and asserts, per scenario, BOTH of the engine's two
 * artifacts:
 *   1. the structured `diff` JSON -- hunk presence (the right words land on the
 *      right side of the diff) and the column-reconstruction invariant; and
 *   2. the emitted `redlineBytes` -- the native Word tracked-change markup
 *      (`<w:ins>` / `<w:del>`), reparsed with the engine's OWN `parseDocx` and
 *      counted through the exported `elName` / `elChildren` / `elAttrs` helpers.
 *
 * Scenarios covered (the seven AAP-mandated fixture pairs, §0.2.3 / §0.6.2):
 *   word insert, word delete, mixed edit, added paragraph, deleted paragraph,
 *   identical (no change), and already-tracked input (accept-all end-to-end).
 *
 * Strict isolation (AAP §0.8.1): this file imports ONLY from the engine under
 * test (`../index`, `../parseDocx`), Node built-ins (`node:fs`, `node:path`),
 * and `vitest`. It never imports `jszip`, `fast-xml-parser`, `../../lib/*`,
 * `../../routes/*`, or any assistant/chat/doc-processing/tabular/workflow code,
 * and it inspects the emitted redline's XML ONLY through the helpers exported by
 * `parseDocx` -- never through a direct XML-parser import.
 *
 * Build note: the backend `tsconfig` compiles `src/**` (including this test) and
 * does NOT register `vitest/globals` types, so the vitest globals are imported
 * explicitly to keep `npm run build` (tsc) green under `strict`.
 *
 * Robustness note: the diff-presence checks use SUBSTRING (`toContain`) matching,
 * deliberately tolerant of how the engine snaps word boundaries or coalesces
 * adjacent same-type segments. Exact hunk-array assertions are intentionally the
 * job of the hand-traced `wordDiff.test.ts` / `diffJson.test.ts`; byte-identity
 * determinism is the job of `determinism.test.ts`. This file proves functional
 * correctness across the real scenarios, not the exact internal segmentation.
 */

import { describe, it, expect, beforeAll } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { runComparison } from "../index";
import { parseDocx, elName, elChildren, elAttrs } from "../parseDocx";

// ---------------------------------------------------------------------------
// Fixture loader
// ---------------------------------------------------------------------------
//
// Fixtures are committed binary `.docx` archives that live alongside this test
// in `./fixtures/`. They are loaded synchronously via `fs`; a missing fixture
// throws here, which is the correct loud-failure signal (the golden files are a
// hard prerequisite for these end-to-end tests). `__dirname` is available
// because the backend compiles to CommonJS (no `"type": "module"`).
const fixturesDir = join(__dirname, "fixtures");
const load = (name: string): Buffer => readFileSync(join(fixturesDir, name));

// Author / date threaded into the emitter. `date` MUST be non-empty (the
// orchestrator rejects an empty date and the emitter derives the output zip's
// timestamps from it); a fixed value keeps the run reproducible.
const OPTS = { author: "Compare Bot", date: "2026-01-01T00:00:00Z" };

// ---------------------------------------------------------------------------
// Fixture contract (base -> revised; the words below are what the assertions
// search for, and MUST match the committed fixtures/ folder spec):
//   word-insert       "The quick brown fox."          -> "The quick brown lazy fox."   (+ "lazy")
//   word-delete       "The quick brown fox."          -> "The quick fox."              (- "brown")
//   mixed-edit        "The quick brown fox jumps."     -> "The slow brown fox leaps."   (- quick/jumps, + slow/leaps)
//   added-paragraph   [Alpha, Gamma]                  -> [Alpha, Beta, Gamma]          (+ "Beta" paragraph)
//   deleted-paragraph [Alpha, Beta, Gamma]            -> [Alpha, Gamma]                (- "Beta" paragraph)
//   identical         [three clauses]                 -> [same three clauses]          (no change)
//   tracked-input     base already carries <w:ins>/<w:del>; accepted view == clean revised (no change)
// ---------------------------------------------------------------------------

/** Edit pairs whose comparison MUST surface at least one tracked change. */
const EDIT_PAIRS = [
  "word-insert",
  "word-delete",
  "mixed-edit",
  "added-paragraph",
  "deleted-paragraph",
] as const;

/** Pairs whose comparison MUST surface NO new changes (all-equal diff). */
const NO_CHANGE_PAIRS = ["identical", "tracked-input"] as const;

/** Every fixture pair, in a stable order. */
const ALL_PAIRS = [...EDIT_PAIRS, ...NO_CHANGE_PAIRS] as const;

// ---------------------------------------------------------------------------
// Read-side helpers (built ONLY on the exported parseDocx helpers)
// ---------------------------------------------------------------------------

/**
 * Recursively count, across an ordered preserve-order node array, how many
 * elements carry the given tag name (e.g. `"w:ins"` / `"w:del"`). Walks the
 * whole subtree via the engine's own `elName` / `elChildren`, so it never needs
 * a direct XML-parser import.
 */
function countTag(nodes: ReturnType<typeof elChildren>, name: string): number {
  let n = 0;
  for (const node of nodes) {
    if (elName(node) === name) n++;
    n += countTag(elChildren(node), name);
  }
  return n;
}

/**
 * Recursively collect the attribute maps of every element with the given tag
 * name, appending them (in document order) to `out`. Used to assert that the
 * caller-supplied `w:author` / `w:date` were threaded onto emitted changes.
 */
function collectAttrs(
  nodes: ReturnType<typeof elChildren>,
  name: string,
  out: Record<string, string>[],
): void {
  for (const node of nodes) {
    if (elName(node) === name) out.push(elAttrs(node));
    collectAttrs(elChildren(node), name, out);
  }
}

/** The exact resolved artifacts of a single `runComparison` invocation. */
type ComparisonResult = Awaited<ReturnType<typeof runComparison>>;
/** The structured-diff artifact (`{ hunks: [...] }`). */
type Diff = ComparisonResult["diff"];

/**
 * Concatenate (space-joined) the text of every hunk of a given `type`. Typed
 * structurally so this helper needs no diff-type import; used for tolerant
 * substring presence checks (`ins` side vs `del` side).
 */
const joinByType = (
  d: { hunks: { type: string; text: string }[] },
  t: string,
): string =>
  d.hunks
    .filter((h) => h.type === t)
    .map((h) => h.text)
    .join(" ");

/**
 * Reconstruct the BASE column plaintext from the ordered hunks: base column =
 * `equal` + `del` text (everything that is NOT a pure insertion).
 */
const reconstructBase = (d: Diff): string =>
  d.hunks
    .filter((h) => h.type !== "ins")
    .map((h) => h.text)
    .join("");

/**
 * Reconstruct the REVISED column plaintext from the ordered hunks: revised
 * column = `equal` + `ins` text (everything that is NOT a pure deletion).
 */
const reconstructRevised = (d: Diff): string =>
  d.hunks
    .filter((h) => h.type !== "del")
    .map((h) => h.text)
    .join("");

// ---------------------------------------------------------------------------
// Pre-warm: run the full pipeline for every pair ONCE, up front
// ---------------------------------------------------------------------------
//
// `runComparison` internally calls the (best-effort, never-fatal) LibreOffice
// render probe in `validate`, whose FIRST invocation can cold-start `soffice`.
// Computing all seven pairs inside a single generously-timed `beforeAll` keeps
// that cost out of the per-test 5s default and lets each `it` read a cached
// result instantly. The map is keyed by pair name; a missing entry is a loud
// programmer error (see `getResult`).
const results = new Map<string, ComparisonResult>();

beforeAll(async () => {
  for (const pair of ALL_PAIRS) {
    const result = await runComparison(
      load(`${pair}.base.docx`),
      load(`${pair}.revised.docx`),
      OPTS,
    );
    results.set(pair, result);
  }
}, 300_000);

/** Return the pre-computed comparison result for a pair (throws if absent). */
function getResult(pair: string): ComparisonResult {
  const result = results.get(pair);
  if (!result) {
    throw new Error(`runComparison result for "${pair}" was not precomputed`);
  }
  return result;
}

// ---------------------------------------------------------------------------
// Phase 2 -- Artifact shape + per-hunk validity invariants
// ---------------------------------------------------------------------------

describe("runComparison: artifact shape and per-hunk validity", () => {
  it("returns a non-empty Buffer redline and a hunks array for every pair", () => {
    for (const pair of ALL_PAIRS) {
      const { redlineBytes, diff } = getResult(pair);
      expect(Buffer.isBuffer(redlineBytes)).toBe(true);
      expect(redlineBytes.length).toBeGreaterThan(0);
      expect(Array.isArray(diff.hunks)).toBe(true);
    }
  });

  it("emits only well-formed hunks with type-consistent range nullability", () => {
    for (const pair of ALL_PAIRS) {
      const { diff } = getResult(pair);
      for (const h of diff.hunks) {
        expect(["ins", "del", "equal"]).toContain(h.type);
        expect(typeof h.text).toBe("string");
        // Range nullability is fully determined by the hunk type:
        //   del   -> base-only  (revisedRange null, baseRange present)
        //   ins   -> revised-only (baseRange null, revisedRange present)
        //   equal -> both columns (both ranges present)
        if (h.type === "del") {
          expect(h.revisedRange).toBeNull();
          expect(h.baseRange).not.toBeNull();
        } else if (h.type === "ins") {
          expect(h.baseRange).toBeNull();
          expect(h.revisedRange).not.toBeNull();
        } else {
          expect(h.baseRange).not.toBeNull();
          expect(h.revisedRange).not.toBeNull();
        }
      }
    }
  });

  it("produces a redline that reparses into a non-empty OOXML tree", async () => {
    for (const pair of ALL_PAIRS) {
      const { redlineBytes } = getResult(pair);
      const re = await parseDocx(redlineBytes);
      expect(re.tree.length).toBeGreaterThan(0);
    }
  });
});

// ---------------------------------------------------------------------------
// Phase 3 -- Per-scenario diff-JSON presence (substring-tolerant)
// ---------------------------------------------------------------------------

describe("runComparison: per-scenario diff-JSON presence", () => {
  it("word-insert: the inserted word appears on the ins side only", () => {
    const { diff } = getResult("word-insert");
    expect(joinByType(diff, "ins")).toContain("lazy");
    expect(joinByType(diff, "del")).not.toContain("lazy");
  });

  it("word-delete: the deleted word appears on the del side", () => {
    const { diff } = getResult("word-delete");
    expect(joinByType(diff, "del")).toContain("brown");
  });

  it("mixed-edit: inserted words on ins, deleted words on del", () => {
    const { diff } = getResult("mixed-edit");
    const ins = joinByType(diff, "ins");
    const del = joinByType(diff, "del");
    expect(ins).toContain("slow");
    expect(ins).toContain("leaps");
    expect(del).toContain("quick");
    expect(del).toContain("jumps");
    // A mixed edit must surface at least one hunk on each side.
    expect(diff.hunks.filter((h) => h.type === "ins").length).toBeGreaterThanOrEqual(1);
    expect(diff.hunks.filter((h) => h.type === "del").length).toBeGreaterThanOrEqual(1);
  });

  it("added-paragraph: the added paragraph text appears on the ins side", () => {
    const { diff } = getResult("added-paragraph");
    expect(joinByType(diff, "ins")).toContain("Beta");
  });

  it("deleted-paragraph: the deleted paragraph text appears on the del side", () => {
    const { diff } = getResult("deleted-paragraph");
    expect(joinByType(diff, "del")).toContain("Beta");
  });
});

// ---------------------------------------------------------------------------
// Phase 4 -- Column reconstruction invariant
// ---------------------------------------------------------------------------
//
// The ordered hunks must rebuild each column consistently: the base column is
// `equal` + `del` text and the revised column is `equal` + `ins` text. These
// are CONTAINS checks (not strict equality) because the whole-document text
// assembly/join is an engine-internal detail; the meaningful invariant is that
// base-only tokens land in the base reconstruction and revised-only tokens land
// in the revised reconstruction (and NOT the other way around).

describe("runComparison: column reconstruction invariant", () => {
  it("word-insert: revised column gains 'lazy'; base column does not", () => {
    const { diff } = getResult("word-insert");
    expect(reconstructRevised(diff)).toContain("lazy");
    expect(reconstructBase(diff)).not.toContain("lazy");
  });

  it("word-delete: base column keeps 'brown'; revised column drops it", () => {
    const { diff } = getResult("word-delete");
    expect(reconstructBase(diff)).toContain("brown");
    expect(reconstructRevised(diff)).not.toContain("brown");
  });

  it("mixed-edit: base carries the base-only words, revised carries the revised-only words", () => {
    const { diff } = getResult("mixed-edit");
    const baseText = reconstructBase(diff);
    const revisedText = reconstructRevised(diff);
    expect(baseText).toContain("quick");
    expect(baseText).toContain("jumps");
    expect(revisedText).toContain("slow");
    expect(revisedText).toContain("leaps");
    // The cross-checks prove the tokens landed on the CORRECT side.
    expect(revisedText).not.toContain("quick");
    expect(baseText).not.toContain("slow");
  });

  it("added-paragraph: revised column gains 'Beta'; base column does not", () => {
    const { diff } = getResult("added-paragraph");
    expect(reconstructRevised(diff)).toContain("Beta");
    expect(reconstructBase(diff)).not.toContain("Beta");
  });

  it("deleted-paragraph: base column keeps 'Beta'; revised column drops it", () => {
    const { diff } = getResult("deleted-paragraph");
    expect(reconstructBase(diff)).toContain("Beta");
    expect(reconstructRevised(diff)).not.toContain("Beta");
  });
});

// ---------------------------------------------------------------------------
// Phase 5 -- Redline tracked-change markup per edit scenario (DIRECTIONAL)
// ---------------------------------------------------------------------------
//
// Every edit pair must carry tracked markup (`w:ins` + `w:del` > 0). The
// per-side checks are DIRECTIONAL ONLY: an insertion scenario must emit
// `w:ins`, and a deletion scenario must emit `w:del`. We intentionally do NOT
// assert the ABSENCE of the opposite tag: a single-paragraph edit
// (word-insert / word-delete) is aligned as a whole-paragraph delete + insert,
// so its redline legitimately carries BOTH `<w:ins>` and `<w:del>`. Asserting
// e.g. `w:del === 0` for word-insert would contradict the real emitter.

describe("runComparison: redline tracked-change markup (edit pairs)", () => {
  it("every edit pair emits at least one tracked change", async () => {
    for (const pair of EDIT_PAIRS) {
      const { redlineBytes } = getResult(pair);
      const re = await parseDocx(redlineBytes);
      const total = countTag(re.tree, "w:ins") + countTag(re.tree, "w:del");
      expect(total).toBeGreaterThan(0);
    }
  });

  it("insertion scenarios emit <w:ins> markup", async () => {
    // Pairs whose revised side introduces content.
    const insertionPairs = ["word-insert", "mixed-edit", "added-paragraph"];
    for (const pair of insertionPairs) {
      const re = await parseDocx(getResult(pair).redlineBytes);
      expect(countTag(re.tree, "w:ins")).toBeGreaterThan(0);
    }
  });

  it("deletion scenarios emit <w:del> markup", async () => {
    // Pairs whose base side loses content.
    const deletionPairs = ["word-delete", "mixed-edit", "deleted-paragraph"];
    for (const pair of deletionPairs) {
      const re = await parseDocx(getResult(pair).redlineBytes);
      expect(countTag(re.tree, "w:del")).toBeGreaterThan(0);
    }
  });
});

// ---------------------------------------------------------------------------
// Phase 6 -- No-change scenarios (identical + accept-all end-to-end)
// ---------------------------------------------------------------------------

describe("runComparison: no-change scenarios", () => {
  it("identical: all-equal diff and a redline with ZERO tracked markup", async () => {
    const { redlineBytes, diff } = getResult("identical");
    expect(diff.hunks.every((h) => h.type === "equal")).toBe(true);
    // The base is clean (no pre-existing tracked changes), so the emitted
    // redline must carry none either.
    const re = await parseDocx(redlineBytes);
    expect(countTag(re.tree, "w:ins")).toBe(0);
    expect(countTag(re.tree, "w:del")).toBe(0);
  });

  it("tracked-input: accept-all yields an all-equal diff and adds ZERO new changes", async () => {
    const { redlineBytes, diff } = getResult("tracked-input");

    // Accept-all end-to-end proof: the base ALREADY carries tracked changes
    // whose ACCEPTED view equals the clean revised document. After the engine's
    // normalize/accept-all step the two are equal, so the comparison must
    // surface NO new changes -- every hunk is `equal`. If instead any `ins`/`del`
    // hunk appeared, normalization regressed.
    expect(diff.hunks.every((h) => h.type === "equal")).toBe(true);

    // The redline is emitted from the BASE tree, which retains the base's OWN
    // pre-existing `<w:ins>`/`<w:del>` markup. The engine must add NONE of its
    // own, so the emitted markup count equals the base fixture's own count --
    // i.e. ZERO comparison-introduced changes (NOT necessarily zero markup).
    const base = await parseDocx(load("tracked-input.base.docx"));
    const re = await parseDocx(redlineBytes);
    expect(countTag(re.tree, "w:ins")).toBe(countTag(base.tree, "w:ins"));
    expect(countTag(re.tree, "w:del")).toBe(countTag(base.tree, "w:del"));
  });
});

// ---------------------------------------------------------------------------
// Phase 7 -- Author / date propagation (spot check)
// ---------------------------------------------------------------------------

describe("runComparison: author/date propagation", () => {
  it("threads opts.author and opts.date onto emitted tracked changes (mixed-edit)", async () => {
    const { redlineBytes } = getResult("mixed-edit");
    const re = await parseDocx(redlineBytes);

    // Gather the attributes of every emitted change wrapper and confirm the
    // caller's author/date flowed through the orchestrator into emission.
    const attrs: Record<string, string>[] = [];
    collectAttrs(re.tree, "w:ins", attrs);
    collectAttrs(re.tree, "w:del", attrs);
    expect(attrs.length).toBeGreaterThan(0);
    expect(
      attrs.some(
        (a) => a["@_w:author"] === OPTS.author && a["@_w:date"] === OPTS.date,
      ),
    ).toBe(true);
  });
});

