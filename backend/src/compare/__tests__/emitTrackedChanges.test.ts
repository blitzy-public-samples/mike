/**
 * Unit tests for `backend/src/compare/emitTrackedChanges.ts` -- the native
 * Word tracked-changes emitter of the deterministic Document Compare / Redline
 * engine.
 *
 * `emitTrackedChanges` merges a parsed BASE `.docx`, the ordered paragraph
 * alignment, and the per-modified-pair word-diff segments into a single
 * `word/document.xml` carrying NATIVE Word tracked changes, and rezips a valid,
 * byte-deterministic `.docx` {@link Buffer}. These tests build a realistic
 * {@link EmitInput} from committed golden fixtures (the exact
 * parse -> normalize -> align -> word-diff pipeline the orchestrator runs),
 * call `emitTrackedChanges`, then REPARSE the emitted bytes with the engine's
 * OWN `parseDocx` and walk the resulting preserve-order node tree to assert the
 * emitted OOXML structures:
 *
 *   - the redline reparses cleanly (a valid zip + parseable `document.xml`);
 *   - insertions are `<w:ins>` and deletions are `<w:del>`, each carrying a
 *     `w:id` / `w:author` / `w:date`;
 *   - every tracked-change `w:id` is numeric, positive, unique, and assigned in
 *     ascending document order;
 *   - deleted text is emitted as `<w:delText xml:space="preserve">` (never a
 *     plain `<w:t>` inside a `<w:del>`);
 *   - a deleted paragraph mark is `<w:pPr><w:rPr><w:del/></w:rPr></w:pPr>`;
 *   - source `<w:rPr>` run formatting is carried onto emitted runs;
 *   - the output is byte-identical for identical inputs (the determinism
 *     contract) and changes when the caller's `date` changes;
 *   - an identical base/revised pair emits NO tracked changes.
 *
 * Strict isolation (AAP s0.8.1): this file imports ONLY from the engine modules
 * under test (`../emitTrackedChanges`, `../parseDocx`, `../normalize`,
 * `../alignParagraphs`, `../wordDiff`, and the `DiffSegment` type from
 * `../diffJson`), Node built-ins (`node:fs`, `node:path`), and `vitest`. It
 * never imports `jszip`, `fast-xml-parser`, `../../lib/*`, the orchestrator
 * `../index`, or any assistant/chat/doc-processing/tabular/workflow code, and it
 * inspects the parsed XML ONLY through the helpers exported by `parseDocx`
 * (`elName` / `elChildren` / `elAttrs`) -- never through a direct XML parser.
 *
 * Build note: the backend `tsconfig` compiles `src/**` (including this test) and
 * does NOT register `vitest/globals` types, so the vitest test-API symbols are
 * imported explicitly to keep `npm run build` (tsc) green under `strict`.
 *
 * Determinism note: every assertion uses the fixed `AUTHOR` / `DATE` constants
 * (never the wall clock), matching the engine's "author/date come from the
 * caller, never `new Date()`" contract.
 */

import { describe, it, expect, beforeAll } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";

import { emitTrackedChanges, type EmitInput } from "../emitTrackedChanges";
import { parseDocx, elName, elChildren, elAttrs } from "../parseDocx";
import { normalizeParagraphs } from "../normalize";
import { alignParagraphs } from "../alignParagraphs";
import { wordDiff } from "../wordDiff";
import type { DiffSegment } from "../diffJson";

// ---------------------------------------------------------------------------
// Fixture loader
// ---------------------------------------------------------------------------
//
// Fixtures are committed binary `.docx` archives that live alongside this test
// in `./fixtures/`. They are loaded synchronously via `fs`; a missing fixture
// throws here, which is the correct loud-failure signal (the golden files are a
// hard prerequisite). `__dirname` is available because the backend compiles to
// CommonJS (no `"type": "module"`).
const fixturesDir = join(__dirname, "fixtures");
const load = (name: string): Buffer => readFileSync(join(fixturesDir, name));

// ---------------------------------------------------------------------------
// Recursive node walker (built ONLY on the exported parseDocx helpers)
// ---------------------------------------------------------------------------
//
// `collect` returns every element in the subtree whose tag name equals `name`,
// in document (depth-first, pre-order) order. It is built exclusively on
// `elName` / `elChildren`, so it never touches a raw XML parser and gracefully
// ignores text / attribute nodes (for which `elChildren` yields `[]`).
function collect(
  nodes: ReturnType<typeof elChildren>,
  name: string,
): ReturnType<typeof elChildren> {
  const out: ReturnType<typeof elChildren>[number][] = [];
  for (const node of nodes) {
    if (elName(node) === name) {
      out.push(node);
    }
    out.push(...collect(elChildren(node), name));
  }
  return out;
}

// ---------------------------------------------------------------------------
// EmitInput builder -- mirrors the orchestrator (backend/src/compare/index.ts)
// ---------------------------------------------------------------------------
//
// `segmentsByPair` is keyed by BASE paragraph index. This is not a guess: the
// emitter under test looks the map up with the alignment op's `baseIndex`
// (`emitTrackedChanges.ts`: `const i = op.baseIndex; segmentsByPair.get(i)`),
// and its `EmitInput` doc explicitly states the map is "keyed by BASE paragraph
// index". `normalizeParagraphs(doc)` is produced 1:1 with `doc.paragraphs`, so
// `baseNorm[op.baseIndex]` corresponds to `base.paragraphs[op.baseIndex]`.
// Entries are set only for `equal` pairs (both indices non-null); because
// `alignParagraphs` classifies a pair `equal` only when the normalized texts
// are byte-equal, those word-diffs are single `equal` segments -- real edits
// surface as whole-paragraph `del` + `ins` ops. Keying by `baseIndex` keeps
// this unit test faithful to the production wiring the emitter expects.
async function buildEmitInput(
  baseName: string,
  revisedName: string,
  author: string,
  date: string,
): Promise<EmitInput> {
  const base = await parseDocx(load(baseName));
  const revised = await parseDocx(load(revisedName));

  const baseNorm = normalizeParagraphs(base);
  const revNorm = normalizeParagraphs(revised);

  const align = alignParagraphs(baseNorm, revNorm);

  const segmentsByPair = new Map<number, DiffSegment[]>();
  align.forEach((op) => {
    if (
      op.type === "equal" &&
      op.baseIndex !== null &&
      op.revisedIndex !== null
    ) {
      // Key by BASE paragraph index -- the contract the emitter consumes.
      segmentsByPair.set(
        op.baseIndex,
        wordDiff(baseNorm[op.baseIndex].text, revNorm[op.revisedIndex].text),
      );
    }
  });

  return { base, align, segmentsByPair, revised, author, date };
}

// Fixed, clock-independent tracked-change identity used across all assertions.
const AUTHOR = "Compare Bot";
const DATE = "2026-01-01T00:00:00Z";

// ---------------------------------------------------------------------------
// Shared mixed-edit emission (computed once)
// ---------------------------------------------------------------------------
//
// The `mixed-edit` fixture pair (base: "The <b>quick</b> brown fox jumps.";
// revised: "The slow brown fox leaps.") exercises an insertion, a deletion, a
// deleted `<w:delText>`, carried `<w:rPr>` (the bold "quick"), and unique ids in
// a single emitted redline, so most phases below assert against ONE emission of
// it. It is produced once here (`Awaited<ReturnType<typeof parseDocx>>` avoids
// importing the `ParsedDocx` type name while staying fully typed).
let mixedBytes: Buffer;
let mixedReparsed!: Awaited<ReturnType<typeof parseDocx>>;

beforeAll(async () => {
  mixedBytes = await emitTrackedChanges(
    await buildEmitInput(
      "mixed-edit.base.docx",
      "mixed-edit.revised.docx",
      AUTHOR,
      DATE,
    ),
  );
  mixedReparsed = await parseDocx(mixedBytes);
});

// ===========================================================================
// Phase 2 -- Emitted redline is valid OOXML (reparse round-trip)
// ===========================================================================
describe("emitTrackedChanges: emitted redline is valid OOXML", () => {
  it("returns a non-empty Buffer", () => {
    expect(Buffer.isBuffer(mixedBytes)).toBe(true);
    expect(mixedBytes.length).toBeGreaterThan(0);
  });

  it("reparses cleanly with the engine's own parseDocx", () => {
    // No throw from parseDocx above (in beforeAll) => a valid zip whose
    // `word/document.xml` parses; a non-empty tree confirms real content.
    expect(mixedReparsed.tree.length).toBeGreaterThan(0);
    expect(mixedReparsed.paragraphs.length).toBeGreaterThan(0);
  });
});

// ===========================================================================
// Phase 3 -- <w:ins> / <w:del> presence + attribute correctness
// ===========================================================================
describe("emitTrackedChanges: <w:ins> / <w:del> presence and attributes", () => {
  it("emits at least one insertion and one deletion for a mixed edit", () => {
    const inss = collect(mixedReparsed.tree, "w:ins");
    const dels = collect(mixedReparsed.tree, "w:del");
    expect(inss.length).toBeGreaterThan(0);
    expect(dels.length).toBeGreaterThan(0);
  });

  it("stamps every tracked change with the caller's author, date, and an id", () => {
    const inss = collect(mixedReparsed.tree, "w:ins");
    const dels = collect(mixedReparsed.tree, "w:del");
    // A `w:del` matched here may be the deleted-paragraph-mark `w:del` nested in
    // a `w:rPr`; it too carries id/author/date, so the assertion holds for all.
    for (const node of [...inss, ...dels]) {
      const attrs = elAttrs(node);
      expect(attrs["@_w:author"]).toBe(AUTHOR);
      expect(attrs["@_w:date"]).toBe(DATE);
      expect(attrs["@_w:id"]).toBeDefined();
    }
  });
});

// ===========================================================================
// Phase 4 -- w:id uniqueness + monotonicity
// ===========================================================================
describe("emitTrackedChanges: tracked-change w:id values", () => {
  it("are numeric, positive, unique, and ascending in document order", () => {
    // Collect <w:ins> AND <w:del> in a SINGLE depth-first, pre-order walk so the
    // resulting sequence is in TRUE document order. Collecting all <w:ins> first
    // and all <w:del> second groups nodes by element type, not document
    // position: within a replace the emitter writes <w:del> before its paired
    // <w:ins>, so a type-grouped list is intentionally NOT ascending even though
    // the ids are assigned monotonically in document order.
    const collectChanges = (
      nodes: ReturnType<typeof elChildren>,
    ): ReturnType<typeof elChildren> => {
      const out: ReturnType<typeof elChildren>[number][] = [];
      for (const node of nodes) {
        const tag = elName(node);
        if (tag === "w:ins" || tag === "w:del") out.push(node);
        out.push(...collectChanges(elChildren(node)));
      }
      return out;
    };
    const ids = collectChanges(mixedReparsed.tree).map((node) =>
      Number(elAttrs(node)["@_w:id"]),
    );

    // Every id parses to a finite number.
    expect(ids.every((n) => Number.isFinite(n))).toBe(true);
    // Ids are unique across the whole document.
    expect(new Set(ids).size).toBe(ids.length);
    // Ids are positive (they start above the largest pre-existing tracked id).
    expect(ids.every((n) => n >= 1)).toBe(true);
    // Ids are assigned monotonically, so their document-order sequence is
    // already sorted ascending.
    const inDocOrder = ids.slice();
    const ascending = [...ids].sort((a, b) => a - b);
    expect(inDocOrder).toEqual(ascending);
  });
});

// ===========================================================================
// Phase 5 -- Deletions use <w:delText xml:space="preserve"> (never <w:t>)
// ===========================================================================
describe("emitTrackedChanges: deletions emit <w:delText>", () => {
  it("emits at least one <w:delText>, each preserving whitespace", () => {
    const delTexts = collect(mixedReparsed.tree, "w:delText");
    expect(delTexts.length).toBeGreaterThan(0);
    for (const dt of delTexts) {
      expect(elAttrs(dt)["@_xml:space"]).toBe("preserve");
    }
  });

  it("never nests a plain <w:t> inside a <w:del> (deleted text is <w:delText>)", () => {
    const dels = collect(mixedReparsed.tree, "w:del");
    // The load-bearing property: deleted content is carried as <w:delText>, so
    // no <w:t> should appear anywhere beneath a <w:del>.
    let wtUnderDel = 0;
    for (const del of dels) {
      wtUnderDel += collect(elChildren(del), "w:t").length;
    }
    expect(wtUnderDel).toBe(0);
  });

  it("emits a preserved <w:delText> for a word-level deletion fixture", async () => {
    // `word-delete` (base "The quick brown fox." -> revised "The quick fox.")
    // exercises the deletion path on a second, independent fixture.
    const bytes = await emitTrackedChanges(
      await buildEmitInput(
        "word-delete.base.docx",
        "word-delete.revised.docx",
        AUTHOR,
        DATE,
      ),
    );
    const reparsed = await parseDocx(bytes);
    const delTexts = collect(reparsed.tree, "w:delText");
    expect(delTexts.length).toBeGreaterThan(0);
    for (const dt of delTexts) {
      expect(elAttrs(dt)["@_xml:space"]).toBe("preserve");
    }
  });
});

// ===========================================================================
// Phase 6 -- Deleted paragraph mark (<w:pPr><w:rPr><w:del/></w:rPr></w:pPr>)
// ===========================================================================
describe("emitTrackedChanges: deleted paragraph mark", () => {
  it("marks a wholly-deleted paragraph's mark via w:pPr>w:rPr>w:del", async () => {
    // `deleted-paragraph` removes the middle of three paragraphs, so the
    // emitter must delete that paragraph's end-of-paragraph mark.
    const bytes = await emitTrackedChanges(
      await buildEmitInput(
        "deleted-paragraph.base.docx",
        "deleted-paragraph.revised.docx",
        AUTHOR,
        DATE,
      ),
    );
    const reparsed = await parseDocx(bytes);

    const pprs = collect(reparsed.tree, "w:pPr");
    const hasDeletedParaMark = pprs.some((ppr) =>
      collect(elChildren(ppr), "w:rPr").some(
        (rpr) => collect(elChildren(rpr), "w:del").length > 0,
      ),
    );
    expect(hasDeletedParaMark).toBe(true);
  });
});

// ===========================================================================
// Phase 7 -- Carried source <w:rPr> onto emitted tracked runs
// ===========================================================================
describe("emitTrackedChanges: carried source <w:rPr>", () => {
  it("carries a source run's <w:rPr> onto a run inside a tracked change", () => {
    // `mixed-edit`'s base wraps "quick" in <w:rPr><w:b/>, which is deleted --
    // so at least one <w:r> beneath a <w:ins>/<w:del> must carry a <w:rPr>.
    const trackedRuns = [
      ...collect(mixedReparsed.tree, "w:ins"),
      ...collect(mixedReparsed.tree, "w:del"),
    ].flatMap((node) => collect(elChildren(node), "w:r"));

    const withRpr = trackedRuns.filter(
      (run) => collect(elChildren(run), "w:rPr").length > 0,
    );
    expect(withRpr.length).toBeGreaterThan(0);

    // And where a run carries <w:rPr>, it is the FIRST child (the OOXML-required
    // position), proving the carry-and-order behavior.
    for (const run of withRpr) {
      expect(elName(elChildren(run)[0])).toBe("w:rPr");
    }
  });
});

// ===========================================================================
// Phase 8 -- Emit-level determinism (byte-identical with the same opts)
// ===========================================================================
describe("emitTrackedChanges: byte-level determinism", () => {
  it("produces byte-identical output for identical inputs", async () => {
    const a = await emitTrackedChanges(
      await buildEmitInput(
        "mixed-edit.base.docx",
        "mixed-edit.revised.docx",
        AUTHOR,
        DATE,
      ),
    );
    const b = await emitTrackedChanges(
      await buildEmitInput(
        "mixed-edit.base.docx",
        "mixed-edit.revised.docx",
        AUTHOR,
        DATE,
      ),
    );
    // Byte-identical output guards the zip-entry-timestamp + DEFLATE-level
    // normalization: no wall-clock time may leak into the archive metadata.
    expect(a.equals(b)).toBe(true);
  });

  it("produces different output when the caller's date changes", async () => {
    const a = await emitTrackedChanges(
      await buildEmitInput(
        "mixed-edit.base.docx",
        "mixed-edit.revised.docx",
        AUTHOR,
        DATE,
      ),
    );
    const c = await emitTrackedChanges(
      await buildEmitInput(
        "mixed-edit.base.docx",
        "mixed-edit.revised.docx",
        AUTHOR,
        "2027-02-02T00:00:00Z",
      ),
    );
    // Sanity that `date` flows through into the output (as `w:date` and the
    // pinned zip timestamps), so a real edit to opts is observable.
    expect(a.equals(c)).toBe(false);
  });
});

// ===========================================================================
// Phase 9 -- Identical-document path (no spurious tracked changes)
// ===========================================================================
describe("emitTrackedChanges: identical documents", () => {
  it("emits no tracked changes when base and revised are identical", async () => {
    const bytes = await emitTrackedChanges(
      await buildEmitInput(
        "identical.base.docx",
        "identical.revised.docx",
        AUTHOR,
        DATE,
      ),
    );
    const reparsed = await parseDocx(bytes);
    expect(collect(reparsed.tree, "w:ins").length).toBe(0);
    expect(collect(reparsed.tree, "w:del").length).toBe(0);
  });
});
