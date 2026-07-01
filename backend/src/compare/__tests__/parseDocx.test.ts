/**
 * Unit tests for `backend/src/compare/parseDocx.ts` -- the read side of the
 * deterministic document-compare engine.
 *
 * `parseDocx` loads a `.docx`, reads `word/document.xml` in `preserveOrder`
 * mode, and produces an ordered list of body paragraphs (each carrying its
 * text-bearing runs and their preserved `<w:rPr>`), plus the raw parsed `tree`,
 * the loaded `zip` archive, and a whole-document `bodyPlainText` sanity string.
 *
 * These tests exercise that contract against committed binary `.docx` golden
 * fixtures loaded from `./fixtures/` via Node's `fs`:
 *   - ordered paragraph/run extraction and run/text consistency,
 *   - the `bodyPlainText` newline-join (clean docs only),
 *   - `<w:rPr>` preservation (a load-bearing property for the emitter),
 *   - the structural fields (`tree`, `zip`), and
 *   - fatal-parse behavior on a non-`.docx` buffer,
 *   - deterministic, stable parsing.
 *
 * Strict isolation (AAP s0.8.1): this file imports ONLY from the engine under
 * test (`../parseDocx`), Node built-ins (`node:fs`, `node:path`), and `vitest`.
 * It never imports `jszip`, `fast-xml-parser`, `../../lib/*`, `../../routes/*`,
 * or any assistant/chat/doc-processing/tabular/workflow code, and it inspects
 * parsed XML ONLY through the helpers exported by `parseDocx` (`elName`,
 * `elChildren`) -- never through a direct XML-parser import.
 *
 * Build note: the backend `tsconfig` compiles `src/**` (including this test)
 * and does NOT register `vitest/globals` types, so the vitest globals are
 * imported explicitly to keep `npm run build` (tsc) green under `strict`.
 *
 * Key behavioral note: `parseDocx` deliberately does NOT flatten
 * `<w:ins>`/`<w:del>` (accept-all normalization is `normalize.ts`'s single
 * responsibility). Consequently the text-equality invariants
 * (`runs.join("") === paragraph.text` and
 * `bodyPlainText === paragraphs.join("\n")`) are asserted ONLY against CLEAN
 * fixtures that carry no pre-existing tracked changes.
 */

import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseDocx, elName, elChildren } from "../parseDocx";

// ---------------------------------------------------------------------------
// Fixture loader
// ---------------------------------------------------------------------------
//
// Fixtures are committed binary `.docx` archives that live alongside this test
// in `./fixtures/`. They are loaded synchronously via `fs`; a missing fixture
// throws here, which is the correct loud-failure signal (the golden files are a
// hard prerequisite for these tests). `__dirname` is available because the
// backend compiles to CommonJS (no `"type": "module"`).
const fixturesDir = join(__dirname, "fixtures");
const load = (name: string): Buffer => readFileSync(join(fixturesDir, name));

// Known, exact content of the CLEAN `identical.base.docx` fixture: three body
// paragraphs, in order, with no tracked changes and no run properties. These
// literals mirror the fixtures/ folder spec and anchor the ordered-extraction
// and determinism assertions below.
const IDENTICAL_BASE_PARAGRAPHS: readonly string[] = [
  "The agreement is effective today.",
  "Each party shall perform its duties.",
  "This is the final clause.",
];

// The CLEAN `formatted.docx` fixture is a single paragraph whose runs are, in
// order: "Bold" (bold), " and " (plain), "italic" (italic), " text." (plain).
// Its concatenated plaintext is therefore this literal.
const FORMATTED_FIRST_PARAGRAPH = "Bold and italic text.";

// ---------------------------------------------------------------------------
// Phase 2 -- Ordered paragraph/run extraction (clean doc)
// ---------------------------------------------------------------------------

describe("parseDocx: ordered paragraph/run extraction (clean doc)", () => {
  it("extracts body paragraphs in document order with exact text", async () => {
    const doc = await parseDocx(load("identical.base.docx"));

    expect(Array.isArray(doc.paragraphs)).toBe(true);
    expect(doc.paragraphs.length).toBe(3);
    expect(doc.paragraphs.map((p) => p.text)).toEqual([
      ...IDENTICAL_BASE_PARAGRAPHS,
    ]);
  });

  it("returns a well-formed shape for every paragraph", async () => {
    const doc = await parseDocx(load("identical.base.docx"));

    for (const p of doc.paragraphs) {
      expect(typeof p.text).toBe("string");
      expect(Array.isArray(p.runs)).toBe(true);
      // The underlying `<w:p>` node is retained (by reference) for the
      // downstream tracked-changes emitter to rewrite.
      expect(p.node && typeof p.node === "object").toBe(true);
    }
  });

  it("keeps concatenated run text equal to paragraph text (no ins/del wrappers)", async () => {
    const doc = await parseDocx(load("identical.base.docx"));

    // Holds for a CLEAN doc: every run's text is a direct `<w:r>`/`<w:t>`
    // contribution, so joining the run texts reconstructs the paragraph
    // text exactly. (This would NOT hold for a doc with pre-existing
    // `<w:ins>`/`<w:del>`, which `parseDocx` purposely leaves un-flattened.)
    for (const p of doc.paragraphs) {
      expect(p.runs.map((r) => r.text).join("")).toBe(p.text);
    }
  });
});

// ---------------------------------------------------------------------------
// Phase 3 -- `bodyPlainText` join (clean doc only)
// ---------------------------------------------------------------------------

describe("parseDocx: bodyPlainText join (clean doc only)", () => {
  it("newline-joins the body paragraph texts", async () => {
    const doc = await parseDocx(load("identical.base.docx"));

    // `bodyPlainText` comes from `extractDocxBodyText`, which `\n`-joins the
    // flattened body paragraphs. On a CLEAN doc the flattened text equals
    // each paragraph's run-concatenated text, so the two agree exactly.
    expect(doc.bodyPlainText).toBe(
      doc.paragraphs.map((p) => p.text).join("\n"),
    );
  });
});

// ---------------------------------------------------------------------------
// Phase 4 -- `<w:rPr>` preservation (formatted doc)
// ---------------------------------------------------------------------------

describe("parseDocx: <w:rPr> preservation (formatted doc)", () => {
  it("carries the run properties node on at least one run", async () => {
    const fdoc = await parseDocx(load("formatted.docx"));

    const runsWithRpr = fdoc.paragraphs
      .flatMap((p) => p.runs)
      .filter((r) => r.rPr !== null);

    // At least one run must carry preserved properties -- the emitter later
    // clones this node onto tracked-change runs to keep formatting fidelity.
    expect(runsWithRpr.length).toBeGreaterThan(0);

    // The preserved reference is the `<w:rPr>` element itself...
    expect(elName(runsWithRpr[0].rPr)).toBe("w:rPr");

    // ...and it retains at least one recognizable formatting child. Presence
    // of any formatting element is sufficient; exact styling is not asserted.
    const kids = elChildren(runsWithRpr[0].rPr).map(elName);
    expect(
      kids.some(
        (n) =>
          n === "w:b" ||
          n === "w:i" ||
          n === "w:rFonts" ||
          n === "w:sz",
      ),
    ).toBe(true);
  });

  it("still concatenates runs to the known paragraph text", async () => {
    const fdoc = await parseDocx(load("formatted.docx"));

    // The fixture is CLEAN, so run concatenation reconstructs the paragraph
    // text regardless of the intervening `<w:rPr>` boundaries.
    expect(fdoc.paragraphs[0].text).toBe(FORMATTED_FIRST_PARAGRAPH);
    expect(
      fdoc.paragraphs[0].runs.map((r) => r.text).join(""),
    ).toBe(FORMATTED_FIRST_PARAGRAPH);
  });
});

// ---------------------------------------------------------------------------
// Phase 5 -- Structural fields + fatal-parse behavior
// ---------------------------------------------------------------------------

describe("parseDocx: structural fields and fatal-parse behavior", () => {
  it("returns the raw preserve-order tree and the loaded archive", async () => {
    const doc = await parseDocx(load("identical.base.docx"));

    // The full parsed tree is handed to the emitter to rebuild the doc.
    expect(Array.isArray(doc.tree)).toBe(true);
    expect(doc.tree.length).toBeGreaterThan(0);

    // The loaded JSZip archive is returned so the emitter can rezip the SAME
    // archive (preserving styles/rels/etc.).
    expect(doc.zip).toBeDefined();
  });

  it("rejects a non-docx buffer (parse failure is fatal)", async () => {
    // A buffer that is not a valid zip cannot be opened by JSZip (and has no
    // `word/document.xml`), so `parseDocx` must reject per its contract.
    await expect(
      parseDocx(Buffer.from("this is not a zip")),
    ).rejects.toThrow();
  });
});

// ---------------------------------------------------------------------------
// Phase 6 -- Determinism (parse is stable)
// ---------------------------------------------------------------------------

describe("parseDocx: determinism", () => {
  it("produces identical output for identical input", async () => {
    const d1 = await parseDocx(load("identical.base.docx"));
    const d2 = await parseDocx(load("identical.base.docx"));

    // Identical inputs must yield an identical ordered paragraph model and
    // identical whole-document plaintext -- there is no AI or nondeterminism
    // anywhere in the parse path.
    expect(d1.paragraphs.map((p) => p.text)).toEqual(
      d2.paragraphs.map((p) => p.text),
    );
    expect(d1.bodyPlainText).toBe(d2.bodyPlainText);
  });
});
