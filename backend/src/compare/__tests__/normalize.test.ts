/**
 * Unit tests for `backend/src/compare/normalize.ts` -- the accept-all
 * normalization step of the deterministic document-compare engine.
 *
 * `normalizeParagraphs(doc)` reduces every parsed body paragraph to its
 * "accepted view": pre-existing insertions (`<w:ins>`) are ACCEPTED (their inner
 * `<w:r>` runs become normal runs, so their text is kept) and pre-existing
 * deletions (`<w:del>`) are DISCARDED (their `<w:delText>` content is excluded).
 * The result is the plaintext a reader sees when Word / Google Docs display the
 * document in "all changes accepted" view. This is the prompt's "normalize
 * first / accept-all on base" step, which lets the engine produce a clean
 * two-way diff even when an input already carries redlines from a prior round.
 *
 * These tests exercise that contract against committed binary `.docx` golden
 * fixtures loaded from `./fixtures/` via Node's `fs`, parsed with the engine's
 * own `parseDocx` (never a direct XML-parser import):
 *   - no-op behavior on a CLEAN document (no tracked changes): normalized text
 *     equals the parsed paragraph text, 1:1 and in order;
 *   - accept-all behavior on a document that ALREADY carries tracked changes:
 *     `<w:ins>` text kept, `<w:del>` text dropped, plain runs passed through;
 *   - determinism (repeated calls are identical) and the 1:1 paragraph mapping.
 *
 * Strict isolation (AAP s0.8.1): this file imports ONLY from the engine under
 * test (`../normalize`, `../parseDocx`), Node built-ins (`node:fs`, `node:path`),
 * and `vitest`. It never imports `jszip`, `fast-xml-parser`, `../../lib/*`,
 * `../../routes/*`, or any assistant/chat/doc-processing/tabular/workflow code.
 *
 * Build note: the backend `tsconfig` compiles `src/**` (including this test)
 * and does NOT register `vitest/globals` types, so the vitest globals are
 * imported explicitly to keep `npm run build` (tsc) green under `strict`.
 */

import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseDocx } from "../parseDocx";
import { normalizeParagraphs } from "../normalize";

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
// paragraphs, in order, with no tracked changes. Because the fixture carries no
// `<w:ins>`/`<w:del>`, accept-all normalization is a no-op and must reproduce
// these strings verbatim.
const IDENTICAL_BASE_PARAGRAPHS: readonly string[] = [
    "The agreement is effective today.",
    "Each party shall perform its duties.",
    "This is the final clause.",
];

// Known, exact accepted-view content of the `tracked-input.base.docx` fixture,
// which ALREADY carries tracked changes. Its first paragraph's body children
// are, in document order: `<w:ins>` wrapping "New ", `<w:del>` wrapping "Old ",
// then a plain run "tail". Accept-all therefore keeps "New ", drops "Old ", and
// keeps "tail" -> "New tail" across exactly two accepted-view runs. Its second
// paragraph is a plain (clean) run reading "Second paragraph.".
const TRACKED_ACCEPTED_FIRST_TEXT = "New tail";
const TRACKED_ACCEPTED_FIRST_RUNS: readonly string[] = ["New ", "tail"];
const TRACKED_SECOND_PARAGRAPH = "Second paragraph.";

// ---------------------------------------------------------------------------
// Phase 2 -- No-op on a clean document
// ---------------------------------------------------------------------------

describe("normalizeParagraphs: no-op on a clean document", () => {
    it("returns one normalized paragraph per body paragraph (1:1, in order)", async () => {
        const doc = await parseDocx(load("identical.base.docx"));
        const norm = normalizeParagraphs(doc);

        // Accept-all produces exactly one NormalizedParagraph per parsed body
        // paragraph, preserving document order.
        expect(norm.length).toBe(doc.paragraphs.length);
    });

    it("leaves the paragraph text unchanged (no ins/del to resolve)", async () => {
        const doc = await parseDocx(load("identical.base.docx"));
        const norm = normalizeParagraphs(doc);

        // With no `<w:ins>`/`<w:del>` wrappers, every accepted-view run is a
        // direct `<w:r>` contribution, so the normalized text matches the
        // parsed paragraph text exactly, paragraph for paragraph.
        expect(norm.map((n) => n.text)).toEqual(
            doc.paragraphs.map((p) => p.text),
        );
    });

    it("matches the known clean-fixture literals", async () => {
        const doc = await parseDocx(load("identical.base.docx"));
        const norm = normalizeParagraphs(doc);

        expect(norm.map((n) => n.text)).toEqual([...IDENTICAL_BASE_PARAGRAPHS]);
    });
});

// ---------------------------------------------------------------------------
// Phase 3 -- Accept-all on a doc with pre-existing tracked changes
// ---------------------------------------------------------------------------

describe("normalizeParagraphs: accept-all on a doc with pre-existing tracked changes", () => {
    it("accepts the insertion and drops the deletion in paragraph 1", async () => {
        const tdoc = await parseDocx(load("tracked-input.base.docx"));
        const tnorm = normalizeParagraphs(tdoc);

        // Insertion accepted ("New " kept), deletion discarded ("Old " gone),
        // trailing plain run kept ("tail") -> the accepted view is "New tail".
        expect(tnorm[0].text).toBe(TRACKED_ACCEPTED_FIRST_TEXT);
    });

    it("excludes the deleted text from the accepted view", async () => {
        const tdoc = await parseDocx(load("tracked-input.base.docx"));
        const tnorm = normalizeParagraphs(tdoc);

        // The `<w:delText>` content ("Old ") must not survive normalization.
        expect(tnorm[0].text).not.toContain("Old");
    });

    it("reconstructs the accepted-view text from the accepted-view runs", async () => {
        const tdoc = await parseDocx(load("tracked-input.base.docx"));
        const tnorm = normalizeParagraphs(tdoc);

        // Joining the accepted-view run texts must reproduce the paragraph's
        // accepted-view plaintext -- this is the robust proof of the accept/drop
        // semantics, independent of how runs are grouped.
        expect(tnorm[0].runs.map((r) => r.text).join("")).toBe(
            TRACKED_ACCEPTED_FIRST_TEXT,
        );
    });

    it("keeps the accepted insertion and the plain run as two accepted-view runs", async () => {
        const tdoc = await parseDocx(load("tracked-input.base.docx"));
        const tnorm = normalizeParagraphs(tdoc);

        // The `<w:ins>` wrapper unwraps to its single inner `<w:r>` ("New ") and
        // the trailing direct `<w:r>` ("tail") is kept as-is, yielding exactly
        // these two accepted-view runs in order. (If a future engine merges or
        // re-splits runs, the join assertion above still proves the semantics.)
        expect(tnorm[0].runs.map((r) => r.text)).toEqual([
            ...TRACKED_ACCEPTED_FIRST_RUNS,
        ]);
    });

    it("passes the clean second paragraph through unchanged", async () => {
        const tdoc = await parseDocx(load("tracked-input.base.docx"));
        const tnorm = normalizeParagraphs(tdoc);

        // The second paragraph carries no tracked changes, so it is emitted
        // verbatim by the accept-all pass.
        expect(tnorm[1].text).toBe(TRACKED_SECOND_PARAGRAPH);
    });
});

// ---------------------------------------------------------------------------
// Phase 4 -- Determinism + 1:1 mapping on the tracked-input doc
// ---------------------------------------------------------------------------

describe("normalizeParagraphs: determinism and 1:1 mapping", () => {
    it("produces identical output across repeated calls (no nondeterminism)", async () => {
        const tdoc = await parseDocx(load("tracked-input.base.docx"));
        const tnorm = normalizeParagraphs(tdoc);

        // `normalizeParagraphs` is pure (no I/O, clock, or randomness), so a
        // second invocation on the same parsed document yields the same
        // accepted-view text and the same run breakdown, paragraph for paragraph.
        const again = normalizeParagraphs(tdoc);
        expect(again.map((n) => n.text)).toEqual(tnorm.map((n) => n.text));
        expect(again.map((n) => n.runs.map((r) => r.text))).toEqual(
            tnorm.map((n) => n.runs.map((r) => r.text)),
        );
    });

    it("emits one normalized paragraph per parsed paragraph (1:1)", async () => {
        const tdoc = await parseDocx(load("tracked-input.base.docx"));
        const tnorm = normalizeParagraphs(tdoc);

        // The accept-all pass never adds or drops paragraphs -- it only rewrites
        // each paragraph's runs -- so the output length tracks the parsed input.
        expect(tnorm.length).toBe(tdoc.paragraphs.length);
    });
});
