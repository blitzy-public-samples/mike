/**
 * Unit tests for `backend/src/compare/validate.ts` -- the post-emission
 * validation gate of the deterministic document-compare / redline engine
 * (AAP s0.6.2 "Build the diff JSON and validate", s0.8.1).
 *
 * `validate(redlineBytes)` performs two INDEPENDENT and asymmetric checks:
 *   1. Reparse (FATAL). The emitted bytes are reloaded as a zip and their
 *      `word/document.xml` is reparsed into a well-formed OOXML tree containing
 *      `<w:document><w:body>`. A failure here means the engine produced an
 *      invalid `.docx`, so the promise REJECTS -- a malformed redline must never
 *      be persisted to R2 or streamed to a user. `reparsed` is therefore always
 *      `true` whenever `validate` RESOLVES.
 *   2. LibreOffice render smoke test (SOFT / best-effort). The emitted bytes are
 *      optionally handed to `docxToPdf` (which `validate.ts` wraps internally --
 *      this test never imports `../../lib/convert`). `docxToPdf` throws when the
 *      LibreOffice/soffice binary is unavailable, and `validate` SWALLOWS that
 *      error and resolves with `renderChecked: false`. The render probe never
 *      fails validation.
 *
 * The asymmetry is deliberate and load-bearing, and this suite encodes it:
 *   - reparse == fatal  -> a corrupt buffer MUST reject;
 *   - render  == soft   -> the tests MUST tolerate `renderChecked` being either
 *     `true` (a machine WITH LibreOffice) or `false` (CI / any host WITHOUT it),
 *     and must NEVER require `renderChecked === true`. Encoding that tolerance is
 *     what keeps the compare suite green on runners that lack LibreOffice.
 *
 * A real, engine-produced redline (via `runComparison` over the committed
 * `mixed-edit` golden fixtures) is used as the primary positive input, making
 * this a meaningful end-to-end check that the engine's OWN output passes its OWN
 * validator -- rather than hand-crafting `.docx` bytes.
 *
 * Strict isolation (AAP s0.8.1): this file imports ONLY from the engine under
 * test (`../validate`, and `../index` for `runComparison`), Node built-ins
 * (`node:fs`, `node:path`), and `vitest`. It never imports `../../lib/convert`,
 * `jszip`, `libreoffice-convert`, or any assistant/chat/doc-processing/tabular/
 * workflow code path.
 *
 * Build note: the backend `tsconfig` compiles `src/**` (including this test) and
 * does NOT register `vitest/globals` types, so the vitest globals are imported
 * explicitly to keep `npm run build` (tsc) green under `strict`. `__dirname` is
 * available because the backend compiles to CommonJS (no `"type": "module"`).
 */

import { describe, it, expect, beforeAll } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { validate } from "../validate";
import { runComparison } from "../index";

// ---------------------------------------------------------------------------
// Fixture loader
// ---------------------------------------------------------------------------
//
// Fixtures are committed binary `.docx` archives that live alongside this test
// in `./fixtures/`. They are loaded synchronously via `fs`; a missing fixture
// throws here, which is the correct loud-failure signal (the golden files are a
// hard prerequisite for these tests).
const fixturesDir = join(__dirname, "fixtures");
const load = (name: string): Buffer => readFileSync(join(fixturesDir, name));

// Deterministic tracked-change metadata for the redline produced in `beforeAll`.
// Supplying an explicit author/date (never the wall clock) is what makes the
// engine's output byte-deterministic; the exact values are irrelevant to the
// reparse/render checks under test here.
const AUTHOR = "Compare Bot";
const DATE = "2026-01-01T00:00:00Z";

// A single well-formed redline, produced ONCE from the `mixed-edit` golden pair
// and reused across the positive cases below. Typed as `Buffer` (not
// `Buffer | undefined`) because `beforeAll` always assigns it before any test
// body runs; a failure inside `runComparison` would fail the hook and surface
// loudly rather than yielding an undefined redline.
let redline: Buffer;

beforeAll(async () => {
  // Produce a genuine engine redline: this is the artifact `validate` is meant
  // to gate in production, so validating it here is a true end-to-end check
  // that the engine's own output passes its own validator.
  const result = await runComparison(
    load("mixed-edit.base.docx"),
    load("mixed-edit.revised.docx"),
    { author: AUTHOR, date: DATE },
  );
  redline = result.redlineBytes;
});

// ---------------------------------------------------------------------------
// Positive case 1: a well-formed redline resolves (reparse passes)
// ---------------------------------------------------------------------------

describe("validate: well-formed redline resolves (reparse passes)", () => {
  it("resolves with reparsed === true for a real engine-produced redline", async () => {
    const result = await validate(redline);

    // Reparse is the fatal gate: resolving at all means the emitted bytes
    // reopened as a zip and their document.xml reparsed into <w:document><w:body>.
    expect(result.reparsed).toBe(true);

    // `renderChecked` is environment-dependent (see soft-skip test below); only
    // its TYPE is guaranteed.
    expect(typeof result.renderChecked).toBe("boolean");

    // The ONLY permitted shape assertion on `renderChecked`: membership in the
    // boolean set. This passes whether or not LibreOffice exists on the host,
    // satisfying the soft-skip requirement. We MUST NOT assert
    // `renderChecked === true` -- CI has no soffice binary.
    expect([true, false]).toContain(result.renderChecked);
  });
});

// ---------------------------------------------------------------------------
// Positive case 2: a raw (non-redline) valid .docx also validates
// ---------------------------------------------------------------------------

describe("validate: a raw valid .docx also passes the reparse gate", () => {
  it("validates a plain (non-tracked-change) valid .docx fixture", async () => {
    // `validate` is a GENERIC OOXML-validity gate, not coupled to the presence
    // of `<w:ins>`/`<w:del>` tracked changes: a clean `.docx` with no redlines
    // must pass the reparse check exactly as an emitted redline does.
    const result = await validate(load("identical.base.docx"));

    expect(result.reparsed).toBe(true);
    expect(typeof result.renderChecked).toBe("boolean");
    // Same soft-skip tolerance as above: never require `true`.
    expect([true, false]).toContain(result.renderChecked);
  });
});

// ---------------------------------------------------------------------------
// Negative case: corrupt input rejects (reparse is fatal)
// ---------------------------------------------------------------------------

describe("validate: corrupt input rejects (reparse is fatal)", () => {
  it("rejects a buffer that is not a .docx zip", async () => {
    // A non-zip buffer cannot be reopened by JSZip, so the fatal reparse gate
    // rejects. Asserted as an ASYNC rejection (`.rejects.toThrow()`), never a
    // synchronous throw, because `validate` returns a promise.
    await expect(
      validate(Buffer.from("this is definitely not a .docx zip")),
    ).rejects.toThrow();
  });

  it("rejects an empty buffer", async () => {
    // An empty buffer is likewise not a loadable archive; the reparse gate is
    // fatal for it too.
    await expect(validate(Buffer.alloc(0))).rejects.toThrow();
  });
});

// ---------------------------------------------------------------------------
// Soft-skip semantics: the render probe never fails validation
// ---------------------------------------------------------------------------

describe("validate: render probe is best-effort (soft-skip)", () => {
  it("render check is soft-skipped when LibreOffice is unavailable", async () => {
    const result = await validate(redline);

    // Documented intent: the render probe (`docxToPdf`) is a SOFT smoke test.
    // In CI there is typically no `soffice` binary, so `docxToPdf` throws and
    // `validate` catches it and resolves with `renderChecked: false`. On a host
    // WITH LibreOffice the probe may run and yield `renderChecked: true`. This
    // test therefore asserts ONLY the fatal-gate outcome (`reparsed === true`)
    // and the TYPE of `renderChecked`; it must tolerate BOTH boolean values and
    // must NEVER require `true`, so the compare suite stays green everywhere.
    expect(result.reparsed).toBe(true);
    expect(typeof result.renderChecked).toBe("boolean");
  });
});
