/**
 * Unit tests for `backend/src/compare/storageKeys.ts` — the pure, I/O-free R2
 * object-key helpers for the Document Compare feature.
 *
 * Why these tests exist
 * ---------------------
 * The returned key strings are a cross-cutting contract shared by two code
 * paths that must agree byte-for-byte:
 *   - write path: the create handler calls `uploadFile(comparisonRedlineKey(...))`
 *     and `uploadFile(comparisonDiffKey(...))`;
 *   - read path: the self-contained streaming download handler calls
 *     `downloadFile(comparisonRedlineKey(...))`.
 * If the format drifted between these two sides, the redline download would
 * resolve a non-existent object and 404. This suite therefore locks the EXACT
 * key layout under the `comparisons/` prefix so any accidental change (a
 * pluralized segment, a reordered path, a renamed artifact) fails loudly.
 *
 * Isolation (AAP §0.8.1)
 * ----------------------
 * This suite imports ONLY from "vitest" and the module under test
 * ("../storageKeys"). It deliberately does NOT pull in `../../lib/storage`, the
 * AWS SDK, `jszip`, or any other code: the helpers are pure string formatters
 * with no I/O, and the compare engine test suite is strictly isolated.
 *
 * Build note
 * ----------
 * Explicit `vitest` imports are mandatory. The backend `tsconfig.json` compiles
 * `src/**` (including tests) with `types: ["node", "express", "cors", "multer"]`
 * and no `vitest/globals` ambient types, so relying on bare `describe`/`it`/
 * `expect` globals would break `npm run build`.
 */

import { describe, it, expect } from "vitest";
import { comparisonRedlineKey, comparisonDiffKey } from "../storageKeys";

describe("storageKeys — R2 comparison key helpers", () => {
  // -------------------------------------------------------------------------
  // Exact-string contract (the locked format)
  //
  // These are the single source of truth for both the upload (write) and the
  // self-contained download (read) paths. The strings below are byte-exact.
  // -------------------------------------------------------------------------
  describe("exact key strings", () => {
    it("builds the redline key as comparisons/{userId}/{comparisonId}/redline.docx", () => {
      expect(comparisonRedlineKey("u1", "c1")).toBe(
        "comparisons/u1/c1/redline.docx",
      );
    });

    it("builds the diff key as comparisons/{userId}/{comparisonId}/diff.json", () => {
      expect(comparisonDiffKey("u1", "c1")).toBe("comparisons/u1/c1/diff.json");
    });

    it("interpolates realistic UUID-like ids verbatim (redline)", () => {
      expect(comparisonRedlineKey("user-123", "abc-def-456")).toBe(
        "comparisons/user-123/abc-def-456/redline.docx",
      );
    });

    it("interpolates realistic UUID-like ids verbatim (diff)", () => {
      expect(comparisonDiffKey("user-123", "abc-def-456")).toBe(
        "comparisons/user-123/abc-def-456/diff.json",
      );
    });
  });

  // -------------------------------------------------------------------------
  // Structural properties
  //
  // Beyond the exact strings, assert the invariants the rest of the engine and
  // the storage layer rely on: a shared prefix, a shared per-comparison
  // directory, stable filenames, and a fixed segment shape.
  // -------------------------------------------------------------------------
  describe("structural properties", () => {
    it("prefixes both keys with comparisons/", () => {
      expect(comparisonRedlineKey("u", "c").startsWith("comparisons/")).toBe(
        true,
      );
      expect(comparisonDiffKey("u", "c").startsWith("comparisons/")).toBe(true);
    });

    it("shares the same comparisons/{userId}/{comparisonId}/ directory portion", () => {
      const dir = "comparisons/u/c/";
      // Both artifacts live under the same per-comparison directory ...
      expect(comparisonRedlineKey("u", "c").slice(0, dir.length)).toBe(dir);
      expect(comparisonDiffKey("u", "c").slice(0, dir.length)).toBe(dir);
      // ... and each full key is exactly that directory + its own filename.
      expect(comparisonRedlineKey("u", "c")).toBe(dir + "redline.docx");
      expect(comparisonDiffKey("u", "c")).toBe(dir + "diff.json");
    });

    it("suffixes the redline key with /redline.docx and the diff key with /diff.json", () => {
      expect(comparisonRedlineKey("u", "c").endsWith("/redline.docx")).toBe(
        true,
      );
      expect(comparisonDiffKey("u", "c").endsWith("/diff.json")).toBe(true);
    });

    it("produces two distinct keys for the same comparison", () => {
      expect(comparisonRedlineKey("u", "c")).not.toBe(
        comparisonDiffKey("u", "c"),
      );
    });

    it("splits into exactly [comparisons, userId, comparisonId, filename]", () => {
      expect(comparisonRedlineKey("u", "c").split("/")).toEqual([
        "comparisons",
        "u",
        "c",
        "redline.docx",
      ]);
      expect(comparisonDiffKey("u", "c").split("/")).toEqual([
        "comparisons",
        "u",
        "c",
        "diff.json",
      ]);
    });
  });

  // -------------------------------------------------------------------------
  // Determinism (pure formatter — no clock, no randomness)
  //
  // Identical inputs must always yield identical output; there is no hidden
  // state, timestamp, or random component in the key helpers.
  // -------------------------------------------------------------------------
  describe("determinism", () => {
    it("returns identical output for identical input (redline)", () => {
      expect(comparisonRedlineKey("u1", "c1")).toBe(
        comparisonRedlineKey("u1", "c1"),
      );
    });

    it("returns identical output for identical input (diff)", () => {
      expect(comparisonDiffKey("u1", "c1")).toBe(comparisonDiffKey("u1", "c1"));
    });
  });
});
