// Vitest configuration scoped EXCLUSIVELY to the isolated Document Compare engine.
//
// Strict isolation (AAP §0.8.1): the `include` glob below is the single mechanism
// that keeps the compare test suite self-contained. It MUST match only tests under
// `src/compare/**` — never broaden it to other backend code or unrelated suites.
//
// Watch mode is intentionally NOT configured here; it is controlled by the CLI via
// the `vitest run` command wired into the package.json "test:compare" script, which
// performs a single, non-watching CI run. This config is safe for one-shot runs.
//
// No coverage thresholds, external reporters, or global setup are wired in: the
// compare engine tests are deterministic, offline, and golden-file based (no DB,
// network, R2, or LibreOffice dependency at config time).
//
// This file lives at the backend root (outside `src/`), so the app build (`tsc`,
// whose tsconfig `include` is `src/**/*`) does not compile it — only Vitest consumes
// it. Vitest auto-discovers this file when invoked from the backend directory, so the
// "test:compare" script needs no `--config` flag.
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Discover ONLY the compare-module tests. This covers the canonical
    // `src/compare/__tests__/*.test.ts` files as well as any test co-located
    // elsewhere under `src/compare/`. Do not widen this pattern.
    include: ["src/compare/**/*.test.ts"],
    // The compare engine runs in pure Node (no DOM), matching the Node 20+ runtime.
    environment: "node",
    // Expose describe/it/expect as globals so the compare test files stay free of
    // per-file test-framework imports.
    globals: true,
  },
});
