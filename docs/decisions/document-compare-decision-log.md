# Document Compare — Decision Log

This decision log is the single source of truth for the rationale behind Mike's
Document Compare / Redline feature, satisfying the Explainability rule for the
feature. Code comments across the feature are deliberately brief and factual and
omit rationale, so this log — not the code — explains *why* each choice was made.
The diff path is deterministic and contains no AI; the reasoning for that and for
every other non-trivial decision is recorded below.

## Related Docs

- [engineering reference](../document-compare.md) — the deep-dive on *how* the
  compare engine, data structures, and API work.
- [README](../../README.md) — onboarding, setup, environment, and high-level
  usage of the Compare tab.
- [executive summary](../presentations/document-compare-executive-summary.html) —
  the self-contained deck for non-technical leadership.

This log owns the *why*; the [engineering reference](../document-compare.md) owns
the *how*, and the [README](../../README.md) owns onboarding. Deep engineering
detail is linked, not duplicated here, and rationale is never duplicated into
code comments.

## Decision Log

Every non-trivial decision is recorded as one row below. The four mandated
columns are **What was decided**, **Alternatives considered**, **Why**, and
**Risks**; the leading `ID` column exists only so decisions can be referenced as
D1–D10 from elsewhere.

| ID | What was decided | Alternatives considered | Why | Risks |
|----|------------------|-------------------------|-----|-------|
| D1 | Expose `POST /projects/:projectId/comparisons`, `GET /comparisons/:id`, and `GET /comparisons/:id/download` with **no `/api` prefix**. | Keep the prompt's literal `/api/...` paths. | The repository mounts every router without an `/api` prefix in `backend/src/index.ts`; matching the established convention keeps the router consistent with `projectChatRouter` and the other existing routers. | Client/route drift.<br>Mitigated: `frontend/src/app/lib/mikeApi.ts` centralizes paths via `API_BASE`, so each path is defined once. |
| D2 | Run `runComparison` **synchronously** inside the `POST /projects/:projectId/comparisons` handler. | A job queue, background worker, or scheduler. | No queue, worker, or scheduler exists in the backend. The `status` column (`pending` / `processing` / `complete` / `error`) plus the `GET /comparisons/:id` polling endpoint keep the contract async-ready for a future worker without building that infrastructure now. | Long-running requests for very large documents.<br>Planned mitigation: the future async worker (see the README "next tasks"). |
| D3 | `GET /comparisons/:id/download` performs its own `checkProjectAccess` and sends the redline `.docx` directly. In V1 it **buffers** the object via `downloadFile` and returns it with `res.send(Buffer.from(raw))`, mirroring the existing `downloads.ts` behavior (the whole codebase reads objects fully buffered). | Route the download through the existing signed `/download/:token` mechanism; or add a streaming storage helper and pipe the object body. | That generic route resolves **only** `document_versions` rows, so a `comparisons/`-prefixed redline key would `404` there. A self-contained handler leaves `backend/src/routes/downloads.ts` untouched, preserving the signed-download contract, and buffering keeps V1 consistent with the existing (buffered) download path rather than introducing a new streaming primitive. | Minor send-logic duplication with `downloads.ts`; buffering a large redline holds it fully in memory.<br>Accepted for V1 to keep the signed-download contract intact and avoid editing a must-remain-untouched file; a streaming storage helper is a suggested next task in the README. |
| D4 | Enforce access in the routes via `checkProjectAccess` (denials masked as `404`, never `403`) plus `REVOKE ALL PRIVILEGES ON TABLE public.document_comparisons FROM anon, authenticated`, with RLS enabled as defense-in-depth. | Per-row `create policy` RLS statements. | Existing project-scoped tables (`projects`, `documents`, `document_versions`, `document_edits`, `chats`, `tabular_reviews`) use grant revocation plus app-level checks, not per-row policies; the service-role client bypasses RLS anyway, so matching the established pattern keeps the table consistent. | Relies on every route performing the access check.<br>Mitigated: every handler (`POST`, `GET` status, `GET` download) is guarded with `checkProjectAccess` / `ensureDocAccess`. |
| D5 | Implement OOXML parsing, normalization, and emission **locally** in `backend/src/compare/`; the engine modules own their read-side and write-side helpers. | Import the internal helpers from `backend/src/lib/docxTrackedChanges.ts`. | That module's helpers are private / non-exported; only the public `extractDocxBodyText` is reused. A local implementation keeps the engine isolated, independently unit-testable, and free of coupling to the assistant / editing code paths. | Some conceptual overlap with `docxTrackedChanges.ts`.<br>Accepted in exchange for strict isolation and testability. |
| D6 | Reject `.doc` and PDF inputs with a clear `400` error; compare only native `.docx`. | LibreOffice convert-first (auto-convert `.doc` / PDF to `.docx`). | V1 scope; rejecting keeps the path deterministic and simple. LibreOffice is reused only for the **optional** redline render validation (`docxToPdf`), never for input conversion. | Users with `.doc` / PDF must convert manually first.<br>Documented as a pitfall and a suggested next task in the README. |
| D7 | Enforce **deterministic fail-fast resource limits** in the engine before the O(n·m) alignment allocates its DP table: input `.docx` bytes (`MAX_INPUT_DOCX_BYTES` = 50 MB, `index.ts`), uncompressed `word/document.xml` size (`MAX_DOCUMENT_XML_BYTES` = 100 MB, `parseDocx.ts`), paragraph count (`MAX_PARAGRAPHS` / `MAX_ALIGN_PARAGRAPHS_PER_DOC` = 50,000), run count (`MAX_RUNS` = 500,000), and LCS cell count (`MAX_LCS_CELLS` = 25,000,000, `alignParagraphs.ts`). Exceeding any bound throws a controlled comparison `Error`. | Trust inputs (no limits); stream/chunk the parse; switch the alignment to a Myers diff with O(n) memory. | The alignment allocated an unbounded `(base+1)*(revised+1)` DP table, so a large or adversarial document could exhaust memory/CPU before any controlled error was returned (review finding M5). Fixed integer constants keep every check deterministic, satisfying the no-nondeterminism mandate (AAP §0.8.1). The thresholds sit far above realistic V1 contract sizes (hundreds to low-thousands of paragraphs), so genuine inputs never trip them. | Legitimately enormous documents are rejected and must be split (documented as a pitfall + next task in the README). A zip bomb that inflates during decompression is bounded by the input-bytes and uncompressed-XML-size guards but not fully eliminated in V1 (JSZip decompresses the entry before its size can be measured) — a documented V1 boundary; a future size-from-central-directory / streaming check is a suggested next task. |
| D8 | Support **two entry flows** through one endpoint: (a) compare two different project documents (their active versions), and (b) compare two explicit **versions of a single document** (the prior-version flow). The create body accepts optional `baseVersionId` / `revisedVersionId`; when supplied, the same `documentId` is allowed on both sides **iff** the two version ids differ, and each version is resolved via `loadActiveVersion(documentId, db, versionId)`. Resolved version ids are persisted on the row (`base_version_id` / `revised_version_id`). | A separate second endpoint for version-vs-version compare; or resolving only each document's active version (flow (a) only). | `loadActiveVersion` already accepts an optional `versionId` and verifies the version belongs to the document, so flow (b) needs no new infrastructure. One endpoint with optional version ids keeps the API surface minimal and async-ready, and the frontend simply toggles between a two-document picker and a single-document/two-version picker. | Same-document requests must be gated so a document is never compared against itself.<br>Mitigated: both the `zod` schema and `canRun` require the two version ids to differ when a single document is selected. |
| D9 | Bind every selected document to the **route** `:projectId`: `resolveDocxInput` rejects (and masks as `404`) any document whose `project_id` differs from the route project, **before** any access check, byte download, or row insert. | Rely on `ensureDocAccess` alone (per-document access), without checking the document's project against the route project. | `ensureDocAccess` proves the caller can reach a document, but not that the document belongs to the project the comparison is being created under. Without the project binding, a user could create a comparison in project A that embeds content from a document in project B, and project A collaborators could then read B's content via the artifacts. Checking `doc.project_id === projectId` first closes this cross-project leak (review finding CRITICAL / Authorization). | A shared/cross-project document model would need this relaxed deliberately.<br>Accepted: V1 documents are project-scoped, and the check is covered by regression tests (`comparisonsAuthorization.test.ts`) asserting cross-project rejection touches no bytes and inserts no row. |
| D10 | Diff **modified paragraphs at word level** with a deterministic two-stage aligner: stage 1 is the exact-anchor paragraph LCS; stage 2 walks each maximal delete/insert change-region and runs a **secondary similarity LCS** that pairs a deleted and an inserted paragraph as a single `equal` (modified) pair when their word-token **Sørensen–Dice** similarity is `≥ 0.5`. Matched pairs then flow through the existing `wordDiff` → `emitTrackedChanges` path; unmatched paragraphs remain whole-paragraph delete / insert. | Match only on exact text equality (the prior behavior — modified paragraphs became whole-paragraph delete + insert, so `wordDiff` never ran); or use a fuzzy/AI matcher. | Exact-only matching produced legally noisy redlines for realistic single-paragraph edits (review finding MAJOR). A fixed Dice threshold with the **same** deterministic tie-break as stage 1 (prefer the deletion branch) keeps the aligner fully deterministic — no AI, no randomness, no time — so byte-identical output is preserved. A region-size guardrail (`MAX_REGION_PAIR_CELLS`) bounds the secondary DP. | The threshold is a heuristic: heavily rewritten paragraphs below `0.5` fall back to delete + insert.<br>Accepted for V1: the threshold is documented, deterministic, and tunable; the fallback is always a correct (if coarser) redline. |

## Deviations from the Literal Prompt

Each deviation from the literal prompt is recorded explicitly below, stating what
differs and why.

- **(a) `created_by` column name.** The prompt's table sketch uses `created_by`,
  and this is logged because several sibling project-scoped tables use `user_id`
  instead. The new table `document_comparisons` deliberately uses `created_by`
  (a nullable `text` column) to match the prompt's column list exactly. There is
  no behavioral impact; access is enforced by `project_id` plus
  `checkProjectAccess`.
- **(b) Router mounted at the app root.** `comparisonsRouter` is mounted via
  `app.use("/", comparisonsRouter)` in `backend/src/index.ts` (absolute paths, no
  `mergeParams`) because its routes span **two path families** —
  `/projects/:projectId/comparisons` (project-scoped create) and
  `/comparisons/:id` (item-level status and download). A single root mount cleanly
  serves both, rather than splitting the feature into two routers.
- **(c) `SideBySideDiff` omits the optional in-diff `TextSearchWidget` in V1.**
  The design system lists `TextSearchWidget` as an optional reuse for in-diff
  search; the V1 `SideBySideDiff` component intentionally does not render it, and
  change navigation is provided via next/previous buttons instead. This keeps the
  component minimal per the minimal-change clause and avoids an unused-import lint
  risk.
- **(d) Redline `w:date` uses the run/request time.** The live route passes
  `new Date().toISOString()` as `opts.date` (written into `w:date`) and the
  request user as `opts.author`, so two runs of the *same* documents at
  *different* times produce redlines that differ in the `w:date` attribute. The
  determinism guarantee is therefore precise: **identical inputs (including
  `opts.date` and `opts.author`) produce byte-identical output** — which the
  golden-file tests pin by passing a fixed `opts`. The engine itself never calls
  `Date.now()` and uses no randomness; time enters only via `opts`.

## Bidirectional Traceability Matrix

The mapping below is **bidirectional and at 100% coverage**: every AAP
requirement is enumerated as its own row and maps to the file(s) that actually
implement it, and every listed file traces back to a requirement — there are no
orphan files and no uncovered requirements. Each row names only files that
implement the row's requirement (reused-unchanged infrastructure is listed
separately in the next section).

| # | AAP Requirement | Implementing file(s) | Notes |
|---|-----------------|----------------------|-------|
| R1a | Entry flow (a): compare two different project documents (their active versions) | `frontend/src/app/components/compare/CompareView.tsx` (documents mode)<br>`frontend/src/app/lib/mikeApi.ts` (`createComparison`)<br>`backend/src/routes/comparisons.ts` (`createComparisonSchema`, `resolveDocxInput`) | Two-document picker; body `{ baseDocumentId, revisedDocumentId }`; each side resolved via `loadActiveVersion` (active version). |
| R1b | Entry flow (b): upload / choose a new revised version of an existing document and compare against the prior version | `frontend/src/app/components/compare/CompareView.tsx` (versions mode + inline `uploadDocumentVersion`)<br>`frontend/src/app/lib/mikeApi.ts` (`CreateComparisonInput` version ids)<br>`frontend/src/app/components/shared/types.ts` (`CreateComparisonInput`)<br>`backend/src/routes/comparisons.ts` (optional `baseVersionId` / `revisedVersionId`; same-doc iff versions differ) | D8. Single-document picker + base/revised version pickers (default prior→revised current); `loadActiveVersion(documentId, db, versionId)` resolves each explicit version. |
| R2 | Deterministic diff path, no AI | `backend/src/compare/index.ts` (`runComparison`)<br>`backend/src/compare/parseDocx.ts`<br>`backend/src/compare/normalize.ts`<br>`backend/src/compare/alignParagraphs.ts`<br>`backend/src/compare/wordDiff.ts`<br>`backend/src/compare/emitTrackedChanges.ts`<br>`backend/src/compare/diffJson.ts`<br>`backend/src/compare/storageKeys.ts` | No AI, no randomness; time enters only via `opts.date` (deviation (d)); byte-identical output for identical inputs+opts. |
| R3 | Accept-all normalization of inputs already carrying tracked changes | `backend/src/compare/normalize.ts` | Diff runs against clean (accepted-view) text. |
| R4 | Paragraph LCS alignment | `backend/src/compare/alignParagraphs.ts` (stage 1: exact-anchor LCS) | Deterministic tie-break (prefer deletion branch). |
| R5 | Word-level diff within matched (modified) paragraphs — end to end | `backend/src/compare/alignParagraphs.ts` (stage 2: similarity refinement pairs modified paragraphs as `equal`)<br>`backend/src/compare/wordDiff.ts` (`fast-diff` word segments)<br>`backend/src/compare/index.ts` (modified-pair `wordDiff` branch)<br>`backend/src/compare/emitTrackedChanges.ts` (`rewriteParagraphToModified`) | D10. Sørensen–Dice ≥ 0.5 pairing makes the `wordDiff` branch reachable; a modified paragraph emits word-level `<w:ins>`/`<w:del>` rather than whole-paragraph delete+insert. |
| R6 | Native tracked-changes `.docx` redline (`w:ins` / `w:del` / `w:delText`, deleted paragraph mark, unique `w:id`, carried `w:rPr`) | `backend/src/compare/emitTrackedChanges.ts` | Rezips a valid `.docx`; monotonic unique `w:id`. |
| R7 | Structured diff JSON — ordered hunks `{ type, text, baseRange, revisedRange }` | `backend/src/compare/diffJson.ts`<br>`frontend/src/app/components/shared/types.ts` (`DiffHunk` / `DiffJson`)<br>`frontend/src/app/components/compare/SideBySideDiff.tsx` | Backend/frontend type parity. |
| R8 | Output validation: reparse + optional LibreOffice render | `backend/src/compare/validate.ts` | Reparse fatal; `docxToPdf` render soft. |
| R9 | Deterministic fail-fast resource limits | `backend/src/compare/index.ts`<br>`backend/src/compare/parseDocx.ts`<br>`backend/src/compare/alignParagraphs.ts` | D7. Fixed integer bounds (input bytes, XML size, paragraph/run/LCS counts, region-pair cells). |
| R10 | HTTP API: `zod`-validated create + synchronous run + status polling | `backend/src/routes/comparisons.ts` | D2. `POST /projects/:projectId/comparisons`; `GET /comparisons/:id` polling; artifacts uploaded before `complete`. |
| R11 | Self-contained, access-checked download that bypasses `/download/:token` (buffered V1) | `backend/src/routes/comparisons.ts` (`GET /comparisons/:id/download`) | D3. Own `checkProjectAccess`; buffers via `downloadFile` + `res.send`, mirroring `downloads.ts` (which stays untouched). |
| R12 | Cross-project document authorization binding | `backend/src/routes/comparisons.ts` (`resolveDocxInput` project-binding check)<br>`backend/src/compare/__tests__/comparisonsAuthorization.test.ts` | D9. Rejects/masks (404) any doc whose `project_id` ≠ route project before byte download or row insert; regression-tested. |
| R13 | Router mount + rate limiters behind `requireAuth` | `backend/src/index.ts` | Mounts `comparisonsRouter`; declares `compareLimiter` + `downloadLimiter`. |
| R14 | Project-scoped persistence (incl. resolved version ids) | `backend/src/lib/documentComparisons.ts`<br>`frontend/src/app/components/shared/types.ts` (`Comparison`) | Service-role client; row carries `base_version_id` / `revised_version_id`. |
| R15 | Exactly one dated migration + matching `schema.sql`; no existing table edits | `backend/migrations/20260616_01_document_comparisons.sql`<br>`backend/schema.sql` | D4. Idempotent `create table if not exists` + `REVOKE ALL PRIVILEGES ... FROM anon, authenticated`; nullable `base_version_id` / `revised_version_id` (FK → `document_versions ... on delete set null`); parity block for fresh DBs. |
| R16 | R2 key prefix `comparisons/{userId}/{comparisonId}/{redline.docx,diff.json}` | `backend/src/compare/storageKeys.ts` | Server-derived keys. |
| R17 | Frontend Compare tab reusing the existing workspace shell/layout | `frontend/src/app/(pages)/projects/[id]/compare/page.tsx`<br>`frontend/src/app/components/projects/ProjectPageParts.tsx`<br>`frontend/src/app/components/projects/ProjectWorkspace.tsx` | `"compare"` union member; tab item + `onChange` + URL-segment helpers. |
| R18 | Inline redline + side-by-side diff rendering (tokenized, accessible, responsive) | `frontend/src/app/components/compare/RedlineDocxView.tsx` (inline via `docx-preview`)<br>`frontend/src/app/components/compare/SideBySideDiff.tsx` (two-column + change nav) | Semantic tokens (destructive/brand-blue/muted/border), `Button` nav, `aria-pressed`, small-screen stacking; no `TextSearchWidget` (deviation (c)). |
| R19 | Tests + scoped runner | `backend/vitest.config.ts`<br>`backend/src/compare/__tests__/**` | Vitest scoped to `src/compare/**`; unit + integration + authorization tests and `.docx` golden fixtures. |
| R20 | Exactly one new dependency + scoped script | `backend/package.json` | `vitest ^4.1.9` dev-dependency and `test:compare` — the only new dependency. |
| R21 | Rule deliverables (decision log / onboarding / executive deck / engineering reference) | `docs/decisions/document-compare-decision-log.md`<br>`README.md`<br>`docs/presentations/document-compare-executive-summary.html`<br>`docs/document-compare.md` | Rule 1 (this file, the *why*); Rule 2 onboarding; Rule 3 deck; engineering *how* reference. |

## Reused (Unchanged) Infrastructure

The feature reuses existing infrastructure without modifying it, and reads a
small set of reference-only files without changing them:

- Consumed unchanged: `requireAuth`, `checkProjectAccess` / `ensureDocAccess`,
  `loadActiveVersion`, `downloadFile` / `uploadFile`, `createServerSupabase`,
  `docxToPdf`, and the public `extractDocxBodyText`.
- Reference-only (read, never modified): `backend/src/routes/downloads.ts`,
  `backend/src/lib/docxTrackedChanges.ts`,
  `frontend/src/app/components/shared/DocxView.tsx`, and
  `frontend/src/app/components/tabular/AddNewTRModal.tsx`.
- Exactly one new dependency was added across the whole feature: `vitest` (a
  backend dev-dependency scoped to the engine's tests).
