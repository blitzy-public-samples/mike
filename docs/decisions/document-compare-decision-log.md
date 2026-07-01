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
D1–D6 from elsewhere.

| ID | What was decided | Alternatives considered | Why | Risks |
|----|------------------|-------------------------|-----|-------|
| D1 | Expose `POST /projects/:projectId/comparisons`, `GET /comparisons/:id`, and `GET /comparisons/:id/download` with **no `/api` prefix**. | Keep the prompt's literal `/api/...` paths. | The repository mounts every router without an `/api` prefix in `backend/src/index.ts`; matching the established convention keeps the router consistent with `projectChatRouter` and the other existing routers. | Client/route drift.<br>Mitigated: `frontend/src/app/lib/mikeApi.ts` centralizes paths via `API_BASE`, so each path is defined once. |
| D2 | Run `runComparison` **synchronously** inside the `POST /projects/:projectId/comparisons` handler. | A job queue, background worker, or scheduler. | No queue, worker, or scheduler exists in the backend. The `status` column (`pending` / `processing` / `complete` / `error`) plus the `GET /comparisons/:id` polling endpoint keep the contract async-ready for a future worker without building that infrastructure now. | Long-running requests for very large documents.<br>Planned mitigation: the future async worker (see the README "next tasks"). |
| D3 | `GET /comparisons/:id/download` performs its own `checkProjectAccess` and streams the redline `.docx` directly. | Route the download through the existing signed `/download/:token` mechanism. | That generic route resolves **only** `document_versions` rows, so a `comparisons/`-prefixed redline key would `404` there. A self-contained handler leaves `backend/src/routes/downloads.ts` untouched, preserving the signed-download contract. | Minor stream-logic duplication with `downloads.ts`.<br>Accepted to keep the signed-download contract intact and avoid editing a must-remain-untouched file. |
| D4 | Enforce access in the routes via `checkProjectAccess` (denials masked as `404`, never `403`) plus `REVOKE ALL PRIVILEGES ON TABLE public.document_comparisons FROM anon, authenticated`, with RLS enabled as defense-in-depth. | Per-row `create policy` RLS statements. | Existing project-scoped tables (`projects`, `documents`, `document_versions`, `document_edits`, `chats`, `tabular_reviews`) use grant revocation plus app-level checks, not per-row policies; the service-role client bypasses RLS anyway, so matching the established pattern keeps the table consistent. | Relies on every route performing the access check.<br>Mitigated: every handler (`POST`, `GET` status, `GET` download) is guarded with `checkProjectAccess` / `ensureDocAccess`. |
| D5 | Implement OOXML parsing, normalization, and emission **locally** in `backend/src/compare/`; the engine modules own their read-side and write-side helpers. | Import the internal helpers from `backend/src/lib/docxTrackedChanges.ts`. | That module's helpers are private / non-exported; only the public `extractDocxBodyText` is reused. A local implementation keeps the engine isolated, independently unit-testable, and free of coupling to the assistant / editing code paths. | Some conceptual overlap with `docxTrackedChanges.ts`.<br>Accepted in exchange for strict isolation and testability. |
| D6 | Reject `.doc` and PDF inputs with a clear `400` error; compare only native `.docx`. | LibreOffice convert-first (auto-convert `.doc` / PDF to `.docx`). | V1 scope; rejecting keeps the path deterministic and simple. LibreOffice is reused only for the **optional** redline render validation (`docxToPdf`), never for input conversion. | Users with `.doc` / PDF must convert manually first.<br>Documented as a pitfall and a suggested next task in the README. |

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

The mapping below is **bidirectional and at 100% coverage**: every requirement
maps to its implementing file(s), and every listed file traces back to a
requirement — there are no orphan files and no uncovered requirements.

| Requirement | Implementing file(s) | Notes |
|-------------|----------------------|-------|
| Deterministic compare engine | `backend/src/compare/index.ts` (orchestrator `runComparison`)<br>`backend/src/compare/parseDocx.ts`<br>`backend/src/compare/normalize.ts`<br>`backend/src/compare/alignParagraphs.ts`<br>`backend/src/compare/wordDiff.ts`<br>`backend/src/compare/emitTrackedChanges.ts`<br>`backend/src/compare/diffJson.ts`<br>`backend/src/compare/validate.ts`<br>`backend/src/compare/storageKeys.ts` | Deterministic diff path, no AI; `opts.date` is the sole time source; isolated module. |
| HTTP API (create / status / download) | `backend/src/routes/comparisons.ts` | `zod`-validated create; synchronous run (D2); self-contained access-checked download stream (D3). |
| Router mount + rate limiters | `backend/src/index.ts` | Mounts `comparisonsRouter` behind `requireAuth`; declares `compareLimiter` and `downloadLimiter`. |
| Persistence helpers (insert / select / update) | `backend/src/lib/documentComparisons.ts` | Uses the `createServerSupabase` service-role client. |
| Database table `document_comparisons` | `backend/migrations/20260616_01_document_comparisons.sql`<br>`backend/schema.sql` | Single dated, idempotent migration with `REVOKE ALL PRIVILEGES ... FROM anon, authenticated` (D4); matching parity block appended to `schema.sql` for fresh DBs. |
| Tests | `backend/vitest.config.ts`<br>`backend/src/compare/__tests__/**` | Vitest scoped to `src/compare/**`; unit + integration tests and `.docx` golden fixtures. |
| Test dependency + script | `backend/package.json` | Adds `vitest ^4.1.9` dev-dependency and the `test:compare` script — the only new dependency. |
| Frontend Compare tab — route page | `frontend/src/app/(pages)/projects/[id]/compare/page.tsx` | Compare page rendered inside the existing workspace shell. |
| Frontend Compare components | `frontend/src/app/components/compare/CompareView.tsx`<br>`frontend/src/app/components/compare/RedlineDocxView.tsx`<br>`frontend/src/app/components/compare/SideBySideDiff.tsx` | Pickers/run/status/result; inline redline via `docx-preview`; two-column diff driven by the diff JSON hunks (deviation (c): no `TextSearchWidget`). |
| Frontend plumbing (tab wiring, API client, shared types) | `frontend/src/app/components/projects/ProjectPageParts.tsx`<br>`frontend/src/app/components/projects/ProjectWorkspace.tsx`<br>`frontend/src/app/lib/mikeApi.ts`<br>`frontend/src/app/components/shared/types.ts` | `"compare"` union member; tab item + `onChange` + URL-segment helpers; `createComparison` / `getComparison` / download client fns; comparison interfaces. |
| Docs / Rules deliverables | `README.md`<br>`docs/decisions/document-compare-decision-log.md`<br>`docs/presentations/document-compare-executive-summary.html`<br>`docs/document-compare.md` | Rule 2 onboarding; Rule 1 (this file, the *why*); Rule 3 executive deck; engineering *how* reference. |

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
