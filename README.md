# Mike

Mike is a legal document assistant with a Next.js frontend, an Express backend, Supabase Auth/Postgres, and Cloudflare R2-compatible object storage.

Website: [mikeoss.com](https://mikeoss.com)

## Contents

- `frontend/` - Next.js application
- `backend/` - Express API, Supabase access, document processing, and database schema
- `backend/schema.sql` - Supabase schema for fresh databases
- `backend/migrations/` - dated, incremental schema migrations; on an existing database, apply the files dated after the Mike version you deployed
- `backend/src/compare/` - deterministic document compare / redline engine (tracked-changes .docx + diff JSON)

## Prerequisites

- Node.js 20 or newer
- npm
- git
- A Supabase project
- A Cloudflare R2 bucket, MinIO bucket, or another S3-compatible bucket
- At least one supported model provider API key: Anthropic, Google Gemini, or OpenAI
- Optional: a CourtListener API token for case law lookup and citation verification
- LibreOffice installed locally if you need DOC/DOCX to PDF conversion, and to validate the rendered redline during a document comparison

## Database Setup

For a new Supabase database, open the Supabase SQL editor and run:

```sql
-- copy and run the contents of:
-- backend/schema.sql
```

The schema file is for fresh deployments and already includes the latest database shape.

For an existing database, do not run the full schema file over production data. Instead, apply the incremental files in `backend/migrations/`: run the migrations dated **after** the version of Mike you currently have deployed, in filename order. Each file is named `YYYYMMDD_<name>.sql` (the date is also recorded in a comment at the top of the file) and is written to be safe to re-run, so when unsure you can re-apply the most recent migrations without harm.

The Document Compare feature adds a single `document_comparisons` table. Fresh databases already include it via `backend/schema.sql`; existing deployments apply the new dated migration in `backend/migrations/` (named `YYYYMMDD_document_comparisons.sql`).

## Environment

Create local env files:

```bash
touch backend/.env
touch frontend/.env.local
```

Create `backend/.env`:

```bash
PORT=3001
FRONTEND_URL=http://localhost:3000
DOWNLOAD_SIGNING_SECRET=replace-with-a-random-32-byte-hex-string
SUPABASE_URL=https://your-project.supabase.co
SUPABASE_SECRET_KEY=your-supabase-service-role-key

R2_ENDPOINT_URL=https://your-account-id.r2.cloudflarestorage.com
R2_ACCESS_KEY_ID=your-r2-access-key
R2_SECRET_ACCESS_KEY=your-r2-secret-key
R2_BUCKET_NAME=mike

GEMINI_API_KEY=your-gemini-key
ANTHROPIC_API_KEY=your-anthropic-key
OPENAI_API_KEY=your-openai-key
RESEND_API_KEY=your-resend-key
USER_API_KEYS_ENCRYPTION_SECRET=your-long-random-secret

# Optional: enables CourtListener case law and citation tools.
COURTLISTENER_API_TOKEN=your-courtlistener-token

# Optional: use locally imported CourtListener bulk data for faster case reads.
COURTLISTENER_BULK_DATA_ENABLED=false
```

Create `frontend/.env.local`:

```bash
NEXT_PUBLIC_SUPABASE_URL=https://your-project.supabase.co
NEXT_PUBLIC_SUPABASE_PUBLISHABLE_DEFAULT_KEY=your-supabase-anon-key
NEXT_PUBLIC_API_BASE_URL=http://localhost:3001
```

Supabase values come from the project dashboard. Use the project URL for `SUPABASE_URL` / `NEXT_PUBLIC_SUPABASE_URL`, the service role key for the backend `SUPABASE_SECRET_KEY`, and the anon/public key for `NEXT_PUBLIC_SUPABASE_PUBLISHABLE_DEFAULT_KEY`. If your Supabase project shows multiple key formats, use the legacy JWT-style anon and service role keys expected by the Supabase client libraries.

Provider keys are only needed for the models, legal research, and email features you plan to use. Model provider keys and the CourtListener token can be configured in `backend/.env` for the whole instance, or per user in **Account > Models & API Keys**. If a provider key is present in `backend/.env`, that provider is available by default and the matching browser API key field is read-only.

## CourtListener Integration

Mike can use CourtListener for US case law citation verification, case fetching, targeted opinion search, and case-law panels in assistant responses.

To enable live CourtListener access, set `COURTLISTENER_API_TOKEN` in `backend/.env` and restart the backend. Users can also add their own CourtListener token from **Account > Models & API Keys** when the instance does not provide one globally.

Fresh databases created from `backend/schema.sql` already include the CourtListener support tables. Existing deployments should apply the matching dated migration in `backend/migrations/` before enabling the feature.

Bulk data is optional. When `COURTLISTENER_BULK_DATA_ENABLED=true`, Mike first tries local Supabase/R2 data before falling back to CourtListener's API:

- citation metadata is read from `public.courtlistener_citation_index`
- case cluster metadata is read from `public.courtlistener_opinion_cluster_index`
- cached opinion JSON is read from the R2 prefix `courtlistener/opinions/by-cluster/{clusterId}/{opinionId}.json`

If you do not import bulk data, leave `COURTLISTENER_BULK_DATA_ENABLED=false`; live CourtListener tools still work with a valid token, subject to CourtListener rate limits.

## Document Compare

Mike can produce a deterministic redline between two versions of a contract. Given a base `.docx` and a revised `.docx`, the compare engine computes their differences and returns two artifacts:

- A downloadable Microsoft Word `.docx` redline carrying **native Word tracked changes** (`<w:ins>` / `<w:del>` with `w:author` / `w:date`) that opens cleanly in both Microsoft Word and Google Docs.
- A structured **diff JSON** of ordered hunks that powers the in-app inline and side-by-side redline views.

The diff is **deterministic**: there is no AI in the diff path, and the engine uses no randomness and no wall-clock time internally. Time and authorship enter only through the caller-supplied `opts` (`opts.author` → `w:author`, `opts.date` → `w:date`). The determinism guarantee is therefore precise: **identical inputs together with identical `opts` (`author` and `date`) always produce byte-identical output** — which the golden-file tests pin by passing a fixed `opts`. The live route sets `opts.author` to the requesting user and `opts.date` to the request time, so two runs of the same documents at different times differ only in the `w:date` attribute.

### Entry flows

- **Compare two project documents.** Inside a project, open the new **Compare** tab (in *"two documents"* mode), pick a base document and a revised document from the project's documents, run Compare, view the result **inline** and **side-by-side**, then download the tracked-changes `.docx`.
- **Compare against a prior version.** In the **Compare** tab, switch to *"two versions of one document"* mode, pick a single document, then choose a **base** and a **revised** version (the pickers default to the prior version → the current version). You can also **upload a new revised version inline** from this view; the newly uploaded version is selected as *revised* and the previous version as *base*. Run Compare to redline the two versions of that one document.

### Where it lives

- Engine: `backend/src/compare/`
- Routes: `backend/src/routes/comparisons.ts`
- Persistence: `backend/src/lib/documentComparisons.ts`
- Table: `document_comparisons` (fresh databases via `backend/schema.sql`; existing databases via the dated migration in `backend/migrations/`)
- Frontend views: `frontend/src/app/components/compare/`, reached via the project **Compare** tab

### Endpoints

Compare routes follow Mike's no-`/api`-prefix convention, mount behind authentication, and enforce project access:

- `POST /projects/:projectId/comparisons` - create and run a comparison; body `{ baseDocumentId, revisedDocumentId, baseVersionId?, revisedVersionId? }`. Supply the optional `baseVersionId` / `revisedVersionId` to compare two explicit versions of a single document (entry flow (b)); the same document id is allowed on both sides only when the two version ids differ. Both selected documents must belong to the route project.
- `GET /comparisons/:id` - poll the comparison status and result
- `GET /comparisons/:id/download` - download the redline `.docx` (V1 reads the object fully buffered through the existing storage helper and sends it, mirroring the generic download path)

The redline `.docx` and `diff.json` are written to object storage under the `comparisons/` key prefix, reusing the existing R2/S3 bucket (no new storage layer).

### Extending the compare engine

The engine is isolated and composed of focused modules in `backend/src/compare/` - `parseDocx`, `normalize`, `alignParagraphs`, `wordDiff`, `emitTrackedChanges`, `diffJson`, `validate`, and `storageKeys` - orchestrated by `index.ts` (`runComparison`). It reuses existing libraries (`jszip`, `fast-xml-parser`, `fast-diff`) and helpers (`loadActiveVersion`, `downloadFile` / `uploadFile`, `docxToPdf`, `extractDocxBodyText`). Keep this module isolated: it must not be embedded in the assistant/chat, document-processing, tabular-review, or workflow code paths.

Run the scoped engine tests (Vitest config scoped to `src/compare/**`):

```bash
npm run test:compare --prefix backend
```

### Suggested next tasks

These are intentionally out of scope for V1 and are good follow-ups:

- Move detection (currently rendered as a delete plus an insert).
- Formatting-only change tracking (`<w:rPrChange>` / `<w:pPrChange>`).
- An async worker/queue (V1 computes the diff synchronously within the create request, while the `status` column and `GET /comparisons/:id` polling keep the contract async-ready).
- `.doc` / PDF convert-first support (V1 accepts `.docx` only).
- True streaming downloads (V1 buffers the redline through the existing storage helper and sends it, mirroring the generic download path; a streaming storage helper would avoid holding large redlines fully in memory).

## Install

Install each app package:

```bash
npm install --prefix backend
npm install --prefix frontend
```

## Run Locally

Start the backend:

```bash
npm run dev --prefix backend
```

Start the main app:

```bash
npm run dev --prefix frontend
```

Open `http://localhost:3000`.

## First Run

1. Sign up in the app.
2. If you did not set provider keys in `backend/.env`, open **Account > Models & API Keys** and add an Anthropic, Gemini, or OpenAI API key.
3. To use legal research tools, add a CourtListener token in `backend/.env` or **Account > Models & API Keys**.
4. Create or open a project and start chatting with documents.

## Troubleshooting

**Sign-up confirmation email never arrives.** Confirmation emails are sent by Supabase Auth, not by Mike. For local development, the simplest fix is to disable email confirmation in **Supabase > Authentication > Providers > Email**. For production, configure custom SMTP in Supabase; the built-in mailer is heavily rate-limited and may be restricted on newer projects.

**The model picker shows a missing-key warning.** Add a key for that provider in **Account > Models & API Keys**, or configure the provider key in `backend/.env` and restart the backend.

**CourtListener tools say the API token is missing.** Set `COURTLISTENER_API_TOKEN` in `backend/.env`, or add a CourtListener token in **Account > Models & API Keys** for the signed-in user. Restart the backend after changing `.env`.

**CourtListener bulk lookup is not returning local results.** Confirm `COURTLISTENER_BULK_DATA_ENABLED=true`, the two CourtListener tables have been populated, and opinion JSON exists in R2 under `courtlistener/opinions/by-cluster/`. If bulk data is unavailable, Mike falls back to the live API when a token is configured.

**DOC or DOCX conversion fails.** Install LibreOffice locally and restart the backend so document conversion commands are available on the process path.

**A comparison fails or the redline does not render.** Ensure LibreOffice is installed and on the process path (the same prerequisite as DOC/DOCX to PDF conversion), then restart the backend.

**Compare rejects a selected document.** V1 supports `.docx` only; legacy `.doc` and PDF inputs are rejected with a clear error, so convert them to `.docx` first. Pick the base and revised documents from the project's ready documents.

## Useful Checks

```bash
npm run build --prefix backend
npm run build --prefix frontend
npm run lint --prefix frontend
npm run test:compare --prefix backend
```
