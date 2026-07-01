/**
 * Document Compare / Redline — HTTP router.
 *
 * The HTTP surface for the Document Compare feature. It exposes three
 * endpoints:
 *   - POST /projects/:projectId/comparisons  create a comparison (synchronous
 *                                            V1 compute)
 *   - GET  /comparisons/:id                  poll status / result (+ diff once
 *                                            complete)
 *   - GET  /comparisons/:id/download         download the redline `.docx`
 *                                            (buffered send; see the handler)
 *
 * Mounted at the app root in `index.ts` (`app.use("/", comparisonsRouter)`), so
 * every path here is ABSOLUTE — no `mergeParams`, no `/api` prefix. All compare
 * logic lives in the isolated engine (`../compare`); this router only wires the
 * engine and persistence to HTTP using existing infra helpers as-is. Access
 * denials are masked as 404 (mirroring `downloads.ts`); the download handler is
 * self-contained and does NOT route through the generic `/download/:token`
 * endpoint (which resolves only `document_versions`).
 *
 * Two create flows are supported (AAP §0.1.1): (a) compare two DIFFERENT project
 * documents by id; (b) compare two DIFFERENT versions of the SAME document by
 * passing `baseVersionId` / `revisedVersionId` (e.g. a freshly uploaded revised
 * version vs. the prior one). When a version id is omitted the document's active
 * version is used. Every selected document is bound to the route `:projectId`
 * (cross-project isolation) before its bytes are ever downloaded.
 */

import { Router } from "express";
import { z } from "zod";
import { requireAuth } from "../middleware/auth";
import { createServerSupabase } from "../lib/supabase";
import { checkProjectAccess, ensureDocAccess } from "../lib/access";
import { loadActiveVersion } from "../lib/documentVersions";
import { downloadFile, uploadFile, buildContentDisposition } from "../lib/storage";
import { runComparison, type DiffJson } from "../compare";
import { comparisonRedlineKey, comparisonDiffKey } from "../compare/storageKeys";
import {
    insertComparison,
    getComparisonById,
    updateComparison,
} from "../lib/documentComparisons";
import { safeErrorMessage, safeErrorLog } from "../lib/safeError";

export const comparisonsRouter = Router(); // absolute paths; mounted at app root in index.ts — NO mergeParams, NO /api prefix

// Document compare: OOXML .docx mime, reused for the upload guard + the (buffered) download response.
const DOCX_MIME =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

// Document compare: request body contract for creating a comparison.
// `baseVersionId` / `revisedVersionId` are OPTIONAL: when present the caller
// pins an explicit `document_versions` row (flow (b) — compare two versions of
// the same document); when omitted the document's active version is used
// (flow (a) — compare two different documents at their current versions).
const createComparisonSchema = z.object({
    baseDocumentId: z.string().min(1),
    revisedDocumentId: z.string().min(1),
    baseVersionId: z.string().min(1).optional(),
    revisedVersionId: z.string().min(1).optional(),
});

// Document compare: derive a lowercased extension (no dot) from the active
// version's file_type, falling back to the filename's extension.
function extensionOf(fileType: string | null, filename: string | null): string {
    const ft = (fileType ?? "").toLowerCase().replace(/^\./, "");
    if (ft) return ft;
    const name = (filename ?? "").toLowerCase();
    const dot = name.lastIndexOf(".");
    return dot >= 0 ? name.slice(dot + 1) : "";
}

// Document compare: outcome of resolving a single input document to a specific
// version's `.docx` bytes. A discriminated union keeps the res early-returns in
// the handler rather than returning `res` from a helper. On success it also
// returns `versionId` — the resolved `document_versions.id` — so the create
// handler can persist exactly which version of each document was compared.
type DocxInput =
    | { ok: true; bytes: Buffer; filename: string | null; versionId: string }
    | { ok: false; status: number; detail: string };

// Document compare: run the per-document project binding + access check,
// version resolution, `.docx` guard, and byte download for one input document.
//
// Exported for targeted authorization regression tests (the cross-project
// binding below is a CRITICAL isolation control); it is otherwise internal to
// this router.
export async function resolveDocxInput(
    db: ReturnType<typeof createServerSupabase>,
    projectId: string,
    documentId: string,
    userId: string,
    userEmail: string | undefined,
    versionId?: string | null,
): Promise<DocxInput> {
    const { data: doc } = await db
        .from("documents")
        .select("id, user_id, project_id")
        .eq("id", documentId)
        .single();
    if (!doc) return { ok: false, status: 404, detail: "Document not found" };

    // SECURITY — cross-project isolation (do not remove): the selected document
    // MUST belong to the SAME project as the route's `:projectId`. `ensureDocAccess`
    // alone is insufficient — it grants access via the document's OWN project, so
    // a caller with access to project A and a document in project B could otherwise
    // create a comparison under A that carries B's content, which A collaborators
    // could then poll/download without any project B access. Bind the document to
    // the route project first, and mask a mismatch as 404 (same as an access denial).
    if ((doc as { project_id: string | null }).project_id !== projectId)
        return { ok: false, status: 404, detail: "Document not found" };

    const access = await ensureDocAccess(
        doc as { user_id: string; project_id: string | null },
        userId,
        userEmail,
        db,
    );
    if (!access.ok)
        return { ok: false, status: 404, detail: "Document not found" };

    // Resolve the requested version (flow (b)) or fall back to the document's
    // active version (flow (a)). `loadActiveVersion` verifies a supplied
    // `versionId` actually belongs to `documentId` (returns null otherwise), so
    // a caller cannot pull a version row from a different document.
    const active = await loadActiveVersion(documentId, db, versionId);
    if (!active)
        return {
            ok: false,
            status: 404,
            detail: versionId
                ? "Requested version not available for this document"
                : "No version available for this document",
        };

    // Document compare: V1 accepts ONLY native .docx — reject legacy .doc / PDF
    // clearly (no conversion, no silent failure).
    if (extensionOf(active.file_type, active.filename) !== "docx")
        return {
            ok: false,
            status: 400,
            detail: "Only .docx documents can be compared in this version.",
        };

    const ab = await downloadFile(active.storage_path);
    if (!ab)
        return { ok: false, status: 404, detail: "Document file not found" };

    return {
        ok: true,
        bytes: Buffer.from(ab),
        filename: active.filename,
        versionId: active.id,
    };
}

// POST /projects/:projectId/comparisons — create a comparison (synchronous V1
// compute). Steps run in order: parse body → checkProjectAccess → resolve both
// inputs (access + .docx guard + bytes) → insert a processing row → run the
// engine → upload redline + diff → update row to complete → respond 201.
comparisonsRouter.post(
    "/projects/:projectId/comparisons",
    requireAuth,
    async (req, res) => {
        const userId = res.locals.userId as string;
        const userEmail = res.locals.userEmail as string | undefined;
        const { projectId } = req.params;

        const parsed = createComparisonSchema.safeParse(req.body);
        if (!parsed.success)
            return void res.status(400).json({
                detail: "baseDocumentId and revisedDocumentId are required.",
            });
        const {
            baseDocumentId,
            revisedDocumentId,
            baseVersionId,
            revisedVersionId,
        } = parsed.data;

        // A comparison must be between two DISTINCT inputs. Flow (a) compares two
        // different documents; flow (b) compares two different VERSIONS of the
        // same document. Reject the degenerate "same document, same version"
        // request (which would diff a file against itself).
        if (
            baseDocumentId === revisedDocumentId &&
            (!baseVersionId ||
                !revisedVersionId ||
                baseVersionId === revisedVersionId)
        )
            return void res.status(400).json({
                detail:
                    "Select two different documents, or two different versions of the same document, to compare.",
            });

        const db = createServerSupabase();

        // Mask project denials as 404.
        const projectAccess = await checkProjectAccess(
            projectId,
            userId,
            userEmail,
            db,
        );
        if (!projectAccess.ok)
            return void res.status(404).json({ detail: "Project not found" });

        // Resolve + project-bind + access-check + .docx-guard + download BOTH
        // inputs (passing the route projectId so each document is confirmed to
        // belong to THIS project, and the optional per-input version id).
        const base = await resolveDocxInput(
            db,
            projectId,
            baseDocumentId,
            userId,
            userEmail,
            baseVersionId,
        );
        if (!base.ok)
            return void res.status(base.status).json({ detail: base.detail });
        const revised = await resolveDocxInput(
            db,
            projectId,
            revisedDocumentId,
            userId,
            userEmail,
            revisedVersionId,
        );
        if (!revised.ok)
            return void res
                .status(revised.status)
                .json({ detail: revised.detail });

        // Persist a processing row up-front (status-polling keeps the contract
        // async-ready even though V1 computes synchronously). Record the
        // resolved version ids so the row captures exactly which versions were
        // compared (essential when both inputs are the same document).
        const row = await insertComparison(db, {
            project_id: projectId,
            base_document_id: baseDocumentId,
            revised_document_id: revisedDocumentId,
            base_version_id: base.versionId,
            revised_version_id: revised.versionId,
            created_by: userId,
            status: "processing",
        });

        try {
            // w:author must be a valid, non-empty string; res.locals.userEmail
            // may be "" so fall back to userId. w:date is the run time.
            const author = userEmail || userId;
            const date = new Date().toISOString();

            const { redlineBytes, diff } = await runComparison(
                base.bytes,
                revised.bytes,
                { author, date },
            );

            const redlineKey = comparisonRedlineKey(userId, row.id);
            const diffKey = comparisonDiffKey(userId, row.id);

            // Pass a clean, exactly-sized ArrayBuffer for each artifact: a fresh
            // Uint8Array copy for the redline (avoids a pooled Buffer's offset)
            // and TextEncoder for the diff JSON.
            await uploadFile(
                redlineKey,
                new Uint8Array(redlineBytes).buffer,
                DOCX_MIME,
            );
            await uploadFile(
                diffKey,
                new TextEncoder().encode(JSON.stringify(diff)).buffer,
                "application/json",
            );

            const updated = await updateComparison(db, row.id, {
                status: "complete",
                redline_storage_path: redlineKey,
                diff_storage_path: diffKey,
            });

            return void res.status(201).json({ ...(updated ?? row), diff });
        } catch (err) {
            console.error("[comparisons/create] error:", safeErrorLog(err));
            const errored = await updateComparison(db, row.id, {
                status: "error",
                error: safeErrorMessage(err),
            }).catch(() => null);
            return void res
                .status(500)
                .json(errored ?? { detail: safeErrorMessage(err) });
        }
    },
);

// GET /comparisons/:id — poll status / result. Returns the snake_case row and,
// when complete, the parsed diff JSON fetched from storage.
comparisonsRouter.get("/comparisons/:id", requireAuth, async (req, res) => {
    const userId = res.locals.userId as string;
    const userEmail = res.locals.userEmail as string | undefined;
    const { id } = req.params;

    const db = createServerSupabase();

    const comparison = await getComparisonById(db, id);
    if (!comparison)
        return void res.status(404).json({ detail: "Comparison not found" });

    const access = await checkProjectAccess(
        comparison.project_id,
        userId,
        userEmail,
        db,
    );
    if (!access.ok)
        return void res.status(404).json({ detail: "Comparison not found" });

    // Response mirrors the DB row (snake_case) = the frontend Comparison type.
    const payload: Record<string, unknown> = { ...comparison };

    if (comparison.status === "complete" && comparison.diff_storage_path) {
        const raw = await downloadFile(comparison.diff_storage_path);
        if (raw) {
            try {
                payload.diff = JSON.parse(
                    Buffer.from(raw).toString("utf-8"),
                ) as DiffJson;
            } catch {
                // Omit diff if it cannot be parsed; status/result still returned.
            }
        }
    }

    return void res.status(200).json(payload);
});

// GET /comparisons/:id/download — self-contained, access-checked redline
// download. It buffers the artifact via `downloadFile` (which returns the whole
// object as an ArrayBuffer) and sends it with `res.send` — the SAME buffered
// response pattern the generic `downloads.ts` uses for document versions. It is
// deliberately self-contained and does NOT route through the generic
// `/download/:token` endpoint (which resolves only `document_versions`, so a
// `comparisons/`-prefixed redline would 404 there). V1 is buffer-based; true
// object-stream piping is a documented future enhancement (see docs/decisions).
comparisonsRouter.get(
    "/comparisons/:id/download",
    requireAuth,
    async (req, res) => {
        const userId = res.locals.userId as string;
        const userEmail = res.locals.userEmail as string | undefined;
        const { id } = req.params;

        const db = createServerSupabase();

        const comparison = await getComparisonById(db, id);
        if (!comparison)
            return void res
                .status(404)
                .json({ detail: "Comparison not found" });

        const access = await checkProjectAccess(
            comparison.project_id,
            userId,
            userEmail,
            db,
        );
        if (!access.ok)
            return void res
                .status(404)
                .json({ detail: "Comparison not found" });

        if (
            comparison.status !== "complete" ||
            !comparison.redline_storage_path
        )
            return void res.status(404).json({ detail: "Redline not ready" });

        const raw = await downloadFile(comparison.redline_storage_path);
        if (!raw)
            return void res.status(404).json({ detail: "Redline not found" });

        res.setHeader("Content-Type", DOCX_MIME);
        res.setHeader(
            "Content-Disposition",
            buildContentDisposition("attachment", "comparison-redline.docx"),
        );
        res.send(Buffer.from(raw));
    },
);
