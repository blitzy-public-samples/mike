/**
 * Document Compare / Redline — HTTP router.
 *
 * The HTTP surface for the Document Compare feature. It exposes three
 * endpoints:
 *   - POST /projects/:projectId/comparisons  create a comparison (synchronous
 *                                            V1 compute)
 *   - GET  /comparisons/:id                  poll status / result (+ diff once
 *                                            complete)
 *   - GET  /comparisons/:id/download         stream the redline `.docx`
 *
 * Mounted at the app root in `index.ts` (`app.use("/", comparisonsRouter)`), so
 * every path here is ABSOLUTE — no `mergeParams`, no `/api` prefix. All compare
 * logic lives in the isolated engine (`../compare`); this router only wires the
 * engine and persistence to HTTP using existing infra helpers as-is. Access
 * denials are masked as 404 (mirroring `downloads.ts`); the download handler is
 * self-contained and does NOT route through the generic `/download/:token`
 * endpoint (which resolves only `document_versions`).
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

// Document compare: OOXML .docx mime, reused for upload + download streaming.
const DOCX_MIME =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

// Document compare: request body contract for creating a comparison.
const createComparisonSchema = z.object({
    baseDocumentId: z.string().min(1),
    revisedDocumentId: z.string().min(1),
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

// Document compare: outcome of resolving a single input document to its active
// `.docx` bytes. A discriminated union keeps the res early-returns in the
// handler rather than returning `res` from a helper.
type DocxInput =
    | { ok: true; bytes: Buffer; filename: string | null }
    | { ok: false; status: number; detail: string };

// Document compare: run the per-document access check, active-version
// resolution, `.docx` guard, and byte download for one input document.
async function resolveDocxInput(
    db: ReturnType<typeof createServerSupabase>,
    documentId: string,
    userId: string,
    userEmail: string | undefined,
): Promise<DocxInput> {
    const { data: doc } = await db
        .from("documents")
        .select("id, user_id, project_id")
        .eq("id", documentId)
        .single();
    if (!doc) return { ok: false, status: 404, detail: "Document not found" };

    const access = await ensureDocAccess(
        doc as { user_id: string; project_id: string | null },
        userId,
        userEmail,
        db,
    );
    if (!access.ok)
        return { ok: false, status: 404, detail: "Document not found" };

    const active = await loadActiveVersion(documentId, db);
    if (!active)
        return {
            ok: false,
            status: 404,
            detail: "No version available for this document",
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

    return { ok: true, bytes: Buffer.from(ab), filename: active.filename };
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
        const { baseDocumentId, revisedDocumentId } = parsed.data;

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

        // Resolve + access-check + .docx-guard + download BOTH inputs.
        const base = await resolveDocxInput(db, baseDocumentId, userId, userEmail);
        if (!base.ok)
            return void res.status(base.status).json({ detail: base.detail });
        const revised = await resolveDocxInput(
            db,
            revisedDocumentId,
            userId,
            userEmail,
        );
        if (!revised.ok)
            return void res
                .status(revised.status)
                .json({ detail: revised.detail });

        // Persist a processing row up-front (status-polling keeps the contract
        // async-ready even though V1 computes synchronously).
        const row = await insertComparison(db, {
            project_id: projectId,
            base_document_id: baseDocumentId,
            revised_document_id: revisedDocumentId,
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

// GET /comparisons/:id/download — self-contained, access-checked redline stream.
// Replicates the streaming response shape of downloads.ts WITHOUT importing it.
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
