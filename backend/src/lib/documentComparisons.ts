/**
 * Persistence helpers for the `document_comparisons` table.
 *
 * Provides insert / select-by-id / update access for the Document Compare /
 * Redline feature through the service-role Supabase client. Consumed only by
 * the comparisons router; this module is pure persistence and contains no
 * diff / OOXML / redline logic.
 */

import type { createServerSupabase } from "./supabase";

type Db = ReturnType<typeof createServerSupabase>;

const TABLE = "document_comparisons";

export type ComparisonStatus = "pending" | "processing" | "complete" | "error";

export interface DocumentComparisonRow {
    id: string; // uuid pk
    project_id: string; // uuid, FK -> projects(id)
    base_document_id: string; // uuid, FK -> documents(id)
    revised_document_id: string; // uuid, FK -> documents(id)
    created_by: string | null; // text (nullable)
    status: ComparisonStatus; // text, CHECK pending|processing|complete|error, default 'pending'
    redline_storage_path: string | null; // text (nullable until complete)
    diff_storage_path: string | null; // text (nullable until complete)
    error: string | null; // text (nullable; set when status='error')
    created_at: string; // timestamptz (ISO string)
    updated_at: string; // timestamptz (ISO string)
}

/**
 * Insert a new comparison row and return the persisted record. When omitted,
 * `status` defaults to "processing" and is always written explicitly. Throws
 * if the insert fails.
 */
export async function insertComparison(
    db: Db,
    input: {
        project_id: string;
        base_document_id: string;
        revised_document_id: string;
        created_by: string | null;
        status?: ComparisonStatus;
    },
): Promise<DocumentComparisonRow> {
    const { data, error } = await db
        .from(TABLE)
        .insert({
            project_id: input.project_id,
            base_document_id: input.base_document_id,
            revised_document_id: input.revised_document_id,
            created_by: input.created_by,
            status: input.status ?? "processing",
        })
        .select()
        .single();
    if (error) throw new Error(`insertComparison failed: ${error.message}`);
    return data as DocumentComparisonRow;
}

/**
 * Fetch a single comparison row by id, returning null when no row matches.
 */
export async function getComparisonById(
    db: Db,
    id: string,
): Promise<DocumentComparisonRow | null> {
    const { data } = await db
        .from(TABLE)
        .select("*")
        .eq("id", id)
        .maybeSingle();
    return (data as DocumentComparisonRow | null) ?? null;
}

/**
 * Apply a partial update to the mutable columns of a comparison row, always
 * refreshing `updated_at`. Throws if the update fails; returns the updated
 * row, or null if no row matched.
 */
export async function updateComparison(
    db: Db,
    id: string,
    patch: Partial<
        Pick<
            DocumentComparisonRow,
            "status" | "redline_storage_path" | "diff_storage_path" | "error"
        >
    >,
): Promise<DocumentComparisonRow | null> {
    const { data, error } = await db
        .from(TABLE)
        .update({ ...patch, updated_at: new Date().toISOString() })
        .eq("id", id)
        .select()
        .single();
    if (error) throw new Error(`updateComparison failed: ${error.message}`);
    return (data as DocumentComparisonRow | null) ?? null;
}
