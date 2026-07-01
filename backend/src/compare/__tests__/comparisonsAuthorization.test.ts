/**
 * Authorization regression tests for the Document Compare router's input
 * resolver (`resolveDocxInput` in `../../routes/comparisons`).
 *
 * SCOPE NOTE (intentional): unlike the sibling ENGINE suites in this directory
 * (which import only the engine + `vitest`), this file exercises the compare
 * FEATURE's HTTP-layer authorization. It therefore imports the router module
 * and mocks the reused backend libraries (`access`, `documentVersions`,
 * `storage`). It lives here so the compare-scoped Vitest `include` discovers it
 * WITHOUT broadening the config to unrelated backend suites — the comparisons
 * router is part of this feature, not an unrelated suite.
 *
 * These tests pin the CRITICAL cross-project isolation control added to
 * `resolveDocxInput`: a selected document MUST belong to the route `:projectId`
 * before ANY of its bytes are resolved/downloaded, so a user cannot pull a
 * document from another project into a comparison. They also cover the
 * version-aware resolution (flow (b)) and the existing not-found / non-.docx
 * guards.
 *
 * Build note: the backend `tsc` build compiles this file WITHOUT the
 * `vitest/globals` ambient types, so the test API (`describe`/`it`/`expect`/
 * `vi`) is imported explicitly to keep `npm run build` green under `strict`.
 */

import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the reused backend libraries the resolver depends on. Only the pieces
// `resolveDocxInput` touches need real behaviour; the rest are inert stubs so
// importing the router module (which also pulls these in) stays side-effect free.
vi.mock("../../lib/access", () => ({
  ensureDocAccess: vi.fn(),
  checkProjectAccess: vi.fn(),
}));
vi.mock("../../lib/documentVersions", () => ({
  loadActiveVersion: vi.fn(),
}));
vi.mock("../../lib/storage", () => ({
  downloadFile: vi.fn(),
  uploadFile: vi.fn(),
  buildContentDisposition: vi.fn(() => "attachment"),
}));

import { resolveDocxInput } from "../../routes/comparisons";
import { ensureDocAccess } from "../../lib/access";
import type { ActiveVersion } from "../../lib/documentVersions";
import { loadActiveVersion } from "../../lib/documentVersions";
import { downloadFile } from "../../lib/storage";

// ---------------------------------------------------------------------------
// Fixtures / builders
// ---------------------------------------------------------------------------

/** Minimal `documents` row as selected by the resolver. */
type DocRow = { id: string; user_id: string; project_id: string | null };

/**
 * Build a fake service-role client whose single query
 * (`.from("documents").select(...).eq(...).single()`) resolves to `docRow`
 * (or `{ data: null }` when `docRow` is null). Cast to the resolver's `db`
 * parameter type so no real Supabase client is constructed.
 */
function fakeDb(docRow: DocRow | null): Parameters<typeof resolveDocxInput>[0] {
  return {
    from() {
      return {
        select() {
          return {
            eq() {
              return {
                async single() {
                  return { data: docRow };
                },
              };
            },
          };
        },
      };
    },
  } as unknown as Parameters<typeof resolveDocxInput>[0];
}

/** Build a full `ActiveVersion` with sensible `.docx` defaults. */
function activeVersion(overrides: Partial<ActiveVersion> = {}): ActiveVersion {
  return {
    id: "ver-1",
    storage_path: "documents/doc-1/ver-1.docx",
    pdf_storage_path: null,
    version_number: 1,
    filename: "contract.docx",
    source: "upload",
    file_type: "docx",
    size_bytes: 1024,
    page_count: 1,
    ...overrides,
  };
}

beforeEach(() => {
  vi.clearAllMocks();
  // Default happy-path stubs; individual tests override as needed.
  vi.mocked(ensureDocAccess).mockResolvedValue({ ok: true, isOwner: true });
  vi.mocked(loadActiveVersion).mockResolvedValue(activeVersion());
  vi.mocked(downloadFile).mockResolvedValue(new Uint8Array([1, 2, 3]).buffer);
});

// ---------------------------------------------------------------------------
// CRITICAL — cross-project isolation
// ---------------------------------------------------------------------------

describe("resolveDocxInput: cross-project isolation (CRITICAL)", () => {
  it("rejects (404) a document whose project_id differs from the route projectId", async () => {
    const db = fakeDb({
      id: "docB",
      user_id: "owner-b",
      project_id: "project-B",
    });

    const result = await resolveDocxInput(
      db,
      "project-A", // route project
      "docB",
      "user-1",
      "user1@example.com",
    );

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.status).toBe(404);
  });

  it("does NOT access-check, resolve a version, or download bytes for a cross-project document", async () => {
    const db = fakeDb({
      id: "docB",
      user_id: "owner-b",
      project_id: "project-B",
    });

    await resolveDocxInput(db, "project-A", "docB", "user-1", "user1@example.com");

    // The project bind rejects BEFORE any content is touched, so none of these
    // downstream steps run for a document belonging to another project.
    expect(ensureDocAccess).not.toHaveBeenCalled();
    expect(loadActiveVersion).not.toHaveBeenCalled();
    expect(downloadFile).not.toHaveBeenCalled();
  });

  it("rejects cross-project even when an explicit versionId is supplied", async () => {
    const db = fakeDb({
      id: "docB",
      user_id: "owner-b",
      project_id: "project-B",
    });

    const result = await resolveDocxInput(
      db,
      "project-A",
      "docB",
      "user-1",
      "user1@example.com",
      "ver-from-B",
    );

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.status).toBe(404);
    expect(loadActiveVersion).not.toHaveBeenCalled();
    expect(downloadFile).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------
// Same-project resolution + version-aware flow
// ---------------------------------------------------------------------------

describe("resolveDocxInput: same-project resolution", () => {
  it("resolves a same-project .docx document and returns its bytes + resolved versionId", async () => {
    vi.mocked(loadActiveVersion).mockResolvedValue(
      activeVersion({ id: "ver-active" }),
    );
    const db = fakeDb({
      id: "docA",
      user_id: "owner-a",
      project_id: "project-A",
    });

    const result = await resolveDocxInput(
      db,
      "project-A",
      "docA",
      "user-1",
      "user1@example.com",
    );

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(Buffer.isBuffer(result.bytes)).toBe(true);
      expect(result.versionId).toBe("ver-active");
      expect(result.filename).toBe("contract.docx");
    }
    expect(downloadFile).toHaveBeenCalledTimes(1);
  });

  it("passes an explicit versionId straight through to loadActiveVersion (flow b)", async () => {
    vi.mocked(loadActiveVersion).mockResolvedValue(
      activeVersion({ id: "ver-prior" }),
    );
    const db = fakeDb({
      id: "docA",
      user_id: "owner-a",
      project_id: "project-A",
    });

    const result = await resolveDocxInput(
      db,
      "project-A",
      "docA",
      "user-1",
      "user1@example.com",
      "ver-prior",
    );

    expect(loadActiveVersion).toHaveBeenCalledWith("docA", db, "ver-prior");
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.versionId).toBe("ver-prior");
  });
});

// ---------------------------------------------------------------------------
// Existing guards (not-found, access denied, non-.docx)
// ---------------------------------------------------------------------------

describe("resolveDocxInput: guards", () => {
  it("returns 404 when the document row does not exist", async () => {
    const result = await resolveDocxInput(
      fakeDb(null),
      "project-A",
      "missing",
      "user-1",
      "user1@example.com",
    );
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.status).toBe(404);
  });

  it("returns 404 when document access is denied within the same project", async () => {
    vi.mocked(ensureDocAccess).mockResolvedValue({ ok: false });
    const db = fakeDb({
      id: "docA",
      user_id: "owner-a",
      project_id: "project-A",
    });

    const result = await resolveDocxInput(
      db,
      "project-A",
      "docA",
      "user-1",
      "user1@example.com",
    );

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.status).toBe(404);
    expect(downloadFile).not.toHaveBeenCalled();
  });

  it("rejects a non-.docx active version with 400 (no conversion in V1)", async () => {
    vi.mocked(loadActiveVersion).mockResolvedValue(
      activeVersion({ file_type: "pdf", filename: "brief.pdf" }),
    );
    const db = fakeDb({
      id: "docA",
      user_id: "owner-a",
      project_id: "project-A",
    });

    const result = await resolveDocxInput(
      db,
      "project-A",
      "docA",
      "user-1",
      "user1@example.com",
    );

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.status).toBe(400);
    expect(downloadFile).not.toHaveBeenCalled();
  });

  it("returns 404 when the requested version cannot be resolved (belongs to another document)", async () => {
    // loadActiveVersion returns null when the versionId does not belong to the
    // document — the resolver surfaces that as a 404 with a version-specific detail.
    vi.mocked(loadActiveVersion).mockResolvedValue(null);
    const db = fakeDb({
      id: "docA",
      user_id: "owner-a",
      project_id: "project-A",
    });

    const result = await resolveDocxInput(
      db,
      "project-A",
      "docA",
      "user-1",
      "user1@example.com",
      "ver-of-other-doc",
    );

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.status).toBe(404);
  });
});
