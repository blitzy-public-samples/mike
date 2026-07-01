"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
    ArrowLeftRight,
    ChevronDown,
    Download,
    Loader2,
    Upload,
} from "lucide-react";
import type {
    Comparison,
    CreateComparisonInput,
    Document,
} from "@/app/components/shared/types";
import {
    createComparison,
    downloadComparisonRedline,
    getComparison,
    listDocumentVersions,
    uploadDocumentVersion,
    type DocumentVersion,
} from "@/app/lib/mikeApi";
import { useProjectWorkspace } from "@/app/components/projects/ProjectWorkspace";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import {
    DropdownMenu,
    DropdownMenuContent,
    DropdownMenuRadioGroup,
    DropdownMenuRadioItem,
    DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { cn } from "@/lib/utils";
import { RedlineDocxView } from "./RedlineDocxView";
import { SideBySideDiff } from "./SideBySideDiff";

type ResultView = "inline" | "sidebyside";

// The two AAP-mandated entry flows:
//   - "documents": pick two DIFFERENT project documents (compare their active
//     versions) — flow (a);
//   - "versions": pick ONE document and compare two of its versions, optionally
//     uploading a new revised version first — flow (b).
type CompareMode = "documents" | "versions";

function statusBadgeVariant(
    status: Comparison["status"],
): "default" | "secondary" | "destructive" {
    if (status === "complete") return "default";
    if (status === "error") return "destructive";
    return "secondary";
}

function statusLabel(status: Comparison["status"]): string {
    if (status === "complete") return "Complete";
    if (status === "error") return "Error";
    if (status === "processing") return "Processing…";
    return "Pending…";
}

/** One selectable option for the shared {@link SelectField}. */
interface SelectOption {
    id: string;
    label: string;
}

/**
 * A token-styled dropdown select built on the shared `DropdownMenu` primitive
 * (the design system has no `Select` component — see AAP §0.5.3). Used for both
 * the document pickers and the version pickers so the two entry flows share one
 * consistent, accessible control.
 */
function SelectField({
    label,
    options,
    value,
    onChange,
    disabledId,
    placeholder,
    widthClass = "w-64",
}: {
    label: string;
    options: SelectOption[];
    value: string;
    onChange: (id: string) => void;
    disabledId?: string;
    placeholder: string;
    widthClass?: string;
}) {
    const selected = options.find((o) => o.id === value);
    return (
        <div className="flex flex-col gap-1">
            <span className="text-xs font-medium text-foreground">{label}</span>
            <DropdownMenu>
                <DropdownMenuTrigger asChild>
                    <Button
                        variant="outline"
                        className={cn(widthClass, "justify-between font-normal")}
                    >
                        <span
                            className={cn(
                                "truncate",
                                !selected && "text-muted-foreground",
                            )}
                        >
                            {selected ? selected.label : placeholder}
                        </span>
                        <ChevronDown className="h-4 w-4 shrink-0 text-muted-foreground" />
                    </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent
                    className={cn("max-h-64 overflow-y-auto", widthClass)}
                >
                    <DropdownMenuRadioGroup
                        value={value}
                        onValueChange={onChange}
                    >
                        {options.map((o) => (
                            <DropdownMenuRadioItem
                                key={o.id}
                                value={o.id}
                                disabled={o.id === disabledId}
                            >
                                {o.label}
                            </DropdownMenuRadioItem>
                        ))}
                    </DropdownMenuRadioGroup>
                </DropdownMenuContent>
            </DropdownMenu>
        </div>
    );
}

/** Human-readable label for a document version option. */
function versionLabel(v: DocumentVersion): string {
    const n = v.version_number != null ? `v${v.version_number}` : "version";
    const name = v.filename ?? "untitled";
    return `${n} · ${name}`;
}

export function CompareView() {
    const { project, projectId } = useProjectWorkspace();

    const docOptions: Document[] = useMemo(
        () =>
            (project?.documents ?? []).filter(
                (d) => d.status === "ready" && d.file_type === "docx",
            ),
        [project?.documents],
    );

    // Entry-flow mode + per-mode selection state.
    const [mode, setMode] = useState<CompareMode>("documents");

    // Flow (a): two different documents.
    const [baseDocumentId, setBaseDocumentId] = useState("");
    const [revisedDocumentId, setRevisedDocumentId] = useState("");

    // Flow (b): one document, two versions.
    const [singleDocId, setSingleDocId] = useState("");
    const [versions, setVersions] = useState<DocumentVersion[]>([]);
    const [versionsLoading, setVersionsLoading] = useState(false);
    const [baseVersionId, setBaseVersionId] = useState("");
    const [revisedVersionId, setRevisedVersionId] = useState("");
    const [uploading, setUploading] = useState(false);

    const [comparison, setComparison] = useState<Comparison | null>(null);
    const [running, setRunning] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [view, setView] = useState<ResultView>("inline");

    const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);
    const fileInputRef = useRef<HTMLInputElement | null>(null);

    // Clear any in-flight poll on unmount.
    useEffect(() => {
        return () => {
            if (pollRef.current) clearInterval(pollRef.current);
        };
    }, []);

    // Flow (b): (re)load the selected document's versions and default the
    // pickers to "compare the current version against the prior one" — the
    // AAP's upload-new-version / prior-version comparison. Guarded against races
    // with a `cancelled` flag so a slow response for an earlier document never
    // overwrites a newer selection.
    useEffect(() => {
        if (mode !== "versions" || !singleDocId) {
            setVersions([]);
            setBaseVersionId("");
            setRevisedVersionId("");
            return;
        }
        let cancelled = false;
        setVersionsLoading(true);
        (async () => {
            try {
                const { current_version_id, versions: list } =
                    await listDocumentVersions(singleDocId);
                if (cancelled) return;
                setVersions(list);
                applyDefaultVersionSelection(list, current_version_id);
            } catch (e) {
                if (cancelled) return;
                console.error("listDocumentVersions failed", e);
                setVersions([]);
                setError("Failed to load version history.");
            } finally {
                if (!cancelled) setVersionsLoading(false);
            }
        })();
        return () => {
            cancelled = true;
        };
    }, [mode, singleDocId]);

    // Default the revised picker to the current version and the base picker to
    // the immediately-prior version (highest version_number below the current),
    // so the common "compare the latest against the previous" case needs no
    // extra clicks. Versions are sorted by version_number descending.
    const applyDefaultVersionSelection = (
        list: DocumentVersion[],
        currentId: string | null,
    ) => {
        const sorted = [...list].sort(
            (a, b) => (b.version_number ?? 0) - (a.version_number ?? 0),
        );
        const revised =
            (currentId && sorted.find((v) => v.id === currentId)?.id) ||
            sorted[0]?.id ||
            "";
        const prior = sorted.find((v) => v.id !== revised)?.id ?? "";
        setRevisedVersionId(revised);
        setBaseVersionId(prior);
    };

    const isTerminal = (s: Comparison["status"]) =>
        s === "complete" || s === "error";

    const startPolling = (id: string) => {
        if (pollRef.current) clearInterval(pollRef.current);
        pollRef.current = setInterval(async () => {
            try {
                const next = await getComparison(id);
                setComparison(next);
                if (isTerminal(next.status)) {
                    if (pollRef.current) clearInterval(pollRef.current);
                    pollRef.current = null;
                    setRunning(false);
                    if (next.status === "error") {
                        setError(next.error ?? "Comparison failed.");
                    }
                }
            } catch (e) {
                console.error("poll failed", e);
                if (pollRef.current) clearInterval(pollRef.current);
                pollRef.current = null;
                setRunning(false);
                setError("Failed to check comparison status.");
            }
        }, 1000);
    };

    // Upload a new revised version of the selected document (flow (b)), then
    // reload the version list and pre-select the new upload as the revised side
    // with the previous current version as the base.
    const handleUploadNewVersion = async (
        e: React.ChangeEvent<HTMLInputElement>,
    ) => {
        const file = e.target.files?.[0];
        // Allow re-uploading the same file name later by clearing the input.
        e.target.value = "";
        if (!file || !singleDocId || uploading) return;
        setUploading(true);
        setError(null);
        try {
            const created = await uploadDocumentVersion(singleDocId, file);
            const { current_version_id, versions: list } =
                await listDocumentVersions(singleDocId);
            setVersions(list);
            // Revised = the just-uploaded version; base = the prior current.
            setRevisedVersionId(created.id);
            const prior =
                (current_version_id &&
                    current_version_id !== created.id &&
                    current_version_id) ||
                [...list]
                    .sort(
                        (a, b) =>
                            (b.version_number ?? 0) - (a.version_number ?? 0),
                    )
                    .find((v) => v.id !== created.id)?.id ||
                "";
            setBaseVersionId(prior);
        } catch (err) {
            console.error("uploadDocumentVersion failed", err);
            setError(
                err instanceof Error
                    ? err.message
                    : "Failed to upload the new version.",
            );
        } finally {
            setUploading(false);
        }
    };

    const handleRun = async () => {
        if (!canRun) return;
        setRunning(true);
        setError(null);
        setComparison(null);
        setView("inline");

        // Build the request per entry flow. Flow (b) sends the SAME document id
        // for both sides plus the two version ids; flow (a) sends two document
        // ids and no version ids (backend falls back to each active version).
        const payload: CreateComparisonInput =
            mode === "versions"
                ? {
                      baseDocumentId: singleDocId,
                      revisedDocumentId: singleDocId,
                      baseVersionId,
                      revisedVersionId,
                  }
                : { baseDocumentId, revisedDocumentId };

        try {
            const created = await createComparison(projectId, payload);
            setComparison(created);
            if (isTerminal(created.status)) {
                setRunning(false);
                if (created.status === "error") {
                    setError(created.error ?? "Comparison failed.");
                }
            } else {
                startPolling(created.id);
            }
        } catch (e) {
            console.error("createComparison failed", e);
            setRunning(false);
            setError(
                e instanceof Error ? e.message : "Failed to start comparison.",
            );
        }
    };

    // Lazily fetch the diff JSON if side-by-side needs it.
    useEffect(() => {
        if (
            view !== "sidebyside" ||
            !comparison ||
            comparison.status !== "complete" ||
            comparison.diff
        ) {
            return;
        }
        let cancelled = false;
        (async () => {
            try {
                const full = await getComparison(comparison.id);
                if (!cancelled) setComparison(full);
            } catch (e) {
                console.error("diff fetch failed", e);
            }
        })();
        return () => {
            cancelled = true;
        };
    }, [view, comparison]);

    const handleDownload = async () => {
        if (!comparison) return;
        try {
            const { blob, filename } = await downloadComparisonRedline(
                comparison.id,
            );
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = filename ?? "comparison-redline.docx";
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 1000);
        } catch (e) {
            console.error("download failed", e);
            setError("Failed to download redline.");
        }
    };

    // Run is enabled once the active flow has a valid, distinct pair selected.
    const canRun =
        !running &&
        (mode === "documents"
            ? !!baseDocumentId &&
              !!revisedDocumentId &&
              baseDocumentId !== revisedDocumentId
            : !!singleDocId &&
              !!baseVersionId &&
              !!revisedVersionId &&
              baseVersionId !== revisedVersionId);

    const isComplete = comparison?.status === "complete";

    const versionOptions: SelectOption[] = versions.map((v) => ({
        id: v.id,
        label: versionLabel(v),
    }));
    const docSelectOptions: SelectOption[] = docOptions.map((d) => ({
        id: d.id,
        label: d.filename,
    }));

    return (
        <div className="flex-1 flex flex-col min-h-0 gap-4 p-4">
            {/* Entry-flow mode toggle. flex-wrap so the two mode buttons wrap
                rather than overflowing the viewport at narrow (<= 375px) widths
                (QA FIN-C Issue 2). */}
            <div
                className="flex flex-wrap items-center gap-1"
                role="group"
                aria-label="Comparison mode"
            >
                <Button
                    variant={mode === "documents" ? "default" : "ghost"}
                    size="sm"
                    aria-pressed={mode === "documents"}
                    onClick={() => setMode("documents")}
                >
                    Two documents
                </Button>
                <Button
                    variant={mode === "versions" ? "default" : "ghost"}
                    size="sm"
                    aria-pressed={mode === "versions"}
                    onClick={() => setMode("versions")}
                >
                    Two versions of one document
                </Button>
            </div>

            {/* Pickers + Run + status */}
            <div className="flex flex-wrap items-end gap-3">
                {mode === "documents" ? (
                    <>
                        <SelectField
                            label="Base version"
                            options={docSelectOptions}
                            value={baseDocumentId}
                            onChange={setBaseDocumentId}
                            disabledId={revisedDocumentId}
                            placeholder="Select a document"
                        />
                        <SelectField
                            label="Revised version"
                            options={docSelectOptions}
                            value={revisedDocumentId}
                            onChange={setRevisedDocumentId}
                            disabledId={baseDocumentId}
                            placeholder="Select a document"
                        />
                    </>
                ) : (
                    <>
                        <SelectField
                            label="Document"
                            options={docSelectOptions}
                            value={singleDocId}
                            onChange={setSingleDocId}
                            placeholder="Select a document"
                        />
                        <SelectField
                            label="Base (prior) version"
                            options={versionOptions}
                            value={baseVersionId}
                            onChange={setBaseVersionId}
                            disabledId={revisedVersionId}
                            placeholder={
                                versionsLoading
                                    ? "Loading versions…"
                                    : "Select base version"
                            }
                            widthClass="w-56"
                        />
                        <SelectField
                            label="Revised version"
                            options={versionOptions}
                            value={revisedVersionId}
                            onChange={setRevisedVersionId}
                            disabledId={baseVersionId}
                            placeholder={
                                versionsLoading
                                    ? "Loading versions…"
                                    : "Select revised version"
                            }
                            widthClass="w-56"
                        />
                        <div className="flex flex-col gap-1">
                            <span className="text-xs font-medium text-foreground">
                                &nbsp;
                            </span>
                            <Button
                                variant="outline"
                                disabled={!singleDocId || uploading}
                                onClick={() => fileInputRef.current?.click()}
                            >
                                {uploading ? (
                                    <>
                                        <Loader2 className="h-4 w-4 animate-spin" />
                                        Uploading…
                                    </>
                                ) : (
                                    <>
                                        <Upload className="h-4 w-4" />
                                        Upload new revised version
                                    </>
                                )}
                            </Button>
                            <input
                                ref={fileInputRef}
                                type="file"
                                accept=".docx"
                                className="hidden"
                                onChange={handleUploadNewVersion}
                            />
                        </div>
                    </>
                )}

                <Button onClick={handleRun} disabled={!canRun}>
                    {running ? (
                        <>
                            <Loader2 className="h-4 w-4 animate-spin" />
                            Comparing…
                        </>
                    ) : (
                        "Compare"
                    )}
                </Button>
                {comparison && (
                    <Badge variant={statusBadgeVariant(comparison.status)}>
                        {statusLabel(comparison.status)}
                    </Badge>
                )}
            </div>

            {mode === "documents" && docOptions.length < 2 && (
                <p className="text-sm text-muted-foreground">
                    You need at least two ready .docx documents in this project
                    to run a comparison.
                </p>
            )}
            {mode === "versions" &&
                !!singleDocId &&
                !versionsLoading &&
                versions.length < 2 && (
                    <p className="text-sm text-muted-foreground">
                        This document has only one version. Upload a new revised
                        version to compare it against the current one.
                    </p>
                )}

            {error && <p className="text-sm text-destructive">{error}</p>}

            {/* Result panel */}
            {isComplete && comparison && (
                <div className="flex flex-1 flex-col min-h-0 overflow-hidden rounded-lg border border-border">
                    {/* flex-wrap + gap-2 lets the Download button wrap below the
                        view toggles instead of being clipped by the parent's
                        overflow-hidden at narrow (<= 375px) widths (QA FIN-C Issue 2). */}
                    <div className="flex flex-wrap items-center justify-between gap-2 border-b border-border px-3 py-2">
                        <div className="flex items-center gap-1">
                            <Button
                                variant={view === "inline" ? "default" : "ghost"}
                                size="sm"
                                aria-pressed={view === "inline"}
                                onClick={() => setView("inline")}
                            >
                                Inline redline
                            </Button>
                            <Button
                                variant={
                                    view === "sidebyside" ? "default" : "ghost"
                                }
                                size="sm"
                                aria-pressed={view === "sidebyside"}
                                onClick={() => setView("sidebyside")}
                            >
                                <ArrowLeftRight className="h-4 w-4" />
                                Side-by-side
                            </Button>
                        </div>
                        <Button
                            variant="outline"
                            size="sm"
                            onClick={handleDownload}
                        >
                            <Download className="h-4 w-4" />
                            Download redline (.docx)
                        </Button>
                    </div>
                    <div className="flex flex-1 min-h-0 overflow-hidden">
                        {view === "inline" ? (
                            <RedlineDocxView comparisonId={comparison.id} />
                        ) : (
                            <SideBySideDiff diff={comparison.diff} />
                        )}
                    </div>
                </div>
            )}
        </div>
    );
}
