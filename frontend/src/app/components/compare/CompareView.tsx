"use client";

import { useEffect, useRef, useState } from "react";
import { ArrowLeftRight, ChevronDown, Download, Loader2 } from "lucide-react";
import type { Comparison, Document } from "@/app/components/shared/types";
import {
    createComparison,
    downloadComparisonRedline,
    getComparison,
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

function DocumentPicker({
    label,
    options,
    value,
    onChange,
    disabledId,
}: {
    label: string;
    options: Document[];
    value: string;
    onChange: (id: string) => void;
    disabledId: string;
}) {
    const selected = options.find((d) => d.id === value);
    return (
        <div className="flex flex-col gap-1">
            <span className="text-xs font-medium text-gray-700">{label}</span>
            <DropdownMenu>
                <DropdownMenuTrigger asChild>
                    <Button
                        variant="outline"
                        className="w-64 justify-between font-normal"
                    >
                        <span
                            className={cn(
                                "truncate",
                                !selected && "text-gray-400",
                            )}
                        >
                            {selected ? selected.filename : "Select a document"}
                        </span>
                        <ChevronDown className="h-4 w-4 shrink-0 text-gray-400" />
                    </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent className="max-h-64 w-64 overflow-y-auto">
                    <DropdownMenuRadioGroup value={value} onValueChange={onChange}>
                        {options.map((d) => (
                            <DropdownMenuRadioItem
                                key={d.id}
                                value={d.id}
                                disabled={d.id === disabledId}
                            >
                                {d.filename}
                            </DropdownMenuRadioItem>
                        ))}
                    </DropdownMenuRadioGroup>
                </DropdownMenuContent>
            </DropdownMenu>
        </div>
    );
}

export function CompareView() {
    const { project, projectId } = useProjectWorkspace();

    const docOptions: Document[] = (project?.documents ?? []).filter(
        (d) => d.status === "ready" && d.file_type === "docx",
    );

    const [baseDocumentId, setBaseDocumentId] = useState("");
    const [revisedDocumentId, setRevisedDocumentId] = useState("");
    const [comparison, setComparison] = useState<Comparison | null>(null);
    const [running, setRunning] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [view, setView] = useState<ResultView>("inline");

    const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);

    // Clear any in-flight poll on unmount.
    useEffect(() => {
        return () => {
            if (pollRef.current) clearInterval(pollRef.current);
        };
    }, []);

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

    const handleRun = async () => {
        if (!baseDocumentId || !revisedDocumentId || running) return;
        setRunning(true);
        setError(null);
        setComparison(null);
        setView("inline");
        try {
            const created = await createComparison(projectId, {
                baseDocumentId,
                revisedDocumentId,
            });
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

    const canRun =
        !!baseDocumentId &&
        !!revisedDocumentId &&
        baseDocumentId !== revisedDocumentId &&
        !running;

    const isComplete = comparison?.status === "complete";

    return (
        <div className="flex-1 flex flex-col min-h-0 gap-4 p-4">
            {/* Pickers + Run + status */}
            <div className="flex flex-wrap items-end gap-3">
                <DocumentPicker
                    label="Base version"
                    options={docOptions}
                    value={baseDocumentId}
                    onChange={setBaseDocumentId}
                    disabledId={revisedDocumentId}
                />
                <DocumentPicker
                    label="Revised version"
                    options={docOptions}
                    value={revisedDocumentId}
                    onChange={setRevisedDocumentId}
                    disabledId={baseDocumentId}
                />
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

            {docOptions.length < 2 && (
                <p className="text-sm text-gray-500">
                    You need at least two ready .docx documents in this project
                    to run a comparison.
                </p>
            )}

            {error && <p className="text-sm text-red-500">{error}</p>}

            {/* Result panel */}
            {isComplete && comparison && (
                <div className="flex flex-1 flex-col min-h-0 overflow-hidden rounded-lg border border-gray-200">
                    <div className="flex items-center justify-between border-b border-gray-200 px-3 py-2">
                        <div className="flex items-center gap-1">
                            <Button
                                variant={view === "inline" ? "default" : "ghost"}
                                size="sm"
                                onClick={() => setView("inline")}
                            >
                                Inline redline
                            </Button>
                            <Button
                                variant={
                                    view === "sidebyside" ? "default" : "ghost"
                                }
                                size="sm"
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
