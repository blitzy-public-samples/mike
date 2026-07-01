"use client";

import { useMemo, useRef, useState } from "react";
import { ChevronLeft, ChevronRight } from "lucide-react";
import type { DiffJson } from "@/app/components/shared/types";
import { cn } from "@/lib/utils";

interface Props {
    diff: DiffJson | null | undefined;
}

export function SideBySideDiff({ diff }: Props) {
    const hunks = diff?.hunks ?? [];

    // Indices (into `hunks`) of changed (del/ins) hunks, in document order.
    const changedIndices = useMemo(
        () =>
            hunks
                .map((h, i) => (h.type === "equal" ? -1 : i))
                .filter((i) => i >= 0),
        [hunks],
    );

    const [currentChange, setCurrentChange] = useState(0);
    // hunk-index -> rendered <span>, for scroll-into-view navigation.
    const hunkRefs = useRef<Record<number, HTMLSpanElement | null>>({});

    const goToChange = (target: number) => {
        if (changedIndices.length === 0) return;
        const count = changedIndices.length;
        const next = ((target % count) + count) % count;
        setCurrentChange(next);
        hunkRefs.current[changedIndices[next]]?.scrollIntoView({
            behavior: "smooth",
            block: "center",
        });
    };

    if (!diff || hunks.length === 0) {
        return (
            <div className="flex flex-1 items-center justify-center p-6">
                <p className="text-sm text-gray-400">
                    No differences to display.
                </p>
            </div>
        );
    }

    const hasChanges = changedIndices.length > 0;
    const currentHunkIdx = hasChanges ? changedIndices[currentChange] : -1;

    return (
        <div className="flex flex-1 flex-col overflow-hidden">
            {/* Change-navigation toolbar */}
            <div className="flex items-center justify-between border-b border-gray-200 px-4 py-2">
                <div className="flex items-center gap-4 text-xs text-gray-500">
                    <span className="flex items-center gap-1.5">
                        <span className="inline-block h-2 w-2 rounded-full bg-red-600" />
                        Base
                    </span>
                    <span className="flex items-center gap-1.5">
                        <span className="inline-block h-2 w-2 rounded-full bg-green-600" />
                        Revised
                    </span>
                </div>
                <div className="flex items-center gap-2">
                    <span className="text-xs text-gray-500">
                        {hasChanges
                            ? `Change ${currentChange + 1} of ${changedIndices.length}`
                            : "No changes"}
                    </span>
                    <button
                        type="button"
                        onClick={() => goToChange(currentChange - 1)}
                        disabled={!hasChanges}
                        aria-label="Previous change"
                        className="rounded-md border border-gray-200 p-1 text-gray-600 transition-colors hover:bg-gray-100 disabled:opacity-40"
                    >
                        <ChevronLeft className="h-4 w-4" />
                    </button>
                    <button
                        type="button"
                        onClick={() => goToChange(currentChange + 1)}
                        disabled={!hasChanges}
                        aria-label="Next change"
                        className="rounded-md border border-gray-200 p-1 text-gray-600 transition-colors hover:bg-gray-100 disabled:opacity-40"
                    >
                        <ChevronRight className="h-4 w-4" />
                    </button>
                </div>
            </div>

            {/* Two independently-scrollable columns */}
            <div className="flex flex-1 min-h-0 divide-x divide-gray-200">
                {/* Base column: equal + del */}
                <div className="flex-1 min-w-0 overflow-auto px-4 py-3">
                    <p className="mb-2 text-xs font-medium uppercase tracking-wide text-gray-400">
                        Base
                    </p>
                    <div className="whitespace-pre-wrap text-sm leading-relaxed text-gray-800">
                        {hunks.map((h, i) => {
                            if (h.type === "ins") return null;
                            if (h.type === "del") {
                                return (
                                    <span
                                        key={i}
                                        ref={(el) => {
                                            hunkRefs.current[i] = el;
                                        }}
                                        className={cn(
                                            "rounded-sm bg-red-50 text-red-600 line-through",
                                            currentHunkIdx === i &&
                                                "ring-2 ring-red-300",
                                        )}
                                    >
                                        {h.text}
                                    </span>
                                );
                            }
                            return <span key={i}>{h.text}</span>;
                        })}
                    </div>
                </div>

                {/* Revised column: equal + ins */}
                <div className="flex-1 min-w-0 overflow-auto px-4 py-3">
                    <p className="mb-2 text-xs font-medium uppercase tracking-wide text-gray-400">
                        Revised
                    </p>
                    <div className="whitespace-pre-wrap text-sm leading-relaxed text-gray-800">
                        {hunks.map((h, i) => {
                            if (h.type === "del") return null;
                            if (h.type === "ins") {
                                return (
                                    <span
                                        key={i}
                                        ref={(el) => {
                                            hunkRefs.current[i] = el;
                                        }}
                                        className={cn(
                                            "rounded-sm bg-green-50 text-green-600",
                                            currentHunkIdx === i &&
                                                "ring-2 ring-green-300",
                                        )}
                                    >
                                        {h.text}
                                    </span>
                                );
                            }
                            return <span key={i}>{h.text}</span>;
                        })}
                    </div>
                </div>
            </div>
        </div>
    );
}
