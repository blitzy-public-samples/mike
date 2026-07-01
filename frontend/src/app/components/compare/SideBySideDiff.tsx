"use client";

import { useMemo, useRef, useState } from "react";
import { ChevronLeft, ChevronRight } from "lucide-react";
import type { DiffJson } from "@/app/components/shared/types";
import { cn } from "@/lib/utils";

interface Props {
    diff: DiffJson | null | undefined;
}

export function SideBySideDiff({ diff }: Props) {
    // Stable reference for the hunks array. A bare `diff?.hunks ?? []` fallback
    // allocates a fresh `[]` on every render, changing identity each render and
    // tripping react-hooks/exhaustive-deps on the `changedIndices` memo below
    // (and defeating its memoization). Memoizing on `diff` keeps the reference
    // stable until the comparison actually changes.
    const hunks = useMemo(() => diff?.hunks ?? [], [diff]);

    // Indices (into `hunks`) of changed (del/ins) hunks, in document order.
    const changedIndices = useMemo(
        () =>
            hunks
                .map((h, i) => (h.type === "equal" ? -1 : i))
                .filter((i) => i >= 0),
        [hunks],
    );

    const [currentChange, setCurrentChange] = useState(0);

    // Reset change navigation to the first change whenever a NEW diff arrives
    // (e.g. the user runs a different comparison). Without this, a `currentChange`
    // left over from a previous, larger comparison would point past the end of
    // the new `changedIndices` and render invalid text like "Change 3 of 1".
    // This is React's recommended "adjust state during render" pattern for
    // resetting state in response to a prop change (no effect / extra frame).
    const [seenDiff, setSeenDiff] = useState(diff);
    if (seenDiff !== diff) {
        setSeenDiff(diff);
        setCurrentChange(0);
    }

    // Defensive clamp into [0, changedIndices.length - 1]: the indicator, the
    // highlighted hunk, and the nav base always reference a VALID changed hunk
    // for the current diff, even during a render where `currentChange` is
    // momentarily out of range. Guarantees `currentHunkIdx` (below) is never
    // undefined. When there are no changes the value is 0 (unused: the buttons
    // are disabled and `currentHunkIdx` is -1).
    const safeChange = Math.max(
        0,
        Math.min(currentChange, changedIndices.length - 1),
    );

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
    const currentHunkIdx = hasChanges ? changedIndices[safeChange] : -1;

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
                            ? `Change ${safeChange + 1} of ${changedIndices.length}`
                            : "No changes"}
                    </span>
                    <button
                        type="button"
                        onClick={() => goToChange(safeChange - 1)}
                        disabled={!hasChanges}
                        aria-label="Previous change"
                        className="rounded-md border border-gray-200 p-1 text-gray-600 transition-colors hover:bg-gray-100 disabled:opacity-40"
                    >
                        <ChevronLeft className="h-4 w-4" />
                    </button>
                    <button
                        type="button"
                        onClick={() => goToChange(safeChange + 1)}
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
