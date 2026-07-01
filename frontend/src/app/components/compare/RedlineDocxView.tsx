"use client";

import { useEffect, useRef, useState } from "react";
import { Loader2 } from "lucide-react";
import { downloadComparisonRedline } from "@/app/lib/mikeApi";

interface Props {
    comparisonId: string;
}

export function RedlineDocxView({ comparisonId }: Props) {
    const scrollRef = useRef<HTMLDivElement>(null);
    const containerRef = useRef<HTMLDivElement>(null);
    const [bytes, setBytes] = useState<ArrayBuffer | null>(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);

    // Fetch the redline .docx bytes for this comparison.
    useEffect(() => {
        let cancelled = false;
        setLoading(true);
        setError(null);
        setBytes(null);
        (async () => {
            try {
                const { blob } = await downloadComparisonRedline(comparisonId);
                const buf = await blob.arrayBuffer();
                if (cancelled) return;
                setBytes(buf);
            } catch (e) {
                if (cancelled) return;
                console.error("redline fetch failed", e);
                setError("Failed to load redline document.");
            } finally {
                if (!cancelled) setLoading(false);
            }
        })();
        return () => {
            cancelled = true;
        };
    }, [comparisonId]);

    // Render the bytes with docx-preview (tracked changes -> <ins>/<del>).
    useEffect(() => {
        let cancelled = false;
        if (!bytes || !containerRef.current) return;
        const containerEl = containerRef.current;
        (async () => {
            try {
                const { renderAsync } = await import("docx-preview");
                if (cancelled) return;
                containerEl.innerHTML = "";
                await renderAsync(bytes, containerEl, undefined, {
                    inWrapper: true,
                    ignoreWidth: false,
                    ignoreHeight: false,
                    renderChanges: true,
                    experimental: true,
                });
            } catch (e) {
                console.error("docx-preview render failed", e);
                if (!cancelled) setError("Failed to render redline document.");
            }
        })();
        return () => {
            cancelled = true;
        };
    }, [bytes]);

    return (
        <div className="relative flex flex-col flex-1 overflow-hidden">
            <div
                ref={scrollRef}
                className="flex-1 overflow-auto bg-muted px-5 pt-5 pb-3"
            >
                {loading && (
                    // role="status" (aria-live="polite" by default) announces the
                    // loading state; the spinner icon is decorative (aria-hidden).
                    <div
                        role="status"
                        className="flex h-full items-center justify-center"
                    >
                        <Loader2
                            className="h-7 w-7 animate-spin text-muted-foreground"
                            aria-hidden="true"
                        />
                        <span className="sr-only">Loading redline document…</span>
                    </div>
                )}
                {error && !loading && (
                    // role="alert" (aria-live="assertive" by default) announces the
                    // fetch/render failure to assistive technology.
                    <div
                        role="alert"
                        className="flex h-full items-center justify-center"
                    >
                        <p className="text-sm text-destructive">{error}</p>
                    </div>
                )}
                <div ref={containerRef} className="docx-view-container" />
            </div>
        </div>
    );
}
