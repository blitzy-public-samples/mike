"use client";

import { use } from "react";
import { ProjectSectionToolbar } from "@/app/components/projects/ProjectWorkspace";
import { CompareView } from "@/app/components/compare/CompareView";

interface Props {
    params: Promise<{ id: string }>;
}

export default function ProjectComparePage({ params }: Props) {
    use(params);
    return (
        <>
            <ProjectSectionToolbar />
            <CompareView />
        </>
    );
}
