"use client";

import React from "react";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { ProjectInfo } from "@/types";

interface ProjectTabProps {
  project: ProjectInfo;
  onUpdateField: (section: "project", field: string, value: any) => void;
}

export const ProjectTab: React.FC<ProjectTabProps> = ({
  project,
  onUpdateField,
}) => {
  return (
    <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em] mb-4">
        Project Parameters
      </h2>
      <div className="space-y-4">
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Contract/Project Title
          </Label>
          <Input
            value={project.projectName}
            onChange={(e) =>
              onUpdateField("project", "projectName", e.target.value)
            }
          />
        </div>
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Transmittal ID
          </Label>
          <Input
            className="font-mono text-[10px]"
            value={project.transmittalNumber}
            onChange={(e) =>
              onUpdateField("project", "transmittalNumber", e.target.value)
            }
            placeholder="Auto-generated (e.g. 202602-0001)"
          />
          <p className="text-[9px] text-slate-400 ml-2">
            Auto-generated but editable. Must be unique across all transmittals.
          </p>
        </div>
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Purpose
          </Label>
          <Input
            value={project.purpose}
            onChange={(e) =>
              onUpdateField("project", "purpose", e.target.value)
            }
          />
        </div>
      </div>
    </div>
  );
};
