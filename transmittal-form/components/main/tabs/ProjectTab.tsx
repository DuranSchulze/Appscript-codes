"use client";

import React from "react";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import {
  Combobox,
  ComboboxContent,
  ComboboxEmpty,
  ComboboxInput,
  ComboboxItem,
  ComboboxList,
} from "@/components/ui/combobox";
import { ProjectInfo } from "@/types";

interface ProjectTabProps {
  project: ProjectInfo;
  projectNameSuggestions: string[];
  departmentSuggestions: string[];
  onUpdateField: (section: "project", field: string, value: any) => void;
}

export const ProjectTab: React.FC<ProjectTabProps> = ({
  project,
  projectNameSuggestions,
  departmentSuggestions,
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
          <Combobox
            value={project.projectName || null}
            onValueChange={(value) =>
              onUpdateField("project", "projectName", String(value || ""))
            }
            inputValue={project.projectName}
            onInputValueChange={(value) =>
              onUpdateField("project", "projectName", value)
            }
          >
            <ComboboxInput className="w-full" placeholder="Enter project title" />
            <ComboboxContent>
              <ComboboxEmpty>No saved project titles yet.</ComboboxEmpty>
              <ComboboxList>
                {projectNameSuggestions.map((suggestion) => (
                  <ComboboxItem key={suggestion} value={suggestion}>
                    {suggestion}
                  </ComboboxItem>
                ))}
              </ComboboxList>
            </ComboboxContent>
          </Combobox>
        </div>
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Project Number
          </Label>
          <Input
            className="font-mono text-[10px]"
            value={project.projectNumber}
            onChange={(e) =>
              onUpdateField("project", "projectNumber", e.target.value)
            }
            placeholder="Enter project reference number"
          />
          <p className="text-[9px] text-slate-400 ml-2">
            Project-specific identifier (separate from Transmittal ID).
          </p>
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
            Auto-generated document number for this transmittal. Editable, but
            must stay unique across all transmittals.
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
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Department
          </Label>
          <div className="flex gap-2">
            <Combobox
              value={project.department || null}
              onValueChange={(value) =>
                onUpdateField("project", "department", String(value || ""))
              }
              inputValue={project.department}
              onInputValueChange={(value) =>
                onUpdateField("project", "department", value)
              }
            >
              <ComboboxInput className="w-full" placeholder="Enter department" />
              <ComboboxContent>
                <ComboboxEmpty>No saved departments yet.</ComboboxEmpty>
                <ComboboxList>
                  {departmentSuggestions.map((suggestion) => (
                    <ComboboxItem key={suggestion} value={suggestion}>
                      {suggestion}
                    </ComboboxItem>
                  ))}
                </ComboboxList>
              </ComboboxContent>
            </Combobox>
            <Button
              type="button"
              variant="outline"
              onClick={() => onUpdateField("project", "department", "Admin")}
              className="whitespace-nowrap"
            >
              Use Admin
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
};
