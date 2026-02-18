"use client";

import React from "react";
import {
  Building2,
  Globe,
  Mail,
  Phone,
  Smartphone,
  Trash2,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectTrigger,
  SelectValue,
  SelectContent,
  SelectItem,
} from "@/components/ui/select";
import { SenderInfo } from "@/types";

interface DbAgency {
  id: string;
  name: string;
}

interface SenderTabProps {
  agencies: DbAgency[];
  selectedAgencyId: string;
  onSelectAgency: (id: string) => void;
  onOpenAgencyModal: (mode: "create" | "update") => void;
  onDeleteAgency: () => void;
  sender: SenderInfo;
}

export const SenderTab: React.FC<SenderTabProps> = ({
  agencies,
  selectedAgencyId,
  onSelectAgency,
  onOpenAgencyModal,
  onDeleteAgency,
  sender,
}) => {
  const hasSelectedAgency = Boolean(selectedAgencyId);
  const readOnlyValue = (value: string) => value?.trim() || "—";

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
        Sender Branding
      </h2>

      <div className="space-y-3">
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Saved Agencies
          </Label>
          <Select
            value={selectedAgencyId}
            onValueChange={(val) => onSelectAgency(val as string)}
            items={agencies.map((agency) => ({
              value: agency.id,
              label: agency.name || "Unnamed agency",
            }))}
          >
            <SelectTrigger className="w-full">
              <SelectValue placeholder="Select agency..." />
            </SelectTrigger>
            <SelectContent>
              {agencies.map((agency) => (
                <SelectItem key={agency.id} value={agency.id}>
                  {agency.name || "Unnamed agency"}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>

        <div className="flex flex-wrap gap-2">
          <Button
            onClick={() =>
              onOpenAgencyModal(hasSelectedAgency ? "update" : "create")
            }
            className="h-11 flex-1 min-w-[120px] rounded-2xl text-[10px] font-black uppercase tracking-[0.2em] shadow-lg"
          >
            {hasSelectedAgency ? "Update" : "Add"}
          </Button>

          {hasSelectedAgency ? (
            <Button
              onClick={() => onOpenAgencyModal("create")}
              variant="outline"
              className="h-11 flex-1 min-w-[120px] rounded-2xl text-[10px] font-black uppercase tracking-[0.2em]"
            >
              Add New
            </Button>
          ) : null}

          <Button
            onClick={onDeleteAgency}
            variant="ghost"
            disabled={!hasSelectedAgency}
            className="h-11 flex-1 min-w-[120px] rounded-2xl text-[10px] font-black uppercase tracking-[0.2em] text-red-600 hover:text-red-700 hover:bg-red-50 disabled:text-slate-300"
          >
            <Trash2 className="mr-1.5 h-3.5 w-3.5" />
            Delete
          </Button>
        </div>
      </div>

      <div className="space-y-5 p-6 bg-slate-50 rounded-[40px] border border-slate-200/60">
        <div className="flex items-center gap-4">
          <div className="w-20 h-20 rounded-3xl border border-white shadow-md flex items-center justify-center bg-white overflow-hidden">
            {sender.logoBase64 ? (
              <img
                src={sender.logoBase64}
                alt="Company Logo"
                className="w-full h-full object-contain"
              />
            ) : (
              <Building2 className="w-7 h-7 text-slate-300" />
            )}
          </div>

          <div>
            <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">
              Active Agency
            </p>
            <p className="text-xs font-black text-slate-800 uppercase tracking-widest mt-1">
              {readOnlyValue(sender.agencyName)}
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div className="space-y-1">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Address Line 1
            </Label>
            <p className="px-2 text-xs font-semibold text-slate-700">
              {readOnlyValue(sender.addressLine1)}
            </p>
          </div>

          <div className="space-y-1">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Address Line 2
            </Label>
            <p className="px-2 text-xs font-semibold text-slate-700">
              {readOnlyValue(sender.addressLine2)}
            </p>
          </div>

          <div className="space-y-1 sm:col-span-2">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Website
            </Label>
            <div className="px-2 flex items-center gap-2 text-xs text-slate-700">
              <Globe className="h-3.5 w-3.5 text-slate-400" />
              <span className="truncate">{readOnlyValue(sender.website)}</span>
            </div>
          </div>

          <div className="space-y-1">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Telephone
            </Label>
            <div className="px-2 flex items-center gap-2 text-xs text-slate-700">
              <Phone className="h-3.5 w-3.5 text-slate-400" />
              <span className="truncate">
                {readOnlyValue(sender.telephone)}
              </span>
            </div>
          </div>

          <div className="space-y-1">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Mobile
            </Label>
            <div className="px-2 flex items-center gap-2 text-xs text-slate-700">
              <Smartphone className="h-3.5 w-3.5 text-slate-400" />
              <span className="truncate">{readOnlyValue(sender.mobile)}</span>
            </div>
          </div>

          <div className="space-y-1 sm:col-span-2">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Email
            </Label>
            <div className="px-2 flex items-center gap-2 text-xs text-slate-700">
              <Mail className="h-3.5 w-3.5 text-slate-400" />
              <span className="truncate">{readOnlyValue(sender.email)}</span>
            </div>
          </div>

          <div className="space-y-1 sm:col-span-2">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Agency Name
            </Label>
            <p className="px-2 text-xs font-semibold text-slate-700">
              {readOnlyValue(sender.agencyName)}
            </p>
          </div>
        </div>

        <p className="text-[10px] font-semibold text-slate-400">
          Brand details are read-only here. Use Add/Update to manage agency
          data.
        </p>
      </div>
    </div>
  );
};
