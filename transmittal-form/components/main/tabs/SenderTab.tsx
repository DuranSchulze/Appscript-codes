"use client";

import React from "react";
import { ImagePlus } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
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
  onOpenAgencyModal: () => void;
  sender: SenderInfo;
  onUpdateField: (section: "sender", field: string, value: any) => void;
  logoInputRef: React.RefObject<HTMLInputElement | null>;
  logoInputKey: number;
  onLogoUpload: (e: React.ChangeEvent<HTMLInputElement>) => void;
}

export const SenderTab: React.FC<SenderTabProps> = ({
  agencies,
  selectedAgencyId,
  onSelectAgency,
  onOpenAgencyModal,
  sender,
  onUpdateField,
  logoInputRef,
  logoInputKey,
  onLogoUpload,
}) => {
  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
        Sender Branding
      </h2>

      <div className="flex flex-col sm:flex-row gap-3 items-end">
        <div className="flex-1 space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Saved Agencies
          </Label>
          <Select
            value={selectedAgencyId}
            onValueChange={(val) => onSelectAgency(val as string)}
          >
            <SelectTrigger className="w-full">
              <SelectValue placeholder="Select agency..." />
            </SelectTrigger>
            <SelectContent>
              {agencies.map((agency) => (
                <SelectItem key={agency.id} value={agency.id}>
                  {agency.name}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
        <Button
          onClick={onOpenAgencyModal}
          className="h-[52px] px-5 rounded-2xl text-[10px] font-black uppercase tracking-[0.2em] shadow-lg"
        >
          Add
        </Button>
      </div>

      <div className="flex flex-col items-center gap-6 p-6 bg-slate-50 rounded-[40px] border border-slate-200/60">
        <div
          className="relative group cursor-pointer"
          onClick={() => logoInputRef.current?.click()}
        >
          <div className="w-32 h-32 rounded-full border-4 border-white shadow-xl flex items-center justify-center bg-white overflow-hidden transition-transform group-hover:scale-105">
            {sender.logoBase64 ? (
              <img
                src={sender.logoBase64}
                alt="Company Logo"
                className="w-full h-full object-contain"
              />
            ) : (
              <div className="flex flex-col items-center text-slate-300">
                <ImagePlus className="w-8 h-8 mb-1" />
                <span className="text-[8px] font-black uppercase tracking-widest">
                  Add Logo
                </span>
              </div>
            )}
          </div>
          <input
            key={logoInputKey}
            type="file"
            ref={logoInputRef}
            className="hidden"
            accept="image/*"
            onChange={onLogoUpload}
          />
        </div>
        <div className="text-center">
          <p className="text-[10px] font-black text-slate-800 uppercase tracking-widest">
            {sender.agencyName}
          </p>
        </div>
      </div>

      <div className="space-y-5">
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Agency Name
          </Label>
          <Input
            value={sender.agencyName}
            onChange={(e) =>
              onUpdateField("sender", "agencyName", e.target.value)
            }
          />
        </div>
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Website
          </Label>
          <Input
            value={sender.website}
            onChange={(e) => onUpdateField("sender", "website", e.target.value)}
          />
        </div>
        <div className="grid grid-cols-2 gap-4">
          <div className="space-y-1">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Telephone
            </Label>
            <Input
              value={sender.telephone}
              onChange={(e) =>
                onUpdateField("sender", "telephone", e.target.value)
              }
            />
          </div>
          <div className="space-y-1">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Mobile
            </Label>
            <Input
              value={sender.mobile}
              onChange={(e) =>
                onUpdateField("sender", "mobile", e.target.value)
              }
            />
          </div>
        </div>
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Email
          </Label>
          <Input
            value={sender.email}
            onChange={(e) => onUpdateField("sender", "email", e.target.value)}
          />
        </div>
      </div>
    </div>
  );
};
