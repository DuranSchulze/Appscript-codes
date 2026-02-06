"use client";

import React from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { SenderInfo } from "../../types";

interface AgencyPresetModalProps {
  isOpen: boolean;
  onClose: () => void;
  draft: SenderInfo;
  onChange: (next: SenderInfo) => void;
  onLogoUpload: (e: React.ChangeEvent<HTMLInputElement>) => void;
  onSave: () => void;
}

export const AgencyPresetModal: React.FC<AgencyPresetModalProps> = ({
  isOpen,
  onClose,
  draft,
  onChange,
  onLogoUpload,
  onSave,
}) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[102] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-md animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-2xl rounded-[32px] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300 border border-white">
        <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
          <div>
            <h3 className="text-2xl font-black text-slate-800 font-display">
              Add Agency / Brand
            </h3>
            <p className="text-xs text-slate-400 mt-1 uppercase tracking-widest font-bold">
              Save presets for quick reuse
            </p>
          </div>
          <Button
            onClick={onClose}
            variant="ghost"
            size="icon"
            className="rounded-full"
          >
            ✕
          </Button>
        </div>

        <div className="p-8 space-y-5">
          <div className="flex items-center gap-4">
            <div className="w-20 h-20 rounded-3xl border border-slate-200 bg-white shadow-sm flex items-center justify-center overflow-hidden">
              {draft.logoBase64 ? (
                <img
                  src={draft.logoBase64}
                  alt="Agency logo"
                  className="w-full h-full object-contain"
                />
              ) : (
                <span className="text-[10px] font-black uppercase tracking-widest text-slate-300">
                  Logo
                </span>
              )}
            </div>
            <div className="flex-1">
              <Label className="text-[9px] font-black text-slate-400 uppercase tracking-widest ml-2">
                Upload Logo
              </Label>
              <input
                type="file"
                accept="image/*"
                onChange={onLogoUpload}
                className="block w-full mt-1 text-xs font-semibold text-slate-600 file:mr-4 file:py-2 file:px-4 file:rounded-2xl file:border-0 file:text-[10px] file:font-black file:uppercase file:tracking-widest file:bg-slate-900 file:text-white hover:file:bg-slate-800"
              />
            </div>
          </div>

          <div className="space-y-4">
            <div className="space-y-1">
              <Label className="text-[9px] font-black text-slate-400 uppercase tracking-widest ml-2">
                Agency Name
              </Label>
              <Input
                value={draft.agencyName}
                onChange={(e) =>
                  onChange({ ...draft, agencyName: e.target.value })
                }
              />
            </div>
            <div className="space-y-1">
              <Label className="text-[9px] font-black text-slate-400 uppercase tracking-widest ml-2">
                Address Line 1
              </Label>
              <Input
                value={draft.addressLine1}
                onChange={(e) =>
                  onChange({ ...draft, addressLine1: e.target.value })
                }
              />
            </div>
            <div className="space-y-1">
              <Label className="text-[9px] font-black text-slate-400 uppercase tracking-widest ml-2">
                Address Line 2
              </Label>
              <Input
                value={draft.addressLine2}
                onChange={(e) =>
                  onChange({ ...draft, addressLine2: e.target.value })
                }
              />
            </div>
            <div className="space-y-1">
              <Label className="text-[9px] font-black text-slate-400 uppercase tracking-widest ml-2">
                Website
              </Label>
              <Input
                value={draft.website}
                onChange={(e) =>
                  onChange({ ...draft, website: e.target.value })
                }
              />
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <Label className="text-[9px] font-black text-slate-400 uppercase tracking-widest ml-2">
                  Telephone
                </Label>
                <Input
                  value={draft.telephone}
                  onChange={(e) =>
                    onChange({ ...draft, telephone: e.target.value })
                  }
                />
              </div>
              <div className="space-y-1">
                <Label className="text-[9px] font-black text-slate-400 uppercase tracking-widest ml-2">
                  Mobile
                </Label>
                <Input
                  value={draft.mobile}
                  onChange={(e) =>
                    onChange({ ...draft, mobile: e.target.value })
                  }
                />
              </div>
            </div>
            <div className="space-y-1">
              <Label className="text-[9px] font-black text-slate-400 uppercase tracking-widest ml-2">
                Email
              </Label>
              <Input
                value={draft.email}
                onChange={(e) => onChange({ ...draft, email: e.target.value })}
              />
            </div>
          </div>

          <div className="flex gap-4 pt-2">
            <Button
              onClick={onClose}
              variant="ghost"
              className="flex-1 py-4 rounded-2xl font-bold"
            >
              Cancel
            </Button>
            <Button
              onClick={onSave}
              disabled={!draft.agencyName.trim()}
              className="flex-[2] py-4 bg-brand-600 text-white rounded-2xl text-[10px] font-black uppercase tracking-[0.2em] shadow-lg hover:bg-brand-500 disabled:opacity-50 transition-all active:scale-95"
            >
              Save Agency
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
};
