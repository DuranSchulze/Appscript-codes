"use client";

import React from "react";
import { ChevronDown } from "lucide-react";

interface SidebarHeaderProps {
  isDriveReady: boolean;
  showPreview: boolean;
  onTogglePreview: (show: boolean) => void;
}

export const SidebarHeader: React.FC<SidebarHeaderProps> = ({
  isDriveReady,
  showPreview,
  onTogglePreview,
}) => {
  return (
    <div className="px-8 pt-6 pb-3 bg-white/40">
      <div className="flex justify-between items-center">
        <h1 className="font-display font-black text-2xl text-slate-900 tracking-tighter">
          Smart Transmittal
        </h1>
        <div className="flex items-center gap-3">
          <button
            onClick={() => onTogglePreview(true)}
            className="lg:hidden text-slate-400 hover:text-brand-600 transition-colors"
            title="Minimize Sidebar"
          >
            <ChevronDown className="w-5 h-5" />
          </button>
        </div>
      </div>
      <p className="text-[10px] text-slate-400 font-black uppercase tracking-widest mt-1">
        {isDriveReady ? "✓ Drive Connected" : "Enterprise Controller v2.1"}
      </p>
    </div>
  );
};
