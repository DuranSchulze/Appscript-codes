"use client";

import React from "react";
import { Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";

interface DriveFileModalProps {
  isOpen: boolean;
  onClose: () => void;
  files: Array<{ id: string; name: string; mimeType: string }>;
  searchValue: string;
  isLoading: boolean;
  error: string;
  selectedIds: Record<string, boolean>;
  isAllSelected: boolean;
  onSearchChange: (value: string) => void;
  onSearch: () => void;
  onToggle: (id: string) => void;
  onToggleAll: () => void;
  onAddSelected: () => void;
  isImporting?: boolean;
  processingMessage?: string;
}

export const DriveFileModal: React.FC<DriveFileModalProps> = ({
  isOpen,
  onClose,
  files,
  searchValue,
  isLoading,
  error,
  selectedIds,
  isAllSelected,
  onSearchChange,
  onSearch,
  onToggle,
  onToggleAll,
  onAddSelected,
  isImporting = false,
  processingMessage = "",
}) => {
  if (!isOpen) return null;
  const selectedCount = Object.values(selectedIds).filter(Boolean).length;
  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-md animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-3xl rounded-[32px] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300 border border-white">
        <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
          <div>
            <h3 className="text-2xl font-black text-slate-800 font-display">
              Browse Drive Files
            </h3>
            <p className="text-xs text-slate-400 mt-1 uppercase tracking-widest font-bold">
              Search and select documents to import
            </p>
          </div>
          <Button
            onClick={onClose}
            variant="ghost"
            size="icon"
            className="rounded-full"
            disabled={isImporting}
          >
            ✕
          </Button>
        </div>
        <div className="p-8 space-y-5">
          <div className="flex flex-col sm:flex-row gap-3">
            <div className="flex-1">
              <Input
                value={searchValue}
                onChange={(e) => onSearchChange(e.target.value)}
                placeholder="Search by file name"
                disabled={isImporting}
              />
            </div>
            <Button onClick={onSearch} disabled={isLoading || isImporting}>
              {isImporting ? (
                <>
                  <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                  Working...
                </>
              ) : (
                "Search"
              )}
            </Button>
          </div>

          <div className="border border-slate-100 rounded-[24px] overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 bg-slate-50/70 border-b border-slate-100">
              <label className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-widest text-slate-500">
                <input
                  type="checkbox"
                  checked={isAllSelected}
                  onChange={onToggleAll}
                  disabled={files.length === 0 || isImporting}
                  className="w-4 h-4 rounded-md text-brand-600 border-slate-300"
                />
                Select All
              </label>
              <Badge variant="secondary">
                {files.length} file{files.length === 1 ? "" : "s"}
              </Badge>
            </div>
            <div className="max-h-[320px] overflow-y-auto">
              {isLoading ? (
                <div className="p-6 text-xs font-semibold text-slate-400">
                  Loading Drive files...
                </div>
              ) : files.length === 0 ? (
                <div className="p-6 text-xs font-semibold text-slate-400">
                  {error || "No files found."}
                </div>
              ) : (
                <ul className="divide-y divide-slate-100">
                  {files.map((file) => (
                    <li key={file.id} className="flex items-center gap-3 p-4">
                      <input
                        type="checkbox"
                        checked={!!selectedIds[file.id]}
                        onChange={() => onToggle(file.id)}
                        disabled={isImporting}
                        className="w-4 h-4 rounded-md text-brand-600 border-slate-300"
                      />
                      <div className="flex-1">
                        <p className="text-xs font-semibold text-slate-800">
                          {file.name}
                        </p>
                        <p className="text-[10px] text-slate-400">
                          {file.mimeType}
                        </p>
                      </div>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          </div>

          {error && files.length > 0 && (
            <div className="p-3 rounded-2xl text-[10px] font-bold border bg-red-50 border-red-100 text-red-600">
              {error}
            </div>
          )}

          {isImporting ? (
            <div className="p-3 rounded-2xl text-[10px] font-bold border bg-brand-50 border-brand-100 text-brand-700 flex items-center gap-2">
              <Loader2 className="w-3.5 h-3.5 animate-spin" />
              {processingMessage || "Importing selected Drive files..."}
            </div>
          ) : null}

          <div className="flex flex-col sm:flex-row gap-3">
            <Button
              onClick={onClose}
              variant="ghost"
              className="flex-1 py-3 rounded-2xl font-bold"
              disabled={isImporting}
            >
              Cancel
            </Button>
            <Button
              onClick={onAddSelected}
              disabled={selectedCount === 0 || isImporting}
              className="flex-1 py-3 rounded-2xl bg-brand-600 text-white text-[10px] font-black uppercase tracking-[0.2em] shadow-lg hover:bg-brand-500 disabled:opacity-50"
            >
              {isImporting ? (
                <>
                  <Loader2 className="w-3.5 h-3.5 mr-2 animate-spin" />
                  Importing...
                </>
              ) : (
                `Add Selected (${selectedCount})`
              )}
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
};
