"use client";

import React, { useState, useEffect } from "react";
import {
  FileText,
  Trash2,
  Loader2,
  FolderOpen,
  CalendarDays,
  X,
} from "lucide-react";
import { format, parse } from "date-fns";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Calendar } from "@/components/ui/calendar";
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from "@/components/ui/popover";

export interface TransmittalSummary {
  id: string;
  transmittalNumber: string;
  projectName: string;
  recipientName: string;
  date: string;
  itemCount: number;
}

interface TransmittalListModalProps {
  isOpen: boolean;
  onClose: () => void;
  onOpenTransmittal: (id: string) => void;
  onDeleteTransmittal: (id: string) => Promise<void>;
  apiBaseUrl: string;
}

export const TransmittalListModal: React.FC<TransmittalListModalProps> = ({
  isOpen,
  onClose,
  onOpenTransmittal,
  onDeleteTransmittal,
  apiBaseUrl,
}) => {
  const [transmittals, setTransmittals] = useState<TransmittalSummary[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [search, setSearch] = useState("");
  const [dateFilter, setDateFilter] = useState<Date | undefined>(undefined);
  const [isCalendarOpen, setIsCalendarOpen] = useState(false);
  const [deletingId, setDeletingId] = useState<string | null>(null);

  useEffect(() => {
    if (isOpen) {
      fetchTransmittals();
      setSearch("");
      setDateFilter(undefined);
    }
  }, [isOpen]);

  const fetchTransmittals = async () => {
    setIsLoading(true);
    try {
      const response = await fetch(`${apiBaseUrl}/api/transmittals`, {
        credentials: "include",
      });
      if (!response.ok) return;
      const payload = await response.json();
      const list: TransmittalSummary[] = Array.isArray(payload.transmittals)
        ? payload.transmittals.map((t: any) => {
            const project = t.project || {};
            const recipient = Array.isArray(t.recipients)
              ? t.recipients[0]
              : null;
            return {
              id: t.id,
              transmittalNumber: project.transmittalNumber || "",
              projectName: project.projectName || t.projectName || "Untitled",
              recipientName: recipient?.recipientName || t.recipientName || "",
              date: project.date || t.createdAt?.split("T")[0] || "",
              itemCount: Array.isArray(t.items) ? t.items.length : 0,
            };
          })
        : [];
      setTransmittals(list);
    } catch (error) {
      console.error("Failed to fetch transmittals", error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleDelete = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setDeletingId(id);
    try {
      await onDeleteTransmittal(id);
      setTransmittals((prev) => prev.filter((t) => t.id !== id));
    } catch {
      // handled by parent
    } finally {
      setDeletingId(null);
    }
  };

  if (!isOpen) return null;

  const filtered = transmittals
    .filter((t) => {
      const q = search.toLowerCase();
      const matchesSearch =
        !q ||
        t.transmittalNumber.toLowerCase().includes(q) ||
        t.projectName.toLowerCase().includes(q) ||
        t.recipientName.toLowerCase().includes(q) ||
        t.date.includes(q);
      const dateFilterStr = dateFilter ? format(dateFilter, "yyyy-MM-dd") : "";
      const matchesDate = !dateFilterStr || t.date === dateFilterStr;
      return matchesSearch && matchesDate;
    })
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-md animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-xl rounded-[32px] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300 border border-white flex flex-col max-h-[80vh]">
        <div className="p-8 pb-4 border-b border-slate-100 bg-slate-50/50">
          <div className="flex justify-between items-center mb-4">
            <div>
              <h3 className="text-2xl font-black text-slate-800 font-display">
                Open Transmittal
              </h3>
              <p className="text-xs text-slate-400 mt-1 uppercase tracking-widest font-bold">
                {transmittals.length} saved transmittal
                {transmittals.length !== 1 ? "s" : ""}
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
          <div className="flex gap-2">
            <Input
              placeholder="Search by ID, project, or recipient..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              autoFocus
              className="rounded-xl flex-1"
            />
            <Popover open={isCalendarOpen} onOpenChange={setIsCalendarOpen}>
              <PopoverTrigger asChild>
                <Button
                  variant="outline"
                  className={`rounded-xl w-[160px] justify-start text-left text-xs font-normal gap-2 ${
                    !dateFilter ? "text-muted-foreground" : ""
                  }`}
                >
                  <CalendarDays className="w-3.5 h-3.5" />
                  {dateFilter
                    ? format(dateFilter, "MMM dd, yyyy")
                    : "Filter by date"}
                </Button>
              </PopoverTrigger>
              <PopoverContent align="end" className="w-auto">
                <Calendar
                  mode="single"
                  selected={dateFilter}
                  onSelect={(day) => {
                    setDateFilter(day);
                    setIsCalendarOpen(false);
                  }}
                  initialFocus
                />
              </PopoverContent>
            </Popover>
            {dateFilter && (
              <Button
                variant="ghost"
                size="icon"
                className="rounded-xl h-9 w-9 shrink-0"
                onClick={() => setDateFilter(undefined)}
              >
                <X className="w-3.5 h-3.5" />
              </Button>
            )}
          </div>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-2">
          {isLoading ? (
            <div className="flex items-center justify-center py-12 text-slate-400">
              <Loader2 className="w-5 h-5 animate-spin mr-2" />
              <span className="text-sm font-medium">Loading...</span>
            </div>
          ) : filtered.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-12 text-slate-400">
              <FolderOpen className="w-8 h-8 mb-2" />
              <p className="text-sm font-medium">
                {search
                  ? "No matching transmittals"
                  : "No saved transmittals yet"}
              </p>
            </div>
          ) : (
            filtered.map((t) => (
              <div
                key={t.id}
                role="button"
                tabIndex={0}
                onClick={() => {
                  onOpenTransmittal(t.id);
                  onClose();
                }}
                onKeyDown={(e) => {
                  if (e.target !== e.currentTarget) return;
                  if (e.key === "Enter" || e.key === " ") {
                    e.preventDefault();
                    onOpenTransmittal(t.id);
                    onClose();
                  }
                }}
                className="w-full flex items-center gap-4 p-4 rounded-2xl border border-slate-100 hover:border-slate-300 hover:bg-slate-50 transition-all text-left group focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-brand-500"
              >
                <div className="flex-shrink-0 w-10 h-10 rounded-xl bg-slate-100 flex items-center justify-center group-hover:bg-brand-50">
                  <FileText className="w-4 h-4 text-slate-400 group-hover:text-brand-600" />
                </div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-baseline gap-2">
                    <span className="font-mono text-[11px] font-bold text-slate-800">
                      {t.transmittalNumber || "—"}
                    </span>
                    <span className="text-[10px] text-slate-400">{t.date}</span>
                  </div>
                  <p className="text-xs text-slate-600 truncate">
                    {t.projectName}
                  </p>
                  <p className="text-[10px] text-slate-400 truncate">
                    {t.recipientName || "No recipient"} · {t.itemCount} item
                    {t.itemCount !== 1 ? "s" : ""}
                  </p>
                </div>
                <button
                  onClick={(e) => handleDelete(t.id, e)}
                  disabled={deletingId === t.id}
                  className="opacity-0 group-hover:opacity-100 p-2 rounded-xl text-slate-400 hover:text-red-500 hover:bg-red-50 transition-all"
                  title="Delete transmittal"
                >
                  {deletingId === t.id ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Trash2 className="w-3.5 h-3.5" />
                  )}
                </button>
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
};
