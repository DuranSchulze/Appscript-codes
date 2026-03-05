"use client";

import React from "react";
import { ZoomIn, ZoomOut, RotateCcw } from "lucide-react";

const ZOOM_STEPS = [25, 50, 75, 100, 125, 150, 200];

interface PreviewToolbarProps {
  zoomPercent: number;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onZoomReset: () => void;
  onResetColumnWidths: () => void;
  onZoomSet: (percent: number) => void;
  transmittalNumber?: string;
  showSaveNotice?: boolean;
}

export const PreviewToolbar: React.FC<PreviewToolbarProps> = ({
  zoomPercent,
  onZoomIn,
  onZoomOut,
  onZoomReset,
  onResetColumnWidths,
  onZoomSet,
  transmittalNumber,
  showSaveNotice = false,
}) => {
  const displayTransmittalNumber = transmittalNumber?.trim() || "Draft";

  return (
    <div className="sticky top-0 z-10 flex flex-col items-center gap-1.5 py-2 px-4">
      <div className="flex items-center justify-center gap-2">
        <div className="flex items-center gap-2 bg-white/90 backdrop-blur-md border border-slate-200/80 rounded-xl shadow-sm px-3 py-1">
          <span className="text-[9px] font-black uppercase tracking-wider text-slate-500">
            Transmittal ID
          </span>
          <span className="font-mono text-[11px] font-bold text-slate-700">
            {displayTransmittalNumber}
          </span>
        </div>
        <div className="flex items-center gap-1 bg-white/90 backdrop-blur-md border border-slate-200/80 rounded-xl shadow-sm px-2 py-1">
          <button
            onClick={onZoomOut}
            disabled={zoomPercent <= ZOOM_STEPS[0]}
            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100 hover:text-slate-700 disabled:opacity-30 disabled:hover:bg-transparent transition-colors"
            title="Zoom out"
          >
            <ZoomOut className="w-4 h-4" />
          </button>

          <select
            value={zoomPercent}
            onChange={(e) => onZoomSet(Number(e.target.value))}
            className="appearance-none bg-transparent text-center text-[11px] font-semibold text-slate-600 px-2 py-1 rounded-md hover:bg-slate-100 cursor-pointer outline-none min-w-[56px] transition-colors"
          >
            {ZOOM_STEPS.map((step) => (
              <option key={step} value={step}>
                {step}%
              </option>
            ))}
            {!ZOOM_STEPS.includes(zoomPercent) && (
              <option value={zoomPercent}>{zoomPercent}%</option>
            )}
          </select>

          <button
            onClick={onZoomIn}
            disabled={zoomPercent >= ZOOM_STEPS[ZOOM_STEPS.length - 1]}
            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100 hover:text-slate-700 disabled:opacity-30 disabled:hover:bg-transparent transition-colors"
            title="Zoom in"
          >
            <ZoomIn className="w-4 h-4" />
          </button>

          <div className="w-px h-4 bg-slate-200 mx-1" />

          <button
            onClick={onZoomReset}
            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100 hover:text-slate-700 transition-colors"
            title="Fit to width"
          >
            <RotateCcw className="w-3.5 h-3.5" />
          </button>
        </div>
        <button
          onClick={onResetColumnWidths}
          className="inline-flex items-center gap-1.5 bg-white/90 backdrop-blur-md border border-slate-200/80 rounded-xl shadow-sm px-3 py-2 text-[10px] font-bold uppercase tracking-wide text-slate-600 hover:bg-white hover:text-slate-800 transition-colors"
          title="Reset column widths"
        >
          Reset Widths
        </button>
      </div>

      {showSaveNotice ? (
        <div className="inline-flex flex-wrap items-center justify-center gap-1.5 bg-amber-50/95 border border-amber-200 rounded-lg shadow-sm px-3 py-1 text-[10px] font-medium text-amber-900">
          <span>
            Click <span className="font-bold">Save</span> button to keep
            transmittals.
          </span>
          <span className="text-amber-700/80">Shortcut:</span>
          <span className="inline-flex items-center gap-0.5">
            <kbd className="rounded border border-amber-300 bg-white px-1 py-0.5 font-semibold text-[9px] leading-none text-amber-900">
              Ctrl
            </kbd>
            <span>+</span>
            <kbd className="rounded border border-amber-300 bg-white px-1 py-0.5 font-semibold text-[9px] leading-none text-amber-900">
              S
            </kbd>
          </span>
          <span className="text-amber-700/70">/</span>
          <span className="inline-flex items-center gap-0.5">
            <kbd className="rounded border border-amber-300 bg-white px-1 py-0.5 font-semibold text-[9px] leading-none text-amber-900">
              ⌘
            </kbd>
            <span>+</span>
            <kbd className="rounded border border-amber-300 bg-white px-1 py-0.5 font-semibold text-[9px] leading-none text-amber-900">
              S
            </kbd>
          </span>
        </div>
      ) : null}
    </div>
  );
};

export { ZOOM_STEPS };
