"use client";

import React from "react";
import { Button } from "@/components/ui/button";

interface DocxPreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
  html: string;
}

export const DocxPreviewModal: React.FC<DocxPreviewModalProps> = ({
  isOpen,
  onClose,
  html,
}) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[101] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xl animate-in fade-in duration-300">
      <div className="bg-slate-100 w-full max-w-5xl h-[92vh] rounded-[40px] shadow-[0_35px_60px_-15px_rgba(0,0,0,0.5)] flex flex-col animate-in zoom-in-95 duration-300 overflow-hidden border border-white/20">
        <div className="p-6 px-10 border-b border-slate-200 bg-white flex justify-between items-center shrink-0">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-brand-600 rounded-2xl flex items-center justify-center text-white shadow-lg shadow-brand-200">
              <svg className="w-6 h-6" fill="currentColor" viewBox="0 0 24 24">
                <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z" />
              </svg>
            </div>
            <div>
              <h3 className="text-xl font-black text-slate-800 font-display uppercase tracking-tight">
                Word Rendering
              </h3>
              <p className="text-[10px] text-slate-400 uppercase tracking-widest font-black">
                Live DOCX Layout Simulation
              </p>
            </div>
          </div>
          <Button
            onClick={onClose}
            variant="ghost"
            size="icon-lg"
            className="rounded-2xl"
          >
            ✕
          </Button>
        </div>

        <div className="flex-1 overflow-y-auto p-12 bg-slate-200/50 flex justify-center custom-scrollbar">
          <div className="bg-white w-full max-w-[8.5in] shadow-2xl p-16 min-h-full font-sans text-slate-900 preview-content border border-slate-300 rounded-sm">
            {html ? (
              <div
                className="animate-in fade-in duration-700"
                dangerouslySetInnerHTML={{ __html: html }}
              />
            ) : (
              <div className="h-full flex flex-col items-center justify-center text-slate-400 gap-4">
                <div className="w-12 h-12 border-4 border-brand-500 border-t-transparent rounded-full animate-spin"></div>
                <span className="text-sm font-bold uppercase tracking-widest">
                  Generating Live Preview...
                </span>
              </div>
            )}
          </div>
        </div>

        <div className="p-6 px-10 border-t border-slate-200 bg-white flex justify-between items-center shrink-0">
          <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">
            Standard A4/Letter Layout
          </p>
          <Button
            onClick={onClose}
            className="px-12 py-3 rounded-2xl font-bold shadow-xl"
          >
            Close View
          </Button>
        </div>
      </div>
      <style>{`
        .preview-content { font-family: 'Arial', sans-serif; font-size: 10pt; color: #1e293b; line-height: 1.4; }
        .preview-content h1 { font-size: 1.5rem; font-weight: bold; margin-bottom: 2rem; text-align: center; color: #000; text-transform: uppercase; }
        .preview-content table { width: 100% !important; border-collapse: collapse !important; margin-bottom: 1.5rem !important; table-layout: fixed; }
        .preview-content td { border: 1px solid #cbd5e1; padding: 6px 10px; font-size: 9pt; vertical-align: top; }
        .custom-scrollbar::-webkit-scrollbar { width: 8px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
      `}</style>
    </div>
  );
};
