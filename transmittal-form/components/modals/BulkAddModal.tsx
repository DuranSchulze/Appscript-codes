"use client";

import React, { useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

interface BulkAddModalProps {
  isOpen: boolean;
  onClose: () => void;
  onImportDriveLink: (link: string) => Promise<void> | void;
}

export const BulkAddModal: React.FC<BulkAddModalProps> = ({
  isOpen,
  onClose,
  onImportDriveLink,
}) => {
  const [driveLink, setDriveLink] = useState("");
  const [isSubmittingDrive, setIsSubmittingDrive] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string>("");
  const [statusMsg, setStatusMsg] = useState<string>("");
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-md animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-2xl rounded-[32px] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300 border border-white">
        <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
          <div>
            <h3 className="text-2xl font-black text-slate-800 font-display">
              Bulk Import
            </h3>
            <p className="text-xs text-slate-400 mt-1 uppercase tracking-widest font-bold">
              Import from a Google Drive folder
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
        <div className="p-8">
          <div className="space-y-3">
            <Input
              placeholder="Paste Google Drive folder link..."
              value={driveLink}
              onChange={(e) => setDriveLink(e.target.value)}
              autoFocus
            />
            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">
              Imports all files in the folder
            </p>
          </div>

          {errorMsg ? (
            <div className="mt-4 rounded-2xl border border-red-200 bg-red-50 p-4 text-sm text-red-800">
              <div className="font-bold">Import failed</div>
              <div className="mt-1 break-words">{errorMsg}</div>
              <div className="mt-2 text-xs text-red-700">
                Make sure you pasted a folder link (not a file), and that you're
                signed in with Drive access.
              </div>
            </div>
          ) : null}

          {statusMsg ? (
            <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-700">
              {statusMsg}
            </div>
          ) : null}

          <div className="mt-8 flex gap-4">
            <Button
              onClick={onClose}
              variant="ghost"
              className="flex-1 py-4 rounded-2xl font-bold"
            >
              Cancel
            </Button>
            <Button
              onClick={async () => {
                if (!driveLink.trim()) return;
                try {
                  setIsSubmittingDrive(true);
                  setErrorMsg("");
                  setStatusMsg("Fetching files from Drive...");
                  await onImportDriveLink(driveLink);
                  setDriveLink("");
                  setStatusMsg("");
                  onClose();
                } catch (e: any) {
                  setStatusMsg("");
                  setErrorMsg(e?.message || "Import failed.");
                } finally {
                  setIsSubmittingDrive(false);
                }
              }}
              disabled={!driveLink.trim() || isSubmittingDrive}
              className="flex-[2] py-4 rounded-2xl font-bold shadow-xl"
            >
              Import From Drive
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
};
