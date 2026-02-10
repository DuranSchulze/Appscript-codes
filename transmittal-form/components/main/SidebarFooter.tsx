"use client";

import React, { useState } from "react";
import {
  Save,
  FileText,
  FileDown,
  FileSpreadsheet,
  Mail,
  Printer,
  LogOut,
  RotateCcw,
} from "lucide-react";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuContent,
  DropdownMenuGroup,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuLabel,
  DropdownMenuShortcut,
} from "@/components/ui/dropdown-menu";
import {
  AlertDialog,
  AlertDialogContent,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogAction,
  AlertDialogCancel,
} from "@/components/ui/alert-dialog";

export interface SidebarMenuBarProps {
  onSaveTransmittal: () => void;
  onExportPdf: () => void;
  onExportDocx: () => void;
  onExportCsv?: () => void;
  onSendEmail?: () => void;
  onPreviewDocx?: () => void;
  onSignOut: () => void;
  onResetWorkspace: () => void;
  isGeneratingPdf: boolean;
  isGeneratingDocx: boolean;
}

export const SidebarMenuBar: React.FC<SidebarMenuBarProps> = ({
  onSaveTransmittal,
  onExportPdf,
  onExportDocx,
  onExportCsv,
  onSendEmail,
  onPreviewDocx,
  onSignOut,
  onResetWorkspace,
  isGeneratingPdf,
  isGeneratingDocx,
}) => {
  const [showResetDialog, setShowResetDialog] = useState(false);

  return (
    <>
      <div className="flex items-center px-6 pt-1 pb-0 gap-1 bg-white/40">
        {/* File menu */}
        <DropdownMenu>
          <DropdownMenuTrigger className="px-3 py-1 rounded-md text-[11px] font-medium text-slate-500 hover:bg-slate-100 hover:text-slate-700 transition-colors outline-none cursor-default">
            File
          </DropdownMenuTrigger>
          <DropdownMenuContent
            side="bottom"
            sideOffset={4}
            align="start"
            className="min-w-[220px]"
          >
            <DropdownMenuGroup>
              <DropdownMenuItem onClick={onSaveTransmittal}>
                <Save className="mr-2 h-4 w-4" />
                Save to History
                <DropdownMenuShortcut>⌘S</DropdownMenuShortcut>
              </DropdownMenuItem>
            </DropdownMenuGroup>
            <DropdownMenuSeparator />
            <DropdownMenuGroup>
              <DropdownMenuLabel>Export</DropdownMenuLabel>
              <DropdownMenuItem
                onClick={onExportPdf}
                disabled={isGeneratingPdf}
              >
                <FileText className="mr-2 h-4 w-4" />
                {isGeneratingPdf ? "Generating PDF..." : "Export as PDF"}
              </DropdownMenuItem>
              <DropdownMenuItem
                onClick={onExportDocx}
                disabled={isGeneratingDocx}
              >
                <FileDown className="mr-2 h-4 w-4" />
                {isGeneratingDocx ? "Generating..." : "Export as Word"}
              </DropdownMenuItem>
              {onPreviewDocx && (
                <DropdownMenuItem onClick={onPreviewDocx}>
                  <Printer className="mr-2 h-4 w-4" />
                  Preview Word Document
                </DropdownMenuItem>
              )}
              {onExportCsv && (
                <DropdownMenuItem onClick={onExportCsv}>
                  <FileSpreadsheet className="mr-2 h-4 w-4" />
                  Export as CSV
                </DropdownMenuItem>
              )}
            </DropdownMenuGroup>
            {onSendEmail && (
              <>
                <DropdownMenuSeparator />
                <DropdownMenuGroup>
                  <DropdownMenuItem onClick={onSendEmail}>
                    <Mail className="mr-2 h-4 w-4" />
                    Email Recipient
                  </DropdownMenuItem>
                </DropdownMenuGroup>
              </>
            )}
            <DropdownMenuSeparator />
            <DropdownMenuGroup>
              <DropdownMenuItem onClick={() => setShowResetDialog(true)}>
                <RotateCcw className="mr-2 h-4 w-4" />
                Clear Workspace
              </DropdownMenuItem>
              <DropdownMenuItem onClick={onSignOut} variant="destructive">
                <LogOut className="mr-2 h-4 w-4" />
                Sign Out
              </DropdownMenuItem>
            </DropdownMenuGroup>
          </DropdownMenuContent>
        </DropdownMenu>

        <div className="ml-auto">
          <button
            onClick={onSignOut}
            className="flex items-center gap-1.5 px-3 py-1 rounded-md text-[11px] font-medium text-red-600 bg-red-50 hover:bg-red-100 transition-colors"
          >
            <LogOut className="w-3.5 h-3.5" />
            Sign Out
          </button>
        </div>
      </div>

      <AlertDialog open={showResetDialog} onOpenChange={setShowResetDialog}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Clear Workspace?</AlertDialogTitle>
            <AlertDialogDescription>
              This will reset the current form. Your saved History snapshots
              will not be affected.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel onClick={() => setShowResetDialog(false)}>
              Cancel
            </AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                setShowResetDialog(false);
                onResetWorkspace();
              }}
            >
              Reset Form
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </>
  );
};
