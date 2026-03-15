import React, { useRef, useEffect, useState } from "react";
import {
  AppData,
  TransmittalItem,
  Signatories,
  ReceivedBy,
  FooterNotes,
} from "../types";
import { REMARKS_OPTIONS } from "../services/dropdown";
import { formatExportText } from "../services/exportText";

interface Props {
  data: AppData;
  onUpdateItem: (
    index: number,
    field: keyof TransmittalItem,
    value: string,
  ) => void;
  onAdjustItemQty: (index: number, delta: number) => void;
  onRemoveItem: (index: number) => void;
  onMoveItem: (index: number, direction: "up" | "down") => void;
  onReorderItems: (fromIndex: number, toIndex: number) => void;
  onAddItem: () => void;
  onBulkAdd: () => void;
  onUpdateSignatory?: (field: keyof Signatories, value: string) => void;
  onUpdateReceivedBy?: (field: keyof ReceivedBy, value: string) => void;
  onUpdateFooter?: (field: keyof FooterNotes, value: string) => void;
  onUpdateNotes?: (value: string) => void;
  isGeneratingPdf?: boolean;
  columnWidths: {
    qty: number;
    noOfItems: number;
    documentNumber: number;
    description: number;
    remarks: number;
  };
  onResizeDivider: (
    leftField: keyof Props["columnWidths"],
    deltaX: number,
  ) => void;
}

const AutoResizeTextArea = ({
  value,
  onChange,
  className = "",
  align = "left",
}: {
  value: string;
  onChange?: (val: string) => void;
  className?: string;
  align?: "left" | "center" | "right";
}) => {
  if (!onChange) {
    return (
      <span
        className={`block whitespace-pre-wrap break-words [overflow-wrap:anywhere] ${className}`}
        style={{ textAlign: align }}
      >
        {value}
      </span>
    );
  }
  return (
    <div className="grid w-full" style={{ gridTemplateColumns: "1fr" }}>
      <textarea
        rows={1}
        wrap="soft"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className={`w-full min-w-0 bg-transparent outline-none font-medium placeholder-slate-400 resize-none leading-tight overflow-hidden ${className}`}
        style={{
          textAlign: align,
          whiteSpace: "pre-wrap",
          overflowWrap: "anywhere",
          wordBreak: "break-word",
          gridArea: "1 / 1 / 2 / 2",
          height: "auto",
          minHeight: "1.25em",
          fieldSizing: "content" as React.CSSProperties["fieldSizing"],
        }}
      />
    </div>
  );
};

const getExportValue = (value: string, isExporting: boolean) =>
  isExporting ? formatExportText(value) : value;

const ComboBox = ({
  value,
  onChange,
  options,
  className = "",
}: {
  value: string;
  onChange?: (val: string) => void;
  options: string[];
  className?: string;
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (
        containerRef.current &&
        !containerRef.current.contains(event.target as Node)
      ) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  if (!onChange) {
    return (
      <span
        className={`block whitespace-pre-wrap break-all [overflow-wrap:anywhere] ${className}`}
      >
        {value}
      </span>
    );
  }

  return (
    <div className="relative w-full" ref={containerRef}>
      <AutoResizeTextArea
        value={value}
        onChange={onChange}
        className={`${className} pr-6`}
      />
      <button
        className="absolute right-0 top-0 text-slate-400 hover:text-slate-600 p-1"
        onClick={() => setIsOpen(!isOpen)}
        tabIndex={-1}
      >
        <svg
          className="w-3 h-3"
          fill="none"
          stroke="currentColor"
          viewBox="0 0 24 24"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            strokeWidth="2"
            d="M19 9l-7 7-7-7"
          />
        </svg>
      </button>
      {isOpen && (
        <div className="absolute left-0 top-full mt-1 min-w-[200px] w-max bg-white border border-slate-200 rounded-lg shadow-lg z-[100] max-h-40 overflow-y-auto">
          {options.map((option) => (
            <div
              key={option}
              className="px-3 py-2 text-xs text-slate-700 hover:bg-slate-50 cursor-pointer"
              onClick={() => {
                onChange(option);
                setIsOpen(false);
              }}
            >
              {option}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

const ResizableHeader = ({
  width,
  label,
  onResizeDelta,
  className,
}: {
  width?: number;
  label: string;
  onResizeDelta?: (deltaX: number) => void;
  className?: string;
}) => {
  const handleMouseDown = (e: React.MouseEvent) => {
    if (!onResizeDelta || !width) return;
    e.preventDefault();
    let lastX = e.pageX;
    const handleMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.pageX - lastX;
      lastX = moveEvent.pageX;
      onResizeDelta(deltaX);
    };
    const handleMouseUp = () => {
      document.removeEventListener("mousemove", handleMouseMove);
      document.removeEventListener("mouseup", handleMouseUp);
    };
    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", handleMouseUp);
  };
  return (
    <th
      className={`relative ${className}`}
      style={width ? { width: `${width}px` } : {}}
    >
      {label}
      {onResizeDelta && (
        <div
          className="absolute right-0 top-0 bottom-0 w-2.5 cursor-col-resize z-10 group"
          onMouseDown={handleMouseDown}
        >
          <div className="absolute right-0 top-1 bottom-1 w-px bg-slate-300 group-hover:bg-brand-500 group-active:bg-brand-600 transition-colors" />
        </div>
      )}
    </th>
  );
};

export const TransmittalTemplate: React.FC<Props> = ({
  data,
  onUpdateItem,
  onAdjustItemQty,
  onRemoveItem,
  onMoveItem,
  onReorderItems,
  onAddItem,
  onBulkAdd,
  onUpdateSignatory,
  onUpdateReceivedBy,
  onUpdateFooter,
  onUpdateNotes,
  isGeneratingPdf = false,
  columnWidths,
  onResizeDivider,
}) => {
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);
  const handleDragStart = (e: React.DragEvent, index: number) => {
    const target = e.target as HTMLElement;
    if (
      target.tagName === "TEXTAREA" ||
      target.tagName === "BUTTON" ||
      target.tagName === "INPUT"
    ) {
      e.preventDefault();
      return;
    }
    setDraggedIndex(index);
    // Add visual feedback to the dragged row via timeout to allow drag initiation
    setTimeout(() => {
      (e.target as HTMLElement).classList.add("drag-row-active");
    }, 0);
  };
  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    if (draggedIndex === null) return;
    if (dragOverIndex !== index) setDragOverIndex(index);
  };
  const handleDragEnd = (e: React.DragEvent) => {
    setDraggedIndex(null);
    setDragOverIndex(null);
    (e.target as HTMLElement).classList.remove("drag-row-active");
  };
  const handleDrop = (e: React.DragEvent, targetIndex: number) => {
    e.preventDefault();
    if (draggedIndex !== null && draggedIndex !== targetIndex)
      onReorderItems(draggedIndex, targetIndex);
    setDraggedIndex(null);
    setDragOverIndex(null);
  };

  const containerClass = isGeneratingPdf
    ? "bg-white text-slate-800 w-full font-sans relative"
    : "bg-white text-slate-800 w-full max-w-[8.5in] mx-auto shadow-2xl font-sans relative min-h-[11in]";

  const cellClass =
    "border border-slate-300 px-2 py-2 align-top text-sm break-words whitespace-normal [overflow-wrap:anywhere] relative transition-all duration-200";
  const headerCellClass =
    "border border-slate-300 bg-slate-50 px-1 py-2 font-bold text-center text-sm uppercase text-slate-700 align-middle select-none";
  const exactWidthStyle = (width: number) => ({
    width: `${width}px`,
    minWidth: `${width}px`,
    maxWidth: `${width}px`,
  });

  const leftData = [
    { label: "To:", value: data.recipient.to },
    { label: "Company:", value: data.recipient.company },
    { label: "Attention:", value: data.recipient.attention },
    { label: "Address:", value: data.recipient.address },
    { label: "Contact No:", value: data.recipient.contactNumber },
    { label: "Email:", value: data.recipient.email },
  ];

  const rightData = [
    { label: "Project Name:", value: data.project.projectName },
    { label: "Project No:", value: data.project.projectNumber },
    { label: "Engagement Ref #:", value: data.project.engagementRef },
    { label: "Purpose:", value: data.project.purpose },
    { label: "Transmittal No:", value: data.project.transmittalNumber },
    { label: "Department:", value: data.project.department },
    { label: "Date:", value: data.project.date },
    { label: "Time Generated:", value: data.project.timeGenerated },
  ];

  const maxRows = Math.max(leftData.length, rightData.length);
  const metaRows = [];
  for (let i = 0; i < maxRows; i++) {
    metaRows.push({
      left: leftData[i] || { label: "", value: "" },
      right: rightData[i] || { label: "", value: "" },
    });
  }

  return (
    <div className={containerClass}>
      <style>{`
                @media print {
                    thead { display: table-header-group; }
                    tfoot { display: table-footer-group; }
                    body { -webkit-print-color-adjust: exact; }
                    tr, td { page-break-inside: auto; }
                }
                .avoid-break { page-break-inside: avoid !important; }
                textarea { height: auto !important; }
                
                /* Drag & Drop Visual Enhancements */
                .row-item { transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1); position: relative; }
                
                .drag-row-active { 
                    opacity: 0.5; 
                    background-color: #ffffff !important;
                    transform: translateY(-2px) scale(0.985);
                    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -4px rgba(0, 0, 0, 0.1) !important;
                    z-index: 50;
                    filter: grayscale(0.5);
                }
                
                .drop-indicator { 
                    outline: 2px dashed #2563eb !important; 
                    outline-offset: -3px;
                    background-color: #eff6ff !important;
                    z-index: 10;
                }

                .grab-handle { cursor: grab; }
                .grab-handle:active { cursor: grabbing; }
            `}</style>

      <table className="w-full h-full border-collapse">
        <thead className="h-[1.5in]">
          <tr>
            <td colSpan={100} className="align-top p-0 border-none">
              <div className="flex justify-between items-start pt-6 px-10 h-[1.5in] box-border overflow-hidden w-full">
                <div className="w-1/3">
                  {data.sender.logoBase64 ? (
                    <img
                      src={data.sender.logoBase64}
                      alt="Logo"
                      className="h-20 object-contain object-left"
                    />
                  ) : (
                    <div className="h-14 w-40 flex items-center justify-center border border-dashed border-slate-300 text-xs text-slate-400 font-black uppercase tracking-widest">
                      Logo Placeholder
                    </div>
                  )}
                </div>
                <div className="w-2/3 text-right text-[9px] leading-tight text-slate-600 break-words pl-4 mt-1 font-medium">
                  <p>Telephone: {data.sender.telephone}</p>
                  <p>Mobile: {data.sender.mobile}</p>
                  <p>Email: {data.sender.email}</p>
                  <p>{data.sender.addressLine1}</p>
                  <p>{data.sender.addressLine2}</p>
                  <p className="font-bold text-brand-600">
                    {data.sender.website}
                  </p>
                </div>
              </div>
            </td>
          </tr>
        </thead>

        <tfoot className="h-[0.5in]">
          <tr>
            <td colSpan={100} className="align-bottom p-0 border-none">
              <div className="h-[0.5in] w-full"></div>
            </td>
          </tr>
        </tfoot>

        <tbody>
          <tr>
            <td
              colSpan={100}
              className={`align-top pt-2 pb-8 border-none ${isGeneratingPdf ? "px-0" : "px-10"}`}
            >
              <h1 className="text-center font-serif text-2xl font-bold mb-6 tracking-wide text-slate-900 uppercase">
                Transmittal Form
              </h1>

              <div className="border border-slate-300 text-sm mb-8">
                {metaRows.map((row, i) => (
                  <div
                    key={i}
                    className="flex border-b border-slate-300 last:border-b-0 min-h-[30px]"
                  >
                    <div className="w-1/2 flex border-r border-slate-300">
                      <div className="w-[30%] bg-slate-50 p-1.5 font-bold border-r border-slate-300 text-[10px] uppercase text-slate-600 flex items-start pt-2">
                        {row.left.label}
                      </div>
                      <div className="w-[70%] p-1.5 font-semibold text-slate-800 uppercase break-words whitespace-pre-wrap flex items-start pt-2 leading-tight">
                        {row.left.value}
                      </div>
                    </div>
                    <div className="w-1/2 flex">
                      <div className="w-[35%] bg-slate-50 p-1.5 font-bold border-r border-slate-300 text-[10px] uppercase text-slate-600 flex items-start pt-2">
                        {row.right.label}
                      </div>
                      <div className="w-[65%] p-1.5 text-slate-700 break-words whitespace-pre-wrap flex items-start pt-2 leading-tight">
                        {row.right.value}
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              <div className="mb-8">
                <table className="w-full border-collapse border border-slate-300 text-sm table-fixed">
                  <thead>
                    <tr>
                      <ResizableHeader
                        label="No. of Items"
                        width={columnWidths.noOfItems}
                        onResizeDelta={
                          !isGeneratingPdf
                            ? (deltaX) => onResizeDivider("noOfItems", deltaX)
                            : undefined
                        }
                        className={headerCellClass}
                      />
                      <ResizableHeader
                        label="QTY"
                        width={columnWidths.qty}
                        onResizeDelta={
                          !isGeneratingPdf
                            ? (deltaX) => onResizeDivider("qty", deltaX)
                            : undefined
                        }
                        className={headerCellClass}
                      />
                      <ResizableHeader
                        label="Document # / Ref #"
                        width={columnWidths.documentNumber}
                        onResizeDelta={
                          !isGeneratingPdf
                            ? (deltaX) =>
                                onResizeDivider("documentNumber", deltaX)
                            : undefined
                        }
                        className={headerCellClass}
                      />
                      <ResizableHeader
                        label="Description"
                        width={columnWidths.description}
                        onResizeDelta={
                          !isGeneratingPdf
                            ? (deltaX) => onResizeDivider("description", deltaX)
                            : undefined
                        }
                        className={headerCellClass}
                      />
                      <ResizableHeader
                        label="Remarks"
                        width={columnWidths.remarks}
                        className={headerCellClass}
                      />
                      {!isGeneratingPdf && (
                        <th className="w-12 print:hidden bg-slate-50 border border-slate-300"></th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {data.items.length === 0 ? (
                      <tr>
                        <td
                          colSpan={isGeneratingPdf ? 5 : 6}
                          className="text-center py-10 text-slate-400 italic border border-slate-300 bg-slate-50/30"
                        >
                          No items listed. Use sidebar to import or add manual
                          rows.
                        </td>
                      </tr>
                    ) : (
                      data.items.map((item, index) => (
                        <tr
                          key={item.id}
                          draggable={!isGeneratingPdf}
                          onDragStart={(e) => handleDragStart(e, index)}
                          onDragOver={(e) => handleDragOver(e, index)}
                          onDragEnd={handleDragEnd}
                          onDrop={(e) => handleDrop(e, index)}
                          className={`row-item group avoid-break ${draggedIndex === index ? "drag-row-active" : "hover:bg-slate-50/80"} ${dragOverIndex === index ? "drop-indicator" : ""}`}
                        >
                          <td
                            className={cellClass}
                            style={exactWidthStyle(columnWidths.noOfItems)}
                          >
                            <AutoResizeTextArea
                              value={getExportValue(
                                item.noOfItems,
                                isGeneratingPdf,
                              )}
                              className="text-center text-slate-800"
                              align="center"
                            />
                          </td>
                          <td
                            className={cellClass}
                            style={exactWidthStyle(columnWidths.qty)}
                          >
                            {isGeneratingPdf ? (
                              <AutoResizeTextArea
                                value={getExportValue(
                                  item.qty,
                                  isGeneratingPdf,
                                )}
                                className="text-center text-slate-800"
                                align="center"
                              />
                            ) : (
                              <div className="flex items-center gap-1">
                                <button
                                  type="button"
                                  onClick={() => onAdjustItemQty(index, -1)}
                                  className="w-6 h-6 rounded-full bg-slate-700 text-white text-xs font-black flex items-center justify-center shadow-sm hover:bg-brand-600 transition-colors"
                                  aria-label="Decrease quantity"
                                >
                                  -
                                </button>
                                <AutoResizeTextArea
                                  value={item.qty}
                                  onChange={(v) =>
                                    onUpdateItem(index, "qty", v)
                                  }
                                  className="text-center text-slate-800"
                                  align="center"
                                />
                                <button
                                  type="button"
                                  onClick={() => onAdjustItemQty(index, 1)}
                                  className="w-6 h-6 rounded-full bg-slate-700 text-white text-xs font-black flex items-center justify-center shadow-sm hover:bg-brand-600 transition-colors"
                                  aria-label="Increase quantity"
                                >
                                  +
                                </button>
                              </div>
                            )}
                          </td>
                          <td
                            className={cellClass}
                            style={exactWidthStyle(columnWidths.documentNumber)}
                          >
                            <AutoResizeTextArea
                              value={getExportValue(
                                item.documentNumber,
                                isGeneratingPdf,
                              )}
                              onChange={
                                isGeneratingPdf
                                  ? undefined
                                  : (v) =>
                                      onUpdateItem(index, "documentNumber", v)
                              }
                              className="text-slate-800"
                            />
                          </td>
                          <td
                            className={cellClass}
                            style={exactWidthStyle(columnWidths.description)}
                          >
                            <AutoResizeTextArea
                              value={getExportValue(
                                item.description,
                                isGeneratingPdf,
                              )}
                              onChange={
                                isGeneratingPdf
                                  ? undefined
                                  : (v) => onUpdateItem(index, "description", v)
                              }
                              className="text-slate-800"
                            />
                          </td>
                          <td
                            className={cellClass}
                            style={exactWidthStyle(columnWidths.remarks)}
                          >
                            <ComboBox
                              value={getExportValue(
                                item.remarks,
                                isGeneratingPdf,
                              )}
                              onChange={
                                isGeneratingPdf
                                  ? undefined
                                  : (v) => onUpdateItem(index, "remarks", v)
                              }
                              options={REMARKS_OPTIONS}
                              className="text-slate-800"
                            />
                          </td>
                          {!isGeneratingPdf && (
                            <td className="border border-slate-300 p-1 text-center print:hidden align-middle">
                              <div className="flex flex-col gap-1.5 items-center justify-center">
                                <div className="grab-handle text-slate-300 hover:text-brand-600 transition-colors p-1">
                                  <svg
                                    viewBox="0 0 24 24"
                                    width="20"
                                    height="20"
                                    stroke="currentColor"
                                    strokeWidth="2.5"
                                    fill="none"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                  >
                                    <circle cx="9" cy="12" r="1.5"></circle>
                                    <circle cx="9" cy="5" r="1.5"></circle>
                                    <circle cx="9" cy="19" r="1.5"></circle>
                                    <circle cx="15" cy="12" r="1.5"></circle>
                                    <circle cx="15" cy="5" r="1.5"></circle>
                                    <circle cx="15" cy="19" r="1.5"></circle>
                                  </svg>
                                </div>
                                <button
                                  onClick={() => onRemoveItem(index)}
                                  className="text-[10px] text-slate-300 hover:text-red-500 transition-colors px-1 font-bold"
                                >
                                  ✕
                                </button>
                              </div>
                            </td>
                          )}
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              {!isGeneratingPdf && (
                <div className="w-full flex justify-center gap-4 mb-6">
                  <button
                    onClick={onAddItem}
                    className="group flex items-center gap-2 px-4 py-2 bg-white border border-slate-300 rounded-full shadow-sm hover:border-brand-400 hover:text-brand-600 hover:shadow-md transition-all text-xs font-bold text-slate-500 uppercase tracking-wide"
                  >
                    <span className="w-5 h-5 flex items-center justify-center bg-slate-100 rounded-full group-hover:bg-brand-100 transition-colors text-lg leading-none pb-0.5">
                      +
                    </span>
                    Add Row
                  </button>
                  <button
                    onClick={onBulkAdd}
                    className="group flex items-center gap-2 px-4 py-2 bg-slate-800 text-white rounded-full shadow-sm hover:bg-slate-700 hover:shadow-md transition-all text-xs font-bold uppercase tracking-wide"
                  >
                    📋 Bulk Import
                  </button>
                </div>
              )}

              <div className="mt-4">
                <div className="pt-6 border-t border-slate-200">
                  <div className="grid grid-cols-3 gap-12 mb-8 avoid-break">
                    <div>
                      <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-6">
                        Prepared by:
                      </p>
                      <div className="border-b border-slate-800 pb-1 mb-1">
                        <AutoResizeTextArea
                          value={data.signatories.preparedBy}
                          onChange={
                            isGeneratingPdf
                              ? undefined
                              : (v) => onUpdateSignatory?.("preparedBy", v)
                          }
                          className="font-bold text-sm uppercase text-slate-800"
                        />
                      </div>
                      <AutoResizeTextArea
                        value={data.signatories.preparedByRole}
                        onChange={
                          isGeneratingPdf
                            ? undefined
                            : (v) => onUpdateSignatory?.("preparedByRole", v)
                        }
                        className="text-[9px] text-slate-400 uppercase tracking-wider font-bold"
                      />
                    </div>
                    <div>
                      <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-6">
                        Noted by:
                      </p>
                      <div className="border-b border-slate-800 pb-1 mb-1">
                        <AutoResizeTextArea
                          value={data.signatories.notedBy}
                          onChange={
                            isGeneratingPdf
                              ? undefined
                              : (v) => onUpdateSignatory?.("notedBy", v)
                          }
                          className="font-bold text-sm uppercase text-slate-800"
                        />
                      </div>
                      <AutoResizeTextArea
                        value={data.signatories.notedByRole}
                        onChange={
                          isGeneratingPdf
                            ? undefined
                            : (v) => onUpdateSignatory?.("notedByRole", v)
                        }
                        className="text-[9px] text-slate-400 uppercase tracking-wider font-bold"
                      />
                    </div>
                    <div>
                      <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-6">
                        Time Released:
                      </p>
                      <div className="border-b border-slate-800 h-[29px] w-full flex items-end pb-1">
                        <AutoResizeTextArea
                          value={data.signatories.timeReleased}
                          onChange={
                            isGeneratingPdf
                              ? undefined
                              : (v) => onUpdateSignatory?.("timeReleased", v)
                          }
                          className="text-sm font-bold text-slate-800"
                        />
                      </div>
                    </div>
                  </div>

                  <div className="bg-slate-50/50 p-5 rounded-lg border border-slate-200 mb-6 avoid-break">
                    <div className="flex items-center gap-6 text-[10px] font-bold text-slate-700 flex-wrap">
                      <span className="text-slate-900 uppercase tracking-widest">
                        Transmitted via:
                      </span>
                      <div className="flex items-center gap-2">
                        <div className="w-4 h-4 border border-slate-400 flex items-center justify-center bg-white rounded-[2px] relative">
                          {data.transmissionMethod.personalDelivery && (
                            <div className="absolute w-2.5 h-2.5 bg-slate-800 rounded-[1px]"></div>
                          )}
                        </div>
                        <span>Hand Delivered</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <div className="w-4 h-4 border border-slate-400 flex items-center justify-center bg-white rounded-[2px] relative">
                          {data.transmissionMethod.pickUp && (
                            <div className="absolute w-2.5 h-2.5 bg-slate-800 rounded-[1px]"></div>
                          )}
                        </div>
                        <span>Pick-up</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <div className="w-4 h-4 border border-slate-400 flex items-center justify-center bg-white rounded-[2px] relative">
                          {data.transmissionMethod.grabLalamove && (
                            <div className="absolute w-2.5 h-2.5 bg-slate-800 rounded-[1px]"></div>
                          )}
                        </div>
                        <span>Courier App</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <div className="w-4 h-4 border border-slate-400 flex items-center justify-center bg-white rounded-[2px] relative">
                          {data.transmissionMethod.registeredMail && (
                            <div className="absolute w-2.5 h-2.5 bg-slate-800 rounded-[1px]"></div>
                          )}
                        </div>
                        <span>Registered Mail</span>
                      </div>
                    </div>
                  </div>

                  {data.notes && (
                    <div className="mb-6 avoid-break">
                      <p className="text-[10px] font-bold text-slate-700 uppercase tracking-widest mb-1">
                        Notes / Instructions:
                      </p>
                      <div className="border border-slate-300 p-3 rounded-lg bg-white min-h-[40px]">
                        <AutoResizeTextArea
                          value={data.notes}
                          onChange={
                            isGeneratingPdf
                              ? undefined
                              : (v) => onUpdateNotes?.(v)
                          }
                          className="text-sm text-slate-800"
                        />
                      </div>
                    </div>
                  )}

                  <div className="mb-3">
                    <AutoResizeTextArea
                      value={data.footerNotes.acknowledgement}
                      onChange={
                        isGeneratingPdf
                          ? undefined
                          : (v) => onUpdateFooter?.("acknowledgement", v)
                      }
                      className="text-[10px] italic text-slate-500 w-full"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-x-12 gap-y-6 text-[10px] mb-8 bg-slate-50/50 p-6 rounded-lg border border-slate-200 avoid-break">
                    <div className="flex items-end gap-3">
                      <span className="w-24 shrink-0 font-bold text-slate-600 uppercase">
                        Received by:
                      </span>
                      <div className="border-b border-slate-300 w-full flex items-end">
                        <AutoResizeTextArea
                          value={data.receivedBy.name}
                          onChange={
                            isGeneratingPdf
                              ? undefined
                              : (v) => onUpdateReceivedBy?.("name", v)
                          }
                          className="font-bold text-slate-800 pb-0.5"
                        />
                      </div>
                    </div>
                    <div className="flex items-end gap-3">
                      <span className="w-24 shrink-0 font-bold text-slate-600 uppercase">
                        Date Received:
                      </span>
                      <div className="border-b border-slate-300 w-full flex items-end">
                        <AutoResizeTextArea
                          value={data.receivedBy.date}
                          onChange={
                            isGeneratingPdf
                              ? undefined
                              : (v) => onUpdateReceivedBy?.("date", v)
                          }
                          className="font-bold text-slate-800 pb-0.5"
                        />
                      </div>
                    </div>
                    <div className="flex items-end gap-3">
                      <span className="w-24 shrink-0 font-bold text-slate-600 uppercase">
                        Time Received:
                      </span>
                      <div className="border-b border-slate-300 w-full flex items-end">
                        <AutoResizeTextArea
                          value={data.receivedBy.time}
                          onChange={
                            isGeneratingPdf
                              ? undefined
                              : (v) => onUpdateReceivedBy?.("time", v)
                          }
                          className="font-bold text-slate-800 pb-0.5"
                        />
                      </div>
                    </div>
                    <div className="flex items-end gap-3">
                      <span className="w-24 shrink-0 font-bold text-slate-600 uppercase">
                        Remarks:
                      </span>
                      <div className="border-b border-slate-300 w-full flex items-end">
                        <AutoResizeTextArea
                          value={data.receivedBy.remarks}
                          onChange={
                            isGeneratingPdf
                              ? undefined
                              : (v) => onUpdateReceivedBy?.("remarks", v)
                          }
                          className="font-bold text-slate-800 pb-0.5"
                        />
                      </div>
                    </div>
                  </div>

                  <div className="flex items-center justify-center text-[8px] text-slate-400 italic w-full avoid-break">
                    <div className="flex-1 text-center">
                      <AutoResizeTextArea
                        value={data.footerNotes.disclaimer}
                        onChange={
                          isGeneratingPdf
                            ? undefined
                            : (v) => onUpdateFooter?.("disclaimer", v)
                        }
                        className="w-full text-center"
                        align="center"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  );
};
