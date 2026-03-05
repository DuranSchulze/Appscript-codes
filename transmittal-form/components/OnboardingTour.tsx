"use client";

import React, { useEffect, useState, useCallback } from "react";
import {
  X,
  ChevronLeft,
  ChevronRight,
  HelpCircle,
  AlertTriangle,
} from "lucide-react";

export interface TourStepItem {
  id: string;
  label: string;
  shortcut?: string;
  description: string;
  important?: boolean;
  importantNote?: string;
}

export interface TourStep {
  id: string;
  targetSelector: string | null;
  title: string;
  body: string;
  items?: TourStepItem[];
}

interface TargetRect {
  top: number;
  left: number;
  width: number;
  height: number;
}

interface PopoverPosition {
  top: number;
  left: number;
  placement: "top" | "bottom" | "left" | "right" | "center";
}

interface OnboardingTourProps {
  steps: TourStep[];
  isOpen: boolean;
  currentStep: number;
  onNext: () => void;
  onBack: () => void;
  onSkip: () => void;
  onFinish: () => void;
}

const PADDING = 10;
const POPOVER_WIDTH = 320;
const POPOVER_WIDTH_ITEMS = 400;
const POPOVER_APPROX_HEIGHT = 200;

function getTargetRect(selector: string | null): TargetRect | null {
  if (!selector) return null;
  try {
    const el = document.querySelector(selector);
    if (!el) return null;
    const rect = el.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return null;
    return {
      top: rect.top,
      left: rect.left,
      width: rect.width,
      height: rect.height,
    };
  } catch {
    return null;
  }
}

function computePopoverPosition(
  target: TargetRect | null,
  vpWidth: number,
  vpHeight: number,
): PopoverPosition {
  if (!target) {
    return {
      top: vpHeight / 2 - POPOVER_APPROX_HEIGHT / 2,
      left: vpWidth / 2 - POPOVER_WIDTH / 2,
      placement: "center",
    };
  }

  const spaceBelow = vpHeight - (target.top + target.height);
  const spaceAbove = target.top;
  const spaceRight = vpWidth - (target.left + target.width);
  const spaceLeft = target.left;

  let top: number;
  let left: number;
  let placement: PopoverPosition["placement"];

  if (spaceBelow >= POPOVER_APPROX_HEIGHT + PADDING) {
    top = target.top + target.height + PADDING;
    left = Math.min(
      Math.max(target.left, PADDING),
      vpWidth - POPOVER_WIDTH - PADDING,
    );
    placement = "bottom";
  } else if (spaceAbove >= POPOVER_APPROX_HEIGHT + PADDING) {
    top = target.top - POPOVER_APPROX_HEIGHT - PADDING;
    left = Math.min(
      Math.max(target.left, PADDING),
      vpWidth - POPOVER_WIDTH - PADDING,
    );
    placement = "top";
  } else if (spaceRight >= POPOVER_WIDTH + PADDING) {
    top = Math.min(
      Math.max(target.top, PADDING),
      vpHeight - POPOVER_APPROX_HEIGHT - PADDING,
    );
    left = target.left + target.width + PADDING;
    placement = "right";
  } else if (spaceLeft >= POPOVER_WIDTH + PADDING) {
    top = Math.min(
      Math.max(target.top, PADDING),
      vpHeight - POPOVER_APPROX_HEIGHT - PADDING,
    );
    left = target.left - POPOVER_WIDTH - PADDING;
    placement = "left";
  } else {
    top = vpHeight / 2 - POPOVER_APPROX_HEIGHT / 2;
    left = vpWidth / 2 - POPOVER_WIDTH / 2;
    placement = "center";
  }

  return { top, left, placement };
}

export const OnboardingTour: React.FC<OnboardingTourProps> = ({
  steps,
  isOpen,
  currentStep,
  onNext,
  onBack,
  onSkip,
  onFinish,
}) => {
  const [targetRect, setTargetRect] = useState<TargetRect | null>(null);
  const [popoverPos, setPopoverPos] = useState<PopoverPosition>({
    top: 0,
    left: 0,
    placement: "center",
  });
  const [vpSize, setVpSize] = useState({ w: 0, h: 0 });

  const step = steps[currentStep];
  const isLastStep = currentStep === steps.length - 1;
  const hasItems = Boolean(step?.items?.length);
  const popoverWidth = hasItems ? POPOVER_WIDTH_ITEMS : POPOVER_WIDTH;
  const isCentered = !step?.targetSelector || !targetRect;

  const recalculate = useCallback(() => {
    if (!step) return;
    const w = window.innerWidth;
    const h = window.innerHeight;
    setVpSize({ w, h });
    const rect = getTargetRect(step.targetSelector);
    setTargetRect(rect);
    setPopoverPos(computePopoverPosition(rect, w, h));
  }, [step]);

  useEffect(() => {
    if (!isOpen || !step) return;

    const scrollAndRecalc = () => {
      if (step.targetSelector) {
        try {
          const el = document.querySelector(step.targetSelector);
          if (el) {
            el.scrollIntoView({ behavior: "smooth", block: "nearest" });
          }
        } catch {}
      }
      setTimeout(recalculate, 150);
    };

    scrollAndRecalc();

    window.addEventListener("resize", recalculate);
    return () => window.removeEventListener("resize", recalculate);
  }, [isOpen, currentStep, step, recalculate]);

  useEffect(() => {
    if (!isOpen) return;
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onSkip();
      if (e.key === "ArrowRight" && !isLastStep) onNext();
      if (e.key === "ArrowLeft" && currentStep > 0) onBack();
    };
    window.addEventListener("keydown", handleKey);
    return () => window.removeEventListener("keydown", handleKey);
  }, [isOpen, isLastStep, currentStep, onNext, onBack, onSkip]);

  if (!isOpen || !step) return null;

  const highlightPad = 6;
  const highlightTop = targetRect ? targetRect.top - highlightPad : 0;
  const highlightLeft = targetRect ? targetRect.left - highlightPad : 0;
  const highlightWidth = targetRect ? targetRect.width + highlightPad * 2 : 0;
  const highlightHeight = targetRect ? targetRect.height + highlightPad * 2 : 0;

  const shadowSpread = 9999;

  return (
    <div className="fixed inset-0 z-[9000] pointer-events-none">
      {isCentered ? (
        <div className="absolute inset-0 bg-slate-900/60 pointer-events-auto" />
      ) : (
        <div
          className="absolute pointer-events-auto"
          style={{
            top: highlightTop,
            left: highlightLeft,
            width: highlightWidth,
            height: highlightHeight,
            borderRadius: 10,
            boxShadow: `0 0 0 ${shadowSpread}px rgba(15, 23, 42, 0.6)`,
          }}
        />
      )}

      {!isCentered && targetRect && (
        <div
          className="absolute border-2 border-brand-500 rounded-[10px] pointer-events-none z-[9001]"
          style={{
            top: highlightTop,
            left: highlightLeft,
            width: highlightWidth,
            height: highlightHeight,
            boxShadow: "0 0 0 3px rgba(99, 102, 241, 0.25)",
          }}
        />
      )}

      <div
        className="absolute pointer-events-auto z-[9002]"
        style={{
          top: isCentered
            ? Math.max(
                PADDING,
                vpSize.h / 2 -
                  (hasItems ? 300 : POPOVER_APPROX_HEIGHT) / 2 -
                  20,
              )
            : popoverPos.top,
          left: isCentered
            ? Math.max(PADDING, vpSize.w / 2 - popoverWidth / 2)
            : popoverPos.left,
          width: popoverWidth,
          maxWidth: `calc(100vw - ${PADDING * 2}px)`,
        }}
      >
        <div className="bg-white rounded-2xl shadow-2xl border border-slate-200/80 overflow-hidden">
          <div className="bg-gradient-to-br from-slate-900 to-slate-800 px-5 py-4 flex items-start justify-between gap-3">
            <div className="flex items-center gap-2.5">
              <HelpCircle className="w-4 h-4 text-brand-400 shrink-0 mt-0.5" />
              <h3 className="text-white font-bold text-sm leading-snug">
                {step.title}
              </h3>
            </div>
            <button
              onClick={onSkip}
              className="text-slate-400 hover:text-white transition-colors shrink-0 mt-0.5"
              title="Skip tour"
            >
              <X className="w-4 h-4" />
            </button>
          </div>

          {hasItems ? (
            <div className="px-5 py-3">
              {step.body && (
                <p className="text-xs text-slate-500 leading-relaxed mb-3">
                  {step.body}
                </p>
              )}
              <div className="space-y-1 max-h-[280px] overflow-y-auto pr-1">
                {step.items!.map((item) => (
                  <div
                    key={item.id}
                    className={`flex items-start gap-3 px-3 py-2.5 rounded-xl transition-colors ${
                      item.important
                        ? "bg-amber-50 border border-amber-200 hover:bg-amber-100"
                        : "bg-slate-50 hover:bg-slate-100"
                    }`}
                  >
                    {item.important && (
                      <AlertTriangle className="w-3.5 h-3.5 text-amber-500 shrink-0 mt-0.5" />
                    )}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span
                          className={`text-[11px] font-bold ${
                            item.important ? "text-amber-800" : "text-slate-800"
                          }`}
                        >
                          {item.label}
                        </span>
                        {item.important && (
                          <span className="text-[9px] font-black uppercase tracking-wide text-amber-600 bg-amber-100 border border-amber-300 px-1.5 py-0.5 rounded-md">
                            Recommended
                          </span>
                        )}
                        {item.shortcut && (
                          <span className="text-[9px] font-mono font-bold text-slate-400 bg-white border border-slate-200 px-1.5 py-0.5 rounded-md">
                            {item.shortcut}
                          </span>
                        )}
                      </div>
                      <p
                        className={`text-[10px] mt-0.5 leading-relaxed ${
                          item.important ? "text-amber-700" : "text-slate-500"
                        }`}
                      >
                        {item.description}
                      </p>
                      {item.important && item.importantNote && (
                        <p className="text-[10px] mt-1 leading-relaxed text-amber-600 font-medium">
                          {item.importantNote}
                        </p>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ) : (
            <div className="px-5 py-4">
              <p className="text-xs text-slate-600 leading-relaxed">
                {step.body}
              </p>
            </div>
          )}

          <div className="px-5 pb-4 flex items-center justify-between gap-3">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">
              {currentStep + 1} / {steps.length}
            </span>

            <div className="flex items-center gap-2">
              {currentStep > 0 && (
                <button
                  onClick={onBack}
                  className="flex items-center gap-1 px-3 py-1.5 rounded-lg text-[11px] font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 transition-colors uppercase tracking-wide"
                >
                  <ChevronLeft className="w-3.5 h-3.5" />
                  Back
                </button>
              )}
              <button
                onClick={onSkip}
                className="px-3 py-1.5 rounded-lg text-[11px] font-bold text-slate-500 hover:bg-slate-100 transition-colors uppercase tracking-wide"
              >
                Skip
              </button>
              {isLastStep ? (
                <button
                  onClick={onFinish}
                  className="flex items-center gap-1 px-4 py-1.5 rounded-lg text-[11px] font-bold text-white bg-brand-600 hover:bg-brand-700 transition-colors uppercase tracking-wide"
                >
                  Finish
                </button>
              ) : (
                <button
                  onClick={onNext}
                  className="flex items-center gap-1 px-3 py-1.5 rounded-lg text-[11px] font-bold text-white bg-brand-600 hover:bg-brand-700 transition-colors uppercase tracking-wide"
                >
                  Next
                  <ChevronRight className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          </div>

          <div className="px-5 pb-3 flex gap-1">
            {steps.map((_, idx) => (
              <div
                key={idx}
                className={`h-1 rounded-full flex-1 transition-all duration-300 ${
                  idx === currentStep
                    ? "bg-brand-600"
                    : idx < currentStep
                      ? "bg-brand-300"
                      : "bg-slate-200"
                }`}
              />
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};
