"use client";

import React from "react";

export type TabKey =
  | "content"
  | "sender"
  | "recipient"
  | "project"
  | "signatories";

const tabs: { key: TabKey; label: string }[] = [
  { key: "content", label: "Content" },
  { key: "sender", label: "Brand" },
  { key: "recipient", label: "Recipient" },
  { key: "project", label: "Project" },
  { key: "signatories", label: "Sign-off" },
];

interface TabBarProps {
  activeTab: TabKey;
  onTabChange: (tab: TabKey) => void;
}

export const TabBar: React.FC<TabBarProps> = ({ activeTab, onTabChange }) => {
  const tabTourIds: Partial<Record<TabKey, string>> = {
    sender: "tab-brand",
    recipient: "tab-recipient",
    project: "tab-project",
    signatories: "tab-signoff",
  };

  return (
    <div
      data-tour="tab-bar"
      className="flex flex-wrap p-2 bg-slate-100/50 m-5 rounded-2xl gap-2 shrink-0"
    >
      {tabs.map(({ key, label }) => (
        <button
          key={key}
          data-tour={tabTourIds[key]}
          onClick={() => onTabChange(key)}
          className={`flex-1 py-2 px-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all whitespace-nowrap ${
            activeTab === key
              ? "bg-white shadow-sm text-slate-800"
              : "text-slate-400 hover:text-slate-600"
          }`}
        >
          {label}
        </button>
      ))}
    </div>
  );
};
