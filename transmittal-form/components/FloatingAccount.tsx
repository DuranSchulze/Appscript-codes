import React, { useEffect, useState } from "react";

interface FloatingAccountProps {
  user: {
    name?: string | null;
    email?: string | null;
    image?: string | null;
  };
  apiBaseUrl: string;
  isDriveReady: boolean;
  onSignOut: () => void;
}

export const FloatingAccount: React.FC<FloatingAccountProps> = ({
  user,
  apiBaseUrl,
  isDriveReady,
  onSignOut,
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [isAiSettingsOpen, setIsAiSettingsOpen] = useState(false);
  const [geminiApiKey, setGeminiApiKey] = useState("");
  const [showApiKey, setShowApiKey] = useState(false);
  const [hasCustomGeminiKey, setHasCustomGeminiKey] = useState(false);
  const [aiStatus, setAiStatus] = useState("");
  const [isLoadingAiSettings, setIsLoadingAiSettings] = useState(false);
  const [isSavingAiKey, setIsSavingAiKey] = useState(false);
  const [isDeletingAiKey, setIsDeletingAiKey] = useState(false);

  const loadAiSettings = async () => {
    if (!apiBaseUrl) return;
    setIsLoadingAiSettings(true);
    setAiStatus("");
    try {
      const response = await fetch(`${apiBaseUrl}/api/user-ai-settings`, {
        credentials: "include",
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload?.error || "Failed to load AI settings.");
      }
      setHasCustomGeminiKey(Boolean(payload?.hasCustomGeminiKey));
    } catch (error: any) {
      setAiStatus(error?.message || "Failed to load AI settings.");
    } finally {
      setIsLoadingAiSettings(false);
    }
  };

  const openAiSettings = () => {
    setGeminiApiKey("");
    setShowApiKey(false);
    setIsAiSettingsOpen(true);
    void loadAiSettings();
  };

  useEffect(() => {
    const handleOpenAiSettings = () => {
      openAiSettings();
    };

    window.addEventListener("open-ai-settings", handleOpenAiSettings);
    return () => {
      window.removeEventListener("open-ai-settings", handleOpenAiSettings);
    };
  }, []);

  const saveAiKey = async () => {
    const key = geminiApiKey.trim();
    if (!key) {
      setAiStatus("Please enter a Gemini API key.");
      return;
    }
    setIsSavingAiKey(true);
    setAiStatus("");
    try {
      const response = await fetch(`${apiBaseUrl}/api/user-ai-settings`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        credentials: "include",
        body: JSON.stringify({ geminiApiKey: key }),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload?.error || "Failed to save AI key.");
      }
      setHasCustomGeminiKey(Boolean(payload?.hasCustomGeminiKey));
      setGeminiApiKey("");
      setAiStatus(
        payload?.warning
          ? `Saved with warning: ${payload.warning}`
          : "Gemini API key saved.",
      );
    } catch (error: any) {
      setAiStatus(error?.message || "Failed to save AI key.");
    } finally {
      setIsSavingAiKey(false);
    }
  };

  const removeAiKey = async () => {
    setIsDeletingAiKey(true);
    setAiStatus("");
    try {
      const response = await fetch(`${apiBaseUrl}/api/user-ai-settings`, {
        method: "DELETE",
        credentials: "include",
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload?.error || "Failed to remove AI key.");
      }
      setHasCustomGeminiKey(false);
      setGeminiApiKey("");
      setAiStatus("Personal Gemini API key removed. Using system key.");
    } catch (error: any) {
      setAiStatus(error?.message || "Failed to remove AI key.");
    } finally {
      setIsDeletingAiKey(false);
    }
  };

  const initials = user.name
    ? user.name
        .split(" ")
        .map((n) => n[0])
        .join("")
        .toUpperCase()
        .slice(0, 2)
    : user.email?.[0]?.toUpperCase() || "U";

  return (
    <>
      {isOpen && (
        <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm">
          <button
            type="button"
            className="absolute inset-0 cursor-default"
            onClick={() => setIsOpen(false)}
            aria-label="Close account modal"
          />
          <div className="relative z-50 w-[340px] bg-white rounded-3xl shadow-2xl border border-slate-200/60 overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="bg-gradient-to-br from-slate-900 to-slate-800 p-6">
              <div className="flex items-center gap-4">
                {user.image ? (
                  <img
                    src={user.image}
                    alt={user.name || "User"}
                    className="w-16 h-16 rounded-full border-2 border-white/20"
                  />
                ) : (
                  <div className="w-16 h-16 rounded-full bg-gradient-to-br from-brand-500 to-brand-700 flex items-center justify-center text-white font-bold text-xl">
                    {initials}
                  </div>
                )}
                <div className="flex-1 min-w-0">
                  <p className="text-white font-bold truncate text-lg">
                    {user.name || "User"}
                  </p>
                  <p className="text-slate-400 text-xs truncate">
                    {user.email}
                  </p>
                </div>
              </div>
            </div>

            <div className="p-5 space-y-4">
              <div className="flex items-center gap-3 text-sm">
                <div
                  className={`w-2.5 h-2.5 rounded-full ${isDriveReady ? "bg-green-500" : "bg-amber-500"}`}
                />
                <span className="text-slate-600">
                  {isDriveReady
                    ? "Google Drive Connected"
                    : "Drive Not Connected"}
                </span>
              </div>

              <div className="pt-2 border-t border-slate-100">
                <p className="text-[10px] text-slate-400 uppercase tracking-widest mb-2">
                  Stored in Cloud
                </p>
                <div className="flex items-center gap-2">
                  <svg
                    className="w-4 h-4 text-slate-400"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth="2"
                      d="M3 15a4 4 0 004 4h9a5 5 0 10-.1-9.999 5.002 5.002 0 10-9.78 2.096A4.001 4.001 0 003 15z"
                    />
                  </svg>
                  <span className="text-xs text-slate-500">
                    PostgreSQL Database
                  </span>
                </div>
              </div>

              <button
                onClick={openAiSettings}
                className="w-full py-2.5 px-4 bg-slate-100 hover:bg-brand-50 hover:text-brand-700 rounded-xl text-xs font-bold uppercase tracking-widest transition-all"
              >
                AI Key Settings
              </button>

              <button
                onClick={onSignOut}
                className="w-full py-2.5 px-4 bg-slate-100 hover:bg-red-50 hover:text-red-600 rounded-xl text-xs font-bold uppercase tracking-widest transition-all"
              >
                Sign Out
              </button>
            </div>
          </div>
        </div>
      )}

      {isAiSettingsOpen && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <button
            type="button"
            className="absolute inset-0 cursor-default"
            onClick={() => setIsAiSettingsOpen(false)}
            aria-label="Close AI settings modal"
          />
          <div className="relative z-[70] w-full max-w-md bg-white rounded-3xl shadow-2xl border border-slate-200/60 overflow-hidden">
            <div className="p-6 border-b border-slate-100">
              <h3 className="text-sm font-black uppercase tracking-widest text-slate-800">
                Gemini API Key
              </h3>
              <p className="mt-1 text-xs text-slate-500">
                Optional. If set, your personal key is preferred over the system
                key during AI parsing.
              </p>
            </div>
            <div className="p-6 space-y-4">
              <div className="text-xs text-slate-600">
                Current source:{" "}
                <span className="font-bold text-slate-800">
                  {isLoadingAiSettings
                    ? "Loading..."
                    : hasCustomGeminiKey
                      ? "Personal key"
                      : "System key"}
                </span>
              </div>

              <div className="space-y-2">
                <label
                  htmlFor="gemini-api-key"
                  className="text-[10px] font-black uppercase tracking-widest text-slate-500"
                >
                  New / Replace Key
                </label>
                <input
                  id="gemini-api-key"
                  type={showApiKey ? "text" : "password"}
                  value={geminiApiKey}
                  onChange={(e) => setGeminiApiKey(e.target.value)}
                  placeholder="Paste Gemini API key"
                  className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs text-slate-700 outline-none focus:border-brand-500"
                />
                <button
                  type="button"
                  onClick={() => setShowApiKey((prev) => !prev)}
                  className="text-[10px] font-bold uppercase tracking-widest text-slate-500 hover:text-slate-700"
                >
                  {showApiKey ? "Hide Key" : "Show Key"}
                </button>
              </div>

              {aiStatus ? (
                <p className="rounded-xl bg-slate-100 px-3 py-2 text-[11px] text-slate-700 break-words">
                  {aiStatus}
                </p>
              ) : null}

              <div className="grid grid-cols-3 gap-2">
                <button
                  onClick={saveAiKey}
                  disabled={isSavingAiKey || isDeletingAiKey || isLoadingAiSettings}
                  className="col-span-2 rounded-xl bg-brand-600 text-white py-2 text-[10px] font-bold uppercase tracking-widest hover:bg-brand-700 disabled:opacity-50"
                >
                  {isSavingAiKey ? "Saving..." : "Save / Test Key"}
                </button>
                <button
                  onClick={removeAiKey}
                  disabled={
                    isDeletingAiKey ||
                    isSavingAiKey ||
                    isLoadingAiSettings ||
                    !hasCustomGeminiKey
                  }
                  className="rounded-xl bg-slate-100 text-slate-700 py-2 text-[10px] font-bold uppercase tracking-widest hover:bg-red-50 hover:text-red-600 disabled:opacity-50"
                >
                  {isDeletingAiKey ? "..." : "Remove"}
                </button>
              </div>

              <button
                onClick={() => setIsAiSettingsOpen(false)}
                className="w-full rounded-xl bg-slate-100 text-slate-700 py-2 text-[10px] font-bold uppercase tracking-widest hover:bg-slate-200"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="fixed bottom-6 right-6 z-50">
        <button
          onClick={() => setIsOpen(true)}
          className="w-14 h-14 rounded-full shadow-lg flex items-center justify-center transition-all duration-200 bg-gradient-to-br from-brand-500 to-brand-700 hover:scale-105"
        >
          {user.image ? (
            <img
              src={user.image}
              alt={user.name || "User"}
              className="w-full h-full rounded-full border-2 border-white"
            />
          ) : (
            <span className="text-white font-bold">{initials}</span>
          )}
        </button>
      </div>
    </>
  );
};
