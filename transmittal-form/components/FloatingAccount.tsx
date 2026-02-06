import React, { useState } from "react";

interface FloatingAccountProps {
  user: {
    name?: string | null;
    email?: string | null;
    image?: string | null;
  };
  isDriveReady: boolean;
  onSignOut: () => void;
}

export const FloatingAccount: React.FC<FloatingAccountProps> = ({
  user,
  isDriveReady,
  onSignOut,
}) => {
  const [isOpen, setIsOpen] = useState(false);

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
                onClick={onSignOut}
                className="w-full py-2.5 px-4 bg-slate-100 hover:bg-red-50 hover:text-red-600 rounded-xl text-xs font-bold uppercase tracking-widest transition-all"
              >
                Sign Out
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
