"use client";

import React from "react";
import Link from "next/link";
import { Button } from "@/components/ui/button";

interface LoginScreenProps {
  onGoogleSignIn: () => void;
}

export const LoginScreen: React.FC<LoginScreenProps> = ({ onGoogleSignIn }) => {
  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-100 font-sans">
      <div className="bg-white border border-slate-200 rounded-[36px] shadow-2xl w-full max-w-lg p-10 text-center">
        <h1 className="text-2xl font-black text-slate-900 font-display">
          Smart Transmittal
        </h1>
        <p className="text-sm text-slate-500 mt-2">
          This is a private system created for the Internal Document Transmittal
          System. Please sign in
        </p>
        <Button
          onClick={onGoogleSignIn}
          className="mt-8 w-full py-4 rounded-2xl bg-slate-900 text-white font-bold uppercase tracking-widest text-xs shadow-xl hover:bg-slate-800 transition-all"
        >
          Sign in with Google
        </Button>
        {/* <div className="flex justify-center gap-6 mt-6">
          <Link
            href="/legal/privacy-policy"
            target="_blank"
            className="text-[10px] font-bold uppercase tracking-widest text-slate-400 hover:text-slate-600 transition-colors"
          >
            Privacy Policy
          </Link>
          <Link
            href="/legal/terms-of-service"
            target="_blank"
            className="text-[10px] font-bold uppercase tracking-widest text-slate-400 hover:text-slate-600 transition-colors"
          >
            Terms of Service
          </Link>
        </div> */}
      </div>
    </div>
  );
};
