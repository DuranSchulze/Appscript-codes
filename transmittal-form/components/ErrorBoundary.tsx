"use client";

import React, { ErrorInfo, ReactNode, Component } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";

interface ErrorBoundaryProps {
  children?: ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  public state: ErrorBoundaryState = { hasError: false, error: null };

  constructor(props: ErrorBoundaryProps) {
    super(props);
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("Uncaught Error:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen flex items-center justify-center bg-slate-100 p-6 font-sans">
          <Card className="max-w-lg w-full text-center border border-slate-200 rounded-3xl shadow-2xl">
            <CardContent className="p-8">
              <div className="w-20 h-20 bg-red-50 text-red-500 rounded-full flex items-center justify-center mx-auto mb-6 shadow-inner">
                <svg
                  className="w-10 h-10"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"
                  ></path>
                </svg>
              </div>
              <h2 className="text-2xl font-black text-slate-800 mb-2 font-display">
                Dashboard Error
              </h2>
              <p className="text-slate-500 mb-8 px-4 leading-relaxed">
                The application encountered a rendering issue. This is usually
                temporary.
              </p>
              <div className="bg-slate-50 p-4 rounded-xl text-left mb-8 overflow-auto max-h-40 border border-slate-100">
                <code className="text-[10px] text-red-500 font-mono break-all">
                  {this.state.error?.message || "Internal Component Error"}
                </code>
              </div>
              <Button
                onClick={() => window.location.reload()}
                className="w-full py-4 rounded-2xl font-bold shadow-xl"
                size="lg"
              >
                Reload Workspace
              </Button>
            </CardContent>
          </Card>
        </div>
      );
    }
    return this.props.children;
  }
}
