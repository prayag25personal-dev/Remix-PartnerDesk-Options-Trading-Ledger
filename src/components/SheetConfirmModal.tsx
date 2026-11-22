import React, { useEffect, useState } from 'react';
import { AlertTriangle, FileSpreadsheet, X } from 'lucide-react';
import { registerSheetConfirmHandler, type SheetConfirmRequest } from '../lib/sheetConfirm.ts';

interface PendingConfirm extends SheetConfirmRequest {
  resolve: (accepted: boolean) => void;
}

export const SheetConfirmModal: React.FC = () => {
  const [pending, setPending] = useState<PendingConfirm | null>(null);

  useEffect(() => {
    registerSheetConfirmHandler((req) => {
      return new Promise<boolean>((resolve) => {
        setPending({ ...req, resolve });
      });
    });
    return () => {
      registerSheetConfirmHandler(null);
    };
  }, []);

  if (!pending) return null;

  const isDanger = pending.tone === 'danger';

  const closeWith = (accepted: boolean) => {
    const resolver = pending.resolve;
    setPending(null);
    resolver(accepted);
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="sheet-confirm-title"
      className="fixed inset-0 z-50 bg-slate-950/50 backdrop-blur-xs flex items-center justify-center p-4"
      onClick={() => closeWith(false)}
    >
      <div
        className="w-full max-w-md bg-white rounded-3xl border border-slate-900/12 shadow-2xl p-5 sm:p-6 space-y-4"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-3">
          <div className="flex items-start gap-3">
            <div
              className={`w-9 h-9 rounded-2xl flex items-center justify-center shrink-0 ${
                isDanger
                  ? 'bg-rose-100 text-rose-700'
                  : 'bg-[#111318] text-[#E6B949]'
              }`}
            >
              {isDanger ? (
                <AlertTriangle className="w-4.5 h-4.5" />
              ) : (
                <FileSpreadsheet className="w-4.5 h-4.5" />
              )}
            </div>
            <div>
              <h3
                id="sheet-confirm-title"
                className="font-display text-base font-bold text-[#111318]"
              >
                {pending.title}
              </h3>
              <p className="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">
                {pending.description}
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={() => closeWith(false)}
            aria-label="Cancel"
            className="min-w-[36px] min-h-[36px] inline-flex items-center justify-center rounded-full text-slate-400 hover:text-[#111318] hover:bg-slate-100 cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {pending.details && pending.details.length > 0 && (
          <ul className="bg-slate-900/[0.035] border border-slate-900/8 rounded-2xl p-3.5 space-y-1.5 text-xs text-slate-700">
            {pending.details.map((line, idx) => (
              <li key={idx} className="leading-relaxed">
                {line}
              </li>
            ))}
          </ul>
        )}

        <div className="flex items-center justify-end gap-2.5 pt-1">
          <button
            type="button"
            onClick={() => closeWith(false)}
            className="min-h-[42px] px-4 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-full cursor-pointer transition-colors"
          >
            {pending.cancelLabel || 'Cancel'}
          </button>
          <button
            type="button"
            onClick={() => closeWith(true)}
            className={`min-h-[42px] px-5 py-2 text-xs font-semibold text-white rounded-full cursor-pointer shadow-xs transition-colors ${
              isDanger
                ? 'bg-rose-600 hover:bg-rose-700'
                : 'bg-[#111318] hover:bg-[#242934]'
            }`}
          >
            {pending.confirmLabel || 'Confirm Update'}
          </button>
        </div>
      </div>
    </div>
  );
};
