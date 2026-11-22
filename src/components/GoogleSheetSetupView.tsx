import React, { useState } from 'react';
import { motion } from 'motion/react';
import {
  CheckCircle2,
  ExternalLink,
  FileSpreadsheet,
  Link2,
  LogOut,
  PlusCircle,
  ShieldCheck,
  Users,
} from 'lucide-react';
import type { PartnerId } from '../../shared/types.ts';
import { api, type SharedSheetConfig } from '../lib/api.ts';
import { googleSignIn, googleSignOut } from '../lib/googleWorkspace.ts';
import { isSheetActionCancelled } from '../lib/sheetConfirm.ts';
import { GoogleSignInButton } from './GoogleSignInButton.tsx';
import { SunlitGalleryBackdrop } from './SunlitGalleryBackdrop.tsx';

interface GoogleSheetSetupViewProps {
  partnerId: PartnerId;
  googleEmail: string | null;
  hasGoogleToken: boolean;
  existingServerConfig: SharedSheetConfig | null;
  authNotice?: string | null;
  onReady: (config: SharedSheetConfig) => void;
  onSignOutPartner: () => void;
}

export const GoogleSheetSetupView: React.FC<GoogleSheetSetupViewProps> = ({
  partnerId,
  googleEmail,
  hasGoogleToken,
  existingServerConfig,
  authNotice,
  onReady,
  onSignOutPartner,
}) => {
  const [signingInGoogle, setSigningInGoogle] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(authNotice || null);

  const [newSheetTitle, setNewSheetTitle] = useState('PartnerDesk — Options Trading Ledger');
  const [partnerEmail, setPartnerEmail] = useState('');
  const [existingUrlInput, setExistingUrlInput] = useState(
    existingServerConfig?.spreadsheetUrl || '',
  );

  const handleGoogleSignIn = async () => {
    setError(null);
    setSigningInGoogle(true);
    try {
      await googleSignIn();
      if (existingServerConfig) {
        // Verify access directly for the shared spreadsheet already linked on the server
        const connected = await api.connectExistingSharedSheet(
          existingServerConfig.spreadsheetId,
        );
        onReady(connected);
      }
    } catch (err) {
      if (!isSheetActionCancelled(err)) {
        setError(err instanceof Error ? err.message : 'Google Sign-In failed.');
      }
    } finally {
      setSigningInGoogle(false);
    }
  };

  const handleConnectExistingServerSheet = async () => {
    if (!existingServerConfig) return;
    setError(null);
    setBusy(true);
    try {
      const cfg = await api.connectExistingSharedSheet(existingServerConfig.spreadsheetId);
      onReady(cfg);
    } catch (err) {
      if (!isSheetActionCancelled(err)) {
        setError(
          err instanceof Error
            ? err.message
            : 'Could not open the shared Google Sheet. Ask your partner to share Editor access with your Gmail.',
        );
      }
    } finally {
      setBusy(false);
    }
  };

  const handleCreateNewSheet = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setBusy(true);
    try {
      const cfg = await api.createAndLinkSharedSheet({
        title: newSheetTitle.trim() || 'PartnerDesk — Options Trading Ledger',
        partnerEmail: partnerEmail.trim() || undefined,
      });
      onReady(cfg);
    } catch (err) {
      if (!isSheetActionCancelled(err)) {
        setError(err instanceof Error ? err.message : 'Could not create Google Sheet.');
      }
    } finally {
      setBusy(false);
    }
  };

  const handleConnectCustomSheet = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!existingUrlInput.trim()) {
      setError('Paste a Google Sheets link or Spreadsheet ID.');
      return;
    }
    setError(null);
    setBusy(true);
    try {
      const cfg = await api.connectExistingSharedSheet(existingUrlInput.trim());
      onReady(cfg);
    } catch (err) {
      if (!isSheetActionCancelled(err)) {
        setError(err instanceof Error ? err.message : 'Could not connect Google Sheet.');
      }
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="relative min-h-screen flex flex-col justify-between overflow-x-hidden">
      <SunlitGalleryBackdrop />

      <header className="w-full px-4 sm:px-8 lg:px-12 h-16 flex items-center justify-between border-b border-slate-900/8 backdrop-blur-md bg-white/65 z-10">
        <div className="flex items-center gap-3">
          <span className="font-display text-lg font-bold tracking-tight text-[#111318]">
            PartnerDesk
          </span>
          <span className="text-xs font-semibold text-slate-600 bg-slate-900/5 px-2.5 py-1 rounded-full">
            Signed in as {partnerId}
          </span>
        </div>
        <button
          type="button"
          onClick={async () => {
            await googleSignOut();
            onSignOutPartner();
          }}
          className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-700 hover:text-[#111318] bg-white/85 border border-slate-900/12 rounded-full cursor-pointer"
        >
          <LogOut className="w-3.5 h-3.5" />
          <span>Sign Out</span>
        </button>
      </header>

      <main className="flex-1 w-full max-w-4xl mx-auto px-4 sm:px-8 py-8 sm:py-12 flex items-center z-10">
        <motion.div
          initial={{ opacity: 0, y: 14 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.35, ease: [0.16, 1, 0.3, 1] }}
          className="w-full space-y-6"
        >
          <div className="gallery-card rounded-3xl p-6 sm:p-8 space-y-5">
            <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between gap-4">
              <div>
                <div className="inline-flex items-center gap-1.5 text-xs font-semibold text-emerald-700">
                  <ShieldCheck className="w-4 h-4" />
                  <span>Live Shared Google Sheet Storage</span>
                </div>
                <h1 className="font-display text-xl sm:text-2xl font-bold text-[#111318] mt-1">
                  Connect Shared Partnership Google Sheet
                </h1>
                <p className="text-xs sm:text-sm text-slate-600 mt-1.5 leading-relaxed max-w-2xl">
                  All trades, screenshots, deposits, withdrawals, and charge rates are saved directly
                  inside one shared Google Sheet — not locally — so every change made by Prayag or
                  Raj appears identically for both partners on any device.
                </p>
              </div>
            </div>

            {/* Step 1: Google Account Authorization */}
            <div className="p-4 sm:p-5 rounded-2xl bg-slate-900/[0.03] border border-slate-900/10 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
              <div className="space-y-1">
                <div className="text-xs font-semibold text-[#111318] flex items-center gap-2">
                  <span className="w-5 h-5 rounded-full bg-[#111318] text-[#E6B949] text-[11px] font-bold inline-flex items-center justify-center">
                    1
                  </span>
                  <span>Authorize Google Sheets & Drive Access</span>
                </div>
                <p className="text-xs text-slate-600">
                  {hasGoogleToken
                    ? `Connected with Google (${googleEmail || 'Authorized'})`
                    : 'Sign in with Google so PartnerDesk can read and write your shared spreadsheet.'}
                </p>
              </div>

              {hasGoogleToken ? (
                <div className="flex items-center gap-2.5 shrink-0">
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-emerald-500/10 border border-emerald-600/20 text-xs font-semibold text-emerald-800">
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    <span>Google Connected</span>
                  </span>
                  <button
                    type="button"
                    onClick={() => void googleSignOut()}
                    className="px-3 py-1.5 text-xs font-medium text-slate-600 hover:text-[#111318] bg-white border border-slate-900/12 rounded-full cursor-pointer"
                  >
                    Switch Account
                  </button>
                </div>
              ) : (
                <GoogleSignInButton
                  onClick={() => void handleGoogleSignIn()}
                  disabled={signingInGoogle || busy}
                  label={signingInGoogle ? 'Signing in...' : 'Sign in with Google'}
                />
              )}
            </div>

            {error && (
              <div
                role="alert"
                className="p-3.5 text-xs sm:text-sm text-rose-700 bg-rose-50 border border-rose-200 rounded-2xl"
              >
                {error}
              </div>
            )}

            {/* Step 2: Connect existing shared sheet OR create/link one */}
            {hasGoogleToken && (
              <div className="space-y-5 pt-1">
                {existingServerConfig && (
                  <div className="p-4 sm:p-5 rounded-2xl bg-[#111318] text-white space-y-3">
                    <div className="flex flex-wrap items-start justify-between gap-3">
                      <div>
                        <div className="text-xs font-semibold text-[#E6B949] flex items-center gap-1.5">
                          <Users className="w-3.5 h-3.5" />
                          <span>
                            Shared Sheet Linked by {existingServerConfig.linkedBy}
                          </span>
                        </div>
                        <div className="text-base font-bold mt-1">
                          {existingServerConfig.spreadsheetTitle}
                        </div>
                        <div className="text-[11px] text-slate-300 font-mono mt-0.5 break-all">
                          ID: {existingServerConfig.spreadsheetId}
                        </div>
                      </div>
                      <a
                        href={existingServerConfig.spreadsheetUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-white/10 hover:bg-white/20 rounded-full"
                      >
                        <span>Open in Sheets</span>
                        <ExternalLink className="w-3.5 h-3.5" />
                      </a>
                    </div>
                    <button
                      type="button"
                      disabled={busy}
                      onClick={() => void handleConnectExistingServerSheet()}
                      className="min-h-[44px] w-full sm:w-auto inline-flex items-center justify-center gap-2 px-5 py-2.5 bg-[#E6B949] hover:bg-[#d6a938] text-[#111318] text-xs font-bold rounded-full cursor-pointer"
                    >
                      <FileSpreadsheet className="w-4 h-4" />
                      <span>{busy ? 'Opening Shared Sheet...' : 'Open Shared PartnerDesk Ledger'}</span>
                    </button>
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {/* Option A: Create New Shared Google Sheet */}
                  <form
                    onSubmit={handleCreateNewSheet}
                    className="p-5 rounded-2xl bg-white/80 border border-slate-900/12 space-y-3.5"
                  >
                    <div className="flex items-center gap-2 font-display text-sm font-bold text-[#111318]">
                      <PlusCircle className="w-4 h-4 text-[#C89D3C]" />
                      <span>Create New Shared Google Sheet</span>
                    </div>
                    <p className="text-xs text-slate-600 leading-relaxed">
                      Creates a ready-to-use workbook with Summary, Trade_Groups, Order_Legs,
                      Deposits_Ledger, Charge_Rates, Screenshots, and Audit_Logs tabs.
                    </p>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Spreadsheet Name
                      </label>
                      <input
                        type="text"
                        value={newSheetTitle}
                        onChange={(e) => setNewSheetTitle(e.target.value)}
                        className="w-full h-10 px-3 text-xs bg-white border border-slate-900/15 rounded-xl text-[#111318]"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Partner&apos;s Gmail to Auto-Share (Optional)
                      </label>
                      <input
                        type="email"
                        value={partnerEmail}
                        onChange={(e) => setPartnerEmail(e.target.value)}
                        placeholder={
                          partnerId === 'Prayag' ? 'raj@gmail.com' : 'prayag@gmail.com'
                        }
                        className="w-full h-10 px-3 text-xs bg-white border border-slate-900/15 rounded-xl text-[#111318]"
                      />
                    </div>
                    <button
                      type="submit"
                      disabled={busy}
                      className="w-full min-h-[42px] inline-flex items-center justify-center gap-2 px-4 py-2 bg-[#111318] hover:bg-[#242934] disabled:bg-slate-400 text-white text-xs font-semibold rounded-full cursor-pointer"
                    >
                      <FileSpreadsheet className="w-4 h-4 text-[#E6B949]" />
                      <span>{busy ? 'Creating...' : 'Create & Link Google Sheet'}</span>
                    </button>
                  </form>

                  {/* Option B: Connect Existing Google Sheet URL / ID */}
                  <form
                    onSubmit={handleConnectCustomSheet}
                    className="p-5 rounded-2xl bg-white/80 border border-slate-900/12 space-y-3.5 flex flex-col justify-between"
                  >
                    <div className="space-y-3.5">
                      <div className="flex items-center gap-2 font-display text-sm font-bold text-[#111318]">
                        <Link2 className="w-4 h-4 text-[#C89D3C]" />
                        <span>Connect Existing Google Sheet</span>
                      </div>
                      <p className="text-xs text-slate-600 leading-relaxed">
                        Already created a shared Google Sheet with your partner? Paste its full
                        Google Sheets URL or Spreadsheet ID here.
                      </p>
                      <div>
                        <label className="block text-xs font-semibold text-slate-700 mb-1">
                          Google Sheet URL or Spreadsheet ID
                        </label>
                        <input
                          type="text"
                          value={existingUrlInput}
                          onChange={(e) => setExistingUrlInput(e.target.value)}
                          placeholder="https://docs.google.com/spreadsheets/d/..."
                          className="w-full h-10 px-3 text-xs bg-white border border-slate-900/15 rounded-xl font-mono text-[#111318]"
                        />
                      </div>
                    </div>
                    <button
                      type="submit"
                      disabled={busy}
                      className="w-full min-h-[42px] inline-flex items-center justify-center gap-2 px-4 py-2 bg-white hover:bg-slate-50 text-[#111318] border border-slate-900/20 text-xs font-semibold rounded-full cursor-pointer"
                    >
                      <Link2 className="w-4 h-4" />
                      <span>{busy ? 'Connecting...' : 'Connect Existing Sheet'}</span>
                    </button>
                  </form>
                </div>
              </div>
            )}
          </div>
        </motion.div>
      </main>
    </div>
  );
};
