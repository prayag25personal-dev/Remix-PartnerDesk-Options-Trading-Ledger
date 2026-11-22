import React, { useCallback, useEffect, useRef, useState } from 'react';
import { motion } from 'motion/react';
import {
  CheckCircle2,
  Download,
  ExternalLink,
  FileSpreadsheet,
  KeyRound,
  RotateCcw,
  Share2,
  Sliders,
  Unlink,
  Upload,
} from 'lucide-react';
import { DEFAULT_CHARGE_RATES } from '../../shared/calculations.ts';
import type { AuditLogRecord, ChargeRates, PartnerId } from '../../shared/types.ts';
import { api, type SharedSheetConfig } from '../lib/api.ts';
import { formatDateTime } from '../lib/format.ts';
import { isSheetActionCancelled } from '../lib/sheetConfirm.ts';

interface SettingsViewProps {
  partnerId: PartnerId;
  googleEmail: string | null;
  sheetConfig: SharedSheetConfig;
  rates: ChargeRates;
  envPasswordConfigured: boolean;
  onRatesUpdated: (rates: ChargeRates) => void;
  onSheetConfigChanged: (config: SharedSheetConfig | null) => void;
  onStorageChanged: () => void;
  onExportExcel: () => void;
}

function formatRateInput(value: number): string {
  if (!Number.isFinite(value) || value === 0) return '0';
  const fixed = value.toFixed(10).replace(/\.?0+$/, '');
  return fixed === '' ? '0' : fixed;
}

function ratesToInputs(r: ChargeRates): Record<keyof ChargeRates, string> {
  return {
    brokeragePerOrder: formatRateInput(r.brokeragePerOrder),
    sttSellPct: formatRateInput(r.sttSellPct),
    exchangeTxnPct: formatRateInput(r.exchangeTxnPct),
    ipftPct: formatRateInput(r.ipftPct),
    sebiPct: formatRateInput(r.sebiPct),
    stampDutyBuyPct: formatRateInput(r.stampDutyBuyPct),
    gstPct: formatRateInput(r.gstPct),
  };
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  partnerId,
  googleEmail,
  sheetConfig,
  rates,
  envPasswordConfigured,
  onRatesUpdated,
  onSheetConfigChanged,
  onStorageChanged,
  onExportExcel,
}) => {
  const restoreInputRef = useRef<HTMLInputElement | null>(null);

  const [counts, setCounts] = useState<{
    trades: number;
    orders: number;
    screenshots: number;
    deposits: number;
  } | null>(null);
  const [partnerEmail, setPartnerEmail] = useState('');
  const [sharing, setSharing] = useState(false);
  const [sheetMsg, setSheetMsg] = useState<string | null>(null);
  const [sheetErr, setSheetErr] = useState<string | null>(null);
  const [workingBackup, setWorkingBackup] = useState(false);

  const [rateInputs, setRateInputs] = useState<Record<keyof ChargeRates, string>>(() =>
    ratesToInputs(rates),
  );
  const [recalcExisting, setRecalcExisting] = useState(false);
  const [savingRates, setSavingRates] = useState(false);
  const [ratesMsg, setRatesMsg] = useState<string | null>(null);
  const [ratesErr, setRatesErr] = useState<string | null>(null);

  const [currentPassword, setCurrentPassword] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [savingPwd, setSavingPwd] = useState(false);
  const [pwdMsg, setPwdMsg] = useState<string | null>(null);
  const [pwdErr, setPwdErr] = useState<string | null>(null);

  const [auditLogs, setAuditLogs] = useState<AuditLogRecord[]>([]);
  const [loadingLogs, setLoadingLogs] = useState(true);

  useEffect(() => {
    setRateInputs(ratesToInputs(rates));
  }, [rates]);

  const loadStatusAndLogs = useCallback(async () => {
    setLoadingLogs(true);
    try {
      const [backup, logsRes] = await Promise.all([
        api.getBackupPayload(),
        api.listAuditLogs(),
      ]);
      setCounts({
        trades: backup.tradeGroups.length,
        orders: backup.orders.length,
        screenshots: backup.uploads.length,
        deposits: backup.deposits.length,
      });
      setAuditLogs(logsRes.logs);
    } catch (err) {
      setSheetErr(err instanceof Error ? err.message : 'Unable to read Google Sheet details.');
    } finally {
      setLoadingLogs(false);
    }
  }, []);

  useEffect(() => {
    void loadStatusAndLogs();
  }, [loadStatusAndLogs]);

  const handleShareSheet = async (e: React.FormEvent) => {
    e.preventDefault();
    setSheetMsg(null);
    setSheetErr(null);
    if (!partnerEmail.trim()) {
      setSheetErr('Enter your partner Gmail address to grant Editor access.');
      return;
    }
    setSharing(true);
    try {
      await api.shareActiveSheetWithPartner(partnerEmail.trim());
      setSheetMsg(
        `Shared "${sheetConfig.spreadsheetTitle}" as Editor with ${partnerEmail.trim()}. Ask them to sign in with Google on PartnerDesk to edit the exact same live spreadsheet.`,
      );
      setPartnerEmail('');
    } catch (err) {
      if (!isSheetActionCancelled(err)) {
        setSheetErr(err instanceof Error ? err.message : 'Could not share Google Sheet.');
      }
    } finally {
      setSharing(false);
    }
  };

  const handleDisconnectSheet = async () => {
    setSheetMsg(null);
    setSheetErr(null);
    setWorkingBackup(true);
    try {
      await api.clearSharedSheetConfig();
      onSheetConfigChanged(null);
    } catch (err) {
      setSheetErr(err instanceof Error ? err.message : 'Failed to unlink spreadsheet.');
    } finally {
      setWorkingBackup(false);
    }
  };

  const handleDownloadBackup = async () => {
    setSheetMsg(null);
    setSheetErr(null);
    try {
      const payload = await api.getBackupPayload();
      api.downloadBackupJson(payload);
      setSheetMsg('Downloaded PartnerDesk JSON backup from the shared Google Sheet.');
    } catch (err) {
      setSheetErr(err instanceof Error ? err.message : 'Failed to export backup.');
    }
  };

  const handleRestoreBackupFile = async (file: File) => {
    setSheetMsg(null);
    setSheetErr(null);
    setWorkingBackup(true);
    try {
      const text = await file.text();
      let parsed: Record<string, unknown>;
      try {
        parsed = JSON.parse(text) as Record<string, unknown>;
      } catch {
        throw new Error('Invalid backup JSON file.');
      }
      if (!parsed || !Array.isArray(parsed.deposits) || !Array.isArray(parsed.tradeGroups)) {
        throw new Error('Selected file is not a valid PartnerDesk backup.');
      }
      const res = await api.mergeBackupPayload(parsed);
      setSheetMsg(
        `Merged backup into shared Google Sheet for both Prayag & Raj: +${res.addedTrades} trade(s), +${res.addedDeposits} deposit(s), +${res.addedScreenshots} screenshot(s).`,
      );
      onStorageChanged();
      await loadStatusAndLogs();
    } catch (err) {
      if (!isSheetActionCancelled(err)) {
        setSheetErr(err instanceof Error ? err.message : 'Failed to import backup.');
      }
    } finally {
      setWorkingBackup(false);
    }
  };

  const handleSaveRates = async (e: React.FormEvent) => {
    e.preventDefault();
    setRatesMsg(null);
    setRatesErr(null);

    const parsedRates: Partial<ChargeRates> = {};
    for (const key of Object.keys(rateInputs) as (keyof ChargeRates)[]) {
      const raw = rateInputs[key].trim();
      if (raw === '') {
        setRatesErr('Please fill in all charge rate fields (use 0 if a charge is zero).');
        return;
      }
      const num = Number(raw);
      if (!Number.isFinite(num) || num < 0) {
        setRatesErr(`Enter a valid non-negative number for ${key}.`);
        return;
      }
      parsedRates[key] = num;
    }

    setSavingRates(true);
    try {
      const res = await api.updateRates({
        ...(parsedRates as ChargeRates),
        recalculateExisting: recalcExisting,
      });
      onRatesUpdated(res.rates);
      setRateInputs(ratesToInputs(res.rates));
      setRatesMsg('Charge rates updated in Google Sheet for both Prayag & Raj.');
      void loadStatusAndLogs();
    } catch (err) {
      if (!isSheetActionCancelled(err)) {
        setRatesErr(err instanceof Error ? err.message : 'Failed to update rates.');
      }
    } finally {
      setSavingRates(false);
    }
  };

  const handleUpdatePassword = async (e: React.FormEvent) => {
    e.preventDefault();
    setPwdMsg(null);
    setPwdErr(null);
    if (newPassword.length < 6) {
      setPwdErr('New password must be at least 6 characters.');
      return;
    }
    if (newPassword !== confirmPassword) {
      setPwdErr('New passwords do not match.');
      return;
    }
    setSavingPwd(true);
    try {
      await api.updatePassword(currentPassword, newPassword);
      setCurrentPassword('');
      setNewPassword('');
      setConfirmPassword('');
      setPwdMsg(
        `Login password updated for ${partnerId}. Both accounts continue sharing the exact same Google Sheet ledger.`,
      );
      void loadStatusAndLogs();
    } catch (err) {
      setPwdErr(err instanceof Error ? err.message : 'Failed to update password.');
    } finally {
      setSavingPwd(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Shared Google Sheet Storage & Collaboration Panel */}
      <motion.section
        initial={{ opacity: 0, y: 10 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.3, ease: [0.16, 1, 0.3, 1] }}
        className="gallery-card rounded-3xl p-4 sm:p-6 space-y-5"
      >
        <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between gap-3">
          <div>
            <h2 className="font-display text-base font-bold text-[#111318] flex items-center gap-2">
              <span className="w-7 h-7 rounded-xl bg-[#111318] text-[#E6B949] flex items-center justify-center shrink-0">
                <FileSpreadsheet className="w-3.5 h-3.5" />
              </span>
              <span>Shared Google Sheet Storage (Live Partnership Workbook)</span>
            </h2>
            <p className="text-xs text-slate-500 mt-1.5 leading-relaxed">
              Every trade, screenshot, deposit, withdrawal, and charge rate is saved directly to{' '}
              <span className="font-semibold text-[#111318]">{sheetConfig.spreadsheetTitle}</span>{' '}
              in Google Sheets. Both Prayag and Raj read and update the exact same spreadsheet from
              any device.
            </p>
          </div>
          <span className="self-start inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-full bg-emerald-500/10 border border-emerald-600/20 text-xs font-semibold text-emerald-800 whitespace-nowrap">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-700" />
            <span>Live Google Sheet Sync</span>
          </span>
        </div>

        {/* Sheet Metrics & Info Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
          <div className="gallery-card-dark rounded-2xl p-3.5 text-white">
            <div className="text-slate-300">Saved Trades</div>
            <div className="mt-1 text-lg font-bold font-mono tabular-nums text-white">
              {counts ? counts.trades : '—'}
            </div>
            <div className="text-[11px] text-[#E6B949] font-mono tabular-nums">
              {counts ? `${counts.orders} order legs` : 'Synced'}
            </div>
          </div>

          <div className="bg-slate-900/[0.03] border border-slate-900/10 rounded-2xl p-3.5">
            <div className="text-slate-500">Screenshots & Deposits</div>
            <div className="mt-1 text-lg font-bold font-mono tabular-nums text-[#111318]">
              {counts ? `${counts.screenshots} / ${counts.deposits}` : '—'}
            </div>
            <div className="text-[11px] text-slate-500">Stored in Sheet tabs</div>
          </div>

          <div className="bg-slate-900/[0.03] border border-slate-900/10 rounded-2xl p-3.5">
            <div className="text-slate-500">Linked Google Account</div>
            <div className="mt-1 text-sm font-semibold text-[#111318] truncate">
              {googleEmail || 'Connected'}
            </div>
            <div className="text-[11px] text-slate-500 truncate">
              Linked by {sheetConfig.linkedBy}
            </div>
          </div>

          <div className="bg-slate-900/[0.03] border border-slate-900/10 rounded-2xl p-3.5 flex flex-col justify-between">
            <div>
              <div className="text-slate-500">Spreadsheet ID</div>
              <div className="mt-1 text-xs font-mono text-[#111318] truncate">
                {sheetConfig.spreadsheetId}
              </div>
            </div>
            <a
              href={sheetConfig.spreadsheetUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="mt-1 inline-flex items-center gap-1 text-[11px] font-semibold text-[#9F771B] hover:text-[#111318]"
            >
              <span>Open in Google Sheets</span>
              <ExternalLink className="w-3 h-3" />
            </a>
          </div>
        </div>

        {/* Share with Partner Gmail */}
        <form
          onSubmit={handleShareSheet}
          className="p-4 rounded-2xl bg-slate-900/[0.03] border border-slate-900/10 flex flex-col sm:flex-row sm:items-end gap-3"
        >
          <div className="flex-1">
            <label htmlFor="partner-share-email" className="block text-xs font-semibold text-[#111318] mb-1">
              Share Editor Access with {partnerId === 'Prayag' ? 'Raj' : 'Prayag'} (Gmail)
            </label>
            <input
              id="partner-share-email"
              type="email"
              value={partnerEmail}
              onChange={(e) => setPartnerEmail(e.target.value)}
              placeholder={partnerId === 'Prayag' ? 'raj@gmail.com' : 'prayag@gmail.com'}
              className="w-full h-11 px-3.5 py-2 text-xs bg-white border border-slate-900/15 rounded-2xl text-[#111318] focus:outline-none focus:border-[#111318]"
            />
          </div>
          <motion.button
            whileTap={{ scale: 0.98 }}
            type="submit"
            disabled={sharing}
            className="min-h-[44px] inline-flex items-center justify-center gap-2 px-5 py-2.5 bg-[#111318] hover:bg-[#242934] disabled:bg-slate-400 text-white text-xs font-semibold rounded-full cursor-pointer whitespace-nowrap shadow-xs"
          >
            <Share2 className="w-4 h-4 text-[#E6B949]" />
            <span>{sharing ? 'Sharing...' : 'Share Sheet Access'}</span>
          </motion.button>
        </form>

        {/* Action Buttons */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          <a
            href={sheetConfig.spreadsheetUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="min-h-[46px] inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-[#111318] hover:bg-[#242934] text-white text-xs font-semibold rounded-full shadow-xs"
          >
            <ExternalLink className="w-4 h-4 shrink-0 text-[#E6B949]" />
            <span>Open Live Google Sheet</span>
          </a>

          <motion.button
            whileTap={{ scale: 0.98 }}
            type="button"
            onClick={handleDownloadBackup}
            className="min-h-[46px] inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-white/90 hover:bg-white text-[#111318] border border-slate-900/15 text-xs font-semibold rounded-full cursor-pointer"
          >
            <Download className="w-4 h-4 shrink-0 text-[#C89D3C]" />
            <span>Download JSON Backup</span>
          </motion.button>

          <motion.button
            whileTap={{ scale: 0.98 }}
            type="button"
            onClick={() => restoreInputRef.current?.click()}
            disabled={workingBackup}
            className="min-h-[46px] inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-white/90 hover:bg-white text-[#111318] border border-slate-900/15 text-xs font-semibold rounded-full cursor-pointer"
          >
            <Upload className="w-4 h-4 shrink-0 text-[#111318]" />
            <span>Merge Backup into Sheet</span>
          </motion.button>

          <motion.button
            whileTap={{ scale: 0.98 }}
            type="button"
            onClick={onExportExcel}
            className="min-h-[46px] inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-white/90 hover:bg-white text-[#111318] border border-slate-900/15 text-xs font-semibold rounded-full cursor-pointer"
          >
            <FileSpreadsheet className="w-4 h-4 shrink-0 text-emerald-700" />
            <span>Download Excel (.xlsx)</span>
          </motion.button>

          <input
            ref={restoreInputRef}
            type="file"
            accept=".json,application/json"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void handleRestoreBackupFile(f);
              e.target.value = '';
            }}
          />
        </div>

        <div className="flex flex-wrap items-center justify-between gap-2 pt-3 border-t border-slate-900/8 text-xs text-slate-500">
          <span>
            Linked spreadsheet: <span className="font-semibold text-[#111318]">{sheetConfig.spreadsheetTitle}</span>
          </span>
          <button
            type="button"
            disabled={workingBackup}
            onClick={handleDisconnectSheet}
            className="min-h-[36px] inline-flex items-center gap-1 text-slate-500 hover:text-rose-700 cursor-pointer"
          >
            <Unlink className="w-3.5 h-3.5" />
            <span>Switch / Disconnect Google Sheet</span>
          </button>
        </div>

        {sheetErr && (
          <div className="p-3 text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-2xl">
            {sheetErr}
          </div>
        )}
        {sheetMsg && (
          <div className="p-3 text-xs text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-2xl flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
            <span>{sheetMsg}</span>
          </div>
        )}
      </motion.section>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 items-start">
        {/* Angel One & NSE Charge Rates */}
        <motion.form
          initial={{ opacity: 0, y: 12 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.32, delay: 0.05, ease: [0.16, 1, 0.3, 1] }}
          onSubmit={handleSaveRates}
          className="gallery-card rounded-3xl p-4 sm:p-6 space-y-4"
        >
          <div className="flex items-start justify-between gap-2">
            <div>
              <h2 className="font-display text-base font-bold text-[#111318] flex items-center gap-2">
                <span className="w-7 h-7 rounded-xl bg-[#111318] text-[#E6B949] flex items-center justify-center shrink-0">
                  <Sliders className="w-3.5 h-3.5" />
                </span>
                <span>Angel One & Exchange Charge Rates</span>
              </h2>
              <p className="text-xs text-slate-500 mt-1.5">
                Configurable official NSE Equity/Index Options charge structure.
              </p>
            </div>
            <button
              type="button"
              onClick={() => {
                setRateInputs(ratesToInputs(DEFAULT_CHARGE_RATES));
                setRatesErr(null);
                setRatesMsg(null);
              }}
              className="min-h-[38px] px-3 inline-flex items-center gap-1.5 text-xs font-semibold text-[#111318] bg-slate-900/5 hover:bg-slate-900/10 rounded-full cursor-pointer whitespace-nowrap transition-colors"
            >
              <RotateCcw className="w-3.5 h-3.5 text-[#C89D3C]" />
              <span>Official Defaults</span>
            </button>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
            {(
              [
                ['brokeragePerOrder', 'Brokerage / Executed Order (₹)'],
                ['sttSellPct', 'STT on Sell Premium (%)'],
                ['exchangeTxnPct', 'NSE Exchange Txn Charge (%)'],
                ['ipftPct', 'NSE IPFT Charge — Buy & Sell (%)'],
                ['sebiPct', 'SEBI Turnover Charge (%)'],
                ['stampDutyBuyPct', 'Stamp Duty on Buy (%)'],
                ['gstPct', 'GST on Brokerage+Exch+IPFT+SEBI (%)'],
              ] as const
            ).map(([field, label]) => (
              <div key={field}>
                <label className="block text-slate-600 font-medium mb-1">{label}</label>
                <input
                  type="text"
                  inputMode="decimal"
                  value={rateInputs[field]}
                  onChange={(e) => {
                    const val = e.target.value;
                    if (val === '' || /^\d*\.?\d*$/.test(val)) {
                      setRateInputs((prev) => ({ ...prev, [field]: val }));
                    }
                  }}
                  placeholder="0"
                  className="w-full h-11 px-3.5 py-2 bg-white/90 border border-slate-900/15 rounded-2xl font-mono text-[#111318] focus:outline-none focus:ring-2 focus:ring-slate-900/15 focus:border-[#111318]"
                />
              </div>
            ))}
          </div>

          <label className="flex items-center gap-2.5 text-xs text-slate-700 cursor-pointer pt-1">
            <input
              type="checkbox"
              checked={recalcExisting}
              onChange={(e) => setRecalcExisting(e.target.checked)}
              className="w-4 h-4 rounded border-slate-300 accent-[#111318]"
            />
            <span>Recalculate existing Estimated trades (Contract Note trades remain untouched)</span>
          </label>

          {ratesErr && (
            <div className="p-3 text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-2xl">
              {ratesErr}
            </div>
          )}
          {ratesMsg && (
            <div className="p-3 text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-2xl flex items-center gap-1.5">
              <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
              <span>{ratesMsg}</span>
            </div>
          )}

          <motion.button
            whileTap={{ scale: 0.98 }}
            type="submit"
            disabled={savingRates}
            className="w-full sm:w-auto min-h-[44px] px-5 py-2.5 bg-[#111318] hover:bg-[#242934] disabled:bg-slate-400 text-white text-xs font-semibold rounded-full cursor-pointer shadow-xs"
          >
            {savingRates ? 'Saving Rates...' : 'Save Charge Rates'}
          </motion.button>
        </motion.form>

        {/* Partner Login Password Settings */}
        <motion.form
          initial={{ opacity: 0, y: 12 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.32, delay: 0.08, ease: [0.16, 1, 0.3, 1] }}
          onSubmit={handleUpdatePassword}
          className="gallery-card rounded-3xl p-4 sm:p-6 space-y-4"
        >
          <div>
            <h2 className="font-display text-base font-bold text-[#111318] flex items-center gap-2">
              <span className="w-7 h-7 rounded-xl bg-[#111318] text-[#E6B949] flex items-center justify-center shrink-0">
                <KeyRound className="w-3.5 h-3.5" />
              </span>
              <span>{partnerId}&apos;s Login Password</span>
            </h2>
            <p className="text-xs text-slate-500 mt-1.5">
              Update the password for {partnerId}. Changing your password only updates {partnerId}
              &apos;s login; all trades, funds, and entries remain shared in the Google Sheet with{' '}
              {partnerId === 'Prayag' ? 'Raj' : 'Prayag'}.
            </p>
          </div>

          {envPasswordConfigured ? (
            <div className="p-4 text-xs bg-slate-900/[0.04] border border-slate-900/12 rounded-2xl text-[#111318]">
              {partnerId}&apos;s password is currently configured via environment secrets (
              <span className="font-mono font-semibold">
                {partnerId === 'Prayag' ? 'PRAYAG_PASSWORD' : 'RAJ_PASSWORD'}
              </span>
              ). Update it in Settings → Secrets to change it.
            </div>
          ) : (
            <>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Current Password ({partnerId})
                </label>
                <input
                  type="password"
                  value={currentPassword}
                  onChange={(e) => setCurrentPassword(e.target.value)}
                  className="w-full h-11 px-3.5 py-2 text-xs bg-white/90 border border-slate-900/15 rounded-2xl text-[#111318] focus:outline-none focus:ring-2 focus:ring-slate-900/15 focus:border-[#111318]"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  New Password for {partnerId}
                </label>
                <input
                  type="password"
                  value={newPassword}
                  onChange={(e) => setNewPassword(e.target.value)}
                  className="w-full h-11 px-3.5 py-2 text-xs bg-white/90 border border-slate-900/15 rounded-2xl text-[#111318] focus:outline-none focus:ring-2 focus:ring-slate-900/15 focus:border-[#111318]"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Confirm New Password
                </label>
                <input
                  type="password"
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                  className="w-full h-11 px-3.5 py-2 text-xs bg-white/90 border border-slate-900/15 rounded-2xl text-[#111318] focus:outline-none focus:ring-2 focus:ring-slate-900/15 focus:border-[#111318]"
                />
              </div>

              {pwdErr && (
                <div className="p-3 text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-2xl">
                  {pwdErr}
                </div>
              )}
              {pwdMsg && (
                <div className="p-3 text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-2xl flex items-center gap-1.5">
                  <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                  <span>{pwdMsg}</span>
                </div>
              )}

              <motion.button
                whileTap={{ scale: 0.98 }}
                type="submit"
                disabled={savingPwd}
                className="w-full sm:w-auto min-h-[44px] px-5 py-2.5 bg-[#111318] hover:bg-[#242934] disabled:bg-slate-400 text-white text-xs font-semibold rounded-full cursor-pointer shadow-xs"
              >
                {savingPwd ? 'Updating...' : `Update ${partnerId}'s Password`}
              </motion.button>
            </>
          )}
        </motion.form>
      </div>

      {/* Shared Audit History */}
      <motion.section
        initial={{ opacity: 0, y: 12 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.34, delay: 0.1, ease: [0.16, 1, 0.3, 1] }}
        className="gallery-card rounded-3xl overflow-hidden"
      >
        <div className="px-4 sm:px-5 py-4 border-b border-slate-900/8 flex items-center justify-between gap-3">
          <div>
            <h3 className="font-display text-sm font-bold text-[#111318]">
              Shared Google Sheet Audit History
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Live change log stored inside the Audit_Logs tab for both Prayag and Raj
            </p>
          </div>
          <button
            type="button"
            onClick={() => void loadStatusAndLogs()}
            className="min-h-[38px] px-3.5 text-xs font-semibold text-[#111318] bg-slate-900/5 hover:bg-slate-900/10 rounded-full cursor-pointer transition-colors"
          >
            Refresh
          </button>
        </div>

        {loadingLogs ? (
          <div className="p-6 text-xs text-slate-500">Loading audit history from Google Sheet...</div>
        ) : auditLogs.length === 0 ? (
          <div className="p-6 text-xs text-slate-500">No audit entries yet.</div>
        ) : (
          <div className="overflow-x-auto max-h-96 overflow-y-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="sticky top-0 bg-white/95 backdrop-blur-xs border-b border-slate-900/10 text-slate-600 font-semibold">
                <tr>
                  <th className="py-2.5 px-4">Timestamp</th>
                  <th className="py-2.5 px-4">Partner</th>
                  <th className="py-2.5 px-4">Action</th>
                  <th className="py-2.5 px-4">Summary</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-900/8">
                {auditLogs.map((log) => (
                  <tr key={log.id} className="hover:bg-slate-900/[0.025]">
                    <td className="py-2.5 px-4 font-mono tabular-nums text-slate-500 whitespace-nowrap">
                      {formatDateTime(log.createdAt)}
                    </td>
                    <td className="py-2.5 px-4 font-semibold text-[#111318]">{log.actor}</td>
                    <td className="py-2.5 px-4 font-mono text-[#9B7722] font-semibold">
                      {log.action}
                    </td>
                    <td className="py-2.5 px-4 text-slate-700">{log.summary}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </motion.section>
    </div>
  );
};
