import React, { useCallback, useEffect, useRef, useState } from 'react';
import { motion } from 'motion/react';
import {
  CheckCircle2,
  Download,
  ExternalLink,
  FileSpreadsheet,
  KeyRound,
  Moon,
  RefreshCw,
  RotateCcw,
  Sliders,
  Sun,
  Upload,
} from 'lucide-react';
import { DEFAULT_CHARGE_RATES } from '../../shared/calculations.ts';
import type {
  AuditLogRecord,
  ChargeRates,
  GoogleSyncStatus,
  PartnerId,
} from '../../shared/types.ts';
import { api } from '../lib/api.ts';
import { formatDateTime } from '../lib/format.ts';

interface SettingsViewProps {
  partnerId: PartnerId;
  darkTheme: boolean;
  onToggleDarkTheme: () => void;
  syncStatus: GoogleSyncStatus | null;
  rates: ChargeRates;
  envPasswordConfigured: boolean;
  onRatesUpdated: (rates: ChargeRates) => void;
  onSyncStatusUpdated: (status: GoogleSyncStatus) => void;
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
  darkTheme,
  onToggleDarkTheme,
  syncStatus,
  rates,
  envPasswordConfigured,
  onRatesUpdated,
  onSyncStatusUpdated,
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
  const [syncingNow, setSyncingNow] = useState(false);
  const [workingBackup, setWorkingBackup] = useState(false);
  const [ledgerMsg, setLedgerMsg] = useState<string | null>(null);
  const [ledgerErr, setLedgerErr] = useState<string | null>(null);

  const [rateInputs, setRateInputs] = useState<Record<keyof ChargeRates, string>>(() =>
    ratesToInputs(rates),
  );
  const [recalcExisting, setRecalcExisting] = useState(false);
  const [savingRates, setSavingRates] = useState(false);
  const [ratesMsg, setRatesMsg] = useState<string | null>(null);
  const [ratesErr, setRatesErr] = useState<string | null>(null);

  const [editingPassword, setEditingPassword] = useState(false);
  const [currentPassword, setCurrentPassword] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [savingPwd, setSavingPwd] = useState(false);
  const [pwdMsg, setPwdMsg] = useState<string | null>(null);
  const [pwdErr, setPwdErr] = useState<string | null>(null);

  const [auditLogs, setAuditLogs] = useState<AuditLogRecord[]>([]);
  const [loadingLogs, setLoadingLogs] = useState(false);

  useEffect(() => {
    setRateInputs(ratesToInputs(rates));
  }, [rates]);

  const loadStatusAndLogs = useCallback(async () => {
    setLoadingLogs(true);
    try {
      const [backup, logsRes, syncRes] = await Promise.all([
        api.getBackupPayload(),
        api.listAuditLogs(),
        api.getSyncStatus().catch(() => null),
      ]);
      setCounts({
        trades: backup.tradeGroups.length,
        orders: backup.orders.length,
        screenshots: backup.uploads.length,
        deposits: backup.deposits.length,
      });
      setAuditLogs(logsRes.logs);
      if (syncRes?.status) {
        onSyncStatusUpdated(syncRes.status);
      }
    } catch (err) {
      setLedgerErr(err instanceof Error ? err.message : 'Unable to load shared ledger details.');
    } finally {
      setLoadingLogs(false);
    }
  }, [onSyncStatusUpdated]);

  useEffect(() => {
    void loadStatusAndLogs();
  }, [loadStatusAndLogs]);

  const handleSyncNow = async () => {
    setLedgerMsg(null);
    setLedgerErr(null);
    setSyncingNow(true);
    try {
      const res = await api.triggerSyncNow();
      onSyncStatusUpdated(res.status);
      if (res.status.lastError) {
        setLedgerErr(res.status.lastError);
      } else if (res.status.configured) {
        setLedgerMsg('Master Google Sheet synced with the latest shared cloud data.');
      } else {
        setLedgerMsg('Shared cloud ledger is up to date.');
      }
    } catch (err) {
      setLedgerErr(err instanceof Error ? err.message : 'Failed to sync Google Sheet.');
    } finally {
      setSyncingNow(false);
    }
  };

  const handleDownloadBackup = async () => {
    setLedgerMsg(null);
    setLedgerErr(null);
    try {
      const payload = await api.getBackupPayload();
      api.downloadBackupJson(payload);
      setLedgerMsg('Downloaded PartnerDesk JSON backup from the shared cloud ledger.');
    } catch (err) {
      setLedgerErr(err instanceof Error ? err.message : 'Failed to export backup.');
    }
  };

  const handleRestoreBackupFile = async (file: File) => {
    setLedgerMsg(null);
    setLedgerErr(null);
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
      setLedgerMsg(
        `Merged backup into shared ledger for both Prayag & Raj: +${res.addedTrades} trade(s), +${res.addedDeposits} deposit(s), +${res.addedScreenshots} screenshot(s).`,
      );
      onStorageChanged();
      await loadStatusAndLogs();
    } catch (err) {
      setLedgerErr(err instanceof Error ? err.message : 'Failed to import backup.');
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
      setRatesMsg('Charge rates updated in shared ledger for both Prayag & Raj.');
      void loadStatusAndLogs();
    } catch (err) {
      setRatesErr(err instanceof Error ? err.message : 'Failed to update rates.');
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
      setEditingPassword(false);
      setPwdMsg(
        `Login password updated for ${partnerId}. Both accounts continue sharing the exact same cloud ledger.`,
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
      {/* Appearance — Dark Theme Button */}
      <motion.section
        initial={{ opacity: 0, y: 8 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.25, ease: [0.16, 1, 0.3, 1] }}
        className="gallery-card rounded-3xl p-4 sm:p-5 flex flex-wrap items-center justify-between gap-3"
      >
        <div>
          <h2 className="font-display text-sm sm:text-base font-bold text-[#111318]">
            Appearance
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            {darkTheme ? 'Dark theme enabled' : 'Light theme enabled'}
          </p>
        </div>
        <motion.button
          whileTap={{ scale: 0.98 }}
          type="button"
          onClick={onToggleDarkTheme}
          aria-pressed={darkTheme}
          className="min-h-[44px] inline-flex items-center justify-center gap-2 px-5 py-2.5 bg-[#111318] hover:bg-[#242934] text-white text-xs font-semibold rounded-full cursor-pointer shadow-xs transition-colors"
        >
          {darkTheme ? (
            <Sun className="w-4 h-4 text-[#E6B949]" />
          ) : (
            <Moon className="w-4 h-4 text-[#E6B949]" />
          )}
          <span>{darkTheme ? 'Switch to Light Theme' : 'Dark Theme'}</span>
        </motion.button>
      </motion.section>

      {/* Shared Cloud Ledger & Automatic Google Sheet Sync Status */}
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
              <span>Shared Partnership Cloud & Google Sheet Sync</span>
            </h2>
            <p className="text-xs text-slate-500 mt-1.5 leading-relaxed">
              Every trade, screenshot, deposit, withdrawal, and edit is saved centrally for both
              Prayag and Raj and automatically mirrored to your master Google Sheet in the
              background.
            </p>
          </div>
          <span className="self-start inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-full bg-emerald-500/10 border border-emerald-600/20 text-xs font-semibold text-emerald-800 whitespace-nowrap">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-700" />
            <span>
              {syncStatus?.configured
                ? syncStatus.pendingSync
                  ? 'Google Sheet Syncing...'
                  : 'Live Google Sheet Sync'
                : 'Shared Cloud Active'}
            </span>
          </span>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
          <div className="gallery-card-dark rounded-2xl p-3.5 text-white">
            <div className="text-slate-300">Saved Trades</div>
            <div className="mt-1 text-lg font-bold font-mono tabular-nums text-white">
              {counts ? counts.trades : '—'}
            </div>
            <div className="text-[11px] text-[#E6B949] font-mono tabular-nums">
              {counts ? `${counts.orders} order legs` : 'Shared ledger'}
            </div>
          </div>

          <div className="bg-slate-900/[0.03] border border-slate-900/10 rounded-2xl p-3.5">
            <div className="text-slate-500">Screenshots & Funds</div>
            <div className="mt-1 text-lg font-bold font-mono tabular-nums text-[#111318]">
              {counts ? `${counts.screenshots} / ${counts.deposits}` : '—'}
            </div>
            <div className="text-[11px] text-slate-500">
              {syncStatus?.configured ? 'Synced with Drive & Sheet' : 'Stored in cloud DB'}
            </div>
          </div>

          <div className="bg-slate-900/[0.03] border border-slate-900/10 rounded-2xl p-3.5">
            <div className="text-slate-500">Master Workbook</div>
            <div className="mt-1 text-sm font-semibold text-[#111318] truncate">
              {syncStatus?.spreadsheetTitle ||
                (syncStatus?.configured ? 'Connected Google Sheet' : 'PartnerDesk Cloud')}
            </div>
            <div className="text-[11px] text-slate-500 truncate">
              {syncStatus?.lastSyncedAt
                ? `Synced ${formatDateTime(syncStatus.lastSyncedAt)}`
                : 'Prayag & Raj live access'}
            </div>
          </div>

          <div className="bg-slate-900/[0.03] border border-slate-900/10 rounded-2xl p-3.5 flex flex-col justify-between">
            <div>
              <div className="text-slate-500">Background Sync</div>
              <div className="mt-1 text-xs font-semibold text-[#111318]">
                {syncStatus?.configured
                  ? syncStatus.lastError
                    ? 'Queued for retry'
                    : 'Server-Authorized'
                  : 'Local Server Storage'}
              </div>
            </div>
            {syncStatus?.spreadsheetUrl ? (
              <a
                href={syncStatus.spreadsheetUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="mt-1 inline-flex items-center gap-1 text-[11px] font-semibold text-[#9F771B] hover:text-[#111318]"
              >
                <span>Open Google Sheet</span>
                <ExternalLink className="w-3 h-3" />
              </a>
            ) : (
              <span className="mt-1 text-[11px] text-slate-500">No Google login needed</span>
            )}
          </div>
        </div>

        {/* Action Buttons */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {syncStatus?.spreadsheetUrl && (
            <a
              href={syncStatus.spreadsheetUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="min-h-[46px] inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-[#111318] hover:bg-[#242934] text-white text-xs font-semibold rounded-full shadow-xs"
            >
              <ExternalLink className="w-4 h-4 shrink-0 text-[#E6B949]" />
              <span>Open Live Google Sheet</span>
            </a>
          )}

          {syncStatus?.configured && (
            <motion.button
              whileTap={{ scale: 0.98 }}
              type="button"
              disabled={syncingNow}
              onClick={handleSyncNow}
              className="min-h-[46px] inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-white/90 hover:bg-white text-[#111318] border border-slate-900/15 text-xs font-semibold rounded-full cursor-pointer"
            >
              <RefreshCw className={`w-4 h-4 shrink-0 text-[#C89D3C] ${syncingNow ? 'animate-spin' : ''}`} />
              <span>{syncingNow ? 'Syncing Sheet...' : 'Sync Google Sheet Now'}</span>
            </motion.button>
          )}

          <motion.button
            whileTap={{ scale: 0.98 }}
            type="button"
            onClick={onExportExcel}
            className="min-h-[46px] inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-[#111318] hover:bg-[#242934] text-white text-xs font-semibold rounded-full cursor-pointer shadow-xs"
          >
            <FileSpreadsheet className="w-4 h-4 shrink-0 text-[#E6B949]" />
            <span>Download Excel (.xlsx)</span>
          </motion.button>

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
            <span>{workingBackup ? 'Merging...' : 'Merge JSON Backup'}</span>
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

        {(syncStatus?.lastError || syncStatus?.driveWarning) && (
          <div className="p-3 text-xs text-amber-800 bg-amber-50 border border-amber-200 rounded-2xl">
            {syncStatus.lastError || syncStatus.driveWarning}
          </div>
        )}

        {ledgerErr && (
          <div className="p-3 text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-2xl">
            {ledgerErr}
          </div>
        )}
        {ledgerMsg && (
          <div className="p-3 text-xs text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-2xl flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
            <span>{ledgerMsg}</span>
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
          {!editingPassword ? (
            <div className="space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div>
                  <h2 className="font-display text-base font-bold text-[#111318] flex items-center gap-2">
                    <span className="w-7 h-7 rounded-xl bg-[#111318] text-[#E6B949] flex items-center justify-center shrink-0">
                      <KeyRound className="w-3.5 h-3.5" />
                    </span>
                    <span>{partnerId}&apos;s Login Password</span>
                  </h2>
                  <p className="text-xs text-slate-500 mt-1.5">
                    Update {partnerId}&apos;s password. All shared data stays unchanged.
                  </p>
                </div>

                <motion.button
                  whileTap={{ scale: 0.98 }}
                  type="button"
                  onClick={() => {
                    setEditingPassword(true);
                    setPwdErr(null);
                    setPwdMsg(null);
                  }}
                  className="min-h-[44px] inline-flex items-center justify-center gap-2 px-5 py-2.5 bg-[#111318] hover:bg-[#242934] text-white text-xs font-semibold rounded-full cursor-pointer shadow-xs transition-colors"
                >
                  <KeyRound className="w-4 h-4 text-[#E6B949]" />
                  <span>Change Password</span>
                </motion.button>
              </div>

              {pwdMsg && (
                <div className="p-3 text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-2xl flex items-center gap-1.5">
                  <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                  <span>{pwdMsg}</span>
                </div>
              )}
            </div>
          ) : (
            <>
              <div>
                <h2 className="font-display text-base font-bold text-[#111318] flex items-center gap-2">
                  <span className="w-7 h-7 rounded-xl bg-[#111318] text-[#E6B949] flex items-center justify-center shrink-0">
                    <KeyRound className="w-3.5 h-3.5" />
                  </span>
                  <span>Change {partnerId}&apos;s Password</span>
                </h2>
                <p className="text-xs text-slate-500 mt-1.5">
                  Changing your password only updates {partnerId}&apos;s login; all trades and funds
                  remain shared with {partnerId === 'Prayag' ? 'Raj' : 'Prayag'}.
                </p>
              </div>

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

              <div className="flex flex-wrap items-center gap-2.5 pt-1">
                <motion.button
                  whileTap={{ scale: 0.98 }}
                  type="submit"
                  disabled={savingPwd}
                  className="min-h-[44px] px-5 py-2.5 bg-[#111318] hover:bg-[#242934] disabled:bg-slate-400 text-white text-xs font-semibold rounded-full cursor-pointer shadow-xs"
                >
                  {savingPwd ? 'Updating...' : 'Proceed'}
                </motion.button>
                <button
                  type="button"
                  disabled={savingPwd}
                  onClick={() => {
                    setEditingPassword(false);
                    setCurrentPassword('');
                    setNewPassword('');
                    setConfirmPassword('');
                    setPwdErr(null);
                  }}
                  className="min-h-[44px] px-4 py-2.5 text-xs font-medium text-slate-700 bg-white border border-slate-900/15 rounded-full hover:bg-slate-50 cursor-pointer"
                >
                  Cancel
                </button>
              </div>
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
              Shared Ledger Audit History
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Live change log shared identically for both Prayag and Raj
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
          <div className="p-6 text-xs text-slate-500">Loading audit history...</div>
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
