import type {
  AuditLogRecord,
  ChargeRates,
  DashboardSummary,
  DepositRecord,
  ExtractedOrderLeg,
  FundEntryType,
  PartnerId,
  SettlementPreview,
  SettlementRecord,
  TradeGroupRecord,
} from '../../shared/types.ts';
import {
  googleSheetsLedger,
  type SheetBackupPayload,
} from './googleSheetsLedger.ts';
import { confirmSheetAction } from './sheetConfirm.ts';

let memoryToken: string | null = null;
let memoryPartnerId: PartnerId | null = null;
let memorySheetConfig: SharedSheetConfig | null = null;

const SYNC_CHANNEL_NAME = 'partnerdesk-sheet-sync';

function currentActor(): PartnerId {
  return memoryPartnerId ?? 'Prayag';
}

function requireSheetConfig(): SharedSheetConfig {
  if (!memorySheetConfig?.spreadsheetId) {
    throw new Error('Please connect your shared Google Sheet before accessing the ledger.');
  }
  return memorySheetConfig;
}

function notifySharedLedgerChanged(): void {
  if (typeof BroadcastChannel === 'undefined') return;
  try {
    const ch = new BroadcastChannel(SYNC_CHANNEL_NAME);
    ch.postMessage({ type: 'ledger-updated', ts: Date.now() });
    ch.close();
  } catch {
    // Ignore unsupported BroadcastChannel contexts
  }
}

export function subscribeSharedLedgerChanges(onUpdate: () => void): () => void {
  if (typeof BroadcastChannel === 'undefined') return () => {};
  try {
    const ch = new BroadcastChannel(SYNC_CHANNEL_NAME);
    const handler = () => onUpdate();
    ch.addEventListener('message', handler);
    return () => {
      ch.removeEventListener('message', handler);
      ch.close();
    };
  } catch {
    return () => {};
  }
}

export function setAuthToken(token: string | null, partnerId?: PartnerId | null): void {
  memoryToken = token;
  memoryPartnerId = token ? (partnerId ?? memoryPartnerId) : null;
  if (!token) {
    memorySheetConfig = null;
  }
}

export function getAuthToken(): string | null {
  return memoryToken;
}

export function setActiveSheetConfig(config: SharedSheetConfig | null): void {
  memorySheetConfig = config;
}

export function getActiveSheetConfig(): SharedSheetConfig | null {
  return memorySheetConfig;
}

async function apiFetch<T>(url: string, options: RequestInit = {}): Promise<T> {
  const headers: Record<string, string> = {
    ...(options.headers as Record<string, string> | undefined),
  };
  if (options.body && !headers['Content-Type']) {
    headers['Content-Type'] = 'application/json';
  }
  if (memoryToken) {
    headers.Authorization = `Bearer ${memoryToken}`;
  }

  const res = await fetch(url, { ...options, headers });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || `Request failed (${res.status})`);
  }
  return data as T;
}

export interface SharedSheetConfig {
  spreadsheetId: string;
  spreadsheetUrl: string;
  spreadsheetTitle: string;
  linkedBy: PartnerId;
  linkedAt: string;
}

export interface AuthStatusResponse {
  needsInitialSetup: boolean;
  envPasswordConfigured: boolean;
  prayagEnvConfigured?: boolean;
  rajEnvConfigured?: boolean;
  persistent?: boolean;
}

export interface AuthSessionResponse {
  token: string;
  partnerId: PartnerId;
}

export interface DashboardResponse {
  summary: DashboardSummary;
  settlementPreview: SettlementPreview;
  recentTrades: TradeGroupRecord[];
  openPositions: TradeGroupRecord[];
  recentDeposits: DepositRecord[];
  recentSettlements: SettlementRecord[];
}

export interface OcrResponse {
  orders: ExtractedOrderLeg[];
  tradingDateFound: boolean;
  notes: string;
}

export const api = {
  getAuthStatus: () => apiFetch<AuthStatusResponse>('/api/auth/status'),

  setupPassword: (partnerId: PartnerId, password: string) =>
    apiFetch<AuthSessionResponse>('/api/auth/setup', {
      method: 'POST',
      body: JSON.stringify({ partnerId, password }),
    }),

  login: (partnerId: PartnerId, password: string) =>
    apiFetch<AuthSessionResponse>('/api/auth/login', {
      method: 'POST',
      body: JSON.stringify({ partnerId, password }),
    }),

  updatePassword: (currentPassword: string, newPassword: string) =>
    apiFetch<{ ok: boolean }>('/api/auth/password', {
      method: 'POST',
      body: JSON.stringify({ currentPassword, newPassword }),
    }),

  // --- Shared Google Sheet Pointer (shared identically between Prayag & Raj) ---
  getSharedSheetConfig: async (): Promise<SharedSheetConfig | null> => {
    const res = await apiFetch<{ config: SharedSheetConfig | null }>('/api/sheets/config');
    memorySheetConfig = res.config;
    return res.config;
  },

  saveSharedSheetConfig: async (payload: {
    spreadsheetId: string;
    spreadsheetUrl?: string;
    spreadsheetTitle?: string;
  }): Promise<SharedSheetConfig> => {
    const res = await apiFetch<{ config: SharedSheetConfig }>('/api/sheets/config', {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
    memorySheetConfig = res.config;
    notifySharedLedgerChanged();
    return res.config;
  },

  clearSharedSheetConfig: async (): Promise<void> => {
    await apiFetch<{ ok: boolean }>('/api/sheets/config', { method: 'DELETE' });
    memorySheetConfig = null;
    notifySharedLedgerChanged();
  },

  createAndLinkSharedSheet: async (params?: {
    title?: string;
    partnerEmail?: string;
  }): Promise<SharedSheetConfig> => {
    const sheetTitle = (params?.title || 'PartnerDesk — Options Trading Ledger').trim();
    const partnerEmail = params?.partnerEmail?.trim() || '';
    const details = [
      `Spreadsheet: "${sheetTitle}"`,
      'Creates tabs: Summary, Trade_Groups, Order_Legs, Deposits_Ledger, Charge_Rates, Audit_Logs, Screenshots, Group_Screenshots',
      'Links this Google Sheet for both Prayag and Raj logins.',
    ];
    if (partnerEmail) {
      details.push(`Grants Editor access to partner email: ${partnerEmail}`);
    }

    await confirmSheetAction({
      title: 'Create New Shared Google Sheet?',
      description:
        'PartnerDesk will create a new Google Sheet in your Google Drive and store all partnership trades, deposits, and settings inside it.',
      confirmLabel: 'Create Google Sheet',
      details,
    });

    const created = await googleSheetsLedger.createPartnerDeskSpreadsheet(sheetTitle);
    if (partnerEmail) {
      await googleSheetsLedger.shareSpreadsheetWithPartner(created.spreadsheetId, partnerEmail);
    }
    return api.saveSharedSheetConfig(created);
  },

  connectExistingSharedSheet: async (rawUrlOrId: string): Promise<SharedSheetConfig> => {
    await confirmSheetAction({
      title: 'Connect Existing Google Sheet?',
      description:
        'PartnerDesk will open the specified Google Sheet and ensure all required ledger tabs and column headers exist before linking it for both Prayag and Raj.',
      confirmLabel: 'Connect Google Sheet',
      details: [`Target Google Sheet: ${rawUrlOrId.trim()}`],
    });

    const connected = await googleSheetsLedger.connectExistingSpreadsheet(rawUrlOrId);
    return api.saveSharedSheetConfig(connected);
  },

  shareActiveSheetWithPartner: async (email: string): Promise<void> => {
    const cfg = requireSheetConfig();
    const cleanEmail = email.trim();
    await confirmSheetAction({
      title: 'Share Google Sheet with Partner?',
      description:
        'This will grant Editor access in Google Drive so your partner can read and update the same Google Sheet from their device.',
      confirmLabel: 'Share Editor Access',
      details: [
        `Spreadsheet: ${cfg.spreadsheetTitle}`,
        `Grant Editor role to: ${cleanEmail}`,
      ],
    });
    await googleSheetsLedger.shareSpreadsheetWithPartner(cfg.spreadsheetId, cleanEmail);
  },

  // --- Shared Google Sheet Ledger Operations (Same live spreadsheet for Prayag & Raj) ---
  getDashboard: async (): Promise<DashboardResponse> => {
    const cfg = requireSheetConfig();
    return googleSheetsLedger.getDashboard(cfg.spreadsheetId);
  },

  extractScreenshot: async (imageDataUrl: string): Promise<OcrResponse> => {
    const cfg = requireSheetConfig();
    const duplicate = await googleSheetsLedger.checkDuplicateScreenshot(
      cfg.spreadsheetId,
      imageDataUrl,
    );
    if (duplicate) {
      throw new Error(
        `Duplicate screenshot detected (already saved in Google Sheet on ${new Date(
          duplicate.uploadedAt,
        ).toLocaleString()} by ${duplicate.uploadedBy}).`,
      );
    }
    return apiFetch<OcrResponse>('/api/trades/ocr', {
      method: 'POST',
      body: JSON.stringify({ imageDataUrl }),
    });
  },

  listTrades: async (): Promise<{ trades: TradeGroupRecord[]; rates: ChargeRates }> => {
    const cfg = requireSheetConfig();
    return googleSheetsLedger.listTrades(cfg.spreadsheetId);
  },

  saveOrders: async (payload: {
    screenshotDataUrl?: string | null;
    originalFilename?: string | null;
    orders: ExtractedOrderLeg[];
  }) => {
    const cfg = requireSheetConfig();
    const symbols = Array.from(
      new Set(
        payload.orders.map(
          (o) => `${o.stockName.trim().toUpperCase()} ${o.strikePrice} ${o.optionType}`,
        ),
      ),
    );
    await confirmSheetAction({
      title: 'Save Trade to Shared Google Sheet?',
      description: `This will write ${payload.orders.length} order leg(s) into "${cfg.spreadsheetTitle}" so both Prayag and Raj see the updated trade and Summary.`,
      confirmLabel: 'Save to Google Sheet',
      details: [
        `Contracts: ${symbols.join(', ')}`,
        payload.screenshotDataUrl
          ? `Screenshot: ${payload.originalFilename || 'Attached image'} (archived in Screenshots tab)`
          : 'Screenshot: None (manual order entry)',
        `Logged by: ${currentActor()}`,
      ],
    });

    const res = await googleSheetsLedger.saveOrders(cfg.spreadsheetId, {
      actor: currentActor(),
      screenshotDataUrl: payload.screenshotDataUrl,
      originalFilename: payload.originalFilename,
      orders: payload.orders,
    });
    notifySharedLedgerChanged();
    return res;
  },

  updateTrade: async (id: number, payload: Record<string, unknown>) => {
    const cfg = requireSheetConfig();
    const isContractNote = Boolean(payload.contractNote);
    await confirmSheetAction({
      title: isContractNote
        ? 'Save Contract Note Charges to Google Sheet?'
        : 'Save Trade Edits to Google Sheet?',
      description: `This will update Trade #${id} and recalculate Summary totals in "${cfg.spreadsheetTitle}".`,
      confirmLabel: 'Confirm Update',
      details: [
        `Spreadsheet: ${cfg.spreadsheetTitle}`,
        `Updated by: ${currentActor()}`,
      ],
    });

    const trade = await googleSheetsLedger.updateTrade(
      cfg.spreadsheetId,
      id,
      currentActor(),
      payload,
    );
    notifySharedLedgerChanged();
    return { trade };
  },

  deleteTrade: async (id: number, label?: string) => {
    const cfg = requireSheetConfig();
    await confirmSheetAction({
      title: 'Delete Trade from Shared Google Sheet?',
      description: `This will permanently remove ${label || `Trade #${id}`} and its order legs from "${cfg.spreadsheetTitle}" for both Prayag and Raj.`,
      confirmLabel: 'Delete Permanently',
      tone: 'danger',
      details: [
        `Spreadsheet: ${cfg.spreadsheetTitle}`,
        `Deleted by: ${currentActor()}`,
      ],
    });

    await googleSheetsLedger.deleteTrade(cfg.spreadsheetId, id, currentActor());
    notifySharedLedgerChanged();
    return { ok: true };
  },

  listDeposits: async (): Promise<{ deposits: DepositRecord[]; summary: DashboardSummary }> => {
    const cfg = requireSheetConfig();
    return googleSheetsLedger.listDeposits(cfg.spreadsheetId);
  },

  addDeposit: async (payload: {
    partner: PartnerId;
    entryType?: FundEntryType;
    amount: number;
    depositDate: string;
    referenceNote?: string;
  }) => {
    const cfg = requireSheetConfig();
    const isWithdrawal = payload.entryType === 'WITHDRAWAL';
    await confirmSheetAction({
      title: isWithdrawal
        ? 'Record Fund Withdrawal in Google Sheet?'
        : 'Record Capital Deposit in Google Sheet?',
      description: `This will add a ${isWithdrawal ? 'withdrawal' : 'deposit'} entry to "${cfg.spreadsheetTitle}" and update live ownership balances.`,
      confirmLabel: isWithdrawal ? 'Confirm Withdrawal' : 'Confirm Deposit',
      tone: isWithdrawal ? 'danger' : 'primary',
      details: [
        `Partner: ${payload.partner}`,
        `Amount: ₹${Number(payload.amount).toLocaleString('en-IN')}`,
        `Date: ${payload.depositDate}`,
        ...(payload.referenceNote ? [`Note: ${payload.referenceNote}`] : []),
      ],
    });

    const deposit = await googleSheetsLedger.addDeposit(cfg.spreadsheetId, {
      ...payload,
      actor: currentActor(),
    });
    notifySharedLedgerChanged();
    return { deposit };
  },

  deleteDeposit: async (id: number, label?: string) => {
    const cfg = requireSheetConfig();
    await confirmSheetAction({
      title: 'Delete Fund Entry from Google Sheet?',
      description: `This will permanently delete ${label || `fund entry #${id}`} from Deposits_Ledger in "${cfg.spreadsheetTitle}".`,
      confirmLabel: 'Delete Entry',
      tone: 'danger',
      details: [
        `Spreadsheet: ${cfg.spreadsheetTitle}`,
        `Removed by: ${currentActor()}`,
      ],
    });

    await googleSheetsLedger.deleteDeposit(cfg.spreadsheetId, id, currentActor());
    notifySharedLedgerChanged();
    return { ok: true };
  },

  getRates: async (): Promise<{ rates: ChargeRates }> => {
    const cfg = requireSheetConfig();
    const rates = await googleSheetsLedger.getRates(cfg.spreadsheetId);
    return { rates };
  },

  updateRates: async (rates: ChargeRates & { recalculateExisting?: boolean }) => {
    const cfg = requireSheetConfig();
    await confirmSheetAction({
      title: 'Update Charge Rates in Google Sheet?',
      description: `This will save updated Angel One / NSE charge rates to "${cfg.spreadsheetTitle}" for both Prayag and Raj.`,
      confirmLabel: 'Save Charge Rates',
      details: [
        `Brokerage per order: ₹${rates.brokeragePerOrder}`,
        rates.recalculateExisting
          ? 'Also recalculating all existing Estimated trades.'
          : 'Applies to future trades (existing trades unchanged).',
      ],
    });

    const updated = await googleSheetsLedger.updateRates(cfg.spreadsheetId, rates, currentActor());
    notifySharedLedgerChanged();
    return { rates: updated };
  },

  listAuditLogs: async (): Promise<{ logs: AuditLogRecord[] }> => {
    const cfg = requireSheetConfig();
    const logs = await googleSheetsLedger.listAuditLogs(cfg.spreadsheetId);
    return { logs };
  },

  getBackupPayload: async (): Promise<SheetBackupPayload> => {
    const cfg = requireSheetConfig();
    return googleSheetsLedger.exportBackupPayload(cfg.spreadsheetId);
  },

  downloadBackupJson: (payload: SheetBackupPayload): void => {
    googleSheetsLedger.downloadJsonFromPayload(payload);
  },

  mergeBackupPayload: async (
    payload: Partial<SheetBackupPayload>,
  ): Promise<{ addedTrades: number; addedDeposits: number; addedScreenshots: number }> => {
    const cfg = requireSheetConfig();
    const tradeCount = Array.isArray(payload.tradeGroups) ? payload.tradeGroups.length : 0;
    const depCount = Array.isArray(payload.deposits) ? payload.deposits.length : 0;
    await confirmSheetAction({
      title: 'Merge JSON Backup into Google Sheet?',
      description: `This will merge non-duplicate records from the selected backup file into "${cfg.spreadsheetTitle}" for both Prayag and Raj.`,
      confirmLabel: 'Merge into Google Sheet',
      details: [
        `Backup contains ${tradeCount} trade group(s) and ${depCount} fund entry(s).`,
        'Existing records with matching keys are preserved without duplication.',
      ],
    });

    const res = await googleSheetsLedger.importMergeBackupPayload(
      cfg.spreadsheetId,
      payload,
      currentActor(),
    );
    notifySharedLedgerChanged();
    return res;
  },
};
