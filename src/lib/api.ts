import type {
  AuditLogRecord,
  ChargeRates,
  DashboardSummary,
  DepositRecord,
  ExtractedOrderLeg,
  FundEntryType,
  GoogleSyncStatus,
  PartnerId,
  SettlementPreview,
  SettlementRecord,
  TradeGroupRecord,
  TradeOrderRecord,
} from '../../shared/types.ts';

let memoryToken: string | null = null;

const SYNC_CHANNEL_NAME = 'partnerdesk-sheet-sync';

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

export function setAuthToken(token: string | null, _partnerId?: PartnerId | null): void {
  memoryToken = token;
}

export function getAuthToken(): string | null {
  return memoryToken;
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

export interface BackupPayload {
  format: string;
  folderName: string;
  exportedAt: string;
  chargeRates: ChargeRates;
  deposits: DepositRecord[];
  uploads: {
    id: number;
    imageHash: string;
    imageDataUrl: string;
    driveFileId?: string | null;
    driveUrl?: string | null;
    originalFilename: string | null;
    uploadedBy: PartnerId;
    uploadedAt: string;
  }[];
  tradeGroups: TradeGroupRecord[];
  orders: TradeOrderRecord[];
  auditLogs: AuditLogRecord[];
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

  getDashboard: () => apiFetch<DashboardResponse>('/api/dashboard'),

  extractScreenshot: (imageDataUrl: string) =>
    apiFetch<OcrResponse>('/api/trades/ocr', {
      method: 'POST',
      body: JSON.stringify({ imageDataUrl }),
    }),

  listTrades: () => apiFetch<{ trades: TradeGroupRecord[]; rates: ChargeRates }>('/api/trades'),

  saveOrders: async (payload: {
    screenshotDataUrl?: string | null;
    originalFilename?: string | null;
    orders: ExtractedOrderLeg[];
  }) => {
    const res = await apiFetch<{
      uploadId: number | null;
      uploadedAt: string;
      affectedTrades: TradeGroupRecord[];
    }>('/api/trades/save', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
    notifySharedLedgerChanged();
    return res;
  },

  updateTrade: async (id: number, payload: Record<string, unknown>) => {
    const res = await apiFetch<{ trade: TradeGroupRecord }>(`/api/trades/${id}`, {
      method: 'PATCH',
      body: JSON.stringify(payload),
    });
    notifySharedLedgerChanged();
    return res;
  },

  deleteTrade: async (id: number, _label?: string) => {
    const res = await apiFetch<{ ok: boolean }>(`/api/trades/${id}`, {
      method: 'DELETE',
    });
    notifySharedLedgerChanged();
    return res;
  },

  listDeposits: () =>
    apiFetch<{ deposits: DepositRecord[]; summary: DashboardSummary }>('/api/deposits'),

  addDeposit: async (payload: {
    partner: PartnerId;
    entryType?: FundEntryType;
    amount: number;
    depositDate: string;
    referenceNote?: string;
  }) => {
    const res = await apiFetch<{ deposit: DepositRecord }>('/api/deposits', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
    notifySharedLedgerChanged();
    return res;
  },

  deleteDeposit: async (id: number, _label?: string) => {
    const res = await apiFetch<{ ok: boolean }>(`/api/deposits/${id}`, {
      method: 'DELETE',
    });
    notifySharedLedgerChanged();
    return res;
  },

  getRates: () => apiFetch<{ rates: ChargeRates }>('/api/settings/rates'),

  updateRates: async (rates: ChargeRates & { recalculateExisting?: boolean }) => {
    const res = await apiFetch<{ rates: ChargeRates }>('/api/settings/rates', {
      method: 'PUT',
      body: JSON.stringify(rates),
    });
    notifySharedLedgerChanged();
    return res;
  },

  listAuditLogs: () => apiFetch<{ logs: AuditLogRecord[] }>('/api/audit-logs'),

  getSyncStatus: () => apiFetch<{ status: GoogleSyncStatus }>('/api/sync/status'),

  triggerSyncNow: () =>
    apiFetch<{ status: GoogleSyncStatus }>('/api/sync/now', { method: 'POST' }),

  getBackupPayload: () => apiFetch<BackupPayload>('/api/sync/backup'),

  downloadBackupJson: (payload: BackupPayload): void => {
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `PartnerDesk_Backup_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  },

  mergeBackupPayload: async (
    payload: Partial<BackupPayload>,
  ): Promise<{ addedTrades: number; addedDeposits: number; addedScreenshots: number }> => {
    const res = await apiFetch<{
      addedTrades: number;
      addedDeposits: number;
      addedScreenshots: number;
    }>('/api/sync/merge', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
    notifySharedLedgerChanged();
    return res;
  },
};
