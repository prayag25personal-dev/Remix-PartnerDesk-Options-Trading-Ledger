import crypto from 'node:crypto';
import type { GoogleSyncStatus } from '../shared/types.ts';
import {
  getFullLedgerSnapshotForSync,
  getSheetSyncMeta,
  listUploadsPendingDriveSync,
  updateSheetSyncMeta,
  updateUploadDriveRef,
} from './db.ts';

const SHEETS_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';
const DRIVE_FILES_BASE = 'https://www.googleapis.com/drive/v3/files';
const DRIVE_UPLOAD_BASE = 'https://www.googleapis.com/upload/drive/v3/files';
const OAUTH_TOKEN_URL = 'https://oauth2.googleapis.com/token';
const SERVICE_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
].join(' ');

const TABS = {
  SUMMARY: 'Summary',
  TRADE_GROUPS: 'Trade_Groups',
  ORDER_LEGS: 'Order_Legs',
  DEPOSITS: 'Deposits_Ledger',
  RATES: 'Charge_Rates',
  AUDIT_LOGS: 'Audit_Logs',
  SCREENSHOTS: 'Screenshots',
  GROUP_SCREENSHOTS: 'Group_Screenshots',
} as const;

const ALL_TAB_NAMES = [
  TABS.SUMMARY,
  TABS.TRADE_GROUPS,
  TABS.ORDER_LEGS,
  TABS.DEPOSITS,
  TABS.RATES,
  TABS.SCREENSHOTS,
  TABS.GROUP_SCREENSHOTS,
  TABS.AUDIT_LOGS,
] as const;

interface ServiceAccountConfig {
  clientEmail: string;
  privateKey: string;
  spreadsheetId: string;
  driveFolderId: string | null;
}

function extractSpreadsheetId(raw: string): string | null {
  const trimmed = raw.trim();
  if (!trimmed) return null;
  const match = /\/spreadsheets\/d\/([a-zA-Z0-9-_]{10,128})/.exec(trimmed);
  if (match) return match[1];
  if (/^[a-zA-Z0-9-_]{10,128}$/.test(trimmed)) return trimmed;
  return null;
}

function extractDriveFolderId(raw: string): string | null {
  const trimmed = raw.trim();
  if (!trimmed) return null;
  const match = /\/folders\/([a-zA-Z0-9-_]{10,128})/.exec(trimmed);
  if (match) return match[1];
  if (/^[a-zA-Z0-9-_]{10,128}$/.test(trimmed)) return trimmed;
  return null;
}

function readServiceAccountConfig(): ServiceAccountConfig | null {
  const spreadsheetId = extractSpreadsheetId(process.env.GOOGLE_SPREADSHEET_ID || '');
  if (!spreadsheetId) return null;

  let clientEmail = (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || '').trim();
  let privateKey = (process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || '')
    .replace(/\\n/g, '\n')
    .trim();

  const rawJson = (process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '').trim();
  if (rawJson) {
    try {
      const parsed = JSON.parse(rawJson) as {
        client_email?: string;
        private_key?: string;
      };
      if (parsed.client_email) clientEmail = parsed.client_email.trim();
      if (parsed.private_key) privateKey = parsed.private_key.replace(/\\n/g, '\n').trim();
    } catch {
      return null;
    }
  }

  if (!clientEmail || !privateKey || !privateKey.includes('BEGIN PRIVATE KEY')) {
    return null;
  }

  const driveFolderId = extractDriveFolderId(process.env.GOOGLE_DRIVE_FOLDER_ID || '');
  return { clientEmail, privateKey, spreadsheetId, driveFolderId };
}

let cachedAccessToken: { token: string; expiresAtMs: number; keyFingerprint: string } | null = null;

async function getServiceAccountAccessToken(cfg: ServiceAccountConfig): Promise<string> {
  const fingerprint = `${cfg.clientEmail}:${cfg.spreadsheetId}`;
  if (
    cachedAccessToken &&
    cachedAccessToken.keyFingerprint === fingerprint &&
    Date.now() < cachedAccessToken.expiresAtMs - 60_000
  ) {
    return cachedAccessToken.token;
  }

  const nowSec = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' }), 'utf8').toString(
    'base64url',
  );
  const claim = Buffer.from(
    JSON.stringify({
      iss: cfg.clientEmail,
      scope: SERVICE_SCOPES,
      aud: OAUTH_TOKEN_URL,
      iat: nowSec,
      exp: nowSec + 3600,
    }),
    'utf8',
  ).toString('base64url');

  const signingInput = `${header}.${claim}`;
  const signer = crypto.createSign('RSA-SHA256');
  signer.update(signingInput);
  signer.end();
  const signature = signer.sign(cfg.privateKey).toString('base64url');
  const assertion = `${signingInput}.${signature}`;

  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    assertion,
  });

  const res = await fetch(OAUTH_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  const data = (await res.json().catch(() => ({}))) as {
    access_token?: string;
    expires_in?: number;
    error_description?: string;
    error?: string;
  };

  if (!res.ok || !data.access_token) {
    throw new Error(
      data.error_description ||
        data.error ||
        `Google service account authorization failed (${res.status}).`,
    );
  }

  const ttlMs = Math.max(300, Number(data.expires_in || 3600)) * 1000;
  cachedAccessToken = {
    token: data.access_token,
    expiresAtMs: Date.now() + ttlMs,
    keyFingerprint: fingerprint,
  };
  return data.access_token;
}

async function googleJsonFetch<T>(
  token: string,
  url: string,
  init: RequestInit = {},
): Promise<T> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    ...(init.headers as Record<string, string> | undefined),
  };
  if (init.body && typeof init.body === 'string' && !headers['Content-Type']) {
    headers['Content-Type'] = 'application/json';
  }
  const res = await fetch(url, { ...init, headers });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const apiMsg =
      (data as { error?: { message?: string } })?.error?.message ||
      `Google API request failed (${res.status}).`;
    if (res.status === 401) {
      cachedAccessToken = null;
    }
    throw new Error(apiMsg);
  }
  return data as T;
}

async function uploadScreenshotToDrive(params: {
  token: string;
  folderId: string | null;
  uploadId: number;
  imageDataUrl: string;
  originalFilename: string | null;
  uploadedBy: string;
}): Promise<{ driveFileId: string; driveUrl: string }> {
  const match =
    /^data:(image\/(?:png|jpeg|jpg|webp|gif|bmp))(?:;[^;]+)*;base64,([\s\S]+)$/i.exec(
      params.imageDataUrl.trim(),
    );
  if (!match) {
    throw new Error(`Unsupported screenshot data format for upload #${params.uploadId}.`);
  }

  const mimeType = match[1].toLowerCase() === 'image/jpg' ? 'image/jpeg' : match[1].toLowerCase();
  const ext = mimeType.split('/')[1] || 'jpg';
  const imageBuffer = Buffer.from(match[2].replace(/\s+/g, ''), 'base64');

  const safeBase = (params.originalFilename || `partnerdesk_trade_${params.uploadId}.${ext}`)
    .replace(/[^\w.\- ]+/g, '_')
    .slice(0, 120);
  const metadata: Record<string, unknown> = {
    name: `${params.uploadedBy}_${params.uploadId}_${safeBase}`,
    description: `PartnerDesk Angel One option trade screenshot #${params.uploadId} uploaded by ${params.uploadedBy}`,
  };
  if (params.folderId) {
    metadata.parents = [params.folderId];
  }

  const boundary = `partnerdesk_${crypto.randomBytes(12).toString('hex')}`;
  const part1 = Buffer.from(
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(
      metadata,
    )}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`,
    'utf8',
  );
  const part3 = Buffer.from(`\r\n--${boundary}--`, 'utf8');
  const body = Buffer.concat([part1, imageBuffer, part3]);

  const res = await fetch(
    `${DRIVE_UPLOAD_BASE}?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${params.token}`,
        'Content-Type': `multipart/related; boundary=${boundary}`,
        'Content-Length': String(body.length),
      },
      body,
    },
  );

  const data = (await res.json().catch(() => ({}))) as {
    id?: string;
    webViewLink?: string;
    error?: { message?: string };
  };

  if (!res.ok || !data.id) {
    throw new Error(data.error?.message || `Google Drive upload failed (${res.status}).`);
  }

  return {
    driveFileId: data.id,
    driveUrl: data.webViewLink || `https://drive.google.com/file/d/${data.id}/view`,
  };
}

function buildTabsValueRanges() {
  return ALL_TAB_NAMES.map((name) => `'${name}'!A:AZ`);
}

async function ensureSheetTabs(
  token: string,
  spreadsheetId: string,
): Promise<{ spreadsheetTitle: string }> {
  const meta = await googleJsonFetch<{
    properties?: { title?: string };
    sheets?: { properties?: { title?: string } }[];
  }>(
    token,
    `${SHEETS_BASE}/${encodeURIComponent(
      spreadsheetId,
    )}?fields=properties.title,sheets.properties.title`,
  );

  const existingTitles = new Set(
    (meta.sheets || []).map((s) => s.properties?.title).filter((t): t is string => Boolean(t)),
  );
  const missing = ALL_TAB_NAMES.filter((name) => !existingTitles.has(name));

  if (missing.length > 0) {
    await googleJsonFetch(token, `${SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({
        requests: missing.map((title) => ({
          addSheet: { properties: { title } },
        })),
      }),
    });
  }

  return {
    spreadsheetTitle: meta.properties?.title?.trim() || 'PartnerDesk Shared Ledger',
  };
}

async function runSingleSyncPass(cfg: ServiceAccountConfig): Promise<void> {
  const token = await getServiceAccountAccessToken(cfg);

  // 1. Archive any pending trade screenshots to Google Drive first
  let driveWarning: string | null = null;
  const pendingUploads = await listUploadsPendingDriveSync();
  for (const upload of pendingUploads) {
    try {
      const ref = await uploadScreenshotToDrive({
        token,
        folderId: cfg.driveFolderId,
        uploadId: upload.id,
        imageDataUrl: upload.imageDataUrl,
        originalFilename: upload.originalFilename,
        uploadedBy: upload.uploadedBy,
      });
      await updateUploadDriveRef(upload.id, ref.driveFileId, ref.driveUrl);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      driveWarning = `Screenshot Drive archive deferred: ${msg}`;
      console.warn('[PartnerDesk GoogleSync] Drive upload warning:', msg);
      // Stop attempting remaining Drive uploads in this cycle if quota/permission fails
      break;
    }
  }

  // 2. Ensure tabs exist inside the single configured Google Sheet
  const { spreadsheetTitle } = await ensureSheetTabs(token, cfg.spreadsheetId);

  // 3. Take an authoritative snapshot from SQLite / Postgres
  const snapshot = await getFullLedgerSnapshotForSync();
  const { summary } = snapshot.dashboard;

  const summaryRows: (string | number)[][] = [
    ['PartnerDesk — Live Options Trading Summary', 'Value'],
    ['Total Partnership Capital (INR)', summary.totalPartnershipCapital],
    ['Angel One Account Balance (INR)', summary.totalAccountBalance],
    ['Prayag Net Capital (INR)', summary.ownerTotalContribution],
    ['Raj Net Capital (INR)', summary.rajTotalContribution],
    ['Prayag Ownership Balance (INR)', summary.ownerOwnershipBalance],
    ['Raj Ownership Balance (INR)', summary.rajOwnershipBalance],
    ['Total Gross P&L (INR)', summary.totalGrossPnl],
    ['Total Brokerage & Taxes (INR)', summary.totalBrokerageAndTaxes],
    ['Total Net P&L (INR)', summary.totalNetPnl],
    ['Prayag 50% Profit Share (INR)', summary.ownerProfitShare],
    ['Raj 50% Profit Share (INR)', summary.rajProfitShare],
    ['Closed Trades Count', summary.totalClosedTrades],
    ['Winning Trades', summary.winningTrades],
    ['Losing Trades', summary.losingTrades],
    ['Open Positions Count', summary.openPositionsCount],
    ['Last Synced At (ISO)', new Date().toISOString()],
  ];

  const tradeGroupRows: (string | number)[][] = [
    [
      'id',
      'groupKey',
      'tradingDate',
      'uploadedAt',
      'stockName',
      'tradingSymbol',
      'expiryDate',
      'strikePrice',
      'optionType',
      'buyQty',
      'sellQty',
      'avgBuyPrice',
      'avgSellPrice',
      'totalBuyAmount',
      'totalSellAmount',
      'realisedQty',
      'openQty',
      'openSide',
      'grossPnl',
      'brokerage',
      'stt',
      'exchangeTxnCharge',
      'sebiCharge',
      'stampDuty',
      'ipftCharge',
      'gst',
      'otherCharges',
      'totalCharges',
      'netPnl',
      'chargesMode',
      'contractNoteRef',
      'notes',
      'updatedAt',
    ],
    ...snapshot.tradeGroups.map((t) => [
      t.id,
      t.groupKey,
      t.tradingDate,
      t.uploadedAt,
      t.stockName,
      t.tradingSymbol,
      t.expiryDate,
      t.strikePrice,
      t.optionType,
      t.buyQty,
      t.sellQty,
      t.avgBuyPrice,
      t.avgSellPrice,
      t.totalBuyAmount,
      t.totalSellAmount,
      t.realisedQty,
      t.openQty,
      t.openSide ?? '',
      t.grossPnl,
      t.brokerage,
      t.stt,
      t.exchangeTxnCharge,
      t.sebiCharge,
      t.stampDuty,
      t.ipftCharge,
      t.gst,
      t.otherCharges,
      t.totalCharges,
      t.netPnl,
      t.chargesMode,
      t.contractNoteRef ?? '',
      t.notes ?? '',
      t.updatedAt,
    ]),
  ];

  const orderRows: (string | number)[][] = [
    [
      'id',
      'tradeGroupId',
      'uploadId',
      'tradingDate',
      'uploadedAt',
      'stockName',
      'tradingSymbol',
      'expiryDate',
      'strikePrice',
      'optionType',
      'side',
      'quantity',
      'lots',
      'executionPrice',
      'amount',
      'executionTime',
      'createdBy',
      'createdAt',
    ],
    ...snapshot.orders.map((o) => [
      o.id,
      o.tradeGroupId,
      o.uploadId ?? '',
      o.tradingDate,
      o.uploadedAt,
      o.stockName,
      o.tradingSymbol,
      o.expiryDate,
      o.strikePrice,
      o.optionType,
      o.side,
      o.quantity,
      o.lots ?? '',
      o.executionPrice,
      o.amount,
      o.executionTime,
      o.createdBy,
      o.createdAt,
    ]),
  ];

  const depositRows: (string | number)[][] = [
    ['id', 'partner', 'entryType', 'amount', 'depositDate', 'referenceNote', 'createdBy', 'createdAt'],
    ...snapshot.deposits.map((d) => [
      d.id,
      d.partner,
      d.entryType || 'DEPOSIT',
      d.amount,
      d.depositDate,
      d.referenceNote ?? '',
      d.createdBy,
      d.createdAt,
    ]),
  ];

  const rateRows: (string | number)[][] = [
    ['key', 'value'],
    ['brokeragePerOrder', snapshot.rates.brokeragePerOrder],
    ['sttSellPct', snapshot.rates.sttSellPct],
    ['exchangeTxnPct', snapshot.rates.exchangeTxnPct],
    ['ipftPct', snapshot.rates.ipftPct],
    ['sebiPct', snapshot.rates.sebiPct],
    ['stampDutyBuyPct', snapshot.rates.stampDutyBuyPct],
    ['gstPct', snapshot.rates.gstPct],
  ];

  const screenshotRows: (string | number)[][] = [
    ['id', 'imageHash', 'driveFileId', 'driveUrl', 'originalFilename', 'uploadedBy', 'uploadedAt'],
    ...snapshot.uploads.map((u) => [
      u.id,
      u.imageHash,
      u.driveFileId ?? '',
      u.driveUrl ?? '',
      u.originalFilename ?? '',
      u.uploadedBy,
      u.uploadedAt,
    ]),
  ];

  const groupScreenshotRows: (string | number)[][] = [
    ['tradeGroupId', 'uploadId'],
    ...snapshot.groupScreenshots.map((l) => [l.tradeGroupId, l.uploadId]),
  ];

  const auditRows: (string | number)[][] = [
    ['id', 'actor', 'action', 'entityType', 'entityId', 'summary', 'detailsJson', 'createdAt'],
    ...snapshot.auditLogs.map((a) => [
      a.id,
      a.actor,
      a.action,
      a.entityType,
      a.entityId ?? '',
      a.summary,
      a.detailsJson ?? '',
      a.createdAt,
    ]),
  ];

  // 4. Clear and rewrite all tabs cleanly in two batch requests to prevent duplicate rows
  await googleJsonFetch(
    token,
    `${SHEETS_BASE}/${encodeURIComponent(cfg.spreadsheetId)}/values:batchClear`,
    {
      method: 'POST',
      body: JSON.stringify({ ranges: buildTabsValueRanges() }),
    },
  );

  await googleJsonFetch(
    token,
    `${SHEETS_BASE}/${encodeURIComponent(cfg.spreadsheetId)}/values:batchUpdate`,
    {
      method: 'POST',
      body: JSON.stringify({
        valueInputOption: 'RAW',
        data: [
          { range: `'${TABS.SUMMARY}'!A1`, values: summaryRows },
          { range: `'${TABS.TRADE_GROUPS}'!A1`, values: tradeGroupRows },
          { range: `'${TABS.ORDER_LEGS}'!A1`, values: orderRows },
          { range: `'${TABS.DEPOSITS}'!A1`, values: depositRows },
          { range: `'${TABS.RATES}'!A1`, values: rateRows },
          { range: `'${TABS.SCREENSHOTS}'!A1`, values: screenshotRows },
          { range: `'${TABS.GROUP_SCREENSHOTS}'!A1`, values: groupScreenshotRows },
          { range: `'${TABS.AUDIT_LOGS}'!A1`, values: auditRows },
        ],
      }),
    },
  );

  await updateSheetSyncMeta({
    lastSyncedAt: new Date().toISOString(),
    pendingSync: false,
    lastError: null,
    driveWarning,
    spreadsheetTitle,
  });
}

// Unused import reference check guard for Drive file deletion helper if needed later
void DRIVE_FILES_BASE;

let syncInFlight: Promise<void> | null = null;
let dirtyQueued = false;
let retryCount = 0;
let retryTimer: ReturnType<typeof setTimeout> | null = null;

function clearRetryTimer(): void {
  if (retryTimer) {
    clearTimeout(retryTimer);
    retryTimer = null;
  }
}

function scheduleRetryTimer(): void {
  if (process.env.VERCEL === '1') return;
  clearRetryTimer();
  const delaysMs = [5000, 15000, 45000, 120000, 300000];
  const delay = delaysMs[Math.min(retryCount, delaysMs.length - 1)];
  retryTimer = setTimeout(() => {
    retryTimer = null;
    void triggerGoogleSync();
  }, delay);
  if (typeof retryTimer.unref === 'function') {
    retryTimer.unref();
  }
}

export function triggerGoogleSync(): Promise<void> {
  const cfg = readServiceAccountConfig();
  if (!cfg) {
    return Promise.resolve();
  }

  dirtyQueued = true;
  if (syncInFlight) {
    return syncInFlight;
  }

  clearRetryTimer();
  syncInFlight = (async () => {
    try {
      while (dirtyQueued) {
        dirtyQueued = false;
        const latestCfg = readServiceAccountConfig();
        if (!latestCfg) break;
        try {
          await runSingleSyncPass(latestCfg);
          retryCount = 0;
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          console.error('[PartnerDesk GoogleSync] Sync error:', msg);
          await updateSheetSyncMeta({
            pendingSync: true,
            lastError: msg,
          }).catch(() => {});
          retryCount += 1;
          scheduleRetryTimer();
          break;
        }
      }
    } finally {
      syncInFlight = null;
      if (dirtyQueued && retryCount === 0) {
        void triggerGoogleSync();
      }
    }
  })();

  return syncInFlight;
}

export async function markSyncPendingAndTrigger(options?: {
  waitBestEffortMs?: number;
}): Promise<void> {
  const cfg = readServiceAccountConfig();
  await updateSheetSyncMeta({ pendingSync: Boolean(cfg) }).catch(() => {});
  if (!cfg) return;

  const p = triggerGoogleSync();
  const waitMs =
    options?.waitBestEffortMs ?? (process.env.VERCEL === '1' ? 3500 : 0);
  if (waitMs > 0) {
    await Promise.race([
      p.catch(() => {}),
      new Promise<void>((resolve) => setTimeout(resolve, waitMs)),
    ]);
  }
}

export async function getGoogleSyncStatus(): Promise<GoogleSyncStatus> {
  const cfg = readServiceAccountConfig();
  const meta = await getSheetSyncMeta().catch(() => ({
    lastSyncedAt: null,
    pendingSync: false,
    lastError: null,
    driveWarning: null,
    spreadsheetTitle: null,
  }));

  return {
    configured: Boolean(cfg),
    spreadsheetId: cfg?.spreadsheetId ?? null,
    spreadsheetUrl: cfg
      ? `https://docs.google.com/spreadsheets/d/${cfg.spreadsheetId}/edit`
      : null,
    spreadsheetTitle: meta.spreadsheetTitle,
    driveFolderConfigured: Boolean(cfg?.driveFolderId),
    lastSyncedAt: meta.lastSyncedAt,
    pendingSync: cfg ? meta.pendingSync || dirtyQueued : false,
    syncingNow: Boolean(syncInFlight),
    lastError: cfg ? meta.lastError : null,
    driveWarning: cfg ? meta.driveWarning : null,
  };
}

export async function initGoogleSyncOnStartup(): Promise<void> {
  const cfg = readServiceAccountConfig();
  if (!cfg) return;
  await triggerGoogleSync().catch(() => {});
}
