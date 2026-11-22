import {
  buildTradeGroupKey,
  computePartnershipAndSettlement,
  computeTradeGroupMetrics,
  DEFAULT_CHARGE_RATES,
  round2,
} from '../../shared/calculations.ts';
import type {
  AuditLogRecord,
  ChargeRates,
  ChargesMode,
  DashboardSummary,
  DepositRecord,
  ExtractedOrderLeg,
  FundEntryType,
  OptionType,
  OrderSide,
  PartnerId,
  SettlementPreview,
  TradeGroupRecord,
  TradeOrderRecord,
} from '../../shared/types.ts';
import { clearGoogleAccessToken, getGoogleAccessToken } from './googleWorkspace.ts';

export interface StoredUploadRecord {
  id: number;
  imageHash: string;
  imageDataUrl: string;
  originalFilename: string | null;
  uploadedBy: PartnerId;
  uploadedAt: string;
}

interface GroupScreenshotLink {
  tradeGroupId: number;
  uploadId: number;
}

export interface SheetBackupPayload {
  format: 'partnerdesk-local-backup-v1';
  folderName: 'PartnerDesk_GoogleSheet';
  exportedAt: string;
  chargeRates: ChargeRates;
  deposits: DepositRecord[];
  uploads: StoredUploadRecord[];
  tradeGroups: TradeGroupRecord[];
  orders: TradeOrderRecord[];
  auditLogs: AuditLogRecord[];
}

export interface SheetDashboardPayload {
  summary: DashboardSummary;
  settlementPreview: SettlementPreview;
  recentTrades: TradeGroupRecord[];
  openPositions: TradeGroupRecord[];
  recentDeposits: DepositRecord[];
  recentSettlements: [];
}

interface SpreadsheetSnapshot {
  spreadsheetTitle: string;
  rates: ChargeRates;
  tradeGroups: TradeGroupRecord[];
  orders: TradeOrderRecord[];
  deposits: DepositRecord[];
  auditLogs: AuditLogRecord[];
  uploads: StoredUploadRecord[];
  groupScreenshots: GroupScreenshotLink[];
}

const SHEETS_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';
const DRIVE_BASE = 'https://www.googleapis.com/drive/v3/files';
const CHUNK_SIZE = 40_000;

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
  TABS.AUDIT_LOGS,
  TABS.SCREENSHOTS,
  TABS.GROUP_SCREENSHOTS,
] as const;

const TRADE_GROUP_HEADERS = [
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
];

const ORDER_LEG_HEADERS = [
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
];

const DEPOSIT_HEADERS = [
  'id',
  'partner',
  'entryType',
  'amount',
  'depositDate',
  'referenceNote',
  'createdBy',
  'createdAt',
];

const RATE_HEADERS = ['key', 'value'];

const AUDIT_HEADERS = [
  'id',
  'actor',
  'action',
  'entityType',
  'entityId',
  'summary',
  'detailsJson',
  'createdAt',
];

const GROUP_SCREENSHOT_HEADERS = ['tradeGroupId', 'uploadId'];

export function extractSpreadsheetId(rawInput: string): string | null {
  const trimmed = rawInput.trim();
  if (!trimmed) return null;
  const urlMatch = /\/spreadsheets\/d\/([a-zA-Z0-9-_]{10,128})/.exec(trimmed);
  if (urlMatch) return urlMatch[1];
  if (/^[a-zA-Z0-9-_]{10,128}$/.test(trimmed)) return trimmed;
  return null;
}

export async function sha256Hex(input: string): Promise<string> {
  const bytes = new TextEncoder().encode(input);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
}

async function googleApiFetch<T>(url: string, init: RequestInit = {}): Promise<T> {
  const token = getGoogleAccessToken();
  if (!token) {
    throw new Error(
      'GOOGLE_AUTH_REQUIRED: Please sign in with Google to read and save data in your shared Google Sheet.',
    );
  }

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    ...(init.headers as Record<string, string> | undefined),
  };
  if (init.body && !headers['Content-Type']) {
    headers['Content-Type'] = 'application/json';
  }

  const res = await fetch(url, { ...init, headers });
  if (res.status === 401) {
    clearGoogleAccessToken();
    throw new Error(
      'GOOGLE_AUTH_EXPIRED: Your Google session expired. Please sign in with Google again.',
    );
  }

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const apiMsg =
      (data as { error?: { message?: string } })?.error?.message ||
      `Google Sheets request failed (${res.status}).`;
    if (res.status === 403) {
      throw new Error(
        `Permission denied on the shared Google Sheet (${apiMsg}). Make sure this Google account has Editor access to the spreadsheet.`,
      );
    }
    throw new Error(apiMsg);
  }
  return data as T;
}

function toNum(v: unknown, fallback = 0): number {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function toNullableNum(v: unknown): number | null {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function toStr(v: unknown, fallback = ''): string {
  if (v === null || v === undefined) return fallback;
  return String(v);
}

function toNullableStr(v: unknown): string | null {
  if (v === null || v === undefined) return null;
  const s = String(v);
  return s === '' ? null : s;
}

function buildSummaryRows(
  trades: TradeGroupRecord[],
  deposits: DepositRecord[],
  actor: string,
): (string | number)[][] {
  const { summary } = buildSummaryFromData(trades, deposits);
  return [
    ['PartnerDesk — Live Options Trading Summary', 'Value'],
    ['Total Partnership Capital (₹)', summary.totalPartnershipCapital],
    ['Angel One Account Balance (₹)', summary.totalAccountBalance],
    ['Prayag Net Capital (₹)', summary.ownerTotalContribution],
    ['Raj Net Capital (₹)', summary.rajTotalContribution],
    ['Prayag Ownership Balance (₹)', summary.ownerOwnershipBalance],
    ['Raj Ownership Balance (₹)', summary.rajOwnershipBalance],
    ['Total Gross P&L (₹)', summary.totalGrossPnl],
    ['Total Brokerage & Taxes (₹)', summary.totalBrokerageAndTaxes],
    ['Total Net P&L (₹)', summary.totalNetPnl],
    ['Prayag 50% Profit Share (₹)', summary.ownerProfitShare],
    ['Raj 50% Profit Share (₹)', summary.rajProfitShare],
    ['Closed Trades Count', summary.totalClosedTrades],
    ['Winning Trades', summary.winningTrades],
    ['Losing Trades', summary.losingTrades],
    ['Open Positions Count', summary.openPositionsCount],
    ['Last Updated By', actor],
    ['Last Updated At (ISO)', new Date().toISOString()],
  ];
}

function buildSummaryFromData(
  trades: TradeGroupRecord[],
  deposits: DepositRecord[],
): { summary: DashboardSummary; settlementPreview: SettlementPreview } {
  const settlementPreview = computePartnershipAndSettlement(trades, deposits, []);
  const closedTrades = trades.filter((t) => t.realisedQty > 0);
  const openTrades = trades.filter((t) => t.openQty > 0);

  const totalBuyAmount = round2(trades.reduce((s, t) => s + t.totalBuyAmount, 0));
  const totalSellAmount = round2(trades.reduce((s, t) => s + t.totalSellAmount, 0));
  const winningTrades = closedTrades.filter((t) => t.netPnl > 0).length;
  const losingTrades = closedTrades.filter((t) => t.netPnl < 0).length;

  const summary: DashboardSummary = {
    totalPartnershipCapital: settlementPreview.totalPartnershipCapital,
    totalAccountBalance: settlementPreview.totalAccountBalance,
    ownerTotalContribution: settlementPreview.ownerTotalDeposits,
    rajTotalContribution: settlementPreview.rajTotalDeposits,
    ownerOwnershipBalance: settlementPreview.ownerOwnershipBalance,
    rajOwnershipBalance: settlementPreview.rajOwnershipBalance,
    ownerPaidOutAmount: settlementPreview.ownerPaidOutSoFar,
    rajPaidOutAmount: settlementPreview.rajPaidOutSoFar,
    eachCapitalObligation: 0,
    ownerOutstandingAmount: 0,
    rajOutstandingAmount: 0,
    totalBuyAmount,
    totalSellAmount,
    totalGrossPnl: settlementPreview.totalRealisedGrossPnl,
    totalBrokerageAndTaxes: settlementPreview.totalRealisedCharges,
    totalNetPnl: settlementPreview.totalRealisedNetPnl,
    ownerProfitShare: round2(settlementPreview.totalRealisedNetPnl / 2),
    rajProfitShare: round2(
      settlementPreview.totalRealisedNetPnl - round2(settlementPreview.totalRealisedNetPnl / 2),
    ),
    totalDistributedAmount: 0,
    pendingSettlementAmount: settlementPreview.unsettledNetPnl,
    totalClosedTrades: closedTrades.length,
    winningTrades,
    losingTrades,
    openPositionsCount: openTrades.length,
  };

  return { summary, settlementPreview };
}

function hydrateTrades(
  groups: TradeGroupRecord[],
  orders: TradeOrderRecord[],
  uploads: StoredUploadRecord[],
  links: GroupScreenshotLink[],
): TradeGroupRecord[] {
  const uploadMap = new Map<number, StoredUploadRecord>();
  for (const u of uploads) uploadMap.set(u.id, u);

  const ordersByGroup = new Map<number, TradeOrderRecord[]>();
  for (const o of orders) {
    const arr = ordersByGroup.get(o.tradeGroupId) || [];
    arr.push(o);
    ordersByGroup.set(o.tradeGroupId, arr);
  }

  const screenshotsByGroup = new Map<number, NonNullable<TradeGroupRecord['screenshots']>>();
  for (const link of links) {
    const u = uploadMap.get(link.uploadId);
    if (!u) continue;
    const arr = screenshotsByGroup.get(link.tradeGroupId) || [];
    if (!arr.some((s) => s.id === u.id)) {
      arr.push({
        id: u.id,
        imageDataUrl: u.imageDataUrl,
        originalFilename: u.originalFilename,
        uploadedBy: u.uploadedBy,
        uploadedAt: u.uploadedAt,
      });
    }
    screenshotsByGroup.set(link.tradeGroupId, arr);
  }

  return groups
    .map((g) => ({
      ...g,
      orders: (ordersByGroup.get(g.id) || []).sort((a, b) => a.id - b.id),
      screenshots: (screenshotsByGroup.get(g.id) || []).sort((a, b) => b.id - a.id),
    }))
    .sort((a, b) => b.tradingDate.localeCompare(a.tradingDate) || b.id - a.id);
}

function recomputeGroupInMemory(
  group: TradeGroupRecord,
  allOrders: TradeOrderRecord[],
  rates: ChargeRates,
): TradeGroupRecord {
  const groupOrders = allOrders
    .filter((o) => o.tradeGroupId === group.id)
    .sort((a, b) => a.id - b.id);

  const override =
    group.chargesMode === 'CONTRACT_NOTE'
      ? {
          brokerage: group.brokerage,
          stt: group.stt,
          exchangeTxnCharge: group.exchangeTxnCharge,
          ipftCharge: group.ipftCharge,
          sebiCharge: group.sebiCharge,
          stampDuty: group.stampDuty,
          gst: group.gst,
          otherCharges: group.otherCharges,
        }
      : null;

  const metrics = computeTradeGroupMetrics(
    groupOrders.map((o) => ({
      side: o.side,
      quantity: o.quantity,
      executionPrice: o.executionPrice,
      amount: o.amount,
    })),
    rates,
    override,
  );

  return {
    ...group,
    buyQty: metrics.buyQty,
    sellQty: metrics.sellQty,
    avgBuyPrice: metrics.avgBuyPrice,
    avgSellPrice: metrics.avgSellPrice,
    totalBuyAmount: metrics.totalBuyAmount,
    totalSellAmount: metrics.totalSellAmount,
    realisedQty: metrics.realisedQty,
    openQty: metrics.openQty,
    openSide: metrics.openSide,
    grossPnl: metrics.grossPnl,
    brokerage: metrics.brokerage,
    stt: metrics.stt,
    exchangeTxnCharge: metrics.exchangeTxnCharge,
    ipftCharge: metrics.ipftCharge,
    sebiCharge: metrics.sebiCharge,
    stampDuty: metrics.stampDuty,
    gst: metrics.gst,
    otherCharges: metrics.otherCharges,
    totalCharges: metrics.totalCharges,
    netPnl: metrics.netPnl,
    updatedAt: new Date().toISOString(),
  };
}

function serializeScreenshots(uploads: StoredUploadRecord[]): (string | number)[][] {
  let maxChunks = 1;
  const chunked = uploads.map((u) => {
    const chunks: string[] = [];
    for (let i = 0; i < u.imageDataUrl.length; i += CHUNK_SIZE) {
      chunks.push(u.imageDataUrl.slice(i, i + CHUNK_SIZE));
    }
    if (chunks.length === 0) chunks.push('');
    maxChunks = Math.max(maxChunks, chunks.length);
    return { upload: u, chunks };
  });

  const headers = [
    'id',
    'imageHash',
    'originalFilename',
    'uploadedBy',
    'uploadedAt',
    'chunkCount',
    ...Array.from({ length: maxChunks }, (_, i) => `chunk_${i}`),
  ];

  const rows: (string | number)[][] = [headers];
  for (const { upload: u, chunks } of chunked) {
    const row: (string | number)[] = [
      u.id,
      u.imageHash,
      u.originalFilename ?? '',
      u.uploadedBy,
      u.uploadedAt,
      chunks.length,
      ...chunks,
    ];
    while (row.length < headers.length) row.push('');
    rows.push(row);
  }
  return rows;
}

function parseTradeGroups(rows: unknown[][]): TradeGroupRecord[] {
  if (rows.length <= 1) return [];
  const out: TradeGroupRecord[] = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length === 0 || !r[0]) continue;
    const id = toNum(r[0]);
    if (id <= 0) continue;
    out.push({
      id,
      groupKey: toStr(r[1]),
      tradingDate: toStr(r[2]),
      uploadedAt: toStr(r[3]),
      stockName: toStr(r[4]),
      tradingSymbol: toStr(r[5]),
      expiryDate: toStr(r[6]),
      strikePrice: toNum(r[7]),
      optionType: r[8] === 'PE' ? 'PE' : 'CE',
      buyQty: toNum(r[9]),
      sellQty: toNum(r[10]),
      avgBuyPrice: toNum(r[11]),
      avgSellPrice: toNum(r[12]),
      totalBuyAmount: toNum(r[13]),
      totalSellAmount: toNum(r[14]),
      realisedQty: toNum(r[15]),
      openQty: toNum(r[16]),
      openSide: r[17] === 'BUY' || r[17] === 'SELL' ? r[17] : null,
      grossPnl: toNum(r[18]),
      brokerage: toNum(r[19]),
      stt: toNum(r[20]),
      exchangeTxnCharge: toNum(r[21]),
      sebiCharge: toNum(r[22]),
      stampDuty: toNum(r[23]),
      ipftCharge: toNum(r[24]),
      gst: toNum(r[25]),
      otherCharges: toNum(r[26]),
      totalCharges: toNum(r[27]),
      netPnl: toNum(r[28]),
      chargesMode: r[29] === 'CONTRACT_NOTE' ? 'CONTRACT_NOTE' : 'ESTIMATED',
      contractNoteRef: toNullableStr(r[30]),
      notes: toNullableStr(r[31]),
      updatedAt: toStr(r[32]),
    });
  }
  return out;
}

function parseOrders(rows: unknown[][]): TradeOrderRecord[] {
  if (rows.length <= 1) return [];
  const out: TradeOrderRecord[] = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length === 0 || !r[0]) continue;
    const id = toNum(r[0]);
    if (id <= 0) continue;
    out.push({
      id,
      tradeGroupId: toNum(r[1]),
      uploadId: toNullableNum(r[2]),
      tradingDate: toStr(r[3]),
      uploadedAt: toStr(r[4]),
      stockName: toStr(r[5]),
      tradingSymbol: toStr(r[6]),
      expiryDate: toStr(r[7]),
      strikePrice: toNum(r[8]),
      optionType: r[9] === 'PE' ? 'PE' : 'CE',
      side: r[10] === 'SELL' ? 'SELL' : 'BUY',
      quantity: toNum(r[11]),
      lots: toNullableNum(r[12]),
      executionPrice: toNum(r[13]),
      amount: toNum(r[14]),
      executionTime: toStr(r[15]),
      createdBy: r[16] === 'Raj' ? 'Raj' : 'Prayag',
      createdAt: toStr(r[17]),
    });
  }
  return out;
}

function parseDeposits(rows: unknown[][]): DepositRecord[] {
  if (rows.length <= 1) return [];
  const out: DepositRecord[] = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length === 0 || !r[0]) continue;
    const id = toNum(r[0]);
    if (id <= 0) continue;
    out.push({
      id,
      partner: r[1] === 'Raj' ? 'Raj' : 'Prayag',
      entryType: r[2] === 'WITHDRAWAL' ? 'WITHDRAWAL' : 'DEPOSIT',
      amount: toNum(r[3]),
      depositDate: toStr(r[4]),
      referenceNote: toNullableStr(r[5]),
      createdBy: r[6] === 'Raj' ? 'Raj' : 'Prayag',
      createdAt: toStr(r[7]),
    });
  }
  return out;
}

function parseRates(rows: unknown[][]): ChargeRates {
  const rates: ChargeRates = { ...DEFAULT_CHARGE_RATES };
  if (rows.length <= 1) return rates;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[0]) continue;
    const key = String(r[0]) as keyof ChargeRates;
    if (key in rates) {
      const val = Number(r[1]);
      if (Number.isFinite(val) && val >= 0) {
        rates[key] = val;
      }
    }
  }
  return rates;
}

function parseAuditLogs(rows: unknown[][]): AuditLogRecord[] {
  if (rows.length <= 1) return [];
  const out: AuditLogRecord[] = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length === 0 || !r[0]) continue;
    const id = toNum(r[0]);
    if (id <= 0) continue;
    out.push({
      id,
      actor: toStr(r[1]),
      action: toStr(r[2]),
      entityType: toStr(r[3]),
      entityId: toNullableStr(r[4]),
      summary: toStr(r[5]),
      detailsJson: toNullableStr(r[6]),
      createdAt: toStr(r[7]),
    });
  }
  return out;
}

function parseScreenshots(rows: unknown[][]): StoredUploadRecord[] {
  if (rows.length <= 1) return [];
  const out: StoredUploadRecord[] = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length === 0 || !r[0]) continue;
    const id = toNum(r[0]);
    if (id <= 0) continue;
    const chunkCount = Math.max(0, toNum(r[5], 0));
    const chunks: string[] = [];
    for (let c = 0; c < chunkCount; c++) {
      chunks.push(toStr(r[6 + c], ''));
    }
    out.push({
      id,
      imageHash: toStr(r[1]),
      originalFilename: toNullableStr(r[2]),
      uploadedBy: r[3] === 'Raj' ? 'Raj' : 'Prayag',
      uploadedAt: toStr(r[4]),
      imageDataUrl: chunks.join(''),
    });
  }
  return out;
}

function parseGroupScreenshots(rows: unknown[][]): GroupScreenshotLink[] {
  if (rows.length <= 1) return [];
  const out: GroupScreenshotLink[] = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length < 2 || !r[0] || !r[1]) continue;
    const tradeGroupId = toNum(r[0]);
    const uploadId = toNum(r[1]);
    if (tradeGroupId > 0 && uploadId > 0) {
      out.push({ tradeGroupId, uploadId });
    }
  }
  return out;
}

function serializeRates(rates: ChargeRates): (string | number)[][] {
  return [
    RATE_HEADERS,
    ...Object.entries(rates).map(([k, v]) => [k, v]),
  ];
}

function serializeTradeGroups(groups: TradeGroupRecord[]): (string | number)[][] {
  return [
    TRADE_GROUP_HEADERS,
    ...groups.map((g) => [
      g.id,
      g.groupKey,
      g.tradingDate,
      g.uploadedAt,
      g.stockName,
      g.tradingSymbol,
      g.expiryDate,
      g.strikePrice,
      g.optionType,
      g.buyQty,
      g.sellQty,
      g.avgBuyPrice,
      g.avgSellPrice,
      g.totalBuyAmount,
      g.totalSellAmount,
      g.realisedQty,
      g.openQty,
      g.openSide ?? '',
      g.grossPnl,
      g.brokerage,
      g.stt,
      g.exchangeTxnCharge,
      g.sebiCharge,
      g.stampDuty,
      g.ipftCharge,
      g.gst,
      g.otherCharges,
      g.totalCharges,
      g.netPnl,
      g.chargesMode,
      g.contractNoteRef ?? '',
      g.notes ?? '',
      g.updatedAt,
    ]),
  ];
}

function serializeOrders(orders: TradeOrderRecord[]): (string | number)[][] {
  return [
    ORDER_LEG_HEADERS,
    ...orders.map((o) => [
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
}

function serializeDeposits(deposits: DepositRecord[]): (string | number)[][] {
  return [
    DEPOSIT_HEADERS,
    ...deposits.map((d) => [
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
}

function serializeAuditLogs(logs: AuditLogRecord[]): (string | number)[][] {
  return [
    AUDIT_HEADERS,
    ...logs.slice(0, 500).map((l) => [
      l.id,
      l.actor,
      l.action,
      l.entityType,
      l.entityId ?? '',
      l.summary,
      l.detailsJson ?? '',
      l.createdAt,
    ]),
  ];
}

function serializeGroupScreenshots(links: GroupScreenshotLink[]): (string | number)[][] {
  return [
    GROUP_SCREENSHOT_HEADERS,
    ...links.map((l) => [l.tradeGroupId, l.uploadId]),
  ];
}

interface SpreadsheetMetadata {
  properties?: { title?: string };
  spreadsheetUrl?: string;
  sheets?: { properties?: { sheetId?: number; title?: string } }[];
}

async function ensureSpreadsheetTabs(spreadsheetId: string): Promise<{
  title: string;
  spreadsheetUrl: string;
}> {
  const meta = await googleApiFetch<SpreadsheetMetadata>(
    `${SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}?fields=properties.title,spreadsheetUrl,sheets.properties`,
  );

  const existingTitles = new Set(
    (meta.sheets || []).map((s) => s.properties?.title).filter((t): t is string => Boolean(t)),
  );
  const missingTabs = ALL_TAB_NAMES.filter((name) => !existingTitles.has(name));

  if (missingTabs.length > 0) {
    await googleApiFetch(`${SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({
        requests: missingTabs.map((title) => ({
          addSheet: {
            properties: {
              title,
              gridProperties: {
                columnCount: title === TABS.SCREENSHOTS ? 40 : 36,
              },
            },
          },
        })),
      }),
    });

    const seedRanges: { range: string; values: (string | number)[][] }[] = [];
    for (const tab of missingTabs) {
      if (tab === TABS.SUMMARY) {
        seedRanges.push({
          range: `${TABS.SUMMARY}!A1`,
          values: buildSummaryRows([], [], 'System'),
        });
      } else if (tab === TABS.TRADE_GROUPS) {
        seedRanges.push({ range: `${TABS.TRADE_GROUPS}!A1`, values: [TRADE_GROUP_HEADERS] });
      } else if (tab === TABS.ORDER_LEGS) {
        seedRanges.push({ range: `${TABS.ORDER_LEGS}!A1`, values: [ORDER_LEG_HEADERS] });
      } else if (tab === TABS.DEPOSITS) {
        seedRanges.push({ range: `${TABS.DEPOSITS}!A1`, values: [DEPOSIT_HEADERS] });
      } else if (tab === TABS.RATES) {
        seedRanges.push({
          range: `${TABS.RATES}!A1`,
          values: serializeRates(DEFAULT_CHARGE_RATES),
        });
      } else if (tab === TABS.AUDIT_LOGS) {
        seedRanges.push({ range: `${TABS.AUDIT_LOGS}!A1`, values: [AUDIT_HEADERS] });
      } else if (tab === TABS.SCREENSHOTS) {
        seedRanges.push({ range: `${TABS.SCREENSHOTS}!A1`, values: serializeScreenshots([]) });
      } else if (tab === TABS.GROUP_SCREENSHOTS) {
        seedRanges.push({
          range: `${TABS.GROUP_SCREENSHOTS}!A1`,
          values: [GROUP_SCREENSHOT_HEADERS],
        });
      }
    }

    if (seedRanges.length > 0) {
      await googleApiFetch(`${SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchUpdate`, {
        method: 'POST',
        body: JSON.stringify({
          valueInputOption: 'RAW',
          data: seedRanges,
        }),
      });
    }
  }

  return {
    title: meta.properties?.title || 'PartnerDesk — Options Trading Ledger',
    spreadsheetUrl:
      meta.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
  };
}

const ensuredSheetsCache = new Set<string>();

async function ensureTabsOnce(spreadsheetId: string): Promise<string> {
  if (!ensuredSheetsCache.has(spreadsheetId)) {
    const info = await ensureSpreadsheetTabs(spreadsheetId);
    ensuredSheetsCache.add(spreadsheetId);
    return info.title;
  }
  return 'PartnerDesk — Options Trading Ledger';
}

async function readFullSnapshot(spreadsheetId: string): Promise<SpreadsheetSnapshot> {
  const title = await ensureTabsOnce(spreadsheetId);
  const ranges = [
    `${TABS.TRADE_GROUPS}!A:AG`,
    `${TABS.ORDER_LEGS}!A:R`,
    `${TABS.DEPOSITS}!A:H`,
    `${TABS.RATES}!A:B`,
    `${TABS.AUDIT_LOGS}!A:H`,
    `${TABS.SCREENSHOTS}!A:ZZ`,
    `${TABS.GROUP_SCREENSHOTS}!A:B`,
  ];
  const params = new URLSearchParams();
  for (const r of ranges) params.append('ranges', r);

  const res = await googleApiFetch<{
    valueRanges?: { range?: string; values?: unknown[][] }[];
  }>(`${SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchGet?${params.toString()}`);

  const vr = res.valueRanges || [];
  return {
    spreadsheetTitle: title,
    tradeGroups: parseTradeGroups(vr[0]?.values || []),
    orders: parseOrders(vr[1]?.values || []),
    deposits: parseDeposits(vr[2]?.values || []),
    rates: parseRates(vr[3]?.values || []),
    auditLogs: parseAuditLogs(vr[4]?.values || []),
    uploads: parseScreenshots(vr[5]?.values || []),
    groupScreenshots: parseGroupScreenshots(vr[6]?.values || []),
  };
}

async function writeSnapshotTabs(
  spreadsheetId: string,
  snapshot: SpreadsheetSnapshot,
  actor: string,
  opts: {
    trades?: boolean;
    deposits?: boolean;
    rates?: boolean;
    screenshots?: boolean;
    audit?: boolean;
  },
): Promise<void> {
  const clearRanges: string[] = [`${TABS.SUMMARY}!A:B`];
  const updates: { range: string; values: (string | number)[][] }[] = [
    {
      range: `${TABS.SUMMARY}!A1`,
      values: buildSummaryRows(snapshot.tradeGroups, snapshot.deposits, actor),
    },
  ];

  if (opts.trades) {
    clearRanges.push(
      `${TABS.TRADE_GROUPS}!A:AG`,
      `${TABS.ORDER_LEGS}!A:R`,
      `${TABS.GROUP_SCREENSHOTS}!A:B`,
    );
    updates.push(
      { range: `${TABS.TRADE_GROUPS}!A1`, values: serializeTradeGroups(snapshot.tradeGroups) },
      { range: `${TABS.ORDER_LEGS}!A1`, values: serializeOrders(snapshot.orders) },
      {
        range: `${TABS.GROUP_SCREENSHOTS}!A1`,
        values: serializeGroupScreenshots(snapshot.groupScreenshots),
      },
    );
  }
  if (opts.deposits) {
    clearRanges.push(`${TABS.DEPOSITS}!A:H`);
    updates.push({ range: `${TABS.DEPOSITS}!A1`, values: serializeDeposits(snapshot.deposits) });
  }
  if (opts.rates) {
    clearRanges.push(`${TABS.RATES}!A:B`);
    updates.push({ range: `${TABS.RATES}!A1`, values: serializeRates(snapshot.rates) });
  }
  if (opts.screenshots) {
    clearRanges.push(`${TABS.SCREENSHOTS}!A:ZZ`);
    updates.push({
      range: `${TABS.SCREENSHOTS}!A1`,
      values: serializeScreenshots(snapshot.uploads),
    });
  }
  if (opts.audit) {
    clearRanges.push(`${TABS.AUDIT_LOGS}!A:H`);
    updates.push({
      range: `${TABS.AUDIT_LOGS}!A1`,
      values: serializeAuditLogs(snapshot.auditLogs),
    });
  }

  await googleApiFetch(`${SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchClear`, {
    method: 'POST',
    body: JSON.stringify({ ranges: clearRanges }),
  });

  await googleApiFetch(`${SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchUpdate`, {
    method: 'POST',
    body: JSON.stringify({
      valueInputOption: 'RAW',
      data: updates,
    }),
  });
}

function createAuditEntry(
  logs: AuditLogRecord[],
  params: {
    actor: string;
    action: string;
    entityType: string;
    entityId?: string | null;
    summary: string;
    details?: unknown;
  },
): AuditLogRecord {
  const maxId = logs.reduce((m, l) => Math.max(m, l.id), 0);
  return {
    id: maxId + 1,
    actor: params.actor,
    action: params.action,
    entityType: params.entityType,
    entityId: params.entityId ?? null,
    summary: params.summary,
    detailsJson: params.details !== undefined ? JSON.stringify(params.details) : null,
    createdAt: new Date().toISOString(),
  };
}

export const googleSheetsLedger = {
  async createPartnerDeskSpreadsheet(
    title = 'PartnerDesk — Options Trading Ledger',
  ): Promise<{ spreadsheetId: string; spreadsheetUrl: string; spreadsheetTitle: string }> {
    const created = await googleApiFetch<{
      spreadsheetId: string;
      spreadsheetUrl: string;
      properties?: { title?: string };
    }>(SHEETS_BASE, {
      method: 'POST',
      body: JSON.stringify({
        properties: { title },
        sheets: ALL_TAB_NAMES.map((tabTitle) => ({
          properties: {
            title: tabTitle,
            gridProperties: {
              columnCount: tabTitle === TABS.SCREENSHOTS ? 40 : 36,
            },
          },
        })),
      }),
    });

    await googleApiFetch(
      `${SHEETS_BASE}/${encodeURIComponent(created.spreadsheetId)}/values:batchUpdate`,
      {
        method: 'POST',
        body: JSON.stringify({
          valueInputOption: 'RAW',
          data: [
            { range: `${TABS.SUMMARY}!A1`, values: buildSummaryRows([], [], 'System') },
            { range: `${TABS.TRADE_GROUPS}!A1`, values: [TRADE_GROUP_HEADERS] },
            { range: `${TABS.ORDER_LEGS}!A1`, values: [ORDER_LEG_HEADERS] },
            { range: `${TABS.DEPOSITS}!A1`, values: [DEPOSIT_HEADERS] },
            { range: `${TABS.RATES}!A1`, values: serializeRates(DEFAULT_CHARGE_RATES) },
            { range: `${TABS.AUDIT_LOGS}!A1`, values: [AUDIT_HEADERS] },
            { range: `${TABS.SCREENSHOTS}!A1`, values: serializeScreenshots([]) },
            { range: `${TABS.GROUP_SCREENSHOTS}!A1`, values: [GROUP_SCREENSHOT_HEADERS] },
          ],
        }),
      },
    );

    ensuredSheetsCache.add(created.spreadsheetId);
    return {
      spreadsheetId: created.spreadsheetId,
      spreadsheetUrl: created.spreadsheetUrl,
      spreadsheetTitle: created.properties?.title || title,
    };
  },

  async connectExistingSpreadsheet(
    rawUrlOrId: string,
  ): Promise<{ spreadsheetId: string; spreadsheetUrl: string; spreadsheetTitle: string }> {
    const spreadsheetId = extractSpreadsheetId(rawUrlOrId);
    if (!spreadsheetId) {
      throw new Error('Paste a valid Google Sheets URL or Spreadsheet ID.');
    }
    const info = await ensureSpreadsheetTabs(spreadsheetId);
    ensuredSheetsCache.add(spreadsheetId);
    return {
      spreadsheetId,
      spreadsheetUrl: info.spreadsheetUrl,
      spreadsheetTitle: info.title,
    };
  },

  async shareSpreadsheetWithPartner(spreadsheetId: string, email: string): Promise<void> {
    const cleanEmail = email.trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanEmail)) {
      throw new Error('Enter a valid Gmail / Google account email address.');
    }
    await googleApiFetch(
      `${DRIVE_BASE}/${encodeURIComponent(spreadsheetId)}/permissions?sendNotificationEmail=true`,
      {
        method: 'POST',
        body: JSON.stringify({
          type: 'user',
          role: 'writer',
          emailAddress: cleanEmail,
        }),
      },
    );
  },

  async checkDuplicateScreenshot(
    spreadsheetId: string,
    imageDataUrl: string,
  ): Promise<StoredUploadRecord | null> {
    const [hash, snap] = await Promise.all([
      sha256Hex(imageDataUrl),
      readFullSnapshot(spreadsheetId),
    ]);
    return snap.uploads.find((u) => u.imageHash === hash) || null;
  },

  async getDashboard(spreadsheetId: string): Promise<SheetDashboardPayload> {
    const snap = await readFullSnapshot(spreadsheetId);
    const trades = hydrateTrades(
      snap.tradeGroups,
      snap.orders,
      snap.uploads,
      snap.groupScreenshots,
    );
    const sortedDeposits = [...snap.deposits].sort(
      (a, b) => b.depositDate.localeCompare(a.depositDate) || b.id - a.id,
    );
    const { summary, settlementPreview } = buildSummaryFromData(trades, sortedDeposits);
    return {
      summary,
      settlementPreview,
      recentTrades: trades.slice(0, 8),
      openPositions: trades.filter((t) => t.openQty > 0),
      recentDeposits: sortedDeposits.slice(0, 6),
      recentSettlements: [],
    };
  },

  async listTrades(
    spreadsheetId: string,
  ): Promise<{ trades: TradeGroupRecord[]; rates: ChargeRates }> {
    const snap = await readFullSnapshot(spreadsheetId);
    const trades = hydrateTrades(
      snap.tradeGroups,
      snap.orders,
      snap.uploads,
      snap.groupScreenshots,
    );
    return { trades, rates: snap.rates };
  },

  async saveOrders(
    spreadsheetId: string,
    params: {
      actor: PartnerId;
      screenshotDataUrl?: string | null;
      originalFilename?: string | null;
      orders: ExtractedOrderLeg[];
    },
  ): Promise<{
    uploadId: number | null;
    uploadedAt: string;
    affectedTrades: TradeGroupRecord[];
  }> {
    const snap = await readFullSnapshot(spreadsheetId);
    const uploadedAt = new Date().toISOString();
    let screenshotsChanged = false;
    let uploadId: number | null = null;

    if (params.screenshotDataUrl) {
      const hash = await sha256Hex(params.screenshotDataUrl);
      const dup = snap.uploads.find((u) => u.imageHash === hash);
      if (dup) {
        throw new Error(
          `Duplicate screenshot detected (already saved in sheet on ${new Date(
            dup.uploadedAt,
          ).toLocaleString()} by ${dup.uploadedBy}).`,
        );
      }
      const maxUploadId = snap.uploads.reduce((m, u) => Math.max(m, u.id), 0);
      uploadId = maxUploadId + 1;
      snap.uploads.push({
        id: uploadId,
        imageHash: hash,
        imageDataUrl: params.screenshotDataUrl,
        originalFilename: params.originalFilename || null,
        uploadedBy: params.actor,
        uploadedAt,
      });
      screenshotsChanged = true;
    }

    let maxGroupId = snap.tradeGroups.reduce((m, g) => Math.max(m, g.id), 0);
    let maxOrderId = snap.orders.reduce((m, o) => Math.max(m, o.id), 0);
    const touchedGroupIds = new Set<number>();

    for (const raw of params.orders) {
      const stockName = raw.stockName.trim().toUpperCase();
      const tradingSymbol = (raw.tradingSymbol || `${stockName} ${raw.strikePrice} ${raw.optionType}`)
        .trim()
        .toUpperCase();
      const expiryDate = raw.expiryDate.trim();
      const tradingDate = raw.tradingDate.trim();
      const strikePrice = Number(raw.strikePrice);
      const optionType: OptionType = raw.optionType === 'PE' ? 'PE' : 'CE';
      const side: OrderSide = raw.side === 'SELL' ? 'SELL' : 'BUY';
      const quantity = Math.round(Number(raw.quantity));
      const lots =
        raw.lots !== null && raw.lots !== undefined && Number(raw.lots) > 0
          ? Number(raw.lots)
          : null;
      const executionPrice = Number(raw.executionPrice);
      const amount = round2(
        Number.isFinite(raw.amount) && raw.amount > 0 ? raw.amount : quantity * executionPrice,
      );
      const executionTime = (raw.executionTime || '09:15 AM').trim();

      const groupKey = buildTradeGroupKey({
        tradingDate,
        stockName,
        expiryDate,
        strikePrice,
        optionType,
      });

      let group = snap.tradeGroups.find((g) => g.groupKey === groupKey);
      let groupId: number;
      if (!group) {
        maxGroupId += 1;
        groupId = maxGroupId;
        group = {
          id: groupId,
          groupKey,
          tradingDate,
          uploadedAt,
          stockName,
          tradingSymbol,
          expiryDate,
          strikePrice,
          optionType,
          buyQty: 0,
          sellQty: 0,
          avgBuyPrice: 0,
          avgSellPrice: 0,
          totalBuyAmount: 0,
          totalSellAmount: 0,
          realisedQty: 0,
          openQty: 0,
          openSide: null,
          grossPnl: 0,
          brokerage: 0,
          stt: 0,
          exchangeTxnCharge: 0,
          sebiCharge: 0,
          stampDuty: 0,
          ipftCharge: 0,
          gst: 0,
          otherCharges: 0,
          totalCharges: 0,
          netPnl: 0,
          chargesMode: 'ESTIMATED',
          contractNoteRef: null,
          notes: null,
          updatedAt: uploadedAt,
        };
        snap.tradeGroups.push(group);
      } else {
        groupId = group.id;
      }

      maxOrderId += 1;
      snap.orders.push({
        id: maxOrderId,
        tradeGroupId: groupId,
        uploadId,
        tradingDate,
        uploadedAt,
        stockName,
        tradingSymbol,
        expiryDate,
        strikePrice,
        optionType,
        side,
        quantity,
        lots,
        executionPrice,
        amount,
        executionTime,
        createdBy: params.actor,
        createdAt: uploadedAt,
      });

      if (
        uploadId !== null &&
        !snap.groupScreenshots.some((l) => l.tradeGroupId === groupId && l.uploadId === uploadId)
      ) {
        snap.groupScreenshots.push({ tradeGroupId: groupId, uploadId });
      }
      touchedGroupIds.add(groupId);
    }

    snap.tradeGroups = snap.tradeGroups.map((g) =>
      touchedGroupIds.has(g.id) ? recomputeGroupInMemory(g, snap.orders, snap.rates) : g,
    );

    const logEntry = createAuditEntry(snap.auditLogs, {
      actor: params.actor,
      action: 'TRADE_UPLOADED',
      entityType: 'trade_group',
      entityId: Array.from(touchedGroupIds).join(','),
      summary: `Saved ${params.orders.length} order leg(s) into ${touchedGroupIds.size} trade group(s) in Google Sheet.`,
      details: { uploadId, filename: params.originalFilename || null },
    });
    snap.auditLogs = [logEntry, ...snap.auditLogs];

    await writeSnapshotTabs(spreadsheetId, snap, params.actor, {
      trades: true,
      screenshots: screenshotsChanged,
      audit: true,
    });

    const hydrated = hydrateTrades(
      snap.tradeGroups,
      snap.orders,
      snap.uploads,
      snap.groupScreenshots,
    );
    return {
      uploadId,
      uploadedAt,
      affectedTrades: hydrated.filter((t) => touchedGroupIds.has(t.id)),
    };
  },

  async updateTrade(
    spreadsheetId: string,
    tradeGroupId: number,
    actor: PartnerId,
    payload: Record<string, unknown>,
  ): Promise<TradeGroupRecord> {
    const snap = await readFullSnapshot(spreadsheetId);
    const idx = snap.tradeGroups.findIndex((g) => g.id === tradeGroupId);
    if (idx === -1) throw new Error('Trade not found in shared Google Sheet.');
    const existing = snap.tradeGroups[idx];

    const tradingDate =
      typeof payload.tradingDate === 'string' && payload.tradingDate.trim()
        ? payload.tradingDate.trim()
        : existing.tradingDate;
    const stockName =
      typeof payload.stockName === 'string' && payload.stockName.trim()
        ? payload.stockName.trim().toUpperCase()
        : existing.stockName;
    const expiryDate =
      typeof payload.expiryDate === 'string' && payload.expiryDate.trim()
        ? payload.expiryDate.trim()
        : existing.expiryDate;
    const strikePrice =
      payload.strikePrice !== undefined ? Number(payload.strikePrice) : existing.strikePrice;
    const optionType: OptionType =
      payload.optionType === 'PE'
        ? 'PE'
        : payload.optionType === 'CE'
          ? 'CE'
          : existing.optionType;
    const tradingSymbol =
      typeof payload.tradingSymbol === 'string' && payload.tradingSymbol.trim()
        ? payload.tradingSymbol.trim().toUpperCase()
        : `${stockName} ${strikePrice} ${optionType}`;
    const notes =
      payload.notes !== undefined
        ? typeof payload.notes === 'string'
          ? payload.notes
          : null
        : existing.notes;

    const groupKey = buildTradeGroupKey({
      tradingDate,
      stockName,
      expiryDate,
      strikePrice,
      optionType,
    });

    const cn = payload.contractNote as Record<string, unknown> | undefined;
    let chargesMode: ChargesMode = existing.chargesMode;
    let contractNoteRef = existing.contractNoteRef;
    let brokerage = existing.brokerage;
    let stt = existing.stt;
    let exchangeTxnCharge = existing.exchangeTxnCharge;
    let ipftCharge = existing.ipftCharge;
    let sebiCharge = existing.sebiCharge;
    let stampDuty = existing.stampDuty;
    let gst = existing.gst;
    let otherCharges = existing.otherCharges;

    if (cn) {
      if (cn.mode === 'CONTRACT_NOTE') {
        chargesMode = 'CONTRACT_NOTE';
        contractNoteRef =
          typeof cn.contractNoteRef === 'string' ? cn.contractNoteRef : existing.contractNoteRef;
        brokerage = Number(cn.brokerage ?? existing.brokerage);
        stt = Number(cn.stt ?? existing.stt);
        exchangeTxnCharge = Number(cn.exchangeTxnCharge ?? existing.exchangeTxnCharge);
        ipftCharge = Number(cn.ipftCharge ?? existing.ipftCharge);
        sebiCharge = Number(cn.sebiCharge ?? existing.sebiCharge);
        stampDuty = Number(cn.stampDuty ?? existing.stampDuty);
        gst = Number(cn.gst ?? existing.gst);
        otherCharges = Number(cn.otherCharges ?? existing.otherCharges);
      } else {
        chargesMode = 'ESTIMATED';
        contractNoteRef = null;
      }
    }

    const updatedGroup: TradeGroupRecord = {
      ...existing,
      groupKey,
      tradingDate,
      stockName,
      tradingSymbol,
      expiryDate,
      strikePrice,
      optionType,
      notes,
      chargesMode,
      contractNoteRef,
      brokerage,
      stt,
      exchangeTxnCharge,
      ipftCharge,
      sebiCharge,
      stampDuty,
      gst,
      otherCharges,
      updatedAt: new Date().toISOString(),
    };

    if (Array.isArray(payload.orders)) {
      let maxOrderId = snap.orders.reduce((m, o) => Math.max(m, o.id), 0);
      const otherOrders = snap.orders.filter((o) => o.tradeGroupId !== tradeGroupId);
      const newOrders: TradeOrderRecord[] = [];
      for (const rawOrd of payload.orders as Record<string, unknown>[]) {
        let oid =
          typeof rawOrd.id === 'number' && rawOrd.id > 0 ? rawOrd.id : 0;
        if (oid === 0) {
          maxOrderId += 1;
          oid = maxOrderId;
        }
        const qty = Math.round(Number(rawOrd.quantity));
        const price = Number(rawOrd.executionPrice);
        const amt = round2(
          Number.isFinite(Number(rawOrd.amount)) && Number(rawOrd.amount) > 0
            ? Number(rawOrd.amount)
            : qty * price,
        );
        newOrders.push({
          id: oid,
          tradeGroupId,
          uploadId: null,
          tradingDate,
          uploadedAt: existing.uploadedAt,
          stockName,
          tradingSymbol,
          expiryDate,
          strikePrice,
          optionType,
          side: rawOrd.side === 'SELL' ? 'SELL' : 'BUY',
          quantity: qty,
          lots:
            rawOrd.lots !== null && rawOrd.lots !== undefined ? Number(rawOrd.lots) : null,
          executionPrice: price,
          amount: amt,
          executionTime: String(rawOrd.executionTime || '09:15 AM'),
          createdBy: actor,
          createdAt: new Date().toISOString(),
        });
      }
      snap.orders = [...otherOrders, ...newOrders];
    } else {
      snap.orders = snap.orders.map((o) =>
        o.tradeGroupId === tradeGroupId
          ? {
              ...o,
              tradingDate,
              stockName,
              tradingSymbol,
              expiryDate,
              strikePrice,
              optionType,
            }
          : o,
      );
    }

    const recomputed = recomputeGroupInMemory(updatedGroup, snap.orders, snap.rates);
    snap.tradeGroups[idx] = recomputed;

    const logEntry = createAuditEntry(snap.auditLogs, {
      actor,
      action: cn ? 'CONTRACT_NOTE_UPDATED' : 'TRADE_UPDATED',
      entityType: 'trade_group',
      entityId: String(tradeGroupId),
      summary: `${recomputed.stockName} ${recomputed.strikePrice} ${recomputed.optionType} updated in Google Sheet (${recomputed.chargesMode}).`,
    });
    snap.auditLogs = [logEntry, ...snap.auditLogs];

    await writeSnapshotTabs(spreadsheetId, snap, actor, {
      trades: true,
      audit: true,
    });

    const hydrated = hydrateTrades(
      snap.tradeGroups,
      snap.orders,
      snap.uploads,
      snap.groupScreenshots,
    );
    const result = hydrated.find((g) => g.id === tradeGroupId);
    if (!result) throw new Error('Trade not found after Google Sheet update.');
    return result;
  },

  async deleteTrade(
    spreadsheetId: string,
    tradeGroupId: number,
    actor: PartnerId,
  ): Promise<void> {
    const snap = await readFullSnapshot(spreadsheetId);
    const existing = snap.tradeGroups.find((g) => g.id === tradeGroupId);
    if (!existing) throw new Error('Trade not found in Google Sheet.');

    snap.tradeGroups = snap.tradeGroups.filter((g) => g.id !== tradeGroupId);
    snap.orders = snap.orders.filter((o) => o.tradeGroupId !== tradeGroupId);
    snap.groupScreenshots = snap.groupScreenshots.filter((l) => l.tradeGroupId !== tradeGroupId);

    const stillReferencedUploadIds = new Set<number>();
    for (const o of snap.orders) {
      if (o.uploadId !== null) stillReferencedUploadIds.add(o.uploadId);
    }
    for (const l of snap.groupScreenshots) {
      stillReferencedUploadIds.add(l.uploadId);
    }
    const prevUploadsCount = snap.uploads.length;
    snap.uploads = snap.uploads.filter((u) => stillReferencedUploadIds.has(u.id));
    const screenshotsChanged = snap.uploads.length !== prevUploadsCount;

    const logEntry = createAuditEntry(snap.auditLogs, {
      actor,
      action: 'TRADE_DELETED',
      entityType: 'trade_group',
      entityId: String(tradeGroupId),
      summary: `Deleted ${existing.stockName} ${existing.strikePrice} ${existing.optionType} (${existing.tradingDate}) from Google Sheet.`,
    });
    snap.auditLogs = [logEntry, ...snap.auditLogs];

    await writeSnapshotTabs(spreadsheetId, snap, actor, {
      trades: true,
      screenshots: screenshotsChanged,
      audit: true,
    });
  },

  async listDeposits(
    spreadsheetId: string,
  ): Promise<{ deposits: DepositRecord[]; summary: DashboardSummary }> {
    const snap = await readFullSnapshot(spreadsheetId);
    const trades = hydrateTrades(
      snap.tradeGroups,
      snap.orders,
      snap.uploads,
      snap.groupScreenshots,
    );
    const sorted = [...snap.deposits].sort(
      (a, b) => b.depositDate.localeCompare(a.depositDate) || b.id - a.id,
    );
    const { summary } = buildSummaryFromData(trades, sorted);
    return { deposits: sorted, summary };
  },

  async addDeposit(
    spreadsheetId: string,
    params: {
      partner: PartnerId;
      entryType?: FundEntryType;
      amount: number;
      depositDate: string;
      referenceNote?: string;
      actor: PartnerId;
    },
  ): Promise<DepositRecord> {
    const snap = await readFullSnapshot(spreadsheetId);
    const maxId = snap.deposits.reduce((m, d) => Math.max(m, d.id), 0);
    const entryType: FundEntryType = params.entryType === 'WITHDRAWAL' ? 'WITHDRAWAL' : 'DEPOSIT';

    const deposit: DepositRecord = {
      id: maxId + 1,
      partner: params.partner,
      entryType,
      amount: round2(params.amount),
      depositDate: params.depositDate,
      referenceNote: params.referenceNote?.trim() || null,
      createdBy: params.actor,
      createdAt: new Date().toISOString(),
    };
    snap.deposits.push(deposit);

    const logEntry = createAuditEntry(snap.auditLogs, {
      actor: params.actor,
      action: entryType === 'WITHDRAWAL' ? 'WITHDRAWAL_RECORDED' : 'DEPOSIT_ADDED',
      entityType: 'deposit',
      entityId: String(deposit.id),
      summary:
        entryType === 'WITHDRAWAL'
          ? `Recorded ₹${deposit.amount.toFixed(2)} fund withdrawal for ${deposit.partner} in Google Sheet.`
          : `Recorded ₹${deposit.amount.toFixed(2)} capital deposit for ${deposit.partner} in Google Sheet.`,
    });
    snap.auditLogs = [logEntry, ...snap.auditLogs];

    await writeSnapshotTabs(spreadsheetId, snap, params.actor, {
      deposits: true,
      audit: true,
    });
    return deposit;
  },

  async deleteDeposit(spreadsheetId: string, id: number, actor: PartnerId): Promise<void> {
    const snap = await readFullSnapshot(spreadsheetId);
    const existing = snap.deposits.find((d) => d.id === id);
    if (!existing) throw new Error('Fund entry not found in Google Sheet.');

    snap.deposits = snap.deposits.filter((d) => d.id !== id);
    const isWithdrawal = existing.entryType === 'WITHDRAWAL';
    const logEntry = createAuditEntry(snap.auditLogs, {
      actor,
      action: isWithdrawal ? 'WITHDRAWAL_DELETED' : 'DEPOSIT_DELETED',
      entityType: 'deposit',
      entityId: String(id),
      summary: `Removed ${isWithdrawal ? 'withdrawal' : 'deposit'} #${id} (₹${existing.amount.toFixed(2)} by ${existing.partner}) from Google Sheet.`,
    });
    snap.auditLogs = [logEntry, ...snap.auditLogs];

    await writeSnapshotTabs(spreadsheetId, snap, actor, {
      deposits: true,
      audit: true,
    });
  },

  async getRates(spreadsheetId: string): Promise<ChargeRates> {
    const snap = await readFullSnapshot(spreadsheetId);
    return snap.rates;
  },

  async updateRates(
    spreadsheetId: string,
    rates: ChargeRates & { recalculateExisting?: boolean },
    actor: PartnerId,
  ): Promise<ChargeRates> {
    const snap = await readFullSnapshot(spreadsheetId);
    const cleanRates: ChargeRates = {
      brokeragePerOrder: Number(rates.brokeragePerOrder),
      sttSellPct: Number(rates.sttSellPct),
      exchangeTxnPct: Number(rates.exchangeTxnPct),
      ipftPct: Number(rates.ipftPct),
      sebiPct: Number(rates.sebiPct),
      stampDutyBuyPct: Number(rates.stampDutyBuyPct),
      gstPct: Number(rates.gstPct),
    };
    snap.rates = cleanRates;

    let tradesChanged = false;
    if (rates.recalculateExisting) {
      snap.tradeGroups = snap.tradeGroups.map((g) => {
        if (g.chargesMode === 'ESTIMATED') {
          tradesChanged = true;
          return recomputeGroupInMemory(g, snap.orders, cleanRates);
        }
        return g;
      });
    }

    const logEntry = createAuditEntry(snap.auditLogs, {
      actor,
      action: 'RATES_UPDATED',
      entityType: 'settings',
      summary: `Updated charge rates in Google Sheet${rates.recalculateExisting ? ' and recalculated estimated trades' : ''}.`,
    });
    snap.auditLogs = [logEntry, ...snap.auditLogs];

    await writeSnapshotTabs(spreadsheetId, snap, actor, {
      rates: true,
      trades: tradesChanged,
      audit: true,
    });
    return cleanRates;
  },

  async listAuditLogs(spreadsheetId: string): Promise<AuditLogRecord[]> {
    const snap = await readFullSnapshot(spreadsheetId);
    return [...snap.auditLogs].sort((a, b) => b.id - a.id);
  },

  async exportBackupPayload(spreadsheetId: string): Promise<SheetBackupPayload> {
    const snap = await readFullSnapshot(spreadsheetId);
    return {
      format: 'partnerdesk-local-backup-v1',
      folderName: 'PartnerDesk_GoogleSheet',
      exportedAt: new Date().toISOString(),
      chargeRates: snap.rates,
      deposits: [...snap.deposits].sort(
        (a, b) => b.depositDate.localeCompare(a.depositDate) || b.id - a.id,
      ),
      uploads: snap.uploads,
      tradeGroups: [...snap.tradeGroups].sort(
        (a, b) => b.tradingDate.localeCompare(a.tradingDate) || b.id - a.id,
      ),
      orders: [...snap.orders].sort((a, b) => a.id - b.id),
      auditLogs: [...snap.auditLogs].sort((a, b) => b.id - a.id),
    };
  },

  downloadJsonFromPayload(payload: SheetBackupPayload): void {
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `PartnerDesk_Sheet_Backup_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  },

  async importMergeBackupPayload(
    spreadsheetId: string,
    payload: Partial<SheetBackupPayload>,
    actor: PartnerId,
  ): Promise<{ addedTrades: number; addedDeposits: number; addedScreenshots: number }> {
    if (!payload || !Array.isArray(payload.deposits) || !Array.isArray(payload.tradeGroups)) {
      throw new Error('Selected file is not a valid PartnerDesk backup.');
    }
    const snap = await readFullSnapshot(spreadsheetId);

    const uploadIdRemap = new Map<number, number>();
    const groupIdRemap = new Map<number, number>();

    let maxUploadId = snap.uploads.reduce((m, u) => Math.max(m, u.id), 0);
    let maxGroupId = snap.tradeGroups.reduce((m, g) => Math.max(m, g.id), 0);
    let maxOrderId = snap.orders.reduce((m, o) => Math.max(m, o.id), 0);
    let maxDepositId = snap.deposits.reduce((m, d) => Math.max(m, d.id), 0);

    const existingUploadByHash = new Map(snap.uploads.map((u) => [u.imageHash, u]));
    const existingGroupByKey = new Map(snap.tradeGroups.map((g) => [g.groupKey, g]));

    let addedScreenshots = 0;
    for (const u of payload.uploads || []) {
      const hash = u.imageHash || (await sha256Hex(u.imageDataUrl));
      const found = existingUploadByHash.get(hash);
      if (found) {
        uploadIdRemap.set(u.id, found.id);
      } else {
        maxUploadId += 1;
        const newU: StoredUploadRecord = { ...u, id: maxUploadId, imageHash: hash };
        snap.uploads.push(newU);
        existingUploadByHash.set(hash, newU);
        uploadIdRemap.set(u.id, maxUploadId);
        addedScreenshots += 1;
      }
    }

    let addedTrades = 0;
    for (const g of payload.tradeGroups || []) {
      const found = existingGroupByKey.get(g.groupKey);
      if (found) {
        groupIdRemap.set(g.id, found.id);
      } else {
        maxGroupId += 1;
        const cleanG: TradeGroupRecord = { ...g, id: maxGroupId };
        delete cleanG.orders;
        delete cleanG.screenshots;
        snap.tradeGroups.push(cleanG);
        existingGroupByKey.set(g.groupKey, cleanG);
        groupIdRemap.set(g.id, maxGroupId);
        addedTrades += 1;
      }
    }

    const orderSignature = (o: TradeOrderRecord) =>
      `${o.tradeGroupId}|${o.side}|${o.quantity}|${o.executionPrice}|${o.amount}|${o.executionTime}`;
    const existingOrderSigs = new Set(snap.orders.map(orderSignature));

    for (const o of payload.orders || []) {
      const targetGroupId = groupIdRemap.get(o.tradeGroupId);
      if (!targetGroupId) continue;
      const mappedUploadId = o.uploadId ? (uploadIdRemap.get(o.uploadId) ?? null) : null;
      const candidate: TradeOrderRecord = {
        ...o,
        id: 0,
        tradeGroupId: targetGroupId,
        uploadId: mappedUploadId,
      };
      const sig = orderSignature(candidate);
      if (!existingOrderSigs.has(sig)) {
        maxOrderId += 1;
        candidate.id = maxOrderId;
        snap.orders.push(candidate);
        existingOrderSigs.add(sig);
      }
      if (
        mappedUploadId &&
        !snap.groupScreenshots.some(
          (l) => l.tradeGroupId === targetGroupId && l.uploadId === mappedUploadId,
        )
      ) {
        snap.groupScreenshots.push({ tradeGroupId: targetGroupId, uploadId: mappedUploadId });
      }
    }

    const depSignature = (d: DepositRecord) =>
      `${d.depositDate}|${d.partner}|${d.entryType || 'DEPOSIT'}|${d.amount}|${d.createdAt}`;
    const existingDepSigs = new Set(snap.deposits.map(depSignature));
    let addedDeposits = 0;

    for (const d of payload.deposits || []) {
      if (existingDepSigs.has(depSignature(d))) continue;
      maxDepositId += 1;
      snap.deposits.push({ ...d, id: maxDepositId });
      existingDepSigs.add(depSignature(d));
      addedDeposits += 1;
    }

    const touchedIds = new Set(groupIdRemap.values());
    snap.tradeGroups = snap.tradeGroups.map((g) =>
      touchedIds.has(g.id) ? recomputeGroupInMemory(g, snap.orders, snap.rates) : g,
    );

    const logEntry = createAuditEntry(snap.auditLogs, {
      actor,
      action: 'SHEET_BACKUP_MERGED',
      entityType: 'backup',
      summary: `Imported backup into Google Sheet (+${addedTrades} trade(s), +${addedDeposits} fund entry(s), +${addedScreenshots} screenshot(s)).`,
    });
    snap.auditLogs = [logEntry, ...snap.auditLogs];

    await writeSnapshotTabs(spreadsheetId, snap, actor, {
      trades: true,
      deposits: true,
      screenshots: addedScreenshots > 0,
      audit: true,
    });

    return { addedTrades, addedDeposits, addedScreenshots };
  },
};
