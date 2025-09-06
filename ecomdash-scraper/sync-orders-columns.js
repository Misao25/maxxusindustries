// sync-orders-columns.js
// Fill Orders tab columns from masterfile (Google Sheets → Google Sheets), no scraping.
// - Uses FIRST match per orderId from masterfile.
// - Ensures headers exist in Orders.
// - Fills values row-aligned by orderId.
// - Default: fills blanks only (set FILL_ONLY_BLANKS=false to overwrite).

require('dotenv').config();
const { google } = require('googleapis');
const dayjs = require('dayjs');

// ───────────────────────────────────────────────────────────────────────────────
// Your updated column rules
// ───────────────────────────────────────────────────────────────────────────────
const COLUMN_RULES = [
  // { destHeader, masterAliases[], transform? }
  { destHeader: 'invoiceDate',          masterAliases: ['InvoiceDate'],          transform: formatDate },
  { destHeader: 'paymentReceivedDate',  masterAliases: ['PaymentReceivedDate'],  transform: formatDate },
  { destHeader: 'completedDate',        masterAliases: ['CompletedDate'],        transform: formatDate },

  { destHeader: 'customerEmail',        masterAliases: ['Email'] },
  { destHeader: 'shipToCity',           masterAliases: ['ShippingCity'] },
  { destHeader: 'shipToState',          masterAliases: ['ShippingState'] },
  { destHeader: 'shipToCountry',        masterAliases: ['ShippingCountry'] },
];

// Add your likely orderId header names here (case/spacing-insensitive matching)
const ORDER_ID_ALIASES = [
  'EcomdashID', 'orderId'
];

const FILL_ONLY_BLANKS = String(process.env.FILL_ONLY_BLANKS || 'true').toLowerCase() === 'true';

// ───────────────────────────────────────────────────────────────────────────────
// ENV & AUTH
// ───────────────────────────────────────────────────────────────────────────────
const credentials = JSON.parse(
  Buffer.from(process.env.GCP_CREDENTIALS_B64, 'base64').toString('utf-8')
);
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

const MASTERFILE_ID     = process.env.MASTERFILE_ID;
const MASTER_RANGE      = process.env.MASTER_RANGE || 'SalesMasterfile!A:AP';
const DESTINATION_ID    = process.env.DESTINATION_ID;
const ORDERS_SHEET_NAME = process.env.ORDERS_SHEET_NAME || 'Orders';

// ───────────────────────────────────────────────────────────────────────────────
// Helpers
// ───────────────────────────────────────────────────────────────────────────────
function normalizeHeader(h) {
  return String(h || '').trim().toLowerCase().replace(/\s+/g, '');
}

function findIndexByAliases(headerRow, aliases) {
  const map = new Map();
  headerRow.forEach((h, i) => map.set(normalizeHeader(h), i));
  for (const a of aliases) {
    const idx = map.get(normalizeHeader(a));
    if (idx !== undefined) return idx;
  }
  return -1;
}

function formatDate(v) {
  if (!v) return '';
  const d = new Date(v);
  if (Number.isNaN(d.getTime())) return String(v);
  return dayjs(d).format('YYYY/MM/DD');
}

function toA1Col(n) {
  // 0 → A, 25 → Z, 26 → AA ...
  let s = '';
  n = Number(n);
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

async function readRange(spreadsheetId, range) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING',
    majorDimension: 'ROWS',
  });
  return res.data.values || [];
}

async function writeHeaderRow(destHeaders) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: DESTINATION_ID,
    range: `${ORDERS_SHEET_NAME}!1:1`,
    valueInputOption: 'RAW',
    requestBody: { values: [destHeaders] },
  });
}

async function batchUpdateRanges(valueRanges) {
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: DESTINATION_ID,
    requestBody: { valueInputOption: 'RAW', data: valueRanges },
  });
}

// Read an entire single column (2..N) as a flat array (strings)
async function readColumn(spreadsheetId, sheetName, colIndex, rowCount) {
  const colLetter = toA1Col(colIndex);
  const range = `${sheetName}!${colLetter}2:${colLetter}${rowCount + 1}`;
  const rows = await readRange(spreadsheetId, range);
  return rows.map(r => (r && r[0] != null ? String(r[0]) : ''));
}

// ───────────────────────────────────────────────────────────────────────────────
// Main
// ───────────────────────────────────────────────────────────────────────────────
(async () => {
  // 1) Read master wide table
  const master = await readRange(MASTERFILE_ID, MASTER_RANGE);
  if (master.length < 2) {
    console.log('No master rows found (need header + data).');
    return;
  }
  const masterHeader = master[0];
  const masterRows = master.slice(1);

  const masterOrderIdIdx = findIndexByAliases(masterHeader, ORDER_ID_ALIASES);
  if (masterOrderIdIdx === -1) {
    console.error('❌ Could not find "orderId" in masterfile (check ORDER_ID_ALIASES).');
    return;
  }

  // Map needed master column indices for rules
  const ruleIndexMap = new Map(); // destHeader -> colIndex
  const missingInMaster = [];
  for (const rule of COLUMN_RULES) {
    const idx = findIndexByAliases(masterHeader, rule.masterAliases);
    if (idx === -1) missingInMaster.push(rule.destHeader);
    else ruleIndexMap.set(rule.destHeader, idx);
  }
  if (missingInMaster.length) {
    console.log('⚠️ Skipping columns not found in master:', missingInMaster.join(', '));
  }

  // Build FIRST-match master map: orderId -> { destHeader: value, ... }
  // We keep the FIRST row encountered per orderId and ignore later duplicates.
  const masterMap = new Map();
  let duplicateCount = 0;
  for (const r of masterRows) {
    const oid = String(r[masterOrderIdIdx] ?? '').trim();
    if (!oid) continue;
    if (masterMap.has(oid)) { duplicateCount++; continue; } // keep first, skip others

    const obj = {};
    for (const rule of COLUMN_RULES) {
      const idx = ruleIndexMap.get(rule.destHeader);
      if (idx === undefined) continue; // not present in master
      const raw = r[idx];
      const val = rule.transform ? rule.transform(raw) : (raw ?? '');
      obj[rule.destHeader] = val;
    }
    masterMap.set(oid, obj);
  }
  if (duplicateCount) {
    console.log(`ℹ️ Detected ${duplicateCount} duplicate master rows (kept first occurrence per orderId).`);
  }

  // 2) Read destination header and ensure orderId column exists
  const destHeaderRow = (await readRange(DESTINATION_ID, `${ORDERS_SHEET_NAME}!1:1`))[0] || [];
  const destHeaders = [...destHeaderRow];
  const destHeaderMap = new Map(destHeaders.map((h, i) => [normalizeHeader(h), i]));

  let destOrderIdCol = destHeaderMap.get(normalizeHeader('orderId'));
  if (destOrderIdCol === undefined) {
    // If no orderId header, add it at the beginning
    destHeaders.unshift('orderId');
    await writeHeaderRow(destHeaders);
    destOrderIdCol = 0;
  }

  // Refresh header map
  const headerMap = new Map(destHeaders.map((h, i) => [normalizeHeader(h), i]));

  // Read all orderId values to align row index → orderId
  const orderIdColLetter = toA1Col(destOrderIdCol);
  const destOrderIdValues = await readRange(
    DESTINATION_ID,
    `${ORDERS_SHEET_NAME}!${orderIdColLetter}2:${orderIdColLetter}`
  );
  const destOrderIds = destOrderIdValues.map(r => String((r && r[0]) || '').trim());
  const rowCount = destOrderIds.length;

  if (!rowCount) {
    console.log('Nothing to update (Orders sheet has no rows beyond header).');
    return;
  }

  // 3) Ensure headers for ALL rules exist (append as needed)
  const newlyAdded = [];
  for (const rule of COLUMN_RULES) {
    const key = normalizeHeader(rule.destHeader);
    if (!headerMap.has(key)) {
      destHeaders.push(rule.destHeader);
      newlyAdded.push(rule.destHeader);
    }
  }
  if (newlyAdded.length) {
    await writeHeaderRow(destHeaders);
    console.log(`➕ Added headers: ${newlyAdded.join(', ')}`);
  }

  // Rebuild header map after header write
  const finalHeaderMap = new Map(destHeaders.map((h, i) => [normalizeHeader(h), i]));

  // 4) Build and write values per target column
  // We process ALL rule columns. If FILL_ONLY_BLANKS=true, we preserve non-blank cells.
  const valueRanges = [];

  for (const rule of COLUMN_RULES) {
    const colIdx = finalHeaderMap.get(normalizeHeader(rule.destHeader));
    if (colIdx === undefined) continue; // shouldn't happen
    const colLetter = toA1Col(colIdx);

    // Read existing column values (to preserve non-blanks if needed)
    const existingCol = await readColumn(DESTINATION_ID, ORDERS_SHEET_NAME, colIdx, rowCount);

    const colValues = [];
    for (let i = 0; i < rowCount; i++) {
      const oid = destOrderIds[i];
      const existing = existingCol[i] ?? '';
      const firstMatch = oid ? masterMap.get(oid) : undefined;
      const candidate = firstMatch ? (firstMatch[rule.destHeader] ?? '') : '';

      if (FILL_ONLY_BLANKS) {
        colValues.push([existing ? existing : candidate]);
      } else {
        colValues.push([candidate]);
      }
    }

    valueRanges.push({
      range: `${ORDERS_SHEET_NAME}!${colLetter}2:${colLetter}${rowCount + 1}`,
      values: colValues,
    });
  }

  if (valueRanges.length) {
    await batchUpdateRanges(valueRanges);
    console.log(
      `✅ Updated columns in Orders (${FILL_ONLY_BLANKS ? 'filled blanks only' : 'overwrote all'}): ` +
      COLUMN_RULES.map(r => r.destHeader).join(', ')
    );
  } else {
    console.log('Nothing to write.');
  }
})().catch((err) => {
  console.error('❌ Error:', err.response?.data || err);
  process.exit(1);
});
