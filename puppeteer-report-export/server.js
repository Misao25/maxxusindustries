// server.js (Batch all updates into one request)

const express = require('express');
const puppeteer = require('puppeteer');
const https = require('https');
const xlsx = require('xlsx');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Auth setup for Google Sheets
const credentials = JSON.parse(
  Buffer.from(process.env.GCP_CREDENTIALS_B64, 'base64').toString('utf-8')
);
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

const SHEET_ID = '1mrw-AMbVWnz1Cp4ksjR0W0eTDz0cUiA-zjThrzcIRnY';

// Convert JS Date back to Excel serial
const jsDateToExcelSerial = (dateObj) => {
  if (!dateObj || isNaN(dateObj)) return '';
  const epoch = Date.UTC(1899, 11, 30); // 1899-12-30 UTC
  let serial = (dateObj.getTime() - epoch) / (86400 * 1000);
  if (serial >= 61) serial += 1; // Excel leap year bug
  return serial;
};

app.get('/generate-report', async (req, res) => {
  const { from, to } = req.query;
  if (!from || !to) return res.status(400).send('Missing from or to date');

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox'],
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 800 });

  try {
    // --- login and navigate ---
    await page.goto('https://app.ecomdash.com/?returnUrl=%2fReporting');
    await page.type('input#UserName', process.env.LOGIN_EMAIL);
    await page.click('input#submit');
    await page.waitForSelector('input#Password');
    await page.type('input#Password', process.env.LOGIN_PASS);
    await page.click('input#submit');
    await page.waitForNavigation({ waitUntil: 'networkidle2' });

    await page.waitForSelector('div#mostPopular-SalesOrdersReport');
    await page.click(
      'div#mostPopular-SalesOrdersReport div.buttonDiv a.albany-btn.albany-btn--primary'
    );
    await page.waitForSelector('form#GenerateReport', { visible: true });

    await page.click('#ReportStartDate', { clickCount: 3 });
    await page.keyboard.press('Backspace');
    await page.type('#ReportStartDate', from);
    await page.$eval('#ReportStartDate', (el) => el.blur());

    await page.click('#ReportEndDate', { clickCount: 3 });
    await page.keyboard.press('Backspace');
    await page.type('#ReportEndDate', to);
    await page.$eval('#ReportEndDate', (el) => el.blur());

    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle2' }),
      page.click('a#GenerateDateRestrictionReport'),
    ]);

    await page.waitForSelector('table');
    const timestampStr = await page.$eval(
      'table tbody tr td:nth-child(1)',
      (el) => el.textContent.trim()
    );

    const historyUrl = 'https://dashboard.ecomdash.com/Support/ReportingHistory';
    let downloadUrl = null;

    for (let i = 0; i < 30; i++) {
      await page.goto(historyUrl, { waitUntil: 'networkidle2' });
      await page.waitForSelector('table');
      const rows = await page.$$('table tbody tr');

      for (const row of rows) {
        const rowTimestamp = await row.$eval(
          'td:nth-child(1)',
          (el) => el.textContent.trim()
        );
        const status = await row.$eval(
          'td:nth-child(4)',
          (el) => el.textContent.trim()
        );

        if (rowTimestamp === timestampStr && status === 'Complete') {
          const linkEl = await row.$('td:nth-child(5) a[href$=".xlsx"]');
          if (linkEl) {
            downloadUrl = await linkEl.evaluate((a) => a.href);
            break;
          }
        }
      }
      if (downloadUrl) break;
      await new Promise((resolve) => setTimeout(resolve, 3000));
    }

    if (!downloadUrl) throw new Error('Report not found');

    // Download XLSX to memory
    const buffer = await new Promise((resolve, reject) => {
      const data = [];
      https.get(downloadUrl, (res) => {
        res.on('data', (chunk) => data.push(chunk));
        res.on('end', () => resolve(Buffer.concat(data)));
        res.on('error', reject);
      });
    });

    // Parse XLSX
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    let rows = xlsx.utils.sheet_to_json(firstSheet, { header: 1 });

    // Columns with dates to convert to serials
    const dateCols = [5, 6, 41]; // Invoice Date, Payment Received Date, Completed Date
    const fixedDateCells = [
      { row: 0, col: 23 },
      { row: 1449, col: 23 },
    ];

    rows = rows.map((row, index) => {
      if (index >= 2) {
        dateCols.forEach((colIndex) => {
          if (row[colIndex]) {
            let dateObj;
            if (typeof row[colIndex] === 'string' && row[colIndex].includes('/')) {
              dateObj = new Date(row[colIndex]);
            } else if (typeof row[colIndex] === 'number') {
              // already a serial, keep it
              return;
            } else {
              dateObj = new Date(row[colIndex]);
            }
            if (dateObj && !isNaN(dateObj)) {
              row[colIndex] = jsDateToExcelSerial(dateObj);
            } else {
              row[colIndex] = '';
            }
          } else {
            row[colIndex] = '';
          }
        });

        // Clean Order Notes
        if (row[36]) {
          row[36] = row[36].toString().replace(/[\r\n]+/g, ' ').trim();
        }
      }
      return row;
    });

    // Format fixed date cells
    fixedDateCells.forEach(({ row, col }) => {
      if (rows[row] && rows[row][col]) {
        let dateObj = new Date(rows[row][col]);
        if (!isNaN(dateObj)) {
          rows[row][col] = jsDateToExcelSerial(dateObj);
        } else {
          rows[row][col] = '';
        }
      }
    });

    // Sort rows by Invoice Date (serial value)
    const headerRows = rows.slice(0, 2);
    const dataRows = rows.slice(2);

    dataRows.sort((a, b) => {
      const dateA = a[5] || 0;
      const dateB = b[5] || 0;
      return dateB - dateA; // Descending
    });

    rows = [...headerRows, ...dataRows];

    // --- Sync with Google Sheets ---

    // Fetch existing data (full rows, not just column A)
    const existingDataResp = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'SalesMasterfile!A:AZ', // adjust range if more cols are used
    });
    const existingData = existingDataResp.data.values || [];

    // Build a map for fast lookup by ID
    const existingMap = new Map();
    existingData.forEach((row, i) => {
      const id = row[0]?.toString().trim();
      if (id) {
        existingMap.set(id, { index: i, row });
      }
    });

    const rowsToAppend = [];
    const rowsToUpdate = [];

    rows.slice(2).forEach((newRow) => {
      const id = newRow[0]?.toString().trim();
      if (!id || id === 'Date TypeCreate') return;

      const existing = existingMap.get(id);

      if (!existing) {
        rowsToAppend.push(newRow);
        return;
      }

      const existingRow = existing.row;

      const normalize = (v) =>
        v === undefined || v === null ? '' : v.toString().trim();
      const existingNormalized = existingRow.map(normalize);
      const newNormalized = newRow.map(normalize);

      const areSame =
        existingNormalized.join('|') === newNormalized.join('|');

      if (!areSame) {
        rowsToUpdate.push({
          rowIndex: existing.index + 1, // +1 since Sheets is 1-indexed
          values: newRow,
        });
      }
    });

    // Add headers if sheet is empty
    if (existingData.length === 0) {
      rowsToAppend.unshift(rows[0], rows[1]);
    }

    // --- Apply changes in bulk ---
    if (rowsToUpdate.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: rowsToUpdate.map(({ rowIndex, values }) => ({
            range: `SalesMasterfile!A${rowIndex}:AZ${rowIndex}`, // adjust AZ if more cols
            values: [values],
          })),
        },
      });
    }

    if (rowsToAppend.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: 'SalesMasterfile!A1',
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: rowsToAppend },
      });
    }

    await browser.close();

    const summary = `✅ Added ${rowsToAppend.length} new rows, updated ${rowsToUpdate.length} rows in SalesMasterfile`;
    console.log(summary);
    res.send(summary);
  } catch (err) {
    console.error(err);
    await browser.close();
    res.status(500).send('Error generating report: ' + err.message);
  }
});

app.listen(PORT, () =>
  console.log(`Server running on port ${PORT}`)
);
