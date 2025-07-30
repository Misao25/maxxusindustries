// server.js (Shift all ecomdash data to the right)

const express = require('express');
const puppeteer = require('puppeteer');
const https = require('https');
const xlsx = require('xlsx');
const { google } = require('googleapis');
const fs = require('fs');
const stream = require('stream');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Auth setup for Google Sheets
const credentials = JSON.parse(Buffer.from(process.env.GCP_CREDENTIALS_B64, 'base64').toString('utf-8'));
const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
});
const sheets = google.sheets({ version: 'v4', auth });

const SHEET_ID = '1mrw-AMbVWnz1Cp4ksjR0W0eTDz0cUiA-zjThrzcIRnY';

// Helper function to create a unique key for duplicate detection
const createRowKey = (row) => {
    // Assuming columns for unique identification (adjust indices as needed)
    // Using order ID (column B/index 1) and date (column F/index 5) as unique identifiers
    const orderId = row[1] || '';
    const date = row[5] || '';
    return `${orderId}_${date}`;
};

// Helper function to add ID column at the beginning
const addIdColumn = (rows) => {
    return rows.map((row, index) => {
        if (index === 0) {
            // Add "ID" header to first row
            return ['ID', ...row];
        } else if (index === 1) {
            // Add empty cell or secondary header to second row
            return ['', ...row];
        } else {
            // Add incremental ID starting from 1 for data rows
            return [index - 1, ...row];
        }
    });
};

app.get('/generate-report', async (req, res) => {
    const { from, to } = req.query;
    if (!from || !to) return res.status(400).send('Missing from or to date');

    const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    try {
        await page.goto('https://app.ecomdash.com/?returnUrl=%2fReporting');
        await page.type('input#UserName', process.env.LOGIN_EMAIL);
        await page.click('input#submit');
        await page.waitForSelector('input#Password');
        await page.type('input#Password', process.env.LOGIN_PASS);
        await page.click('input#submit');
        await page.waitForNavigation({ waitUntil: 'networkidle2' });

        await page.waitForSelector('div#mostPopular-SalesOrdersReport');
        await page.click('div#mostPopular-SalesOrdersReport div.buttonDiv a.albany-btn.albany-btn--primary');
        await page.waitForSelector('form#GenerateReport', { visible: true });

        await page.click('#ReportStartDate', { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await page.type('#ReportStartDate', from);
        await page.$eval('#ReportStartDate', el => el.blur());

        await page.click('#ReportEndDate', { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await page.type('#ReportEndDate', to);
        await page.$eval('#ReportEndDate', el => el.blur());

        await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2' }),
            page.click('a#GenerateDateRestrictionReport')
        ]);

        await page.waitForSelector('table');
        const timestampStr = await page.$eval('table tbody tr td:nth-child(1)', el => el.textContent.trim());

        const historyUrl = 'https://dashboard.ecomdash.com/Support/ReportingHistory';
        let downloadUrl = null;

        for (let i = 0; i < 30; i++) {
            await page.goto(historyUrl, { waitUntil: 'networkidle2' });
            await page.waitForSelector('table');
            const rows = await page.$$('table tbody tr');

            for (const row of rows) {
                const rowTimestamp = await row.$eval('td:nth-child(1)', el => el.textContent.trim());
                const status = await row.$eval('td:nth-child(4)', el => el.textContent.trim());

                if (rowTimestamp === timestampStr && status === 'Complete') {
                    const linkEl = await row.$('td:nth-child(5) a[href$=".xlsx"]');
                    if (linkEl) {
                        downloadUrl = await linkEl.evaluate(a => a.href);
                        break;
                    }
                }
            }
            if (downloadUrl) break;
            await new Promise(resolve => setTimeout(resolve, 3000));
        }

        if (!downloadUrl) throw new Error('Report not found');

        // Download XLSX to memory
        const buffer = await new Promise((resolve, reject) => {
            const data = [];
            https.get(downloadUrl, res => {
                res.on('data', chunk => data.push(chunk));
                res.on('end', () => resolve(Buffer.concat(data)));
                res.on('error', reject);
            });
        });

        // Parse XLSX
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let newRows = xlsx.utils.sheet_to_json(firstSheet, { header: 1 });

        const excelDateToJS = serial => {
            if (!serial || isNaN(serial)) return '';
            const utc_days = Math.floor(serial - 25569);
            const utc_value = utc_days * 86400; 
            const date_info = new Date(utc_value * 1000);
            const fractionalDay = serial % 1;
            const totalSeconds = Math.round(86400 * fractionalDay);
            date_info.setSeconds(totalSeconds);
            return date_info.toLocaleString('en-US', {
                year: 'numeric', month: 'numeric', day: 'numeric',
                hour: 'numeric', minute: 'numeric', second: 'numeric',
                hour12: true
            });
        };

        // Format specified date columns & cells
        const dateCols = [5, 6, 41];
        const fixedDateCells = [{ row: 0, col: 23 }, { row: 1449, col: 23 }];

        newRows = newRows.map((row, index) => {
            if (index >= 2) {
                dateCols.forEach(i => {
                    if (row[i]) row[i] = excelDateToJS(row[i]);
                });

                // Clean line breaks in Order Notes
                if (row[36]) {
                    row[36] = row[36].toString().replace(/[\r\n]+/g, ' ').trim();
                }
            }
            return row;
        });

        // Format fixed cells separately
        fixedDateCells.forEach(({ row, col }) => {
            if (newRows[row] && newRows[row][col]) {
                newRows[row][col] = excelDateToJS(newRows[row][col]);
            }
        });

        // Get existing data from Google Sheets
        let existingRows = [];
        try {
            const existingData = await sheets.spreadsheets.values.get({
                spreadsheetId: SHEET_ID,
                range: 'SalesData!A:Z'
            });
            existingRows = existingData.data.values || [];
        } catch (error) {
            console.log('No existing data found or error reading sheet:', error.message);
        }

        // Create sets for duplicate detection
        const existingKeys = new Set();
        const existingDataRows = existingRows.slice(2); // Skip headers

        // Build existing keys set (now accounting for ID column shift)
        existingDataRows.forEach(row => {
            // Skip the ID column (index 0) when creating keys
            const rowWithoutId = row.slice(1);
            const key = createRowKey(rowWithoutId);
            if (key && key !== '_') { // Avoid empty keys
                existingKeys.add(key);
            }
        });

        // Filter new rows to exclude duplicates
        const newDataRows = newRows.slice(2); // Skip headers
        const uniqueNewRows = newDataRows.filter(row => {
            const key = createRowKey(row);
            return key && key !== '_' && !existingKeys.has(key);
        });

        console.log(`Found ${newDataRows.length} new rows, ${uniqueNewRows.length} are unique`);

        // Combine existing data rows with new unique rows
        let combinedDataRows = [...existingDataRows, ...uniqueNewRows];

        const formatDate = (dateObj) => {
            return dateObj.toLocaleString('en-US', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
                hour12: true
            });
        };

        // Sort combined data by date in column G (index 6, was 5 before ID column), latest first
        combinedDataRows = combinedDataRows
            .map(row => {
                if (row[5]) { // Column F is now index 5 (was column F/index 5, now column G/index 6 after ID)
                    row[5] = new Date(row[5]); // Parse to Date
                }
                return row;
            })
            .sort((a, b) => {
                const dateA = a[5] instanceof Date ? a[5] : new Date(0);
                const dateB = b[5] instanceof Date ? b[5] : new Date(0);
                return dateB - dateA; // Descending (latest first)
            })
            .map(row => {
                if (row[5] instanceof Date) {
                    row[5] = formatDate(row[5]); // Format back to string
                }
                return row;
            });

        // Reconstruct final rows with headers
        let finalRows = [];
        if (existingRows.length >= 2) {
            // Use existing headers (but remove ID column from existing data for merging)
            const existingHeadersWithoutId = existingRows[0].slice(1);
            const existingSecondRowWithoutId = existingRows[1].slice(1);
            const existingDataWithoutId = existingDataRows.map(row => row.slice(1));
            
            finalRows = [existingHeadersWithoutId, existingSecondRowWithoutId, ...existingDataWithoutId, ...uniqueNewRows];
        } else {
            // Use new headers if no existing data
            finalRows = [newRows[0], newRows[1], ...combinedDataRows];
        }

        // Add ID column to all rows (this shifts everything to the right)
        finalRows = addIdColumn(finalRows);

        // Clear old sheet content
        await sheets.spreadsheets.values.clear({ 
            spreadsheetId: SHEET_ID, 
            range: `SalesData!A:Z` 
        });

        // Write updated data
        await sheets.spreadsheets.values.update({
            spreadsheetId: SHEET_ID,
            range: `SalesData!A1`,
            valueInputOption: 'RAW',
            requestBody: { values: finalRows }
        });

        await browser.close();
        
        const totalRows = finalRows.length - 2; // Exclude headers
        const addedRows = uniqueNewRows.length;
        
        res.send(`✅ Report updated successfully!
📊 Total rows: ${totalRows}
➕ New rows added: ${addedRows}
🆔 ID column added (original data shifted right)
📅 Sorted by date (latest to oldest)
🔍 Duplicates removed`);

    } catch (err) {
        console.error(err);
        await browser.close();
        res.status(500).send('Error generating report: ' + err.message);
    }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));