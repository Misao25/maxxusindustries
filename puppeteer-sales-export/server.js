// sales.js (initial deploy)

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

const SHEET_ID = '1nrKVyyv9Lfxe07o_znL7pdX7xMirP2KU0ubv4uaCRHw';

app.get('/generate-sales', async (req, res) => {
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

        await page.waitForSelector('div#mostPopular-SalesWithinDateRange_Category');
        await page.click('div#mostPopular-SalesWithinDateRange_Category div.buttonDiv a.albany-btn.albany-btn--primary');
        await page.waitForSelector('form#GenerateReport', { visible: true });

        await page.click('#ReportStartDate', { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await page.type('#ReportStartDate', from);
        await page.$eval('#ReportStartDate', el => el.blur());

        await page.click('#ReportEndDate', { clickCount: 3 });
        await page.keyboard.press('Backspace');
        await page.type('#ReportEndDate', to);
        await page.$eval('#ReportEndDate', el => el.blur());
        
        await new Promise(resolve => setTimeout(resolve, 3000));

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
        let rows = xlsx.utils.sheet_to_json(firstSheet, { header: 1 });

        // Make sure we have at least one row (header)
        if (!rows.length) throw new Error('XLSX file appears empty');

        // Remove rows without SKU #
        rows = rows.filter((row, index) => index === 0 || (row[2] && row[2].toString().trim() !== ''));

        // Add Report Start/End Dates to each row after header
        const reportStartDate = from;
        const reportEndDate = to;

        // Insert headers for the new columns
        rows[0].push('Start Date', 'End Date', 'Date Generated');

        // Append the same start/end date values to each row
        rows = rows.map((row, index) => {
            if (index === 0) return row; // Skip header row
            return [...row, reportStartDate, reportEndDate, timestampStr];
        });

        // Only keep data rows that actually have SKU or UPC (avoid blanks)
        const headerRows = rows.slice(0, 1);
        const dataRows = rows.slice(1).filter(row => row[2] || row[3]); // SKU# or UPC present

        // Recombine
        rows = [...headerRows, ...dataRows];


        // Fetch existing Ecomdash IDs from column A (SalesMasterfile sheet)
        const existingData = await sheets.spreadsheets.values.get({
            spreadsheetId: SHEET_ID,
            range: 'SalesData!A:A',
        });
        const existingIDs = new Set((existingData.data.values || []).flat().map(id => id?.toString().trim()));

        // Prepare rows to append (skip rows with 'Date TypeCreate' in column A)
        const newRows = rows.slice(2).filter(row => {
            const id = row[0]?.toString().trim();
            return id && id !== 'Date TypeCreate' && !existingIDs.has(id);
        });

        // Add header if sheet is empty
        if ((existingData.data.values || []).length === 0) {
            newRows.unshift(rows[0], rows[1]); // include headers
        }

        // Append new rows to SalesMasterfile
        if (newRows.length > 0) {
            await sheets.spreadsheets.values.append({
                spreadsheetId: SHEET_ID,
                range: 'SalesData!A1',
                valueInputOption: 'USER_ENTERED', // Allows Google Sheets to recognize date formats
                insertDataOption: 'INSERT_ROWS',
                requestBody: { values: newRows },
            });
        }

        await browser.close();
        res.send(`✅ Added ${newRows.length - (existingData.data.values?.length === 0 ? 2 : 0)} new rows to SalesData`);
    }
    catch (err) {
        console.error(err);
        await browser.close();
        res.status(500).send('Error generating report: ' + err.message);
    }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));