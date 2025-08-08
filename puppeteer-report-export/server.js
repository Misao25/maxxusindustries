// server.js (formatted data version - optimized for Google Sheets usability)

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
        let rows = xlsx.utils.sheet_to_json(firstSheet, { header: 1 });

        // Convert Excel serial dates to proper date format
        const excelDateToJS = serial => {
            if (!serial || isNaN(serial) || serial === 0 || serial < 1) return null;
            
            // Handle Excel's 1900 date system bug (treats 1900 as leap year)
            let adjustedSerial = serial;
            if (serial >= 61) {
                adjustedSerial = serial - 1;
            }
            
            const utc_days = Math.floor(adjustedSerial - 25569);
            const utc_value = utc_days * 86400; 
            const date_info = new Date(utc_value * 1000);
            const fractionalDay = adjustedSerial % 1;
            const totalSeconds = Math.round(86400 * fractionalDay);
            date_info.setSeconds(totalSeconds);
            
            // Validate the resulting date (should be after 1900-01-01)
            if (date_info.getFullYear() < 1900) return null;
            
            return date_info;
        };

        // Format date for Google Sheets (YYYY/MM/DD HH:MM:SS AM/PM format)
        const formatDateForSheets = (dateObj) => {
            if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj)) return '';
            
            // Additional validation: reject dates before 1900 (likely Excel artifacts)
            if (dateObj.getFullYear() < 1900) return '';
            
            const year = dateObj.getFullYear();
            const month = String(dateObj.getMonth() + 1).padStart(2, '0');
            const day = String(dateObj.getDate()).padStart(2, '0');
            
            // Format time with AM/PM
            let hours = dateObj.getHours();
            const minutes = String(dateObj.getMinutes()).padStart(2, '0');
            const seconds = String(dateObj.getSeconds()).padStart(2, '0');
            const ampm = hours >= 12 ? 'PM' : 'AM';
            
            // Convert to 12-hour format
            hours = hours % 12;
            hours = hours ? hours : 12; // 0 should be 12
            const formattedHours = String(hours).padStart(2, '0');
            
            return `${year}/${month}/${day} ${formattedHours}:${minutes}:${seconds} ${ampm}`;
        };

        // Process and format data
        const dateCols = [5, 6, 41]; // Invoice Date, Payment Received Date, Completed Date
        const fixedDateCells = [{ row: 0, col: 23 }, { row: 1449, col: 23 }];

        rows = rows.map((row, index) => {
            if (index >= 2) { // Skip header rows
                // Format date columns
                dateCols.forEach(colIndex => {
                    if (row[colIndex] !== undefined && row[colIndex] !== null && row[colIndex] !== '') {
                        // Check if it's already a formatted date string or an Excel serial number
                        let dateObj;
                        
                        if (typeof row[colIndex] === 'string' && row[colIndex].includes('/')) {
                            // Already formatted date string - parse it
                            dateObj = new Date(row[colIndex]);
                            
                            // Check if it's the problematic 1899 date
                            if (dateObj.getFullYear() === 1899) {
                                row[colIndex] = ''; // Set to empty for null dates
                                return;
                            }
                        } else if (typeof row[colIndex] === 'number') {
                            // Excel serial number
                            dateObj = excelDateToJS(row[colIndex]);
                        } else {
                            // Try to parse as date
                            dateObj = new Date(row[colIndex]);
                        }
                        
                        // Format the date
                        row[colIndex] = formatDateForSheets(dateObj);
                    } else {
                        row[colIndex] = ''; // Ensure empty cells are empty strings
                    }
                });

                // Clean Order Notes (remove line breaks)
                if (row[36]) {
                    row[36] = row[36].toString().replace(/[\r\n]+/g, ' ').trim();
                }
            }
            return row;
        });

        // Format fixed date cells separately
        fixedDateCells.forEach(({ row, col }) => {
            if (rows[row] && rows[row][col]) {
                let dateObj;
                
                if (typeof rows[row][col] === 'string' && rows[row][col].includes('/')) {
                    dateObj = new Date(rows[row][col]);
                    // Check for 1899 dates
                    if (dateObj.getFullYear() === 1899) {
                        rows[row][col] = '';
                        return;
                    }
                } else if (typeof rows[row][col] === 'number') {
                    dateObj = excelDateToJS(rows[row][col]);
                } else {
                    dateObj = new Date(rows[row][col]);
                }
                
                rows[row][col] = formatDateForSheets(dateObj);
            }
        });

        // Sort rows by Invoice Date (column F, index 5) with custom sorting logic
        const headerRows = rows.slice(0, 2);
        const dataRows = rows.slice(2);
        
        dataRows.sort((a, b) => {
            const dateA = a[5] || '';
            const dateB = b[5] || '';
            
            // Handle empty dates (null/uncompleted orders should be at top)
            if (!dateA && !dateB) return 0;
            if (!dateA) return -1; // Empty dates go to top
            if (!dateB) return 1;
            
            // Compare actual dates (latest first)
            const parsedDateA = new Date(dateA);
            const parsedDateB = new Date(dateB);
            return parsedDateB - parsedDateA; // Descending order (latest first)
        });

        // Reconstruct rows array
        rows = [...headerRows, ...dataRows];

        // Fetch existing Ecomdash IDs from column A (SalesMasterfile sheet)
        const existingData = await sheets.spreadsheets.values.get({
            spreadsheetId: SHEET_ID,
            range: 'SalesMasterfile!A:A',
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
                range: 'SalesMasterfile!A1',
                valueInputOption: 'USER_ENTERED', // Allows Google Sheets to recognize date formats
                insertDataOption: 'INSERT_ROWS',
                requestBody: { values: newRows },
            });
        }

        await browser.close();
        res.send(`✅ Added ${newRows.length - (existingData.data.values?.length === 0 ? 2 : 0)} new rows to SalesMasterfile`);
    }
    catch (err) {
        console.error(err);
        await browser.close();
        res.status(500).send('Error generating report: ' + err.message);
    }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));