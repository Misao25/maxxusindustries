// puppeteer-report-export.js

const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const https = require('https');
require('dotenv').config();

(async () => {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    // Login
    await page.goto('https://app.ecomdash.com/?returnUrl=%2fReporting');
    await page.type('input#UserName', process.env.LOGIN_EMAIL);
    await page.click('input#submit');
    await page.waitForSelector('input#Password');
    await page.type('input#Password', process.env.LOGIN_PASS);
    await page.click('input#submit');
    await page.waitForNavigation({ waitUntil: 'networkidle2' });

    // Generate report
    await page.waitForSelector('div#mostPopular-SalesWithinDateRange_Category');
    await page.click('div#mostPopular-SalesWithinDateRange_Category div.buttonDiv a.albany-btn.albany-btn--primary');
    await page.waitForSelector('form#GenerateReport', { visible: true });

    const fromDate = new Date();
    fromDate.setDate(fromDate.getDate() - 7);
    const toDate = new Date();

    const fromDateStr = fromDate.toLocaleDateString('en-US');
    const toDateStr = toDate.toLocaleDateString('en-US');

    await page.click('#ReportStartDate', { clickCount: 3 });
    await page.keyboard.press('Backspace');
    await page.type('#ReportStartDate', fromDateStr);
    await page.$eval('#ReportStartDate', el => el.blur());

    await page.click('#ReportEndDate', { clickCount: 3 });
    await page.keyboard.press('Backspace');
    await page.type('#ReportEndDate', toDateStr);
    await page.$eval('#ReportEndDate', el => el.blur());

    // Submit and wait for redirect
    await Promise.all([
        page.waitForNavigation({ waitUntil: 'networkidle2' }),
        page.click('a#GenerateDateRestrictionReport')
    ]);

    // After redirect, get the top timestamp as the target
    await page.waitForSelector('table');
    const timestampStr = await page.$eval('table tbody tr td:nth-child(1)', el => el.textContent.trim());
    console.log('⏳ Will match timestamp:', timestampStr);

    // Poll for matching complete report
    const historyUrl = 'https://dashboard.ecomdash.com/Support/ReportingHistory';
    let downloadUrl = null;

    for (let i = 0; i < 30; i++) {
        await page.goto(historyUrl, { waitUntil: 'networkidle2' });
        await page.waitForSelector('table');

        const rows = await page.$$('table tbody tr');
        console.log(`📄 Checking ${rows.length} rows`);

        for (const row of rows) {
            const rowTimestamp = await row.$eval('td:nth-child(1)', el => el.textContent.trim());
            const status = await row.$eval('td:nth-child(4)', el => el.textContent.trim());
            console.log(`🔍 Row: timestamp="${rowTimestamp}", status="${status}"`);

            if (rowTimestamp === timestampStr && status === 'Complete') {
                const linkEl = await row.$('td:nth-child(5) a[href$=".xlsx"]');
                if (linkEl) {
                    downloadUrl = await linkEl.evaluate(a => a.href);
                    console.log('✅ Found download URL:', downloadUrl);
                    break;
                }
            }
        }

        if (downloadUrl) break;
        await new Promise(resolve => setTimeout(resolve, 3000));
    }

    if (!downloadUrl) throw new Error('❌ No matching report found');

    const filePath = path.join('./reports', path.basename(downloadUrl));
    const file = fs.createWriteStream(filePath);

    https.get(downloadUrl, res => {
        res.pipe(file);
        file.on('finish', () => {
            file.close();
            console.log('✅ Report downloaded to:', filePath);
        });
    });

    await browser.close();
    console.log('🎉 All done!');
})();
