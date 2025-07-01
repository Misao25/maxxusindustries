// server.js (supports batch XLSX downloads for 'generate all')

const express = require('express');
const puppeteer = require('puppeteer');
const https = require('https');
const archiver = require('archiver');
const stream = require('stream');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

app.get('/generate-report', async (req, res) => {
    const { from, to, all } = req.query;
    if (!from || !to) return res.status(400).send('Missing from or to date');

    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
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

        const categories = all === 'true'
        ? [
            'SalesWithinDateRange_Category',
            'SalesSummary_Category',
            'InventoryValue_Category'
            ]
        : ['SalesWithinDateRange_Category'];

        const downloadLinks = [];

        for (const cat of categories) {
        await page.goto('https://dashboard.ecomdash.com/Reporting');
        await page.waitForSelector(`div#mostPopular-${cat}`);
        await page.click(`div#mostPopular-${cat} div.buttonDiv a.albany-btn.albany-btn--primary`);
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

        let downloadUrl = null;
        const historyUrl = 'https://dashboard.ecomdash.com/Support/ReportingHistory';

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

        if (downloadUrl) downloadLinks.push(downloadUrl);
        }

        await browser.close();

        if (all === 'true') {
        res.setHeader('Content-Disposition', 'attachment; filename="all-reports.zip"');
        res.setHeader('Content-Type', 'application/zip');

        const archive = archiver('zip');
        archive.pipe(res);

        for (const url of downloadLinks) {
            const fileName = url.split('/').pop();
            archive.append(https.get(url, stream.PassThrough()), { name: fileName });
        }

        archive.finalize();
        } else {
            const fileUrl = downloadLinks[0];
            const fileName = fileUrl.split('/').pop();
            res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            https.get(fileUrl, fileRes => fileRes.pipe(res)).on('error', err => {
                console.error('Error downloading file:', err);
                res.status(500).send('Download failed');
            });
        }
    } catch (err) {
        console.error(err);
        await browser.close();
        res.status(500).send('Error generating report: ' + err.message);
    }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
