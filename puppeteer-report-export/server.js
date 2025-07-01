// server.js (Railway-ready Puppeteer API)

const express = require('express');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const https = require('https');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

app.get('/generate-report', async (req, res) => {
  const { from, to } = req.query;
  if (!from || !to) return res.status(400).send('Missing from or to date');

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 800 });

  try {
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

    await page.click('#ReportStartDate', { clickCount: 3 });
    await page.keyboard.press('Backspace');
    await page.type('#ReportStartDate', from);
    await page.$eval('#ReportStartDate', el => el.blur());

    await page.click('#ReportEndDate', { clickCount: 3 });
    await page.keyboard.press('Backspace');
    await page.type('#ReportEndDate', to);
    await page.$eval('#ReportEndDate', el => el.blur());

    // Submit and wait for redirect
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle2' }),
      page.click('a#GenerateDateRestrictionReport')
    ]);

    // Capture timestamp
    await page.waitForSelector('table');
    const timestampStr = await page.$eval('table tbody tr td:nth-child(1)', el => el.textContent.trim());
    console.log('⏳ Will match timestamp:', timestampStr);

    // Poll for match
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

    // Ensure reports directory exists
    const reportsDir = path.join(__dirname, 'reports');
    if (!fs.existsSync(reportsDir)) fs.mkdirSync(reportsDir);

    const filePath = path.join(reportsDir, path.basename(downloadUrl));
    const file = fs.createWriteStream(filePath);

    await new Promise((resolve, reject) => {
      https.get(downloadUrl, res => {
        res.pipe(file);
        file.on('finish', () => file.close(resolve));
        res.on('error', reject);
      });
    });

    await browser.close();
    res.download(filePath);
  } catch (err) {
    console.error(err);
    await browser.close();
    res.status(500).send('Error generating report: ' + err.message);
  }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
