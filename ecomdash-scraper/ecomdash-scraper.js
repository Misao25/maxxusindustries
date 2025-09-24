// scraper.js

const puppeteer = require('puppeteer');
require('dotenv').config();
const { google } = require('googleapis');

// --- Google Sheets setup ---
const credentials = JSON.parse(
  Buffer.from(process.env.GCP_CREDENTIALS_B64, 'base64').toString('utf-8')
);
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// ✅ Masterfile (source of orderIds)
const MASTERFILE_ID = '1mrw-AMbVWnz1Cp4ksjR0W0eTDz0cUiA-zjThrzcIRnY';

// ✅ Destination sheet (scraped data)
const DESTINATION_ID = '1CNrLur_7RQkznmoNLZypFflY2gRmgHLg0vMJqdGm3_c';

// ✅ Batch size
const BATCH_SIZE = 100;

// --- Helpers ---
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const formatDate = (dateStr) => {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}/${mm}/${dd}`;
};

function chunkArray(arr, size) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}

async function writeToSheet(sheetName, values) {
  if (!values.length) return;
  await sheets.spreadsheets.values.append({
    spreadsheetId: DESTINATION_ID,
    range: `${sheetName}!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
}

async function getExistingOrderIds() {
  try {
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: DESTINATION_ID,
      range: 'Orders!A2:A',
    });
    const rows = result.data.values || [];
    return new Set(rows.map((r) => r[0]));
  } catch {
    return new Set();
  }
}

async function readOrderIds() {
  const range = 'Distinct_Orders!A2:A';
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: MASTERFILE_ID,
    range,
  });
  const rows = result.data.values || [];
  return rows.map((r) => r[0]).filter(Boolean);
}

// --- Scraper Logic per Batch ---
async function processBatch(orderIds, batchIndex, totalBatches, results) {
  console.log(`\n🚀 Starting batch ${batchIndex + 1} of ${totalBatches}`);

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 800 });

  // --- Login ---
  await page.goto(
    'https://app.ecomdash.com/?returnUrl=%2fSalesOrderModule%2fAllSalesOrders'
  );
  await page.type('input#UserName', process.env.LOGIN_EMAIL);
  await page.click('input#submit');
  await page.waitForSelector('input#Password', { timeout: 10000 });
  await page.type('input#Password', process.env.LOGIN_PASS);
  await page.click('input#submit');
  await page.waitForNavigation({ waitUntil: 'networkidle2' });

  const existingOrderIds = await getExistingOrderIds();

  for (const orderId of orderIds) {
    if (existingOrderIds.has(orderId)) continue;

    const url = `https://dashboard.ecomdash.com/SalesOrderModule/SalesOrderDetails?ID=${orderId}&ReturnItem=AllSO`;
    console.log(`\n🔎 Visiting Order ID: ${orderId}`);

    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

      // --- Extract Order Info ---
      const orderInfo = await page.evaluate(() => {
        const rawOrderNumber =
          document.querySelector('.orderdetail-header__text')?.innerText || '';
        let orderNumber = rawOrderNumber.replace(/\s+/g, ' ').trim();
        orderNumber = orderNumber.replace(/^ORDER/i, '').trim();
        orderNumber = orderNumber.replace(/#+/g, '#');

        let status =
          document
            .querySelector('.orderdetail-header__status')
            ?.innerText.trim() || '';
        status = status.toLowerCase();

        const ecomdashId = document.querySelector('input#ID')?.value || '';
        const orderDate =
          document.querySelector('input#SalesOrderCreateDate')?.value || '';
        let storefront = '';
        try {
          storefront =
            document.querySelector('.orderdetail-header label:nth-of-type(2)')
              ?.nextSibling?.textContent || '';
        } catch {}
        storefront = storefront.replace(/\s+/g, ' ').trim();

        const merchandiseTotal =
          document.querySelector('#ProductTotal')?.value || '';
        const tax1 = document.querySelector('#Tax1')?.value || '';
        const tax2 = document.querySelector('#Tax2')?.value || '';
        const tax3 = document.querySelector('#Tax3')?.value || '';
        const shipping =
          document.querySelector('#ShippingandHandling')?.value || '';
        const discount = document.querySelector('#Discount')?.value || '';
        const otherFees = document.querySelector('#OtherAmount')?.value || '';
        const orderTotal =
          document.querySelector('#SalesOrderTotal')?.value || '';

        return {
          orderNumber,
          status,
          ecomdashId,
          orderDate,
          storefront,
          financials: {
            merchandiseTotal,
            tax1,
            tax2,
            tax3,
            shipping,
            discount,
            otherFees,
            orderTotal,
          },
        };
      });

      orderInfo.orderDate = formatDate(orderInfo.orderDate);

      // --- Products extraction ---
      await page.waitForSelector('#SalesOrderProductList tbody tr', {
        timeout: 10000,
      });

      const allRowHandles = await page.$$('#SalesOrderProductList tbody > tr');
      const products = [];

      for (let r = 0; r < allRowHandles.length; r++) {
        const row = allRowHandles[r];
        const isChildRow = (await row.$('.child-table')) !== null;
        if (isChildRow) continue;

        const product = await (async () => {
          const name =
            (await row
              .$eval('td:nth-child(2) b font', (el) => el.innerText.trim())
              .catch(() => '')) || '';
          const tdText =
            (await row
              .$eval('td:nth-child(2)', (el) => el.innerText)
              .catch(() => '')) || '';
          let sku = '';
          const skuMatch = tdText.match(/SKU:\s*([A-Za-z0-9\-_]+)/);
          if (skuMatch) sku = skuMatch[1].trim();

          const qty =
            (await row
              .$eval('td:nth-child(5) input[type="hidden"]', (el) => el.value)
              .catch(() => '')) || '';

          let price =
            (await row
              .$eval('td:nth-child(6) input[type="hidden"]', (el) => el.value)
              .catch(() => '')) || '';
          if (!price) {
            price =
              (await row
                .$eval('td:nth-child(6) input.order-price', (el) => el.value)
                .catch(() => '')) || '';
          }

          return { name, sku, qty, price };
        })();

        products.push(product);
      }

      // --- Push to Google Sheets ---
      const ordersData = [[
        orderId,
        orderInfo.orderNumber,
        orderInfo.orderDate,
        orderInfo.status,
        orderInfo.storefront,
        orderInfo.financials.merchandiseTotal,
        orderInfo.financials.tax1,
        orderInfo.financials.tax2,
        orderInfo.financials.tax3,
        orderInfo.financials.shipping,
        orderInfo.financials.discount,
        orderInfo.financials.otherFees,
        orderInfo.financials.orderTotal,
      ]];

      const productsData = products.map((p, i) => [
        orderId,
        i + 1,
        p.name,
        p.sku,
        p.qty,
        p.price,
        orderInfo.orderDate,
        orderInfo.storefront,
      ]);

      await writeToSheet('Orders', ordersData);
      await writeToSheet('Products', productsData);

      existingOrderIds.add(orderId);
      results.successes.push(orderId);
      console.log(`✅ Pushed Order ${orderId}`);
    } catch (err) {
      console.error(`❌ Failed on Order ${orderId}:`, err.message);
      results.failures.push({ orderId, error: err.message });
    }

    await sleep(2000);
  }

  await browser.close();
  console.log(`✅ Finished batch ${batchIndex + 1} of ${totalBatches}`);
}

// --- Main Runner (exported only) ---
async function runScraper() {
  const orderIds = await readOrderIds();
  const orderChunks = chunkArray(orderIds, BATCH_SIZE);

  const results = {
    processed: orderIds.length,
    batches: orderChunks.length,
    successes: [],
    failures: [],
  };

  for (let i = 0; i < orderChunks.length; i++) {
    await processBatch(orderChunks[i], i, orderChunks.length, results);
    await sleep(5000);
  }

  console.log('🎉 All done!');
  return results;
}

module.exports = runScraper;
