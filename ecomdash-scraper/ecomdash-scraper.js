// scraper.js

const puppeteer = require('puppeteer');
require('dotenv').config();
const { google } = require('googleapis');

// --- Google Sheets setup ---
const credentials = JSON.parse(Buffer.from(process.env.GCP_CREDENTIALS_B64, 'base64').toString('utf-8'));
const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
});
const sheets = google.sheets({ version: 'v4', auth });

// ✅ Masterfile (source of orderIds)
const MASTERFILE_ID = '1mrw-AMbVWnz1Cp4ksjR0W0eTDz0cUiA-zjThrzcIRnY';

// ✅ Destination sheet (scraped data)
const DESTINATION_ID = '1CNrLur_7RQkznmoNLZypFflY2gRmgHLg0vMJqdGm3_c';

// ✅ Batch size (restart browser every 100 orders)
const BATCH_SIZE = 100;

// --- Helpers ---
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
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
            range: 'Orders!A2:A', // column A = orderId
        });
        const rows = result.data.values || [];
        return new Set(rows.map(r => r[0]));
    } catch {
        return new Set();
    }
}

async function readOrderIds() {
    const range = 'Distinct_Orders!A2:A'; // ✅ safe to pull full col now
    const result = await sheets.spreadsheets.values.get({
        spreadsheetId: MASTERFILE_ID,
        range,
    });
    const rows = result.data.values || [];
    return rows.map(r => r[0]).filter(Boolean);
}

// --- Scraper Logic per Batch ---
async function processBatch(orderIds, batchIndex, totalBatches) {
    console.log(`\n🚀 Starting batch ${batchIndex + 1} of ${totalBatches}`);

    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    // --- Login ---
    await page.goto('https://app.ecomdash.com/?returnUrl=%2fSalesOrderModule%2fAllSalesOrders');
    await page.type('input#UserName', process.env.LOGIN_EMAIL);
    await page.click('input#submit');
    await page.waitForSelector('input#Password', { timeout: 10000 });
    await page.type('input#Password', process.env.LOGIN_PASS);
    await page.click('input#submit');
    await page.waitForNavigation({ waitUntil: 'networkidle2' });

    const existingOrderIds = await getExistingOrderIds();

    for (const orderId of orderIds) {
        if (existingOrderIds.has(orderId)) {
            // console.log(`⏭️ Skipping ${orderId}, already in Orders tab`);
            continue;
        }

        const url = `https://dashboard.ecomdash.com/SalesOrderModule/SalesOrderDetails?ID=${orderId}&ReturnItem=AllSO`;
        console.log(`\n🔎 Visiting Order ID: ${orderId}`);

        try {
            await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

            // --- Extract Order Info ---
            const orderInfo = await page.evaluate(() => {
                const rawOrderNumber = document.querySelector('.orderdetail-header__text')?.innerText || '';
                let orderNumber = rawOrderNumber.replace(/\s+/g, ' ').trim();
                orderNumber = orderNumber.replace(/^ORDER/i, '').trim(); // remove "ORDER"
                orderNumber = orderNumber.replace(/#+/g, '#'); // collapse multiple # to single

                let status = document.querySelector('.orderdetail-header__status')?.innerText.trim() || '';
                status = status.toLowerCase(); // lowercase status

                const ecomdashId = document.querySelector('input#ID')?.value || '';
                const orderDate = document.querySelector('input#SalesOrderCreateDate')?.value || '';
                let storefront = '';
                try {
                    // "Storefront: " label is the 2nd label under header block
                    storefront = document.querySelector('.orderdetail-header label:nth-of-type(2)')?.nextSibling?.textContent || '';
                } catch {}
                storefront = storefront.replace(/\s+/g, ' ').trim(); // normalize spaces/newlines

                const merchandiseTotal = document.querySelector('#ProductTotal')?.value || '';
                const tax1 = document.querySelector('#Tax1')?.value || '';
                const tax2 = document.querySelector('#Tax2')?.value || '';
                const tax3 = document.querySelector('#Tax3')?.value || '';
                const shipping = document.querySelector('#ShippingandHandling')?.value || '';
                const discount = document.querySelector('#Discount')?.value || '';
                const otherFees = document.querySelector('#OtherAmount')?.value || '';
                const orderTotal = document.querySelector('#SalesOrderTotal')?.value || '';

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
                        orderTotal
                    }
                };
            });

            orderInfo.orderDate = formatDate(orderInfo.orderDate);

            // --- Robust product + kit extraction (row-by-row using handles) ---
            await page.waitForSelector('#SalesOrderProductList tbody tr', { timeout: 10000 });

            const allRowHandles = await page.$$('#SalesOrderProductList tbody > tr');

            const products = [];
            let productIndex = 0;

            for (let r = 0; r < allRowHandles.length; r++) {
                const row = allRowHandles[r];

                // Skip child rows (they contain .child-table)
                const isChildRow = await row.$('.child-table') !== null;
                if (isChildRow) continue;

                // Extract product basics from the product row itself
                const product = await (async () => {
                    const name = (await row.$eval('td:nth-child(2) b font', el => el.innerText.trim()).catch(() => '')) || '';

                    const tdText = (await row.$eval('td:nth-child(2)', el => el.innerText).catch(() => '')) || '';
                    let sku = '';
                    const skuMatch = tdText.match(/SKU:\s*([A-Za-z0-9\-_]+)/);
                    if (skuMatch) sku = skuMatch[1].trim();

                    // Qty: hidden input preferred
                    const qty = (await row.$eval('td:nth-child(5) input[type="hidden"]', el => el.value).catch(() => '')) || '';

                    // Price: hidden input first, then visible text input (awaiting shipments)
                    let price = (await row.$eval('td:nth-child(6) input[type="hidden"]', el => el.value).catch(() => '') ) || '';
                    if (!price) {
                        price = (await row.$eval('td:nth-child(6) input.order-price', el => el.value).catch(() => '') ) || '';
                    }

                    return { name, sku, qty, price, kits: [] };
                })();

                // Per-product kit: click expander in THIS row, then read NEXT sibling row's .child-table
                const expander = await row.$('.albany-expander');
                if (expander) {
                    let kitRows = [];
                    try {
                        // Try clicking twice if needed
                        for (let attempt = 0; attempt < 2; attempt++) {
                            await expander.click();
                            try {
                                await page.waitForFunction(
                                    el => el.nextElementSibling && el.nextElementSibling.querySelector('.child-table'),
                                    { timeout: 5000 },
                                    row
                                );
                                // Child table found, extract and break retry loop
                                kitRows = await row.evaluate(el => {
                                    const sib = el.nextElementSibling;
                                    if (!sib) return [];
                                    const table = sib.querySelector('.child-table');
                                    if (!table) return [];
                                    return Array.from(table.querySelectorAll('tbody tr'))
                                        .map(tr => Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim()))
                                        .filter(cols => cols.some(text => text.length > 0)); // skip empty rows
                                });
                                break;
                            } catch {
                                if (attempt === 0) {
                                    console.log(`⚠️ First click didn’t open child-table, retrying...`);
                                }
                            }
                        }

                        // Map if we got kit rows
                        if (kitRows.length > 0) {
                            product.kits = kitRows.map(k => {
                                const [componentName, componentSku, componentQty, componentLocation] = k;
                                return { componentName, componentSku, componentQty, componentLocation };
                            });
                        }
                    } catch (e) {
                        console.log(`⚠️ Failed to expand product row: ${e.message}`);
                    }
                }


                products.push(product);
                productIndex++;
            }
            // --- Get existing order rows so we know where new ones will land ---
            const ordersRange = await sheets.spreadsheets.values.get({
                spreadsheetId: DESTINATION_ID,
                range: 'Orders!A:A'
            });
            const existingOrderRows = ordersRange.data.values ? ordersRange.data.values.length : 0;
            const rowIndex = existingOrderRows + 1;

            // --- Flatten Data ---
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
                `=SUMIFS(Products!I:I, Products!A:A, A${rowIndex})`
            ]];

            const productsData = [];
            const kitsData = [];

            // --- Get existing product rows so we know where new ones will land ---
            const productsRange = await sheets.spreadsheets.values.get({
                spreadsheetId: DESTINATION_ID,
                range: 'Products!A:A'
            });
            const existingProductRows = productsRange.data.values ? productsRange.data.values.length : 0;

            // --- Build Products + Product_Items rows ---
            products.forEach((p, i) => {
                const rowIndex = existingProductRows + productsData.length + 1; // next row index in Products

                // --- Products tab ---
                productsData.push([
                    orderId,
                    i + 1,
                    p.name,
                    p.sku,
                    p.qty,
                    p.price,
                    orderInfo.orderDate,
                    orderInfo.storefront,
                    `=SUMIFS(Product_Items!L:L, Product_Items!A:A, A${rowIndex}, Product_Items!B:B, B${rowIndex})` // 👈 formula for cost
                ]);

                // --- Product_Items tab ---
                if (p.kits.length > 0) {
                    p.kits.forEach((kit, j) => {
                        kitsData.push([
                            orderId,
                            i + 1,
                            j + 1,
                            kit.componentSku || '',
                            kit.componentName || '',
                            kit.componentQty || '',
                            kit.componentLocation || '',
                            orderInfo.orderDate,
                            orderInfo.storefront
                        ]);
                    });
                } else {
                    // Standalone product → add as single itemized row for inventory accuracy
                    kitsData.push([
                        orderId,
                        i + 1,
                        1,
                        p.sku,
                        p.name,
                        p.qty,
                        '',
                        orderInfo.orderDate,
                        orderInfo.storefront
                    ]);
                }
            });


            // --- Push to Google Sheets ---
            await writeToSheet('Orders', ordersData);
            await writeToSheet('Products', productsData);
            await writeToSheet('Product_Items', kitsData);

            // add to in-memory set so if same orderId appears later in this run, we skip
            existingOrderIds.add(orderId);

            console.log(`✅ Pushed Order ${orderId} to Google Sheets`);
        } catch (err) {
            console.error(`❌ Failed on Order ${orderId}:`, err.message);
        }

        await sleep(2000); // small delay between orders
    }

    await browser.close();
    console.log(`✅ Finished batch ${batchIndex + 1} of ${totalBatches}`);
}

// --- Main Runner ---
(async () => {
    const orderIds = await readOrderIds();
    const orderChunks = chunkArray(orderIds, BATCH_SIZE);

    for (let i = 0; i < orderChunks.length; i++) {
        await processBatch(orderChunks[i], i, orderChunks.length);
        await sleep(5000); // wait before next batch
    }

    console.log('🎉 All done!');
})();
