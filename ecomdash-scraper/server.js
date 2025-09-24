// server.js
const express = require('express');
const runScraper = require('./ecomdash-scraper');

const app = express();
const port = process.env.PORT || 3000;

let isRunning = false;
const queue = [];

async function runWithQueue() {
  if (isRunning) return;
  isRunning = true;

  while (queue.length > 0) {
    const { resolve } = queue.shift();
    try {
      const result = await runScraper();
      resolve({ result });
    } catch (err) {
      resolve({
        result: {
          success: false,
          message: `Fatal error in queue: ${err.message}`,
          errors: [err.stack],
        },
      });
    }
  }

  isRunning = false;
}

app.get('/run', async (req, res) => {
  console.log('📡 Triggered by Make.com');

  // Queue the run
  const runPromise = new Promise((resolve) => {
    queue.push({ resolve });
  });

  // Start queue processor if idle
  runWithQueue();

  // Wait for result of this queued run
  const { result } = await runPromise;

  if (result.success) {
    res.status(200).json({ status: 'success', ...result });
  } else {
    res.status(500).json({ status: 'error', ...result });
  }
});

app.get('/', (req, res) => {
  res.send('✅ Scraper service is alive. Call /run to trigger scraping.');
});

app.listen(port, () => console.log(`🚀 Server listening on ${port}`));
