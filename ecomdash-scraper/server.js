// server.js
const express = require('express');
const runScraper = require('./ecomdash-scraper');

const app = express();
const port = process.env.PORT || 3000;

app.get('/run', async (req, res) => {
    try {
        const result = await runScraper();
        res.json({ status: 'success', ...result });
    } catch (err) {
        res.status(500).json({ status: 'error', message: err.message });
    }
});

app.get('/', (req, res) => res.send('✅ Scraper is alive. Hit /run to start.'));

app.listen(port, () => console.log(`🚀 Server listening on port ${port}`));
