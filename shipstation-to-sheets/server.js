import express from "express";
import { google } from "googleapis";

const app = express();

// Middleware
app.use(express.json({ limit: "10mb" }));            // Parse JSON body
app.use(express.static("public"));  // Serve index.html + assets from /public

// Load Google service account from Railway ENV
const serviceAccount = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);

// Google auth
const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

// Spreadsheet ID (set in Railway variables)
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// Push "orders summary"
app.post("/push-orders", async (req, res) => {
  await pushToGoogleSheets(req, res, "SS_Orders!A1");
});

// Push "line items"
app.post("/push-line-items", async (req, res) => {
  await pushToGoogleSheets(req, res, "SS_Items!A1");
});

// Helper function
async function pushToGoogleSheets(req, res, range) {
  try {
    const rows = req.body;
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: "No rows provided" });
    }

    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    const values = rows.map(r => Object.values(r));

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range,
      valueInputOption: "RAW",
      requestBody: { values },
    });

    res.json({ success: true, inserted: rows.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}


// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server running on http://localhost:${PORT}`);
});
