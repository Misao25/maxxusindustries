import express from "express";
import { google } from "googleapis";

const app = express();

// Middleware
app.use(express.json());            // Parse JSON body
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

// API route: push rows into Google Sheets
app.post("/push-to-sheets", async (req, res) => {
  try {
    const rows = req.body; // Expect array of flattened order rows
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ error: "No rows provided" });
    }

    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    // Convert objects → array of values
    const values = rows.map(row => Object.values(row));

    // Append rows
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Orders!A1",  // Change "Sheet1" if your tab name differs
      valueInputOption: "RAW",
      requestBody: { values },
    });

    res.json({ success: true, inserted: rows.length });
  } catch (err) {
    console.error("Error pushing to Sheets:", err);
    res.status(500).json({ error: err.message });
  }
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server running on http://localhost:${PORT}`);
});
