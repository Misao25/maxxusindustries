import express from "express";
import { google } from "googleapis";
import path from "path";

const app = express();
app.use(express.json());
app.use(express.static("public"));

// load service account creds from env (Railway supports secrets)
const serviceAccount = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);

// authenticate
const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

// spreadsheet ID
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// endpoint to receive JSON rows
app.post("/push-to-sheets", async (req, res) => {
  try {
    const rows = req.body; // array of flattened rows

    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    // format rows as 2D array
    const values = rows.map(row => Object.values(row));

    // append to sheet
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sheet1!A1", // adjust your sheet name
      valueInputOption: "RAW",
      requestBody: {
        values,
      },
    });

    res.json({ success: true, inserted: rows.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on ${PORT}`));
