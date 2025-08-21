const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const { google } = require("googleapis");
const path = require("path");

const app = express();
app.use(bodyParser.json());
app.use(cors()); // Allows all origins; adjust if needed

// Serve frontend
app.use(express.static(path.join(__dirname, "public")));

// Google Sheets setup
const serviceAccount = require("./service-account.json");
const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"]
});
const sheets = google.sheets({ version: "v4", auth });

// Spreadsheet ID & sheet name
const SPREADSHEET_ID = "1mcA1-anLx6LZQdqkgUodRPAnHL5oD_KuhqQHFQHqOow";
const SHEET_NAME = "Orders";

app.post("/append-orders", async (req, res) => {
  try {
    const rows = req.body.data; // Expect array of arrays [[orderId, orderNumber, ...], ...]

    const client = await auth.getClient();
    await sheets.spreadsheets.values.append({
      auth: client,
      spreadsheetId: SPREADSHEET_ID,
      range: SHEET_NAME,
      valueInputOption: "RAW",
      resource: { values: rows }
    });

    res.status(200).send("Data appended successfully");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error appending data: " + err.message);
  }
});

// Start server
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
