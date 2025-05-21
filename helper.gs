/** ============================================================================
 * 📦 RESTOCK SYSTEM – CORE SETUP
 * ============================================================================
*/

/** 📧 Credentials and Constants */
const JUBE_EMAIL = ""; //Jubelio Login Email
const JUBE_PASS = ""; //Jubelio Login Password
const JUBE_API_BASE = 'https://api2.jubelio.com';
const JURNAL_USERNAME = ""; //HMAC Auth Credential Client ID
const JURNAL_SECRET = ""; //HMAC Auth Credential Client Secret
const JURNAL_API_BASE = "https://api.mekari.com";
const recipients = [
  "email_1", "email_2", "email_3"
]; // Recipients for email notifier. Add more emails as needed

/** ============================================================================
 * 🧩 UTILITY FUNCTIONS
 * ============================================================================
*/

/** UI Menu Button */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Restocker Tools')
    .addItem('📥 Start from Jubelio Invoices(90d)', 'fetchJubelioInvoices')
    .addItem('📦 Start from Jubelio Invoice Items', 'fetchJubelioInvoiceItems')
    .addItem('📋 Start from Jubelio Stock', 'fetchJubelioStock')
    .addItem('📥 Summarize Jurnal Invoices(90d)', 'fetchJurnalSales90d')
    .addItem('📦 Start from Jurnal Invoice Items(90d)', 'fetchJurnalInvoiceItems')
    .addItem('📦 Send weekly restock email', 'sendWeeklyRestockEmail')
    .addItem('🔁 Reorder Sheets', 'reorderSheets')
    .addToUi();
}

/** Create authorization headers for Jubelio API requests. */
function getJubelioAuthHeaders(token) {
  return {
    Authorization: token,
    muteHttpExceptions: true
  };
}

/** Create authorization headers for Jurnal API requests. */
function getJurnalHmacHeaders(method, endpointPath) {
  const hmac_username = JURNAL_USERNAME; // Replace with your actual HMAC username
  const hmac_secret = JURNAL_SECRET; // Replace with your actual HMAC secret

  const methodUpper = method.toUpperCase();
  const dateString = new Date().toUTCString(); // Required header
  const requestLine = `${methodUpper} ${endpointPath} HTTP/1.1`;

  const signatureRaw = `date: ${dateString}\n${requestLine}`;
  const signatureBytes = Utilities.computeHmacSha256Signature(signatureRaw, hmac_secret);
  const signatureBase64 = Utilities.base64Encode(signatureBytes);

  const hmacHeader = `hmac username="${hmac_username}", algorithm="hmac-sha256", headers="date request-line", signature="${signatureBase64}"`;

  return {
    "Authorization": hmacHeader,
    "Date": dateString
  };
}

/** Get value from Config sheet (skipping header row) */
function getConfigValue(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1]; // Column B = Value
    }
  }

  return null;
}

/** Reorder sheets from clean to raw */
function reorderSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const desiredOrder = ["Restock Dashboard", "Config", "Jurnal_Avg_Price", "Jubelio_Stock", "JURNAL_Sales_Daily", "Jurnal_WH_Transfers", "Jurnal_Sales_90d", "Jubelio_Sales_90d", "Jurnal_Invoice_Items", "Jubelio_Invoice_Items", "Jubelio_Invoices"];

  desiredOrder.forEach((sheetName, index) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.setActiveSheet(sheet); // Required for moveActiveSheet
      ss.moveActiveSheet(index + 1);
    }
  });

  Logger.log("✅ Sheets reordered successfully.");

  // Ensure "Restock Dashboard" is the active sheet after reordering
  const restockDashboardSheet = ss.getSheetByName("Restock Dashboard");
  ss.setActiveSheet(restockDashboardSheet);  
}

/** Helper to format timestamps nicely */
function formatDDMMYYYY(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

/** Helper to format timestamps */
function formatDateTime(input) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || "Asia/Jakarta";
  const date = new Date(input);
  return Utilities.formatDate(date, tz, "dd-MM-yyyy HH:mm:ss");
}

/** Export a sheet to PDF */
function exportSheetAsPDF(spreadsheetId, sheetId, filename) {
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?`;

  const params = {
    format: "pdf",
    portrait: false,
    size: "A4",
    sheetnames: false,
    printtitle: false,
    pagenumbers: false,
    gridlines: false,
    fzr: false,
    gid: sheetId,
    fitw: true
  };

  const queryString = Object.keys(params)
    .map(key => `${key}=${params[key]}`)
    .join("&");

  const url = exportUrl + queryString;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  return response.getBlob().setName(filename);
}
