/** ============================================================================
 * 📧 EMAIL NOTIFICATIONS – MONTHLY & WEEKLY
 * ============================================================================
*/

/** Send MONTHLY STATUS notification email */
function sendMonthlyStatusEmail(statusMessage = "✅ Completed successfully", errorLogs = []) {
  const recipientStr = recipients.join(", ");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetUrl = ss.getUrl();

  const invoiceSheet = ss.getSheetByName("Jubelio_Invoices");
  const itemSheet = ss.getSheetByName("Jubelio_Invoice_Items");
  const summarySheet = ss.getSheetByName("Jubelio_Sales_90d");
  const stockSheet = ss.getSheetByName("Jubelio_Stock");

  const invoiceCount = invoiceSheet.getLastRow() - 1;
  const itemCount = itemSheet.getLastRow() - 1;
  const skuCount = summarySheet.getLastRow() - 1;
  const stockCount = stockSheet.getLastRow() - 1;

  const processStartTime = invoiceSheet.getRange("I2").getValue();
  const processEndTime = new Date();

  // Always use date range from invoice sheet column C (transaction_date)
  let fromDate, toDate;
  const rawDates = invoiceSheet.getRange("C2:C" + invoiceSheet.getLastRow()).getValues().flat();
  const invoiceDates = rawDates
    .map(val => new Date(val))
    .filter(date => !isNaN(date)); // only keep valid dates

  if (invoiceDates.length > 0) {
    fromDate = new Date(Math.min(...invoiceDates.map(d => d.getTime())));
    toDate = new Date(Math.max(...invoiceDates.map(d => d.getTime())));
  } else {
    ss.toast("📅 No invoice dates found – fallback to last 90 days.");
    const today = new Date();
    const ninetyDaysAgo = new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000);
    fromDate = ninetyDaysAgo;
    toDate = today;
  }

  const formattedFromDate = Utilities.formatDate(fromDate, Session.getScriptTimeZone(), "dd-MM-yyyy");
  const formattedToDate = Utilities.formatDate(toDate, Session.getScriptTimeZone(), "dd-MM-yyyy");
  Logger.log(`📅 Final date range used: from ${formattedFromDate} to ${formattedToDate}`);


  const formattedSubjectDate = Utilities.formatDate(processEndTime, Session.getScriptTimeZone(), "dd-MM-yyyy");
  const subject = `📦 ${formattedSubjectDate} Monthly Past 90d Sales Summary – ${statusMessage}`;

  const errorSection = errorLogs.length > 0
    ? `
    <h3>⚠️ Errors Encountered (${errorLogs.length})</h3>
    <ul style="color: #B00020;">
      ${errorLogs.map(err => `<li>${err}</li>`).join('')}
    </ul>
    `
    : "<p>✅ No errors were encountered.</p>";

  const htmlBody = `
    <p>Hi! Your monthly restocker automation has completed. Here's the full summary:</p>

    <h3>🕐 Process Timestamp (dd/mm/yyyy hh:mm:ss)</h3>
    <ul>
      <li><b>Started:</b> ${formatDateTime(processStartTime)}</li>
      <li><b>Ended:</b> ${formatDateTime(processEndTime)}</li>
    </ul>

    <h3>📅 Invoice Date Range</h3>
    <ul>
      <li><b>From:</b> ${formatDateTime(fromDate)}</li>
      <li><b>To:</b> ${formatDateTime(toDate)}</li>
    </ul>

    <h3>📊 Data Summary</h3>
    <ul>
      <li><b>Jubelio Invoices Fetched:</b> ${invoiceCount}</li>
      <li><b>Jubelio Invoice Items Fetched:</b> ${itemCount}</li>
      <li><b>Unique SKUs Sold:</b> ${skuCount}</li>
      <li><b>SKUs with Current Stock Info:</b> ${stockCount}</li>
    </ul>

    ${errorSection}

    <p>🔗 <a href="${spreadsheetUrl}" target="_blank"><b>Click here to view the Google Sheet</b></a></p>
  `;

  const dashboardSheet = ss.getSheetByName("Restock Dashboard");
  const pdfBlob = exportSheetAsPDF(ss.getId(), dashboardSheet.getSheetId(), "Restock_Dashboard.pdf");

  MailApp.sendEmail({
    to: recipientStr,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [pdfBlob]
  });
}


/** Send WEEKLY STATUS notification email */
function sendWeeklyRestockEmail(restockTopList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("Restock Dashboard");

  // Fallback: if no argument is provided, pull from "Restock Dashboard" and slice top 10%
  if (!Array.isArray(restockTopList)) {
    const data = dashboardSheet.getRange("A2:F" + dashboardSheet.getLastRow()).getValues();
    const validRows = data.filter(row => row[0] && typeof row[0] === "string"); // Filter out empty rows
    const topPercentCount = Math.ceil(validRows.length * 0.1); // send the top 10%
    restockTopList = validRows.slice(0, topPercentCount).map(row => ({
      sku: row[0],
      name: row[1],
      qtySold: row[2],
      currentStock: row[3],
      minStock: row[4],
      restockNeeded: row[5]
    }));
  }

  const recipientStr = recipients.join(", ");
  const spreadsheetUrl = ss.getUrl();
  const today = new Date();
  const subject = `📦 Weekly Restock Summary – ${Utilities.formatDate(today, Session.getScriptTimeZone(), "dd MMM yyyy")}`;

  // Get updated timestamps
  const stockSheet = ss.getSheetByName("Jubelio_Stock");
  const summarySheet = ss.getSheetByName("Jubelio_Sales_90d");

  const stockUpdatedAt = stockSheet.getRange("I2").getValue();
  const summaryUpdatedAt = summarySheet.getRange("I2").getValue();

  const stockStr = stockUpdatedAt 
    ? Utilities.formatDate(new Date(stockUpdatedAt), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")
    : "Unknown";
  const summaryStr = summaryUpdatedAt 
    ? Utilities.formatDate(new Date(summaryUpdatedAt), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")
    : "Unknown";

  // Get Restock Dashboard last updated date (bottom of sheet)
  const dashboardLastRow = dashboardSheet.getLastRow();
  let dashboardUpdatedAt = null;

  for (let r = dashboardLastRow; r > 0; r--) {
    const label = dashboardSheet.getRange(r, 1).getValue();
    const value = dashboardSheet.getRange(r, 2).getValue();
    if (typeof label === 'string' && label.trim().toLowerCase() === "last updated:" && value) {
      dashboardUpdatedAt = new Date(value);
      break;
    }
  }

  const dashboardStr = dashboardUpdatedAt instanceof Date && !isNaN(dashboardUpdatedAt)
    ? Utilities.formatDate(dashboardUpdatedAt, Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")
    : "Unknown";

  // Build the email content
  let restockHtml = "";
  if (restockTopList.length === 0) {
    restockHtml = "<p>✅ All stocks are currently above minimum levels. No urgent restocks needed.</p>";
  } else {
    restockHtml = `
      <h3>🛒 Top Items Needing Restock</h3>
      <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
        <tr>
          <th>SKU</th>
          <th>Name</th>
          <th>Sold (90d)</th>
          <th>Stock</th>
          <th>Min</th>
          <th>Restock Needed</th>
        </tr>
        ${restockTopList.map(item => `
          <tr>
            <td>${item.sku}</td>
            <td>${item.name}</td>
            <td>${item.qtySold}</td>
            <td>${item.currentStock}</td>
            <td>${item.minStock}</td>
            <td><b>${item.restockNeeded}</b></td>
          </tr>`).join("")}
      </table>
    `;
  }

  const htmlBody = `
    <p>Hi! Here's your <b>weekly restock summary</b> based on the last 90 days of sales and current stock levels.</p>
    ${restockHtml}
    <p>🔗 <a href="${spreadsheetUrl}" target="_blank"><b>Click here to view the Google Sheet</b></a></p>
    <hr>
    <p><b>Data Calculation Range:</b></p>
    <ul>
      <li><b>Sales Summary:</b> Past 90 days ending at <b>${summaryStr}</b></li>
      <li><b>Current Stock:</b> Based on stock data as of <b>${stockStr}</b></li>
      <li><b>Restock Dashboard:</b> Last updated at <b>${dashboardStr}</b></li>
    </ul>
    <p>🕒 Email generated at: ${Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss")}</p>
  `;

  MailApp.sendEmail({
    to: recipientStr,
    subject,
    htmlBody
  });
}


/** ============================================================================
 * ⚙️ Main Process JUBELIO
 * ============================================================================
*/

/** Fetch Jubelio current stock RAW */
function getAllJubelioStockRaw() {
  const token = loginToJubelioWMS();
  if (!token) return [];

  const pageSize = 200;
  let page = 1;
  let allItems = [];

  while (true) {
    const url = `${JUBE_API_BASE}/inventory/?page=${page}&pageSize=${pageSize}`;
    const options = {
      method: 'get',
      headers: getJubelioAuthHeaders(token)
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    const items = result.data || [];

    if (items.length === 0) break;

    allItems = allItems.concat(items);
    if (items.length < pageSize) break;

    page++;
  }

  return allItems;
}

/** Retrieve Jubelio current stock as a MAP */
function fetchLiveJubelioStockMap() {
  const stockItems = getAllJubelioStockRaw(); // already fetches everything
  const stockMap = {};

  stockItems.forEach(item => {
    if (item.item_id && item.total_stocks?.available != null) {
      stockMap[item.item_id] = item.total_stocks.available;
    }
  });

  return stockMap;
}

/** Step 1. Login to Jubelio */
function loginToJubelioWMS() {
  const url = `${JUBE_API_BASE}/login`;
  const payload = {
    email: JUBE_EMAIL,
    password: JUBE_PASS
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  Logger.log(data.token);
  return data.token;
}

/** Step 2: Fetch all Jubelio invoice IDs and save to Google Sheet (monthly) */
function fetchJubelioInvoices() {
  const sheetName = "Jubelio_Invoices";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();

  // Clear or create Jubelio Invoices sheet
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  sheet.appendRow(["doc_id", "doc_number", "transaction_date", "customer_name", "store_name", "status"]);

  // Add timestamp to I1 and I2
  sheet.getRange("I1").setValue("Last Updated");
  sheet.getRange("I2").setValue(now);

  // Clear Jubelio_Invoice_Items sheet if exists
  const itemSheetName = "Jubelio_Invoice_Items";
  let itemSheet = ss.getSheetByName(itemSheetName);
  if (itemSheet) {
    itemSheet.clear();
    Logger.log("🧹 Cleared Jubelio_Invoice_Items sheet.");
  } else {
    itemSheet = ss.insertSheet(itemSheetName);
    Logger.log("🆕 Created Jubelio_Invoice_Items sheet.");
  }

  ss.toast("Preparing import...");
  const token = loginToJubelioWMS();
  const pageSize = 200;

  // Read from Config sheet
  let fromDate = getConfigValue("fromDate");
  let toDate = getConfigValue("toDate");

  // Fallback to last 90 days if needed
  if (!fromDate || !toDate) {
    const today = new Date();
    const ninetyDaysAgo = new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000);
    fromDate = ninetyDaysAgo.toISOString().split('.')[0];
    toDate = today.toISOString().split('.')[0];
    ss.toast("📅 Using auto date range (last 90 days)");
    Logger.log("📅 Fallback to default range:", fromDate, toDate);
  }

  let page = 1;
  let totalFetched = 0;
  let allRows = [];

  ss.toast("📡 Fetching invoices from Jubelio...");

  while (true) {
    const url = `${JUBE_API_BASE}/sales/invoices/?page=${page}&pageSize=${pageSize}&sortDirection=DESC&sortBy=doc_id&transactionDateFrom=${fromDate}&transactionDateTo=${toDate}`;
    const options = {
      method: 'get',
      headers: getJubelioAuthHeaders(token),
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const result = JSON.parse(response.getContentText());

      if (!result.data || result.data.length === 0) break;

      const rows = result.data.map(invoice => [
        invoice.doc_id,
        invoice.doc_number,
        invoice.transaction_date,
        invoice.customer_name,
        invoice.store_name,
        ""
      ]);

      allRows = allRows.concat(rows);
      totalFetched += rows.length;

      Logger.log(`Fetched page ${page}, total so far: ${totalFetched}`);
      ss.toast(`Fetched page ${page}, total so far: ${totalFetched}`);

      if (totalFetched >= result.totalCount) break;
      page++;

    } catch (error) {
      Logger.log("❌ Error on page " + page + ": " + error);
      ss.toast("❌ Error fetching data from API!");
      break;
    }
  }

  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);
  }

  ss.toast("✅ All Jubelio invoices fetched and written to sheet!");
  Logger.log("✅ Finished fetching Jubelio invoices.");

  // Create 1-minute trigger for item processing
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === 'fetchJubelioInvoiceItems');

  if (!exists) {
    ScriptApp.newTrigger('fetchJubelioInvoiceItems')
      .timeBased()
      .everyMinutes(1)
      .create();
    Logger.log("⏱️ Created 1-minute trigger for Jubelio invoice item batching.");
  } else {
    Logger.log("⚠️ 1-minute trigger already exists. Not creating a duplicate.");
  }
}

/** Step 3: Batched Fetching Jubelio Invoice Items and SKU Summary (monthly) */
function fetchJubelioInvoiceItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const toast = (msg, title = "Fetch Jubelio Invoice Items", timeout = 5) => ss.toast(msg, title, timeout);

  const lock = LockService.getScriptLock();
  const success = lock.tryLock(30000);
  if (!success) {
    toast("⏳ Another execution is already running. Skipping this run.");
    Logger.log("⏳ Another execution is already running. Skipping this run.");
    return;
  }

  const errorLogs = []; // ← Catch errors here

  try {
    toast("🚀 Starting to fetch Jubelio invoice items...");

    const invoiceSheet = ss.getSheetByName("Jubelio_Invoices");
    const itemSheet = ss.getSheetByName("Jubelio_Invoice_Items") || ss.insertSheet("Jubelio_Invoice_Items");

    if (itemSheet.getLastRow() === 0) {
      itemSheet.appendRow(["doc_id", "item_name", "item_code", "qty", "price", "total"]);
    }

    const token = loginToJubelioWMS();
    const dataRange = invoiceSheet.getRange(2, 1, invoiceSheet.getLastRow() - 1, 6); // A to F
    const data = dataRange.getValues();

    const totalInvoices = data.length;
    let allItemRows = [];
    let processedCount = 0;
    const now = new Date();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const docId = row[0];
      const status = row[5];

      if (status === "✅") continue;

      const url = `${JUBE_API_BASE}/sales/invoices/${docId}`;
      const options = {
        method: 'get',
        headers: getJubelioAuthHeaders(token)
      };

      try {
        const response = UrlFetchApp.fetch(url, options);
        const invoiceDetail = JSON.parse(response.getContentText());
        const items = invoiceDetail.items || [];

        if (items.length > 0) {
          const rows = items.map(item => [
            docId,
            item.item_name || "",
            item.item_code || "",
            parseFloat(item.qty_in_base) || 0,
            parseFloat(item.price) || 0,
            parseFloat(item.amount) || 0
          ]);
          allItemRows = allItemRows.concat(rows);
        }

        invoiceSheet.getRange(i + 2, 6).setValue("✅");
        processedCount++;

        if (processedCount >= 200) break; // batch size
      } catch (error) {
        const errMsg = `❌ Error fetching Jubelio invoice ${docId}: ${error}`;
        Logger.log(errMsg);
        errorLogs.push(errMsg); // ← Log error here
      }
    }

    if (allItemRows.length > 0) {
      const lastRow = itemSheet.getLastRow();
      itemSheet.getRange(lastRow + 1, 1, allItemRows.length, 6).setValues(allItemRows);
    }

    itemSheet.getRange("I1").setValue("Last Updated");
    itemSheet.getRange("I2").setValue(now);

    toast(`✅ Fetched ${processedCount} of ${totalInvoices} Jubelio invoices this batch.`);
    Logger.log(`✅ Fetched items from ${processedCount} Jubelio invoices this run.`);

    const remaining = invoiceSheet.getRange(2, 6, invoiceSheet.getLastRow() - 1).getValues().flat();
    const pendingCount = remaining.filter(status => status !== "✅").length;

    if (pendingCount === 0) {
      Logger.log("🎉 All Jubelio invoices processed. Generating SKU summary...");

      const summarySheet = ss.getSheetByName("Jubelio_Sales_90d") || ss.insertSheet("Jubelio_Sales_90d");
      summarySheet.clear();
      summarySheet.appendRow(["Item Name", "SKU", "Total Qty Sold"]);

      const itemData = itemSheet.getRange(2, 2, itemSheet.getLastRow() - 1, 3).getValues();
      const skuMap = {};

      itemData.forEach(([name, sku, qty]) => {
        if (!sku) return;
        if (!skuMap[sku]) {
          skuMap[sku] = { name, qty: 0 };
        }
        skuMap[sku].qty += parseFloat(qty || 0);
      });

      const summaryRows = Object.entries(skuMap)
        .map(([sku, obj]) => [obj.name, sku, obj.qty])
        .sort((a, b) => b[2] - a[2]); // sort by qty

      summarySheet.getRange(2, 1, summaryRows.length, 3).setValues(summaryRows);
      summarySheet.getRange("I1").setValue("Last Updated");
      summarySheet.getRange("I2").setValue(now);

      Logger.log("✅ SKU summary created and sorted by quantity.");
      toast("🎯 All items fetched. SKU summary is ready and sorted!");

      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === "fetchJubelioInvoiceItems") {
          ScriptApp.deleteTrigger(trigger);
          Logger.log("🧹 Deleted 1-minute trigger after finishing batch.");
        }
      });

      const status = errorLogs.length > 0 ? "⚠️ Completed with some errors" : "✅ All good";
      sendMonthlyStatusEmail(status, errorLogs); // ✅ Send status + error list
    }

  } finally {
    lock.releaseLock();
  }
}

/** Step 4: Fetch Jubelio current stock (daily) */
function fetchJubelioStock() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("🔒 Locked. Skipping.");
    return;
  }

  try {
    const stockItems = getAllJubelioStockRaw();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Jubelio_Stock") || ss.insertSheet("Jubelio_Stock");
    sheet.clear();

    sheet.getRange("A1:D1").setValues([["item_id", "item_name", "item_code", "available_stock"]]);
    sheet.getRange("I1").setValue("Last Updated");

    const rows = stockItems.map(item => {
      const shortName = (item.item_name || "").split(" ").slice(0, 7).join(" ");
      return [
        item.item_id || "",
        shortName,
        item.item_code || "",
        item.total_stocks?.available || 0
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }

    sheet.getRange("I2").setValue(new Date());
    generateRestockDashboard();
    SpreadsheetApp.getActive().toast(`✅ Stock updated for ${rows.length} items.`);
  } finally {
    lock.releaseLock();
  }
}

/** Step 5: Generate Restock Dasboard from step 3 and 4 data (daily)
function generateRestockDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Jubelio_Stock");
  const summarySheet = ss.getSheetByName("Jubelio_Sales_90d");
  const configSheet = ss.getSheetByName("Config");
  const dashboardSheet = ss.getSheetByName("Restock Dashboard") || ss.insertSheet("Restock Dashboard");

  const stockData = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 4).getValues(); // A:D
  const skuSummary = summarySheet.getRange(2, 2, summarySheet.getLastRow() - 1, 2).getValues(); // B:C
  const configData = configSheet.getRange(2, 9, configSheet.getLastRow() - 1, 2).getValues(); // I:J

  // Build lookup maps
  const leadTimeMap = {};
  configData.forEach(([sku, leadTime]) => {
    if (sku && leadTime) leadTimeMap[sku] = parseInt(leadTime, 10);
  });

  const stockMap = {};
  stockData.forEach(row => {
    const item_name = row[1];
    const item_code = row[2];
    const available_stock = parseFloat(row[3]) || 0;

    if (item_code) {
      stockMap[item_code] = {
        name: item_name,
        available: available_stock
      };
    }
  });

  const safetyDays = 7;
  const defaultLeadTime = 14;
  const restockList = [];

  const missingLeadTimeSKUs = [];
  const configSKUs = Object.keys(leadTimeMap);
  const unusedLeadTimeSKUs = configSKUs.filter(sku => !(sku in stockMap));

  skuSummary.forEach(([sku, qtySold]) => {
    const stockInfo = stockMap[sku];
    if (!stockInfo) return;

    const totalSold = parseFloat(qtySold) || 0;
    const dailyAvg = totalSold / 90;

    const leadTime = leadTimeMap[sku];
    if (leadTime === undefined) missingLeadTimeSKUs.push(sku);

    const restockPeriod = (leadTime !== undefined ? leadTime : defaultLeadTime) + safetyDays;
    const minStock = Math.ceil(dailyAvg * restockPeriod);
    const restockNeeded = Math.max(0, minStock - stockInfo.available);

    restockList.push({
      sku,
      name: stockInfo.name,
      qtySold: totalSold,
      currentStock: stockInfo.available,
      minStock,
      restockNeeded
    });
  });

  // Sort by priority: need restock + high sales
  restockList.sort((a, b) => {
    if (a.restockNeeded === 0 && b.restockNeeded > 0) return 1;
    if (a.restockNeeded > 0 && b.restockNeeded === 0) return -1;
    const aScore = (a.qtySold * 2) - (a.currentStock);
    const bScore =(b.qtySold * 2) - (b.currentStock);
    return bScore - aScore;
  });

  // Clear and write headers + data
  dashboardSheet.clear();
  const headers = ["Product Name", "SKU", "Qty Sold (90d)", "Current Stock", "Minimum Stock", "Restock Needed"];
  const dataRows = restockList.map(item => [
    item.name,
    item.sku,
    item.qtySold,
    item.currentStock,
    item.minStock,
    item.restockNeeded
  ]);
  dashboardSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (dataRows.length > 0) {
    dashboardSheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
  }

  // Append logs for missing configs below data
  let logRow = dashboardSheet.getLastRow() + 2;
  if (missingLeadTimeSKUs.length > 0) {
    dashboardSheet.getRange(logRow++, 1).setValue("⚠️ SKUs with Sales but Missing Lead Time Config");
    dashboardSheet.getRange(logRow++, 1).setValue("SKU");
    missingLeadTimeSKUs.forEach(sku => {
      dashboardSheet.getRange(logRow++, 1).setValue(sku);
    });
    logRow++;
  }

  if (unusedLeadTimeSKUs.length > 0) {
    dashboardSheet.getRange(logRow++, 1).setValue("⚠️ SKUs in Lead Time Config but Not in Current Stock");
    dashboardSheet.getRange(logRow++, 1).setValue("SKU");
    unusedLeadTimeSKUs.forEach(sku => {
      dashboardSheet.getRange(logRow++, 1).setValue(sku);
    });
    logRow++;
  }

  // Final last updated timestamp
  dashboardSheet.getRange(logRow, 1).setValue("Last Updated:");
  dashboardSheet.getRange(logRow, 2).setValue(new Date());

  // Reorder sheets from clean to raw data sheets
  reorderSheets();

  // Notification email for Restock. delete "//" to enable
  // sendWeeklyRestockEmail();
} */

/**function generateRestockDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Jubelio_Stock");
  const jubelioSheet = ss.getSheetByName("Jubelio_Sales_90d");
  const jurnalSheet = ss.getSheetByName("JURNAL_Sales_90d");
  const configSheet = ss.getSheetByName("Config");
  const dashboardSheet = ss.getSheetByName("Restock Dashboard") || ss.insertSheet("Restock Dashboard");

  const stockData = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 4).getValues(); // A:D
  const jubelioSales = jubelioSheet.getRange(2, 2, jubelioSheet.getLastRow() - 1, 2).getValues(); // B:C
  const jurnalSales = jurnalSheet.getRange(2, 2, jurnalSheet.getLastRow() - 1, 2).getValues(); // B:C
  const configData = configSheet.getRange(2, 9, configSheet.getLastRow() - 1, 2).getValues(); // I:J

  // Build leadTimeMap
  const leadTimeMap = {};
  configData.forEach(([sku, leadTime]) => {
    if (sku && leadTime) leadTimeMap[sku] = parseInt(leadTime, 10);
  });

  // Build stockMap
  const stockMap = {};
  stockData.forEach(row => {
    const item_name = row[1];
    const item_code = row[2];
    const available_stock = parseFloat(row[3]) || 0;

    if (item_code) {
      stockMap[item_code] = {
        name: item_name,
        available: available_stock
      };
    }
  });

  // Combine sales from Jubelio and Jurnal into one map
  const combinedSalesMap = {};

  jubelioSales.forEach(([sku, qty]) => {
    if (sku) {
      combinedSalesMap[sku] = (combinedSalesMap[sku] || 0) + (parseFloat(qty) || 0);
    }
  });

  jurnalSales.forEach(([sku, qty]) => {
    if (sku) {
      combinedSalesMap[sku] = (combinedSalesMap[sku] || 0) + (parseFloat(qty) || 0);
    }
  });

  const safetyDays = 7;
  const defaultLeadTime = 14;
  const restockList = [];

  const missingLeadTimeSKUs = [];
  const configSKUs = Object.keys(leadTimeMap);
  const unusedLeadTimeSKUs = configSKUs.filter(sku => !(sku in stockMap));

  Object.keys(combinedSalesMap).forEach(sku => {
    const stockInfo = stockMap[sku];
    if (!stockInfo) return;

    const totalSold = combinedSalesMap[sku];
    const dailyAvg = totalSold / 90;

    const leadTime = leadTimeMap[sku];
    if (leadTime === undefined) missingLeadTimeSKUs.push(sku);

    const restockPeriod = (leadTime !== undefined ? leadTime : defaultLeadTime) + safetyDays;
    const minStock = Math.ceil(dailyAvg * restockPeriod);
    const restockNeeded = Math.max(0, minStock - stockInfo.available);

    restockList.push({
      sku,
      name: stockInfo.name,
      qtySold: totalSold,
      currentStock: stockInfo.available,
      minStock,
      restockNeeded
    });
  });

  // Sort by priority
  restockList.sort((a, b) => {
    if (a.restockNeeded === 0 && b.restockNeeded > 0) return 1;
    if (a.restockNeeded > 0 && b.restockNeeded === 0) return -1;
    const aScore = (a.qtySold * 2) - (a.currentStock);
    const bScore = (b.qtySold * 2) - (b.currentStock);
    return bScore - aScore;
  });

  // Output to sheet
  dashboardSheet.clear();
  const headers = ["Product Name", "SKU", "Qty Sold (90d)", "Current Stock", "Minimum Stock", "Restock Needed"];
  const dataRows = restockList.map(item => [
    item.name,
    item.sku,
    item.qtySold,
    item.currentStock,
    item.minStock,
    item.restockNeeded
  ]);
  dashboardSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (dataRows.length > 0) {
    dashboardSheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
  }

  // Logs for config issues
  let logRow = dashboardSheet.getLastRow() + 2;
  if (missingLeadTimeSKUs.length > 0) {
    dashboardSheet.getRange(logRow++, 1).setValue("⚠️ SKUs with Sales but Missing Lead Time Config");
    dashboardSheet.getRange(logRow++, 1).setValue("SKU");
    missingLeadTimeSKUs.forEach(sku => {
      dashboardSheet.getRange(logRow++, 1).setValue(sku);
    });
    logRow++;
  }

  if (unusedLeadTimeSKUs.length > 0) {
    dashboardSheet.getRange(logRow++, 1).setValue("⚠️ SKUs in Lead Time Config but Not in Current Stock");
    dashboardSheet.getRange(logRow++, 1).setValue("SKU");
    unusedLeadTimeSKUs.forEach(sku => {
      dashboardSheet.getRange(logRow++, 1).setValue(sku);
    });
    logRow++;
  }

  dashboardSheet.getRange(logRow, 1).setValue("Last Updated:");
  dashboardSheet.getRange(logRow, 2).setValue(new Date());

  reorderSheets();
  // sendWeeklyRestockEmail();
}*/

function generateRestockDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Jubelio_Stock");
  const jubelioSheet = ss.getSheetByName("Jubelio_Sales_90d");
  const jurnalSheet = ss.getSheetByName("JURNAL_Sales_90d");
  const configSheet = ss.getSheetByName("Config");
  const priceSheet = ss.getSheetByName("Jurnal_Avg_Price"); // New sheet for average price
  const dashboardSheet = ss.getSheetByName("Restock Dashboard") || ss.insertSheet("Restock Dashboard");

  const stockData = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 4).getValues(); // A:D
  const jubelioSales = jubelioSheet.getRange(2, 2, jubelioSheet.getLastRow() - 1, 2).getValues(); // B:C
  const jurnalSales = jurnalSheet.getRange(2, 2, jurnalSheet.getLastRow() - 1, 2).getValues(); // B:C
  const configData = configSheet.getRange(2, 9, configSheet.getLastRow() - 1, 2).getValues(); // I:J
  const priceData = priceSheet.getRange(2, 2, priceSheet.getLastRow() - 1, 2).getValues(); // B:C (Product Code, Average Price)

  // Build leadTimeMap
  const leadTimeMap = {};
  configData.forEach(([sku, leadTime]) => {
    if (sku && leadTime) leadTimeMap[sku] = parseInt(leadTime, 10);
  });

  // Build stockMap
  const stockMap = {};
  stockData.forEach(row => {
    const item_name = row[1];
    const item_code = row[2];
    const available_stock = parseFloat(row[3]) || 0;

    if (item_code) {
      stockMap[item_code] = {
        name: item_name,
        available: available_stock
      };
    }
  });

  // Build averagePriceMap
  const averagePriceMap = {};
  priceData.forEach(([productCode, avgPrice]) => {
    if (productCode) {
      averagePriceMap[productCode] = parseFloat(avgPrice) || 0;
    }
  });

  // Combine sales from Jubelio and Jurnal into one map
  const combinedSalesMap = {};

  jubelioSales.forEach(([sku, qty]) => {
    if (sku) {
      combinedSalesMap[sku] = (combinedSalesMap[sku] || 0) + (parseFloat(qty) || 0);
    }
  });

  jurnalSales.forEach(([sku, qty]) => {
    if (sku) {
      combinedSalesMap[sku] = (combinedSalesMap[sku] || 0) + (parseFloat(qty) || 0);
    }
  });

  const safetyDays = 7;
  const defaultLeadTime = 14;
  const restockList = [];

  const missingLeadTimeSKUs = [];
  const configSKUs = Object.keys(leadTimeMap);
  const unusedLeadTimeSKUs = configSKUs.filter(sku => !(sku in stockMap));

  Object.keys(combinedSalesMap).forEach(sku => {
    const stockInfo = stockMap[sku];
    if (!stockInfo) return;

    const totalSold = combinedSalesMap[sku];
    const dailyAvg = totalSold / 90;

    const leadTime = leadTimeMap[sku];
    if (leadTime === undefined) missingLeadTimeSKUs.push(sku);

    const restockPeriod = (leadTime !== undefined ? leadTime : defaultLeadTime) + safetyDays;
    const minStock = Math.ceil(dailyAvg * restockPeriod);
    const restockNeeded = Math.max(0, minStock - stockInfo.available);

    const avgPrice = averagePriceMap[sku] || 0;

    restockList.push({
      sku,
      name: stockInfo.name,
      qtySold: totalSold,
      currentStock: stockInfo.available,
      minStock,
      restockNeeded,
      avgPrice
    });
  });

  // Sort by normalized priority score
  const wQty = 0.6;
  const wRev = 0.4;

  const maxQty = Math.max(...restockList.map(item => item.qtySold));
  const minQty = Math.min(...restockList.map(item => item.qtySold));
  const maxRev = Math.max(...restockList.map(item => item.qtySold * item.avgPrice));
  const minRev = Math.min(...restockList.map(item => item.qtySold * item.avgPrice));

  restockList.sort((a, b) => {
    if (a.restockNeeded === 0 && b.restockNeeded > 0) return 1;
    if (a.restockNeeded > 0 && b.restockNeeded === 0) return -1;

    const aQtyNorm = (maxQty === minQty) ? 0.5 : (a.qtySold - minQty) / (maxQty - minQty);
    const bQtyNorm = (maxQty === minQty) ? 0.5 : (b.qtySold - minQty) / (maxQty - minQty);

    const aRevenue = a.qtySold * a.avgPrice;
    const bRevenue = b.qtySold * b.avgPrice;

    const aRevNorm = (maxRev === minRev) ? 0.5 : (aRevenue - minRev) / (maxRev - minRev);
    const bRevNorm = (maxRev === minRev) ? 0.5 : (bRevenue - minRev) / (maxRev - minRev);

    const aScore = wQty * aQtyNorm + wRev * aRevNorm;
    const bScore = wQty * bQtyNorm + wRev * bRevNorm;

    return bScore - aScore;
  });

  // Output to sheet
  dashboardSheet.clear();
  const headers = ["Product Name", "SKU", "Qty Sold (90d)", "Current Stock", "Minimum Stock", "Restock Needed", "Avg Price"];
  const dataRows = restockList.map(item => [
    item.name,
    item.sku,
    item.qtySold,
    item.currentStock,
    item.minStock,
    item.restockNeeded,
    item.avgPrice
  ]);
  dashboardSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (dataRows.length > 0) {
    dashboardSheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
  }

  // Logs for config issues
  let logRow = dashboardSheet.getLastRow() + 2;
  if (missingLeadTimeSKUs.length > 0) {
    dashboardSheet.getRange(logRow++, 1).setValue("⚠️ SKUs with Sales but Missing Lead Time Config");
    dashboardSheet.getRange(logRow++, 1).setValue("SKU");
    missingLeadTimeSKUs.forEach(sku => {
      dashboardSheet.getRange(logRow++, 1).setValue(sku);
    });
    logRow++;
  }

  if (unusedLeadTimeSKUs.length > 0) {
    dashboardSheet.getRange(logRow++, 1).setValue("⚠️ SKUs in Lead Time Config but Not in Current Stock");
    dashboardSheet.getRange(logRow++, 1).setValue("SKU");
    unusedLeadTimeSKUs.forEach(sku => {
      dashboardSheet.getRange(logRow++, 1).setValue(sku);
    });
    logRow++;
  }

  dashboardSheet.getRange(logRow, 1).setValue("Last Updated:");
  dashboardSheet.getRange(logRow, 2).setValue(new Date());

  reorderSheets();
  // sendWeeklyRestockEmail();
}

/** Deduct Jurnal sales from jubelio stock(daily) */
function postDailyJurnalSubtractionsToJubelio() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JURNAL_Sales_Daily");
  const data = sheet.getRange(2, 8, sheet.getLastRow() - 1, 4).getValues(); // Columns H–K: item_id, item_code, name, quantity

  const liveStockMap = fetchLiveJubelioStockMap();

  const validItems = data
    .filter(row => row[0] && !isNaN(row[0]) && !isNaN(row[3])) // valid item_id and qty
    .map(row => {
      const itemId = Number(row[0]);
      const requestQty = Number(row[3]);
      const availableStock = liveStockMap[itemId] || 0;

      const subtractQty = Math.min(availableStock, requestQty);

      if (subtractQty <= 0) return null; // Skip if no stock to subtract

      return {
        item_adj_detail_id: 0,
        item_id: itemId,
        description: "Subtract Jurnal sales",
        serial_no: null,
        batch_no: null,
        qty_in_base: -subtractQty,
        original_item_adj_detail_id: 0,
        unit: "Buah",
        amount: 0,
        location_id: -1,
        account_id: 75,
        expired_date: null,
        bin_id: 3,
        cost: 0
      };
    })
    .filter(item => item !== null);

  if (validItems.length === 0) {
    Logger.log("🚫 No items to subtract after checking stock availability.");
    return;
  }

  const payload = {
    item_adj_id: 0,
    item_adj_no: "[auto]",
    transaction_date: new Date().toISOString(),
    note: "[AUTO] Daily sales subtraction from offline Jurnal",
    location_id: -1,
    is_opening_balance: false,
    items: validItems
  };

  const token = loginToJubelioWMS();
  if (!token) {
    Logger.log("❌ Failed to authenticate with Jubelio.");
    return;
  }

  const url = `${JUBE_API_BASE}/inventory/adjustments/`;
  const options = {
    method: "post",
    contentType: "application/json",
    headers: getJubelioAuthHeaders(token),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    Logger.log("📦 Payload to Jubelio:");
    Logger.log(JSON.stringify(payload, null, 2));
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    Logger.log("✅ Stock adjustment submitted:");
    Logger.log(result);
  } catch (e) {
    Logger.log("❌ Error submitting adjustment:");
    Logger.log(e.message);
  }
  // ✅ Trigger fresh stock fetch and regenerate dashboard
  fetchJubelioStock();
}
