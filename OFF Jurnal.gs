/** ============================================================================
 * ⚙️ Main Process JURNAL
 * =============================================================================
*/

/** Matches item SKU with Jubelio item_id */
function getItemIdMapFromStockSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jubelio_Stock");
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const codeIdx = header.indexOf("item_code");
  const idIdx = header.indexOf("item_id");

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const code = row[codeIdx];
    const id = row[idIdx];
    if (code && id) {
      map[code] = id;
    }
  }
  return map;
}

/** Fetch all Past 90d Raw Jurnal invoice IDs per items (daily) */
function fetchJurnalInvoiceItems() {
  const sheetName = "Jurnal_Invoice_Items";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Set headers if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Invoice No", "Transaction Date", "Customer", "Product Code", "Product Name",
      "Quantity", "Rate", "Amount"
    ]);
  }

  // Get existing invoice numbers
  let lastRow = sheet.getLastRow();
  let existingInvoiceNos = new Set();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    existingInvoiceNos = new Set(data.map(x => String(x).trim()));
  }

  const page = 1;
  const pageSize = 100;
  const requestPath = `/public/jurnal/api/v1/sales_invoices?page=${page}&page_size=${pageSize}`;
  const fullUrl = JURNAL_API_BASE + requestPath;
  const headers = getJurnalHmacHeaders("GET", requestPath);

  const response = UrlFetchApp.fetch(fullUrl, {
    method: "GET",
    headers,
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log("❌ Failed to fetch: " + response.getContentText());
    SpreadsheetApp.getActive().toast("❌ Error fetching Jurnal invoices. Check logs.");
    return;
  }

  const data = JSON.parse(response.getContentText());
  const invoices = data.sales_invoices || [];
  const newRows = [];
  const today = new Date();
  const ninetyDaysAgo = new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000);

  invoices.forEach(inv => {
    const rawInvoiceNo = String(inv.transaction_no || "").trim();
    const customer = inv.person?.display_name || "";

    if (existingInvoiceNos.has(rawInvoiceNo)) return;
    if (customer === "EVEREST AUTO ACCESORIES") return;

    const rawDate = inv.transaction_date || "";
    const [day, month, year] = rawDate.split("/").map(str => parseInt(str, 10));
    const parsedDate = new Date(year, month - 1, day);

    const lines = inv.transaction_lines_attributes || [];

    lines.forEach(line => {
      const prod = line.product || {};
      const code = prod.product_code || prod.code || "";
      const name = prod.name || "Unnamed";
      const qty = line.quantity || 0;
      const rate = line.rate || "";
      const amount = line.amount || "";

      newRows.push([
        rawInvoiceNo,
        parsedDate,
        customer,
        code,
        name,
        parseFloat(qty),
        parseFloat(rate),
        parseFloat(amount)
      ]);
    });

    existingInvoiceNos.add(rawInvoiceNo);
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  // Format column 2 (Transaction Date)
  const finalRow = sheet.getLastRow();
  if (finalRow > 1) {
    sheet.getRange(2, 2, finalRow - 1).setNumberFormat("dd/mm/yyyy");
    sheet.getRange(2, 1, finalRow - 1, 8).sort([
      { column: 2, ascending: false },
      { column: 1, ascending: false }
    ]);
  }

  // Delete rows older than 90 days
  let oldInvoice = 0;
  const allDates = sheet.getRange(2, 2, sheet.getLastRow() - 1).getValues();
  for (let i = allDates.length - 1; i >= 0; i--) {
    const rowDate = allDates[i][0];
    if (rowDate instanceof Date && rowDate < ninetyDaysAgo) {
      sheet.deleteRow(i + 2);
      oldInvoice++;
    }
  }

  // Last updated timestamp
  sheet.getRange("I1").setValue("Last Updated");
  sheet.getRange("I2").setValue(new Date());
  Logger.log(`✅ Jurnal invoice items synced. New rows: ${newRows.length}`);
  Logger.log(`✅ Jurnal invoice items deleted. Deleted rows: ${oldInvoice}`);
  SpreadsheetApp.getActive().toast(`✅ Jurnal invoice items synced. New rows: ${newRows.length}`);
  generateJurnalDailySalesSummary();
}

/** Summarize SKU Sales Qty per DAY and save to Google Sheet (daily) */
function generateJurnalDailySalesSummary() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jurnal_Invoice_Items");
  const targetSheetName = "JURNAL_Sales_Daily";
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName)
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet(targetSheetName);
  
  targetSheet.clear(); // Reset the daily summary sheet
  targetSheet.appendRow(["Date", "Product Name", "Product Code", "Qty Sold"]);

  if (!sourceSheet || sourceSheet.getLastRow() <= 1) {
    Logger.log("❌ No data in source sheet.");
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const header = data[0];
  const dateIdx = header.indexOf("Transaction Date");
  const nameIdx = header.indexOf("Product Name");
  const codeIdx = header.indexOf("Product Code");
  const qtyIdx = header.indexOf("Quantity");

  if (dateIdx === -1 || nameIdx === -1 || codeIdx === -1 || qtyIdx === -1) {
    Logger.log("❌ Required columns not found.");
    return;
  }

  const salesMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawDate = row[dateIdx];
    const name = row[nameIdx] || "Unnamed Product";
    const code = row[codeIdx] || "";
    const qty = parseFloat(row[qtyIdx]) || 0;

    // Normalize date (dd/mm/yyyy)
    if (!(rawDate instanceof Date)) continue;
    //const day = rawDate.getDate().toString().padStart(2, '0');
    //const month = (rawDate.getMonth() + 1).toString().padStart(2, '0');
    //const year = rawDate.getFullYear();
    //const dateKey = `${day}/${month}/${year}`;
    const dateKey = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    const key = `${dateKey}|||${name}|||${code}`;
    if (!salesMap[key]) {
      salesMap[key] = {
        date: dateKey,
        name,
        code,
        qty: 0
      };
    }
    salesMap[key].qty += qty;
  }

  const rows = Object.values(salesMap).map(s => [
    s.date, s.name, s.code, s.qty
  ]);

  if (rows.length > 0) {
    targetSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    targetSheet.getRange(2, 1, rows.length, 4).sort([
      { column: 1, ascending: false },
      { column: 4, ascending: false }
    ]);
  }

  targetSheet.getRange("F1").setValue("Last Updated");
  targetSheet.getRange("F2").setValue(new Date());
  
  Logger.log(`✅ Jurnal daily summary generated. Rows: ${rows.length}`);
  fetchJurnalWarehouseTransfers();
}

/** Fetch Jurnal warehouse transfers (daily) */
function fetchJurnalWarehouseTransfers() {
  const sheetName = "Jurnal_WH_Transfers";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Add header if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Transfer No", "Date", "From Warehouse", "To Warehouse", "Product Code",
      "Product Name", "Quantity", "Product Unit"
    ]);
  }

  const transferNosCol = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
  const existingNos = new Set(transferNosCol.flat().map(val => String(val).trim().toUpperCase()));

  const pageSize = 30;
  const requestPath = `/public/jurnal/api/v1/warehouse_transfers?page_size=${pageSize}`;
  const fullUrl = JURNAL_API_BASE + requestPath;
  const headers = getJurnalHmacHeaders("GET", requestPath);

  try {
    const response = UrlFetchApp.fetch(fullUrl, {
      method: "GET",
      headers,
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code !== 200) {
      Logger.log(`❌ Error fetching transfers. Code: ${code}, Message: ${response.getContentText()}`);
      SpreadsheetApp.getActive().toast("❌ Error fetching transfers.");
      return;
    }

    const data = JSON.parse(response.getContentText());
    const transfers = data?.warehouse_transfers || [];
    const newRows = [];
    const newTransferNos = new Set();

    for (const transfer of transfers) {
      const transferNoStr = String(transfer?.transaction_no || "").trim().toUpperCase();
      if (!transferNoStr || existingNos.has(transferNoStr)) continue;

      newTransferNos.add(transferNoStr);

      const detailPath = `/public/jurnal/api/v1/warehouse_transfers/${encodeURIComponent(transferNoStr)}`;
      const detailUrl = JURNAL_API_BASE + detailPath;
      const detailHeaders = getJurnalHmacHeaders("GET", detailPath);

      const detailResponse = UrlFetchApp.fetch(detailUrl, {
        method: "GET",
        headers: detailHeaders,
        muteHttpExceptions: true
      });

      const detailCode = detailResponse.getResponseCode();
      if (detailCode !== 200) continue;

      const detailData = JSON.parse(detailResponse.getContentText());
      const detail = detailData?.warehouse_transfer;
      const lines = detail?.warehouse_transfer_line_attributes || [];

      let transferDate;
      if (detail.transaction_date && detail.transaction_date.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        const [dd, mm, yyyy] = detail.transaction_date.split("/").map(Number);
        transferDate = new Date(yyyy, mm - 1, dd);
      } else {
        transferDate = "";
      }

      for (const line of lines) {
        const productCode = String(line.product?.code || "").trim().toUpperCase();
        const productName = String(line.product?.name || "").trim();
        const productUnit = String(line.product?.unit || "").trim();

        newRows.push([
          transferNoStr,
          transferDate,
          detail.from_warehouse?.name || "Unassigned",
          detail.to_warehouse?.name || "Unassigned",
          productCode,
          productName,
          Number(line.quantity) || 0,
          productUnit
        ]);
      }

      Utilities.sleep(5000); // Respect API limits
    }

    if (newRows.length > 0) {
      sheet.insertRowsAfter(1, newRows.length);  // Insert blank rows below the header
      sheet.getRange(2, 1, newRows.length, 8).setValues(newRows);  // Fill them with new data
      sheet.getRange(2, 2, newRows.length).setNumberFormat("dd/mm/yyyy"); // ✅ Fix the date format for the new rows in column B (2nd column)
      Logger.log(`✅ ${newTransferNos.size} new transfer no(s) found, added ${newRows.length} row(s) at the top.`);
      SpreadsheetApp.getActive().toast(`✅ ${newTransferNos.size} new transfers added at top.`);
    } else {
      Logger.log("ℹ️ No new transfers to add.");
      SpreadsheetApp.getActive().toast("ℹ️ No new transfers to add.");
    }

    // ✅ Prune rows older than 90 days (no full clear!)
    const today = new Date();
    const threshold = new Date(today);
    threshold.setDate(today.getDate() - 90);
    const thresholdTime = threshold.getTime();

    let deletedCount = 0;
    for (let i = sheet.getLastRow(); i >= 2; i--) {
      const dateVal = sheet.getRange(i, 2).getValue();
      const date = new Date(dateVal);
      if (date instanceof Date && !isNaN(date) && date.getTime() < thresholdTime) {
        sheet.deleteRow(i);
        deletedCount++;
      }
    }

    const remainingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    const remainingTransferNos = [...new Set(remainingData.map(r => String(r[0]).trim().toUpperCase()))];

    if (remainingData.length > 0) {
      const dates = remainingData.map(r => new Date(r[1]));
      const earliestDate = new Date(Math.min(...dates));
      const latestDate = new Date(Math.max(...dates));

      const formatDate = (d) => `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;

      const sortedTransferNos = remainingTransferNos.slice().sort();
      const smallestTransferNo = sortedTransferNos[0];
      const largestTransferNo = sortedTransferNos[sortedTransferNos.length - 1];

      Logger.log(`✅ Pruned ${deletedCount} outdated rows.`);
      Logger.log(`📦 ${remainingTransferNos.length} unique transfer no(s) remain, ${remainingData.length} rows remaining.`);
      Logger.log(`📅 Date range: ${formatDate(earliestDate)} - ${formatDate(latestDate)}`);
      Logger.log(`🔢 Transfer No range: Smallest: ${smallestTransferNo}, Largest: ${largestTransferNo}`);
    } else {
      Logger.log("ℹ️ No rows remain after pruning.");
    }

  } catch (error) {
    Logger.log("❌ Exception: " + error.message);
    SpreadsheetApp.getActive().toast("❌ Error during transfer fetch.");
    return
  }
  // ✅ Set Last Updated timestamp in I1 and I2
  sheet.getRange("I1:I100").clear();
  sheet.getRange("I1").setValue("Last Updated");
  const now = new Date();
  sheet.getRange("I2").setValue(now);
  sheet.getRange("I2").setNumberFormat("dd/mm/yyyy hh:mm:ss");

  summarizeTodayJurnalOutbound()
}

/** Combine sales and WH transfer to get daily outbound (daily) */
function summarizeTodayJurnalOutbound() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = ss.getSheetByName("Jurnal_Invoice_Items");
  const transferSheet = ss.getSheetByName("Jurnal_WH_Transfers");
  const outputSheet = ss.getSheetByName("JURNAL_Sales_Daily");
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  //const today = new Date();
  //today.setDate(today.getDate() - 1); // Move date back by 1 day
  //const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");

  const itemIdMap = getItemIdMapFromStockSheet(); // { item_code: item_id }

  const outboundMap = {}; // key: code || name → { id, code, name, qty }

  // === 1. Process Jurnal Invoice Items ===
  const invoiceData = invoiceSheet.getDataRange().getValues();
  const invoiceHeaders = invoiceData[0];
  const invoiceRows = invoiceData.slice(1);

  const idxProductCode = invoiceHeaders.indexOf("Product Code");
  const idxName = invoiceHeaders.indexOf("Product Name");
  const idxQty = invoiceHeaders.indexOf("Quantity");
  const idxDate = invoiceHeaders.indexOf("Transaction Date");

  for (const row of invoiceRows) {
    const rawDate = row[idxDate];
    if (!rawDate) continue;

    const dateStr = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
    if (dateStr !== todayStr) continue;

    const code = String(row[idxProductCode] || "").toUpperCase().trim();
    const name = String(row[idxName] || "").trim();
    const qty = Number(row[idxQty]) || 0;
    if (!name || isNaN(qty)) continue;

    const key = code || name; // use code if available, otherwise name
    const id = code ? (itemIdMap[code] || "") : "";

    if (!outboundMap[key]) {
      outboundMap[key] = { id, code, name, qty };
    } else {
      outboundMap[key].qty += qty;
    }
  }

  // === 2. Process Warehouse Transfers (Unassigned → Aneka MGK only) ===
  const transferData = transferSheet.getDataRange().getValues();
  const transferHeaders = transferData[0];
  const transferRows = transferData.slice(1);

  const idxTransferDate = transferHeaders.indexOf("Date");
  const idxFromWh = transferHeaders.indexOf("From Warehouse");
  const idxToWh = transferHeaders.indexOf("To Warehouse");
  const idxTransferCode = transferHeaders.indexOf("Product Code");
  const idxTransferName = transferHeaders.indexOf("Name");
  const idxTransferQty = transferHeaders.indexOf("Quantity");

  for (const row of transferRows) {
    const rawDate = row[idxTransferDate];
    if (!rawDate) continue;

    const dateStr = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
    if (dateStr !== todayStr) continue;

    const fromWh = String(row[idxFromWh]).trim();
    const toWh = String(row[idxToWh]).trim();
    if (!(fromWh === "Unassigned" && toWh === "Aneka MGK")) continue;

    const code = String(row[idxTransferCode] || "").toUpperCase().trim();
    const name = String(row[idxTransferName] || "").trim();
    const qty = Number(row[idxTransferQty]) || 0;
    if (!name || isNaN(qty)) continue;

    const key = code || name;
    const id = code ? (itemIdMap[code] || "") : "";

    if (!outboundMap[key]) {
      outboundMap[key] = { id, code, name, qty };
    } else {
      outboundMap[key].qty += qty;
    }
  }

  // === 3. Output Summary ===
  const summary = Object.values(outboundMap).map(item => [
    item.id, item.code, item.name, item.qty
  ]);

  // Write to columns H–K
  const startCol = 8; // Column H
  outputSheet.getRange(1, startCol, 1, 4).setValues([["item_id", "item_code", "name", "quantity"]]);

  const maxRowsToClear = 100;
  outputSheet.getRange(2, startCol, maxRowsToClear, 4).clearContent();

  if (summary.length > 0) {
    outputSheet.getRange(2, startCol, summary.length, 4).setValues(summary);
  }
  outputSheet.getRange("M1").setValue("Last Updated");
  outputSheet.getRange("M2").setValue(new Date());
  Logger.log(`✅ Outbound summary updated. ${summary.length} SKUs.`);
  postDailyJurnalSubtractionsToJubelio();
}

/** Summarize into 90 DAYS SKU Sales Qty (monthly) */
function fetchJurnalSales90d() {
  const sourceSheetName = "Jurnal_Invoice_Items";
  const targetSheetName = "JURNAL_Sales_90d";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const targetSheet = ss.getSheetByName(targetSheetName) || ss.insertSheet(targetSheetName);
  targetSheet.clear(); // Reset output sheet

  // Write headers
  targetSheet.appendRow(["Product Name", "Product Code", "Total Qty Sold (90d)"]);

  if (!sourceSheet) {
    SpreadsheetApp.getActive().toast("❌ 'Jurnal_Invoice_Items' sheet not found.");
    Logger.log("❌ 'Jurnal_Invoice_Items' sheet not found.");
    targetSheet.getRange("I1").setValue("Last Updated");
    targetSheet.getRange("I2").setValue(new Date());
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getActive().toast("❌ No invoice data to summarize.");
    Logger.log("❌ No invoice data to summarize.");
    targetSheet.getRange("I1").setValue("Last Updated");
    targetSheet.getRange("I2").setValue(new Date());
    return;
  }

  const header = data[0];
  const nameIdx = header.indexOf("Product Name");
  const codeIdx = header.indexOf("Product Code");
  const qtyIdx = header.indexOf("Quantity");

  if (nameIdx === -1 || codeIdx === -1 || qtyIdx === -1) {
    SpreadsheetApp.getActive().toast("❌ Required columns not found in 'Jurnal_Invoice_Items'.");
    Logger.log("❌ Required columns not found in 'Jurnal_Invoice_Items'.");
    targetSheet.getRange("I1").setValue("Last Updated");
    targetSheet.getRange("I2").setValue(new Date());
    return;
  }

  const skuMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameIdx] || "Unnamed Product";
    const code = row[codeIdx] || "";
    const qty = parseFloat(row[qtyIdx]) || 0;

    const key = `${name}|||${code}`;
    if (!skuMap[key]) {
      skuMap[key] = {
        name: name,
        code: code,
        totalQty: 0
      };
    }
    skuMap[key].totalQty += qty;
  }

  const finalRows = Object.values(skuMap).map(sku => [
    sku.name,
    sku.code,
    sku.totalQty
  ]);

  if (finalRows.length > 0) {
    targetSheet.getRange(2, 1, finalRows.length, 3).setValues(finalRows);
    targetSheet.getRange(2, 1, finalRows.length, 3).sort({ column: 3, ascending: false });  //sort from highest to lowest qty sold
  }

  // Last updated timestamp
  targetSheet.getRange("I1").setValue("Last Updated");
  targetSheet.getRange("I2").setValue(new Date());

  fetchJurnalAvgPriceFromSales();
  SpreadsheetApp.getActive().toast("✅ Summarized Jurnal 90-day sales.");
  Logger.log("✅ Summarized Jurnal 90-day sales.");
}

/** Fetch average selling price from Jurnal (monthly) */
function fetchJurnalAvgPriceFromSales() {
  const sheetName = "Jurnal_Avg_Price";

  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(endDate.getDate() - 180);

  const startDateStr = formatDDMMYYYY(startDate);
  const endDateStr = formatDDMMYYYY(endDate);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear(); // Clear old data
  }

  // Set headers
  sheet.getRange("A1:C1").setValues([["Product Name", "Product Code", "Average Price"]]);

  const requestPath = `/public/jurnal/api/v1/sales_by_products?start_date=${encodeURIComponent(startDateStr)}&end_date=${encodeURIComponent(endDateStr)}`;
  const fullUrl = JURNAL_API_BASE + requestPath;
  const headers = getJurnalHmacHeaders("GET", requestPath);

  try {
    const response = UrlFetchApp.fetch(fullUrl, {
      method: "GET",
      headers,
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const now = new Date();

    if (code !== 200) {
      Logger.log("Failed to fetch data: " + code);
      sheet.getRange("I1").setValue("Last Updated");
      sheet.getRange("I2").setValue(now);
      sheet.getRange("I3").setValue("Error");
      sheet.getRange("I4").setValue("HTTP " + code);
      return;
    }

    const data = JSON.parse(response.getContentText());
    const records = data.sales_by_products?.reports?.products || [];

    const output = [];

    for (const item of records) {
      const product = item.product;
      const code = product.product_code;
      //if (!code || code.trim() === "") continue; // Skips product with no product code(SKU)

      const name = product.product_name || "";
      const avgPriceStr = product.average_price || "0";

      // Convert "1.510.764,33" → 1510764.33 → 1510764
      const avgPriceRaw = parseFloat(avgPriceStr.replace(/\./g, '').replace(',', '.')) || 0;
      const avgPrice = Math.round(avgPriceRaw);  // round to nearest whole number

      output.push([name, code, avgPrice]);
    }

    if (output.length > 0) {
      sheet.getRange(2, 1, output.length, 3).setValues(output);
    }

    // Write last updated timestamp and date range used
    sheet.getRange("H1").setValue("Start Date");
    sheet.getRange("H2").setValue(startDate);
    sheet.getRange("I1").setValue("Last Updated");
    sheet.getRange("I2").setValue(now);

  } catch (error) {
    const now = new Date();
    Logger.log("Exception: " + error.message);
    sheet.getRange("I1").setValue("Last Updated");
    sheet.getRange("I2").setValue(now);
    sheet.getRange("I3").setValue("Error");
    sheet.getRange("I4").setValue(error.message);
  }
}

/** Fetch Jurnal Purchases and summarize (daily) */
function syncJurnalDailyPurchases() {
  const targetSheetName = "JURNAL_Purchase_Daily";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName) || ss.insertSheet(targetSheetName);

  // Reset and setup header
  sheet.clear();
  sheet.appendRow(["Date", "Product Name", "Product Code", "Qty Purchased"]);

  const page = 1;
  const pageSize = 100;
  const requestPath = `/public/jurnal/api/v1/purchase_invoices?page=${page}&page_size=${pageSize}`;
  const fullUrl = JURNAL_API_BASE + requestPath;
  const headers = getJurnalHmacHeaders("GET", requestPath); // Assuming you already have this function

  const response = UrlFetchApp.fetch(fullUrl, {
    method: "GET",
    headers,
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log("❌ Failed to fetch purchase data: " + response.getContentText());
    SpreadsheetApp.getActive().toast("❌ Error fetching Jurnal purchases.");
    return;
  }

  const data = JSON.parse(response.getContentText());
  const purchases = data.purchase_invoices || [];
  const today = new Date();
  const ninetyDaysAgo = new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000);

  const purchaseMap = {};

  purchases.forEach(inv => {
    const rawDate = inv.transaction_date || "";
    const [day, month, year] = rawDate.split("/").map(str => parseInt(str, 10));
    const parsedDate = new Date(year, month - 1, day);
    if (parsedDate < ninetyDaysAgo) return;

    const lines = inv.transaction_lines_attributes || [];

    lines.forEach(line => {
      const product = line.product || {};
      const name = product.name || "Unnamed Product";
      const code = product.product_code || product.code || "";
      const qty = parseFloat(line.quantity || 0);

      const dateKey = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      const key = `${dateKey}|||${name}|||${code}`;
      if (!purchaseMap[key]) {
        purchaseMap[key] = {
          date: dateKey,
          name,
          code,
          qty: 0
        };
      }
      purchaseMap[key].qty += qty;
    });
  });

  const rows = Object.values(purchaseMap).map(p => [
    p.date, p.name, p.code, p.qty
  ]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    sheet.getRange(2, 1, rows.length, 4).sort([
      { column: 1, ascending: false },
      { column: 4, ascending: false }
    ]);
  }

  sheet.getRange("F1").setValue("Last Updated");
  sheet.getRange("F2").setValue(new Date());
  Logger.log(`✅ Jurnal daily purchase summary generated. Rows: ${rows.length}`);
  summarizeTodayJurnalInbound()
}

/** Combine purchases and WH transfer to get daily INBOUND (daily) */
function summarizeTodayJurnalInbound() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("JURNAL_Purchase_Daily");
  const transferSheet = ss.getSheetByName("Jurnal_WH_Transfers");
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  // if want to test out previous days
  //const yesterday = new Date();
  //yesterday.setDate(yesterday.getDate() - 1);
  //const todayStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "dd/MM/yyyy");

  const itemIdMap = getItemIdMapFromStockSheet(); // { item_code: item_id }
  const inboundMap = {}; // key: code || name → { id, code, name, qty }

  // === 1. From Purchase Summary Sheet ===
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  const idxDate = header.indexOf("Date");
  const idxName = header.indexOf("Product Name");
  const idxCode = header.indexOf("Product Code");
  const idxQty = header.indexOf("Qty Purchased");

  for (const row of rows) {
    const rawDate = row[idxDate];
    if (!rawDate) continue;
    const dateStr = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
    if (dateStr !== todayStr) continue;

    const name = String(row[idxName] || "").trim();
    const code = String(row[idxCode] || "").toUpperCase().trim();
    const qty = Number(row[idxQty]) || 0;
    const key = code || name;
    const id = code ? (itemIdMap[code] || "") : "";

    if (!inboundMap[key]) {
      inboundMap[key] = { id, code, name, qty };
    } else {
      inboundMap[key].qty += qty;
    }
  }

  // === 2. From Warehouse Transfers (Unassigned → Aneka MGK or Supplier → Aneka MGK) ===
  const transferData = transferSheet.getDataRange().getValues();
  const transferHeaders = transferData[0];
  const transferRows = transferData.slice(1);

  const idxTransferDate = transferHeaders.indexOf("Date");
  const idxFromWh = transferHeaders.indexOf("From Warehouse");
  const idxToWh = transferHeaders.indexOf("To Warehouse");
  const idxTransferCode = transferHeaders.indexOf("Product Code");
  const idxTransferName = transferHeaders.indexOf("Product Name");
  const idxTransferQty = transferHeaders.indexOf("Quantity");

  for (const row of transferRows) {
    const rawDate = row[idxTransferDate];
    if (!rawDate) continue;

    const dateStr = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
    if (dateStr !== todayStr) continue;

    const fromWh = String(row[idxFromWh]).trim();
    const toWh = String(row[idxToWh]).trim();
    if (!(fromWh === "Aneka MGK" && toWh === "Unassigned")) continue;

    const name = String(row[idxTransferName] || "").trim();
    const code = String(row[idxTransferCode] || "").toUpperCase().trim();
    const qty = Number(row[idxTransferQty]) || 0;
    const key = code || name;
    const id = code ? (itemIdMap[code] || "") : "";

    if (!inboundMap[key]) {
      inboundMap[key] = { id, code, name, qty };
    } else {
      inboundMap[key].qty += qty;
    }
  }

  // === 3. Output to columns F–I of the same sheet ===
  const summary = Object.values(inboundMap).map(item => [
    item.id, item.code, item.name, item.qty
  ]);

  const startCol = 8; // Column H
  const maxRowsToClear = 100;
  sheet.getRange(1, startCol, 1, 4).setValues([["item_id", "item_code", "name", "quantity"]]);
  sheet.getRange(2, startCol, maxRowsToClear, 4).clearContent();

  if (summary.length > 0) {
    sheet.getRange(2, startCol, summary.length, 4).setValues(summary);
  }

  sheet.getRange("M1").setValue("Inbound Last Updated");
  sheet.getRange("M2").setValue(new Date());

  Logger.log(`✅ Inbound summary updated. ${summary.length} SKUs.`);
  postDailyJurnalAdditionsToJubelio();
}
