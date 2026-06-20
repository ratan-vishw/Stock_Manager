/**
 * INVENTORY MANAGEMENT SYSTEM - BACKEND LOGIC
 * DATABASE FORMAT: 5-column scheme (Item Name, Stock, Unit, Rate, Item Image)
 * BILLS FORMAT: Exactly matches your 13-column Bills layout
 */
const TRANSACTIONS_SHEET_NAME = 'All_Transactions';
const SITE_LIST_SHEET_NAME = 'SiteNameList';
const UPLOAD_FOLDER_NAME = "Inventory_Bills_Uploads";
const MT_SHEET_PREFIX = 'MT.';
const HEADER_ROW_COUNT = 1;

/**
 * Helper to dynamically fetch folders by name while verifying active (non-trashed) state [1].
 * Creates the folder if it does not already exist.
 */
function getOrCreateActiveFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  let folder;
  while (folders.hasNext()) {
    const candidate = folders.next();
    if (!candidate.isTrashed()) {
      folder = candidate;
      break;
    }
  }
  if (!folder) {
    folder = DriveApp.createFolder(folderName);
  }
  return folder;
}

/**
 * Converts any Google Drive URL to a direct-viewable thumbnail URL.
 * Works with /file/d/ID/..., /open?id=ID, /thumbnail?id=ID, etc.
 * Returns original URL if it's not a Drive link.
 */
function toDriveThumbnailUrl(url, size) {
  if (!url || typeof url !== 'string') return '';
  url = url.trim();
  if (!url) return '';
  if (url.includes('/thumbnail?')) return url;
  let fileId = '';
  let match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match) { fileId = match[1]; }
  else {
    match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match) fileId = match[1];
  }
  if (fileId) {
    return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w' + (size || 400);
  }
  return url;
}

/**
 * ONE-TIME MIGRATION: Run this ONCE from the Apps Script editor to convert
 * all MT.* sheets from 8-column format to 5-column format.
 */
function migrateSheets_To5Columns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let migrated = 0;
  let skipped = 0;

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(MT_SHEET_PREFIX)) return;

    const lastCol = sheet.getLastColumn();

    if (lastCol <= 5) {
      if (lastCol === 5) {
        const h5 = sheet.getRange(1, 5).getValue().toString().trim();
        if (h5 !== 'Image Link' && h5 !== 'Item Image') {
          sheet.getRange(1, 5).setValue('Image Link');
        }
      } else if (lastCol === 4) {
        sheet.getRange(1, 5).setValue('Image Link');
      }
      skipped++;
      return;
    }

    const colsToDelete = lastCol - 4;
    sheet.deleteColumns(5, colsToDelete);

    sheet.insertColumnAfter(4);
    sheet.getRange(1, 5).setValue('Image Link');

    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
    sheet.setFrozenRows(1);

    migrated++;
    console.log('Migrated: ' + name + ' (had ' + lastCol + ' cols → now 5 cols)');
  });

  SpreadsheetApp.flush();
  const msg = '✅ Migration complete!\n\n' +
    '• Migrated: ' + migrated + ' sheet(s)\n' +
    '• Skipped (already correct): ' + skipped + ' sheet(s)';
  console.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// ── BILL MANAGER CONSTANTS ──────────────────────────────────────────────────
const BILLS_SHEET_NAME   = 'Bills';
const BILLS_FOLDER_NAME  = 'Bill Manager Images';

// Retrieve API keys directly from Script Properties to reduce exposure vectors
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Inventory Transaction Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// AI BILL EXTRACTION (GEMINI VISION)
// ============================================================================

function extractBillDataFromImage(base64Data, mimeType) {
  if (!GEMINI_API_KEY) {
    throw new Error("Configuration Error: GEMINI_API_KEY is not defined in Script Properties.");
  }
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;

  const systemPrompt = `
    You are an expert data extraction assistant. I will provide an image of a bill or invoice.
    Extract the line items. Return ONLY a strictly formatted JSON array of objects. Do not wrap the JSON in markdown blocks (no \`\`\`json).
    
    Each object must have the following keys:
    - "itemName" (string): The description of the goods (e.g., "Wacker WN Black 280ml" or "12MM CLEAR TOUGHENED GLASS").
    - "quantity" (number): The total quantity or Sq.Mt. Do not include unit strings here.
    - "rate" (number): The price or rate per unit.
    - "unit" (string): The unit of measurement (e.g., "pcs", "box", "Sq.Mt"). If not found, use "pcs".
    - "invoiceNumber" (string): The document number. Look for "Invoice No", "TAX INVOICE", "Proforma No", or "Work Order No". Keep this consistent across all items in the same bill.
    
    If the document has multiple line items, return multiple objects in the array.
  `;

  const payload = {
    contents: [{
      parts: [
        { text: systemPrompt },
        { inlineData: { mimeType: mimeType, data: base64Data } }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error("Gemini API error " + response.getResponseCode() + ": " + response.getContentText());
    }
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      throw new Error(json.error.message);
    }
    
    if (!json.candidates || json.candidates.length === 0 || 
        !json.candidates[0].content || !json.candidates[0].content.parts || 
        json.candidates[0].content.parts.length === 0) {
      throw new Error("No text or content candidates returned from Gemini API.");
    }
    
    let extractedText = json.candidates[0].content.parts[0].text;
    extractedText = extractedText.replace(/^```json\s*/i, "").replace(/```\s*$/, "").trim();
    return extractedText; 
    
  } catch (e) {
    throw new Error("Failed to extract data: " + e.toString());
  }
}

function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sites = [];
  try {
    const siteSheet = ss.getSheetByName(SITE_LIST_SHEET_NAME);
    if (siteSheet) {
      const lastRow = siteSheet.getLastRow();
      if (lastRow > 0) {
        sites = siteSheet.getRange(1, 1, lastRow, 1).getValues()
          .flat()
          .map(s => s.toString().trim())
          .filter(String);
      }
    }
  } catch (e) {
    console.error("Error fetching sites: " + e.toString());
  }

  const categories = [];
  const itemsMap = {};
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.startsWith(MT_SHEET_PREFIX)) {
      categories.push(sheetName);
      const lastRow = sheet.getLastRow();
      if (lastRow > HEADER_ROW_COUNT) {
        const lastCol = sheet.getLastColumn();
        const numCols = Math.min(5, lastCol > 0 ? lastCol : 5);
        const data = sheet.getRange(HEADER_ROW_COUNT + 1, 1, lastRow - HEADER_ROW_COUNT, numCols).getValues();
        const items = data.map((row) => ({
          name: (row[0] || '').toString().trim(),
          stock: parseFloat(row[1] || 0),
          unit: (row[2] || '').toString().trim() || 'pcs',
          rate: parseFloat(row[3] || 0),
          imageUrl: numCols >= 5 ? toDriveThumbnailUrl((row[4] || '').toString().trim(), 400) : '',
          category: sheetName,
          searchStr: `${row[0]} (Stock: ${row[1]})`
        })).filter(item => item.name);
        itemsMap[sheetName] = items;
      } else {
        itemsMap[sheetName] = [];
      }
    }
  });

  return {
    sites: sites,
    categories: categories,
    itemsByCategory: itemsMap,
    userEmail: Session.getActiveUser().getEmail()
  };
}

function getRecentTransactions(limit = 50) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transSheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
    if (!transSheet) return [];
    const lastRow = transSheet.getLastRow();
    if (lastRow <= HEADER_ROW_COUNT) return [];
    const numRows = Math.min(limit, lastRow - HEADER_ROW_COUNT);
    const startRow = Math.max(HEADER_ROW_COUNT + 1, lastRow - numRows + 1);
    const values = transSheet.getRange(startRow, 1, numRows, 15).getValues();
    return values.map(row => {
      let timestamp = row[1];
      if (timestamp instanceof Date) {
        timestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      }
      return {
        transactionId: row[0],
        timestamp: timestamp,
        site: row[2] || '',
        type: row[4] || '',
        category: row[5] || '',
        itemName: row[6] || '',
        qty: parseFloat(row[7]) || 0,
        unit: row[8] || '',
        rate: parseFloat(row[9]) || 0
      };
    }).reverse();
  } catch (e) {
    console.error('getRecentTransactions failed:', e);
    return [];
  }
}

function addCategory(categoryName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `${MT_SHEET_PREFIX}${categoryName}`;
  if (ss.getSheetByName(sheetName)) throw new Error(`Category "${categoryName}" already exists`);
  const newSheet = ss.insertSheet(sheetName);
  const headerValues = ['Item Name / Description', 'Current Stock', 'Unit', 'Rate (₹)', 'Image Link'];
  newSheet.getRange(1, 1, 1, headerValues.length)
    .setValues([headerValues])
    .setFontWeight('bold')
    .setBackground('#4CAF50')
    .setFontColor('white');
  newSheet.setFrozenRows(1);
  return { success: true, message: `Category "${categoryName}" created`, sheetName: sheetName };
}

function addItem(formData) {
  const { itemName, itemCategory, itemUnit, currentQty, rate, imageBase64, imageMime, imageName } = formData;
  if (!itemName?.trim()) throw new Error('Item name is required.');
  if (!itemCategory?.trim()) throw new Error('Category is required.');
  if (!itemUnit?.trim()) throw new Error('Unit is required.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(itemCategory.trim());
  if (!sheet) throw new Error(`Sheet ${itemCategory} not found`);

  const targetBase = itemName.split('/')[0].trim().toLowerCase();
  const dataRows = sheet.getLastRow() - HEADER_ROW_COUNT;
  const existingItems = dataRows > 0
    ? sheet.getRange(HEADER_ROW_COUNT + 1, 1, dataRows, 1).getValues().flat()
    : [];
  for (let existing of existingItems) {
    if (existing && existing.toString().split('/')[0].trim().toLowerCase() === targetBase) {
      throw new Error(`An item with this name already exists in ${itemCategory}`);
    }
  }

  let imageUrl = "";
  if (imageBase64) {
    try {
      const folder = getOrCreateActiveFolder(UPLOAD_FOLDER_NAME);
      const blob = Utilities.newBlob(Utilities.base64Decode(imageBase64), imageMime, imageName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      imageUrl = file.getUrl();
    } catch(err) {
      imageUrl = "Upload Error: " + err.toString();
    }
  }

  const stock = Math.max(0, parseFloat(currentQty) || 0);
  const parsedRate = Math.max(0, parseFloat(rate) || 0);

  sheet.appendRow([
    itemName.trim(),
    stock,
    itemUnit.trim(),
    parsedRate,
    imageUrl
  ]);
  return {
    success: true,
    item: {
      name: itemName.trim(),
      stock: stock,
      unit: itemUnit.trim(),
      rate: parsedRate,
      imageUrl: imageUrl ? toDriveThumbnailUrl(imageUrl, 400) : '',
      category: itemCategory.trim()
    }
  };
}

function submitForm(transactionList) {
  // Lock mechanism prevents stock corruption issues during rapid/concurrent form submissions
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error("Unable to obtain process update lock. Please try submitting again shortly.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transSheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
    const timeZone = Session.getScriptTimeZone();
    const folder = getOrCreateActiveFolder(UPLOAD_FOLDER_NAME);

    const rowsToAppend = [];
    const stockUpdates = [];
    const baseTransactionId = "TXN" + Utilities.formatDate(new Date(), timeZone, "yyyyMMddHHmmss");

    const runningStock = {};

    transactionList.forEach(txn => {
      const cat        = (txn.category || "").trim();
      const itemNameRaw = (txn.name || "").toString().trim();
      const searchName  = itemNameRaw.toLowerCase();
      const key         = `${cat}|${searchName}`;

      if (runningStock.hasOwnProperty(key)) return; 

      if (txn.isNewItem) {
        if (!cat) throw new Error(`Category is required to create new item "${itemNameRaw}"`);
        const sheet = ss.getSheetByName(cat);
        if (!sheet) throw new Error(`Category sheet "${cat}" not found. Create the category first.`);

        const existingData = sheet.getLastRow() > HEADER_ROW_COUNT
          ? sheet.getRange(HEADER_ROW_COUNT + 1, 1, sheet.getLastRow() - HEADER_ROW_COUNT, 1).getValues().flat()
          : [];
        const alreadyExists = existingData.some(n => (n || '').toString().trim().toLowerCase() === searchName);

        if (!alreadyExists) {
        const unit      = (txn.unitOverride || txn.unit || 'pcs').toString().trim();
        const rate      = parseFloat(txn.rateOverride || txn.rate || 0) || 0;
        let imageUrl    = txn.itemImageUrl || '';

        if (!imageUrl && txn.newItemImageData && txn.newItemImageData.data) {
          try {
            const imgBlob = Utilities.newBlob(
              Utilities.base64Decode(txn.newItemImageData.data),
              txn.newItemImageData.mime || 'image/jpeg',
              txn.newItemImageData.name || 'item.jpg'
            );
            const imgFile = folder.createFile(imgBlob);
            imgFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
            imageUrl = imgFile.getUrl();
          } catch (imgErr) {
            imageUrl = "Upload Error: " + imgErr.toString();
          }
        }

        sheet.appendRow([itemNameRaw, 0, unit, rate, imageUrl]);
      }

        runningStock[key] = 0;

      } else {
        if (cat) {
          const sheet = ss.getSheetByName(cat);
          if (sheet && sheet.getLastRow() > HEADER_ROW_COUNT) {
            const data = sheet.getDataRange().getValues();
            for (let i = 1; i < data.length; i++) {
              if ((data[i][0] || '').toString().trim().toLowerCase() === searchName) {
                const s = parseFloat(data[i][1] || 0);
                runningStock[key] = s < 0 ? 0 : s;
                break;
              }
            }
          }
        }
        if (!runningStock.hasOwnProperty(key)) runningStock[key] = 0;
      }
    });

    const batchFileUrlCache = {};
    const billRowsToInsert = []; 

    transactionList.forEach((txn, index) => {
      let finalFileUrl = (txn.preUploadedFileUrl || "").toString().trim();

      if (!finalFileUrl && txn.fileData && txn.fileData.data) {
        const cacheKey = (txn.fileData.name || "file") + "|" + txn.fileData.data.length;
        if (batchFileUrlCache[cacheKey]) {
          finalFileUrl = batchFileUrlCache[cacheKey];
        } else {
          try {
            const blob = Utilities.newBlob(
              Utilities.base64Decode(txn.fileData.data),
              txn.fileData.mime || 'application/octet-stream',
              txn.fileData.name
            );
            finalFileUrl = folder.createFile(blob).getUrl();
            batchFileUrlCache[cacheKey] = finalFileUrl; 
          } catch (e) {
            finalFileUrl = "Upload Error: " + e.toString();
          }
        }
      }

      const cat         = (txn.category || "").trim();
      const itemNameRaw  = (txn.name || "").toString().trim();
      const searchName   = itemNameRaw.toLowerCase();
      const stockKey     = `${cat}|${searchName}`;

      let currentStock = runningStock[stockKey] || 0;
      let qty          = parseFloat(txn.quantity || 0);
      let newStock     = currentStock;

      if (txn.transactionType === 'In') {
        newStock = currentStock + qty;
      } else if (txn.transactionType === 'Out' || txn.transactionType === 'Repair') {
        if (qty > currentStock) {
          throw new Error(`❌ INSUFFICIENT STOCK for ${itemNameRaw}: ${currentStock} available, trying to remove ${qty}`);
        }
        newStock = currentStock - qty;
      }

      runningStock[stockKey] = newStock; 

      let finalItemName = itemNameRaw;
      if (txn.additionalDescription && txn.additionalDescription.trim() !== "") {
        finalItemName += " | " + txn.additionalDescription.trim();
      }

      const transactionDate = new Date(txn.transactionDateTime);
      const dateString      = Utilities.formatDate(transactionDate, timeZone, "dd/MM/yyyy HH:mm");
      const billDateStr     = Utilities.formatDate(transactionDate, timeZone, "yyyy-MM-dd");

      const effectiveUnit = txn.unitOverride || txn.unit || 'pcs';
      const effectiveRate = parseFloat(txn.rateOverride !== null && txn.rateOverride !== undefined && txn.rateOverride !== ''
        ? txn.rateOverride : txn.rate) || 0;
      const itemTotalAmt  = (parseFloat(txn.quantity) || 0) * effectiveRate;

      rowsToAppend.push([
        baseTransactionId + "-" + (index + 1),
        dateString,
        txn.siteName,
        txn.processedBy,
        txn.transactionType,
        txn.category,
        finalItemName,
        txn.quantity,
        effectiveUnit,
        effectiveRate,
        itemTotalAmt || 0,
        currentStock,
        newStock,
        txn.invoiceNumber || '',
        finalFileUrl
      ]);

      // ── AUTO-SAVE TO "Bills" SHEET ──
      if (txn.saveToBills && finalFileUrl) {
        const billRecordId = "BILL-AUTO-" + Utilities.formatDate(new Date(), timeZone, "yyyyMMddHHmmss") + "-" + (index + 1);
        
        let extractedPiNo = "";
        const searchTarget = ((txn.invoiceNumber || "") + " " + (txn.additionalDescription || "")).toLowerCase();
        const piMatch = searchTarget.match(/\bpi[\s\-\#\.no]*(\d+)/);
        if (piMatch) {
          extractedPiNo = "PI-" + piMatch[1];
        }

        billRowsToInsert.push([
          billRecordId,                               
          txn.siteName || "",                         
          "Direct Inventory Upload",                  
          finalItemName,                              
          txn.quantity || "1",                        
          effectiveUnit,                              
          itemTotalAmt || "0",                        
          finalFileUrl,                               
          billDateStr,                                
          txn.invoiceNumber || "N/A",                 
          extractedPiNo,                              
          "Automatically stored from Inventory Card", 
          txn.additionalDescription || ""             
        ]);
      }

      if (['In', 'Out', 'Repair'].includes(txn.transactionType)) {
        stockUpdates.push({ category: cat, searchName: searchName, name: itemNameRaw, newStock: newStock });
      }
    });

    if (rowsToAppend.length > 0) {
      transSheet.getRange(transSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
        .setValues(rowsToAppend);
    }

    if (billRowsToInsert.length > 0) {
      const billsSheet = ss.getSheetByName(BILLS_SHEET_NAME) || ss.insertSheet(BILLS_SHEET_NAME);
      if (billsSheet.getLastRow() === 0) {
        const headers = [
          "Bill_ID", "Site_Location", "Vendor Name", "item Name", "Qty", "Unit", 
          "Total_Bill_Amount", "Bill_Image", "Bill date", "invoice No.", 
          "Pi No. for glass Bills", "Additional_Data_From_Bill", "Additional_Description"
        ];
        billsSheet.appendRow(headers);
        billsSheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#0f172a')
          .setFontColor('#ffffff');
        billsSheet.setFrozenRows(1);
      }
      billsSheet.getRange(billsSheet.getLastRow() + 1, 1, billRowsToInsert.length, 13)
        .setValues(billRowsToInsert);
    }

    // Batch updates execution loop writes memory changes in one transaction block
    if (stockUpdates.length > 0) {
      const updatesBySheet = {};
      stockUpdates.forEach(u => {
        if (!updatesBySheet[u.category]) updatesBySheet[u.category] = {};
        updatesBySheet[u.category][u.searchName] = u; 
      });

      for (const sheetName in updatesBySheet) {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) continue;
        const lastRow = sheet.getLastRow();
        if (lastRow <= HEADER_ROW_COUNT) continue;
        
        const stockRange = sheet.getRange(HEADER_ROW_COUNT + 1, 1, lastRow - HEADER_ROW_COUNT, 2);
        const stockValues = stockRange.getValues();
        const sheetUpdates = updatesBySheet[sheetName];
        let changed = false;

        for (let i = 0; i < stockValues.length; i++) {
          const rowName = (stockValues[i][0] || '').toString().trim().toLowerCase();
          const update = sheetUpdates[rowName];
          if (update) {
            stockValues[i][1] = Math.max(0, update.newStock);
            changed = true;
          }
        }

        if (changed) {
          stockRange.setValues(stockValues);
        }
      }
    }

    const siteSheet = ss.getSheetByName(SITE_LIST_SHEET_NAME);
    if (siteSheet) {
      const lastSiteRow = siteSheet.getLastRow();
      let existingSites = [];
      if (lastSiteRow > 0) {
        existingSites = siteSheet.getRange(1, 1, lastSiteRow, 1).getValues()
          .flat().map(s => s.toString().trim().toLowerCase()).filter(String);
      }
      const newSites = [];
      transactionList.forEach(txn => {
        const sName = (txn.siteName || "").trim();
        if (sName) {
          const sLower = sName.toLowerCase();
          if (!existingSites.includes(sLower) && !newSites.map(s => s.toLowerCase()).includes(sLower)) {
            newSites.push(sName);
          }
        }
      });
      if (newSites.length > 0) {
        const targetRow = lastSiteRow > 0 ? lastSiteRow + 1 : 1;
        siteSheet.getRange(targetRow, 1, newSites.length, 1).setValues(newSites.map(s => [s]));
      }
    }

    SpreadsheetApp.flush();

    const firstFileUrl = rowsToAppend.length > 0 ? (rowsToAppend[0][14] || '') : '';

    return {
      success: true,
      count: rowsToAppend.length,
      savedBillsCount: billRowsToInsert.length, 
      fileUrl: firstFileUrl
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// BILL MANAGER — GEMINI EXTRACTION
// ============================================================================

function extractBillWithGemini(base64Data, mimeType) {
  if (!GEMINI_API_KEY) {
    throw new Error("Configuration Error: GEMINI_API_KEY is not defined in Script Properties.");
  }
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;

  const prompt =
    "Identify receipt details. Extract: Date (YYYY-MM-DD), invoice/receipt number, " +
    "vendor or merchant name, proforma invoice number (PI No.) as piNo, any extra raw data in additionalData, and extra text in " +
    "additionalDescription. IMPORTANT: Extract EACH purchased item separately. " +
    "Output STRICTLY as a JSON object matching this structure: " +
    '{ "billDate": "YYYY-MM-DD", "invoiceNo": "...", "vendorName": "...", "piNo": "...", ' +
    '"additionalData": "...", "additionalDescription": "...", ' +
    '"items": [ { "itemName": "...", "qty": "...", "unit": "...", "totalAmount": "..." } ] }';

  const payload = {
    contents: [{
      parts: [
        { inlineData: { mimeType: mimeType, data: base64Data } },
        { text: prompt }
      ]
    }],
    generationConfig: { responseMimeType: "application/json" }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error("Gemini API error " + response.getResponseCode() + ": " + response.getContentText());
    }
    const result = JSON.parse(response.getContentText());
    
    if (!result || !result.candidates || result.candidates.length === 0 || 
        !result.candidates[0].content || !result.candidates[0].content.parts || 
        result.candidates[0].content.parts.length === 0) {
      throw new Error("No text or content candidates returned from Gemini API.");
    }
    
    let rawText = result.candidates[0].content.parts[0].text;
    
    // Fallback cleanup if model wraps response in Markdown format blocks
    rawText = rawText.replace(/^```json\s*/i, "").replace(/```\s*$/, "").trim();
    
    try {
      return JSON.parse(rawText);
    } catch (parseErr) {
      throw new Error("JSON parse failed on Gemini response: " + parseErr.message + " | Raw text: " + rawText);
    }
  } catch (e) {
    throw new Error("Bill extraction pipeline failed: " + e.toString());
  }
}

// ============================================================================
// BILL MANAGER — SAVE ROWS TO "Bills" SHEET (ALIGNED TO EXACT 13 HEADERS)
// ============================================================================

function saveBillItems(billData) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const tz    = Session.getScriptTimeZone();
    let   sheet = ss.getSheetByName(BILLS_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(BILLS_SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      const headers = [
        "Bill_ID", "Site_Location", "Vendor Name", "item Name", "Qty", "Unit", 
        "Total_Bill_Amount", "Bill_Image", "Bill date", "invoice No.", 
        "Pi No. for glass Bills", "Additional_Data_From_Bill", "Additional_Description"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#0f172a')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    let driveLink = "";
    if (billData.imageData) {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(billData.imageData),
          billData.imageMime || "image/jpeg",
          billData.imageName || "receipt.jpg"
        );
        const folder = getOrCreateActiveFolder(BILLS_FOLDER_NAME);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
        driveLink = file.getUrl();
      } catch (uploadErr) {
        driveLink = "Upload Error: " + uploadErr.toString();
      }
    }

    const baseId        = "BILL-" + Utilities.formatDate(new Date(), tz, "yyyyMMddHHmmss");
    const items         = billData.items || [];
    const rowsToInsert  = [];

    items.forEach(function(item, idx) {
      rowsToInsert.push([
        baseId + "-" + (idx + 1),               
        billData.siteLocation      || "",       
        billData.vendorName        || "",       
        item.itemName              || "Materials", 
        item.qty                   || "1",      
        item.unit                  || "pcs",    
        item.totalAmount           || "0",      
        driveLink,                              
        billData.billDate          || "",       
        billData.invoiceNo         || "N/A",    
        billData.piNo              || "",       
        billData.additionalData    || "",       
        billData.additionalDescription || ""    
      ]);
    });

    if (rowsToInsert.length > 0) {
      sheet
        .getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, 13)
        .setValues(rowsToInsert);
    }

    SpreadsheetApp.flush();
    return { success: true, count: rowsToInsert.length, driveLink: driveLink };

  } catch (e) {
    throw new Error("Bills sheet error: " + e.toString());
  }
}

function getHistoryLogs() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const tz    = Session.getScriptTimeZone();
    const sheet = ss.getSheetByName(BILLS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];

    // Dynamically query bottom-most 50 elements from raw data rows (newest first)
    const lastRow  = sheet.getLastRow();
    const maxRows  = 50;
    const startRow = Math.max(2, lastRow - maxRows + 1);
    const numRows  = lastRow - startRow + 1;

    const values  = sheet.getRange(startRow, 1, numRows, 13).getValues();
    const logs    = [];

    for (let i = values.length - 1; i >= 0; i--) {
      const row = values[i];
      logs.push({
        id:                   row[0],
        siteLocation:         row[1],
        vendorName:           row[2],
        itemName:             row[3],
        qty:                  row[4],
        unit:                 row[5],
        totalAmount:          row[6],
        driveLink:            row[7],
        billDate:             row[8] instanceof Date
                                ? Utilities.formatDate(row[8], tz, "yyyy-MM-dd")
                                : (row[8] || ""),
        invoiceNo:            row[9],
        piNo:                 row[10]
      });
    }
    return logs;
  } catch (e) {
    return [];
  }
}
