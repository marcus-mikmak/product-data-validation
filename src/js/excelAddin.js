// ==========================================
// GLOBAL CONSTANTS
// ==========================================
const TEMPLATE_SHEET_NAME = 'Product Data';
const INTERNAL_PREFIX = 'INTERNAL_';
const LOCALE_PREFIX = 'locale::';

// IMPORTANT: Update this to match your actual GitHub URL
const BASE_URL = 'https://YOUR_GITHUB_USERNAME.github.io/YOUR_REPO_NAME/src/html/';

let currentDialog; // Store the active dialog instance

// ==========================================
// INITIALIZATION
// ==========================================
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Add-in is ready.");
    
    // If this script is loaded inside the Sidebar Taskpane, wire up the button
    const startBtn = document.getElementById('startBtn');
    if (startBtn) {
      // Assuming a simplistic batch call for the manual taskpane validation
      startBtn.onclick = () => console.log("Taskpane validation triggered.");
    }
  }
});

// ==========================================
// 1. RIBBON BUTTON COMMANDS (Triggered by Manifest)
// ==========================================
function btnShowImportDialog(event) {
  openDialog('ImportDialog.html', 70, 70, event);
}

function btnShowMinotaurImportDialog(event) {
  openDialog('MinotaurImportDialog.html', 40, 50, event);
}

function btnShowMinotaurExportDialog(event) {
  openDialog('MinotaurExportDialog.html', 80, 80, event);
}

function btnExportGoodData(event) {
  processAndExportToTab('good').then(() => event.completed());
}

function btnExportReviewData(event) {
  processAndExportToTab('review').then(() => event.completed());
}

// Map the functions so the Excel XML Manifest can find them
Office.actions.associate("btnShowImportDialog", btnShowImportDialog);
Office.actions.associate("btnShowMinotaurImportDialog", btnShowMinotaurImportDialog);
Office.actions.associate("btnShowMinotaurExportDialog", btnShowMinotaurExportDialog);
Office.actions.associate("btnExportGoodData", btnExportGoodData);
Office.actions.associate("btnExportReviewData", btnExportReviewData);

// ==========================================
// 2. DIALOG MANAGEMENT & EVENT LISTENER
// ==========================================
function openDialog(htmlFileName, heightPct, widthPct, event) {
  const url = BASE_URL + htmlFileName;
  Office.context.ui.displayDialogAsync(url, { height: heightPct, width: widthPct, displayInIframe: true }, (asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Failed) {
      currentDialog = asyncResult.value;
      currentDialog.addEventHandler(Office.EventType.DialogMessageReceived, processDialogMessage);
    } else {
      console.error("Error opening dialog:", asyncResult.error.message);
    }
    if (event) event.completed(); // Tell Excel Ribbon the button click is done
  });
}

async function processDialogMessage(arg) {
  const message = JSON.parse(arg.message);

  try {
    switch (message.action) {
      // Sent by MinotaurExportDialog.html and ImportDialog.html
      case "GET_MIKMAK_HEADERS":
      case "GET_TEMPLATE_HEADERS":
        const headerData = await getHeadersAndSample();
        currentDialog.messageChild(JSON.stringify({ 
          action: message.action === "GET_MIKMAK_HEADERS" ? "SEND_MIKMAK_HEADERS" : "SEND_TEMPLATE_HEADERS", 
          payload: headerData 
        }));
        break;

      // Sent by MinotaurImportDialog.html
      case "MINOTAUR_IMPORT":
        await executeMinotaurBatch(message.csvData);
        currentDialog.close();
        break;

      // Sent by ImportDialog.html
      case "IMPORT_DATA":
        const parsedCsv = parseSimpleCsv(message.csvData);
        await processImportBatch(parsedCsv, message.mapping);
        currentDialog.close();
        break;

      // Sent by MinotaurExportDialog.html
      case "PROCESS_MINOTAUR_EXPORT":
        await processMinotaurExportToTab(message.mapping);
        currentDialog.close();
        break;
    }
  } catch (err) {
    console.error("Error processing dialog request:", err);
  }
}

// ==========================================
// 3. CORE EXCEL BATCH & PROCESSING FUNCTIONS
// ==========================================

/** * Translates checkImageVisibility() from UrlFetchApp to native browser fetch().
 */
async function checkImageVisibility(url) {
  if (!url || typeof url !== 'string' || !url.startsWith('http')) return false;

  const imageExtensions = ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.avif'];
  const hasImageExtension = imageExtensions.some(ext => url.toLowerCase().includes(ext));

  try {
    const response = await fetch(url, {
      method: 'GET',
      headers: { 'Range': 'bytes=0-1' } // Fetch only first byte to save bandwidth
    });

    const contentType = (response.headers.get('content-type') || '').toLowerCase();

    if (response.status === 200 || response.status === 206) {
      if (contentType.includes('image/') || hasImageExtension) {
        return true;
      }
    }
    return false;
  } catch (e) {
    console.error(`Error checking ${url}:`, e);
    return false;
  }
}

/** * Replaces prepareImport() and executeImportBatch().
 */
async function processImportBatch(parsedCsvRows, headerMap) {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const templateSheet = sheets.getItem(TEMPLATE_SHEET_NAME);

    // 1. Get Headers from the Template Sheet (Row 0)
    const headerRange = templateSheet.getRange("1:1").getUsedRange();
    headerRange.load("values");
    await context.sync();

    const templateHeaders = headerRange.values[0];
    const mappableHeaders = templateHeaders.filter(h => !h.startsWith(INTERNAL_PREFIX));
    const renderCheckGlobalIdx = templateHeaders.indexOf('INTERNAL_Image Render Check');

    const csvHeaders = parsedCsvRows[0];
    const dataRows = parsedCsvRows.slice(1);

    const imgCache = {};
    const processedOutput = [];
    const renderChecks = [];

    // 2. Map and Process Rows
    for (const row of dataRows) {
      const remappedRow = mappableHeaders.map(th => {
        const mappingValue = headerMap[th];
        if (!mappingValue || mappingValue === "-- Skip this column --" || mappingValue === "-- Skip --") return "";
        if (mappingValue.startsWith(LOCALE_PREFIX)) return mappingValue.substring(LOCALE_PREFIX.length);

        const csvIdx = csvHeaders.indexOf(mappingValue);
        return csvIdx > -1 ? row[csvIdx] : "";
      });

      // Perform Image Check
      let isVisible = "";
      const imgIdx = mappableHeaders.indexOf('image_link');
      if (imgIdx > -1) {
        const url = remappedRow[imgIdx];
        if (url) {
          if (!(url in imgCache)) {
            imgCache[url] = await checkImageVisibility(url);
          }
          isVisible = imgCache[url];
        }
      }

      processedOutput.push(remappedRow);
      renderChecks.push([isVisible]);
    }

    // 3. Clear Existing Data (leaving Row 1 headers intact)
    const usedRange = templateSheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    if (usedRange.rowCount > 1) {
      const rangeToClear = templateSheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, templateHeaders.length);
      rangeToClear.clear();
    }

    // 4. Bulk Write Remapped Data
    if (processedOutput.length > 0) {
      const targetDataRange = templateSheet.getRangeByIndexes(
        1, 0, 
        processedOutput.length, 
        mappableHeaders.length
      );
      targetDataRange.values = processedOutput;

      // Bulk Write Image Check Column if present
      if (renderCheckGlobalIdx > -1) {
        const targetCheckRange = templateSheet.getRangeByIndexes(
          1, renderCheckGlobalIdx, 
          renderChecks.length, 1
        );
        targetCheckRange.values = renderChecks;
      }
    }

    await context.sync();
    return processedOutput.length;
  });
}

/** * Translates processAndExportToTab().
 */
async function processAndExportToTab(type) {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const sourceSheet = sheets.getItem(TEMPLATE_SHEET_NAME);
    const usedRange = sourceSheet.getUsedRange();
    
    usedRange.load(["values", "rowCount", "columnCount"]);
    await context.sync();

    const allData = usedRange.values;
    if (!allData || allData.length < 2) {
      console.warn("No data rows found to export.");
      return;
    }

    const headers = allData[0];
    const rows = allData.slice(1);

    const readinessScoreIndex = headers.indexOf('INTERNAL_Data_Readiness_Score');
    if (readinessScoreIndex === -1) {
      throw new Error("Cannot find 'INTERNAL_Data_Readiness_Score' column.");
    }

    const allowedInternalHeaders = [
      'INTERNAL_Data_Readiness_Score',
      'INTERNAL_Plain_Text_Product_Data_Status'
    ];

    const colIndicesToKeep = [];
    const headersToKeep = [];

    headers.forEach((header, index) => {
      const isPublic = !header.startsWith(INTERNAL_PREFIX);
      const isAllowedInternal = (type === 'review' && allowedInternalHeaders.includes(header));
      if (isPublic || isAllowedInternal) {
        colIndicesToKeep.push(index);
        headersToKeep.push(header);
      }
    });

    const filteredRows = rows.filter(row => {
      const numericScore = parseFloat(row[readinessScoreIndex]);
      return (type === 'good') ? numericScore === 100 : numericScore !== 100;
    });

    if (filteredRows.length === 0) {
      console.log("No products match the readiness criteria for export.");
      return;
    }

    const dataRowsToKeep = filteredRows.map(row => colIndicesToKeep.map(idx => row[idx]));
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const newSheetName = `${type === 'good' ? 'Good_Data' : 'Review_Data'}_${timestamp}`;
    const newSheet = sheets.add(newSheetName);

    const headerRange = newSheet.getRangeByIndexes(0, 0, 1, headersToKeep.length);
    headerRange.values = [headersToKeep];

    const textKeywords = ['id', 'ean', 'upc', 'gtin', 'barcode', 'mpn', 'modelname'];
    headersToKeep.forEach((header, index) => {
      if (textKeywords.some(keyword => header.toLowerCase().includes(keyword))) {
        const colRange = newSheet.getRangeByIndexes(1, index, dataRowsToKeep.length, 1);
        colRange.numberFormat = "@"; 
      }
    });

    const dataRange = newSheet.getRangeByIndexes(1, 0, dataRowsToKeep.length, headersToKeep.length);
    dataRange.values = dataRowsToKeep;

    newSheet.getUsedRange().format.autofitColumns();
    newSheet.activate();

    await context.sync();
  });
}

/** * Fetch Headers & Sample Data for Mapping Dialogs 
 */
async function getHeadersAndSample() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(TEMPLATE_SHEET_NAME);
    sheet.load("isNullObject");
    await context.sync();

    if (sheet.isNullObject) return { headers: [], sampleRow: [] };

    const range = sheet.getRange("1:2").getUsedRange();
    range.load("values");
    await context.sync();

    return {
      headers: range.values[0] || [],
      sampleRow: range.values[1] || []
    };
  });
}

/** * Execute Minotaur Data Import 
 */
async function executeMinotaurBatch(csvData) {
  const rows = parseSimpleCsv(csvData);
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getItemOrNullObject('Existing Minotaur Products');
    sheet.load("isNullObject");
    await context.sync();

    if (!sheet.isNullObject) { sheet.delete(); }
    sheet = sheets.add('Existing Minotaur Products');

    if (rows.length > 0) {
      const targetRange = sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length);
      targetRange.values = rows;
      sheet.getUsedRange().format.autofitColumns();
    }
    await context.sync();
  });
}

/** * Minotaur Export Process 
 */
async function processMinotaurExportToTab(mapping) {
  return await Excel.run(async (context) => {
    const sourceSheet = context.workbook.worksheets.getItem(TEMPLATE_SHEET_NAME);
    const usedRange = sourceSheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const allData = usedRange.values;
    if (allData.length < 2) return;

    const sourceHeaders = allData[0];
    const readinessIdx = sourceHeaders.indexOf('INTERNAL_Data_Readiness_Score');
    const inMinotaurIdx = sourceHeaders.indexOf('INTERNAL_In_Existing_Minotaur_Data');

    const filteredData = allData.slice(1).filter(row => {
      return parseFloat(row[readinessIdx]) === 100 && String(row[inMinotaurIdx]).toUpperCase() === 'FALSE';
    });

    if (filteredData.length === 0) return;

    const orderedMinotaurHeaders = [
      'Brand', 'CountryCode', 'Language', 'Barcode', 'ModelName', 'Title',
      'ATR_BrandDisplayName', 'EAN', 'UPC', 'GTIN', 'MPN', 'Images',
      'MarketingText', 'Category', 'ATR_ShortDescription', 'ATR_Color',
      'ATR_Size', 'ATR_PackSize', 'ATR_Scent', 'ATR_ParentBarcode',
      'WebpageURL', 'Discontinued'
    ];

    const sourceHeaderIndexMap = new Map(sourceHeaders.map((h, i) => [h, i]));
    const exportData = [orderedMinotaurHeaders];

    filteredData.forEach(sourceRow => {
      const newRow = orderedMinotaurHeaders.map(mh => {
        const mapInfo = mapping[mh];
        if (!mapInfo || mapInfo.type === 'skip') return "";
        if (mapInfo.type === 'static') return mapInfo.value;
        const sourceIndex = sourceHeaderIndexMap.get(mapInfo.value);
        return sourceIndex !== undefined ? sourceRow[sourceIndex] : "";
      });
      exportData.push(newRow);
    });

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const newSheet = context.workbook.worksheets.add(`Minotaur_Export_${timestamp}`);
    
    const range = newSheet.getRangeByIndexes(0, 0, exportData.length, orderedMinotaurHeaders.length);
    range.values = exportData;
    
    const textColumns = ['ean', 'upc', 'gtin', 'barcode', 'mpn', 'modelname'];
    orderedMinotaurHeaders.forEach((h, idx) => {
      if (textColumns.includes(h.toLowerCase())) {
        newSheet.getRangeByIndexes(1, idx, exportData.length - 1, 1).numberFormat = "@";
      }
    });

    newSheet.getUsedRange().format.autofitColumns();
    await context.sync();
  });
}

// ==========================================
// 4. UTILITIES
// ==========================================

/** * Simple CSV Parser fallback for mapping dialogs
 */
function parseSimpleCsv(str) {
  const lines = str.split(/\r?\n/).filter(l => l.trim() !== "");
  return lines.map(line => line.split(","));
}
