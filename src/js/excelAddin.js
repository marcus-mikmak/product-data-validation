// --- GLOBAL CONSTANTS ---
const TEMPLATE_SHEET_NAME = 'Product Data';
const INTERNAL_PREFIX = 'INTERNAL_';
const LOCALE_PREFIX = 'locale::';

/**
 * 1. IMAGE CHECKER
 * Translates checkImageVisibility() from UrlFetchApp to native browser fetch().
 */
async function checkImageVisibility(url) {
  if (!url || typeof url !== 'string' || !url.startsWith('http')) return false;

  const imageExtensions = ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.avif'];
  const hasImageExtension = imageExtensions.some(ext => url.toLowerCase().includes(ext));

  try {
    // Send a Range header request to read only the first byte (saves bandwidth)
    const response = await fetch(url, {
      method: 'GET',
      headers: { 'Range': 'bytes=0-1' }
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

/**
 * 2. BATCH DATA IMPORT & MAPPING ENGINE
 * Replaces prepareImport() and executeImportBatch().
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
        if (!mappingValue || mappingValue === "-- Skip this column --") return "";
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
      // Clear rows starting at index 1 (Row 2) down to the bottom
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

/**
 * 3. DATA EXPORTER
 * Translates processAndExportToTab().
 * Filters data based on readiness score, creates a new sheet, and applies text formatting.
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

    // Filter rows based on score
    const filteredRows = rows.filter(row => {
      const numericScore = parseFloat(row[readinessScoreIndex]);
      return (type === 'good') ? numericScore === 100 : numericScore !== 100;
    });

    if (filteredRows.length === 0) {
      console.log("No products match the readiness criteria for export.");
      return;
    }

    const dataRowsToKeep = filteredRows.map(row => colIndicesToKeep.map(idx => row[idx]));

    // Generate New Sheet Name
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const newSheetName = `${type === 'good' ? 'Good_Data' : 'Review_Data'}_${timestamp}`;
    
    // Add new sheet to workbook
    const newSheet = sheets.add(newSheetName);

    // Write Headers
    const headerRange = newSheet.getRangeByIndexes(0, 0, 1, headersToKeep.length);
    headerRange.values = [headersToKeep];

    // Apply Text Number Format ('@') to explicit columns (EAN, UPC, Barcode, etc.)
    const textKeywords = ['id', 'ean', 'upc', 'gtin', 'barcode', 'mpn', 'modelname'];
    headersToKeep.forEach((header, index) => {
      if (textKeywords.some(keyword => header.toLowerCase().includes(keyword))) {
        // Target the full data column height for this header
        const colRange = newSheet.getRangeByIndexes(1, index, dataRowsToKeep.length, 1);
        colRange.numberFormat = "@"; // Force Excel Text Format
      }
    });

    // Write Filtered Data Rows
    const dataRange = newSheet.getRangeByIndexes(1, 0, dataRowsToKeep.length, headersToKeep.length);
    dataRange.values = dataRowsToKeep;

    // Format & Activate
    newSheet.getUsedRange().format.autofitColumns();
    newSheet.activate();

    await context.sync();
  });
}
