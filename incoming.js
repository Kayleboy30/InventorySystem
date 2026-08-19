/**
 * Incoming Operations & Supplier Directory Handlers
 */

function fetchSupplierDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const supplierSet = new Set();

    // 1. Read from Database sheet
    const dbSheet = ss.getSheetByName(SHEET_NAMES.DATABASE);
    if (dbSheet) {
      const data = dbSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0].map(h => String(h).trim().toLowerCase());
        let supplierColIdx = headers.findIndex(h => h.includes('supplier'));
        if (supplierColIdx === -1) supplierColIdx = 4; // Fallback to column E

        for (let i = 1; i < data.length; i++) {
          const val = String(data[i][supplierColIdx] || '').trim();
          if (val && val !== '-' && val.toLowerCase() !== 'n/a') {
            supplierSet.add(val);
          }
        }
      }
    }

    // 2. Also read from existing Incoming sheet entries
    const incSheet = ss.getSheetByName(SHEET_NAMES.INCOMING);
    if (incSheet) {
      const incData = incSheet.getDataRange().getValues();
      for (let i = 1; i < incData.length; i++) {
        const sup = String(incData[i][2] || '').trim();
        if (sup && sup !== '-' && sup.toLowerCase() !== 'n/a') {
          supplierSet.add(sup);
        }
      }
    }

    const supplierList = Array.from(supplierSet).sort();
    return { supplierList: supplierList };
  } catch (err) {
    return { supplierList: [], error: err.message };
  }
}

function fetchIncomingData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.INCOMING);

    if (!sheet) {
      return {
        totals: {},
        monthlyIncoming: {},
        transactions: [],
        pastRecords: []
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      const emptyTotals = {};
      MATERIAL_COLUMNS.forEach(mat => emptyTotals[mat.key] = 0);

      return {
        totals: emptyTotals,
        monthlyIncoming: {},
        transactions: [],
        pastRecords: []
      };
    }

    // ============================================================
    // READ HEADER ROW
    // ============================================================
    const headers = data[0].map(h =>
      String(h || '').trim().toLowerCase()
    );

    // ============================================================
    // FIND MATERIAL COLUMNS BY HEADER NAME
    // This prevents Desk Calendar from receiving another item's value
    // ============================================================
    const materialColumnMap = {};

    MATERIAL_COLUMNS.forEach(mat => {
      const key = mat.key.toLowerCase().trim();
      const name = mat.name.toLowerCase().trim();

      let columnIndex = headers.findIndex(h => h === key);

      // If key is not found, also try the full material name
      if (columnIndex === -1) {
        columnIndex = headers.findIndex(h => h === name);
      }

      // Store the actual spreadsheet column
      materialColumnMap[mat.key] = columnIndex;
    });

    // ============================================================
    // INITIALIZE TOTALS
    // ============================================================
    const totals = {};

    MATERIAL_COLUMNS.forEach(mat => {
      totals[mat.key] = 0;
    });

    const monthlyIncoming = {};
    const transactions = [];
    const pastRecords = [];

    // ============================================================
    // FIND GENERAL COLUMNS
    // ============================================================
    const idCol = headers.findIndex(h => h === 'id');

    const dateCol = headers.findIndex(h =>
      h === 'date'
    );

    const supplierCol = headers.findIndex(h =>
      h === 'supplier' ||
      h === 'supplier name' ||
      h === 'party'
    );

    const drCol = headers.findIndex(h =>
      h === 'dr #' ||
      h === 'dr no' ||
      h === 'dr number' ||
      h === 'dr' ||
      h.includes('delivery receipt')
    );

    // ============================================================
    // PROCESS EACH INCOMING RECORD
    // ============================================================
    for (let i = 1; i < data.length; i++) {

      const row = data[i];

      // ----------------------------------------------------------
      // DATE
      // ----------------------------------------------------------
      const rawDate =
        dateCol !== -1
          ? row[dateCol]
          : row[1];

      if (!rawDate) continue;

      const dateObj = new Date(rawDate);

      if (isNaN(dateObj.getTime())) continue;

      const formattedDate = Utilities.formatDate(
        dateObj,
        Session.getScriptTimeZone(),
        'M/d/yyyy'
      );

      const monthKey = Utilities.formatDate(
        dateObj,
        Session.getScriptTimeZone(),
        'MMM yy'
      );

      // ----------------------------------------------------------
      // CREATE MONTH ENTRY
      // ----------------------------------------------------------
      if (!monthlyIncoming[monthKey]) {
        monthlyIncoming[monthKey] = {};

        MATERIAL_COLUMNS.forEach(mat => {
          monthlyIncoming[monthKey][mat.key] = 0;
        });
      }

      // ----------------------------------------------------------
      // SUPPLIER
      // ----------------------------------------------------------
      const supplierName =
        supplierCol !== -1
          ? String(row[supplierCol] || '-').trim()
          : String(row[2] || '-').trim();

      // ----------------------------------------------------------
      // DR NUMBER
      // ----------------------------------------------------------
      let drNum = '-';

      if (drCol !== -1) {
        drNum = String(row[drCol] || '-').trim();
      } else {
        drNum = String(row[row.length - 1] || '-').trim();
      }

      // ----------------------------------------------------------
      // MATERIAL QUANTITIES
      // ----------------------------------------------------------
      const recordItems = {};

      MATERIAL_COLUMNS.forEach(mat => {

        const columnIndex = materialColumnMap[mat.key];

        let val = 0;

        // Only read the column if the material was actually found
        if (columnIndex !== -1) {
          val = Number(row[columnIndex]) || 0;
        }

        // Store the quantity under the CORRECT material key
        recordItems[mat.key] = val;

        // Add to total
        totals[mat.key] += val;

        // Add to monthly total
        monthlyIncoming[monthKey][mat.key] += val;
      });

      // ----------------------------------------------------------
      // TRANSACTION
      // ----------------------------------------------------------
      transactions.push({
        type: 'Incoming',
        date: formattedDate,
        rawDate: dateObj.getTime(),

        // Spreadsheet row = encoding order
        rowIndex: i + 1,

        party: supplierName,
        ref: drNum !== '-' ? drNum : 'N/A'
      });

      // ----------------------------------------------------------
      // PAST RECORD
      // ----------------------------------------------------------
      pastRecords.push({
        id: i + 1,

        recordId:
          idCol !== -1 && row[idCol]
            ? row[idCol]
            : ('INC-' + (i + 1)),

        date: formattedDate,

        supplier: supplierName,

        drNumber:
          drNum !== '-'
            ? drNum
            : '',

        items: recordItems
      });
    }

    // Newest records first
    pastRecords.reverse();

    // ============================================================
    // RETURN DATA
    // ============================================================
    return {
      totals: totals,
      monthlyIncoming: monthlyIncoming,
      transactions: transactions,
      pastRecords: pastRecords
    };

  } catch (err) {

    console.error(
      'fetchIncomingData error:',
      err
    );

    return {
      totals: {},
      monthlyIncoming: {},
      transactions: [],
      pastRecords: [],
      error: err.message
    };
  }
}

function recordIncoming(formData) {
  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.INCOMING);

    if (!sheet) {
      throw new Error(
        'Sheet ' + SHEET_NAMES.INCOMING + ' not found.'
      );
    }

    // ============================================================
    // READ SHEET HEADERS
    // ============================================================
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h || '').trim());

    const newRow = new Array(headers.length).fill('');

    // ============================================================
    // CREATE ID
    // ============================================================
    const newId =
      'INC-' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyyMMdd-HHmmss'
      );

    // ============================================================
    // WRITE DATA BASED ON HEADER NAME
    // ============================================================
    headers.forEach((header, index) => {

      const h = header.toLowerCase().trim();

      // ID
      if (h === 'id') {

        newRow[index] = newId;

      }

      // DATE
      else if (h === 'date') {

        newRow[index] =
          formData.date
            ? new Date(formData.date)
            : new Date();

      }

      // SUPPLIER
      else if (
        h === 'supplier' ||
        h === 'supplier name' ||
        h === 'party'
      ) {

        newRow[index] =
          formData.party || '';

      }

      // DR NUMBER
      else if (
        h === 'dr #' ||
        h === 'dr no' ||
        h === 'dr number' ||
        h === 'dr' ||
        h.includes('delivery receipt')
      ) {

        newRow[index] =
          formData.drNumber || '';

      }

      // ==========================================================
      // MATERIAL
      // ==========================================================
      else {

        const material = MATERIAL_COLUMNS.find(mat => {

          return (
            mat.key.toLowerCase().trim() === h ||
            mat.name.toLowerCase().trim() === h
          );

        });

        if (material) {

          const qty =
            formData.items &&
            formData.items[material.key]
              ? Number(formData.items[material.key])
              : 0;

          newRow[index] = qty;
        }
      }
    });

    // ============================================================
    // SAVE ROW
    // ============================================================
    sheet.appendRow(newRow);

    return {
      success: true,
      id: newId
    };

  } catch (err) {

    console.error(
      'recordIncoming error:',
      err
    );

    return {
      success: false,
      error: err.message
    };
  }
}

function updateIncoming(rowIndex, formData) {
  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.INCOMING);

    if (!sheet) {
      throw new Error(
        'Sheet ' + SHEET_NAMES.INCOMING + ' not found.'
      );
    }

    // ============================================================
    // READ HEADERS
    // ============================================================
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h || '').trim());

    // ============================================================
    // GET EXISTING ROW
    // ============================================================
    const existingRow = sheet
      .getRange(
        rowIndex,
        1,
        1,
        headers.length
      )
      .getValues()[0];

    const updatedRow = existingRow.slice();

    // ============================================================
    // UPDATE BASED ON HEADER
    // ============================================================
    headers.forEach((header, index) => {

      const h = header.toLowerCase().trim();

      // ----------------------------------------------------------
      // ID
      // ----------------------------------------------------------
      if (h === 'id') {
        // Keep existing ID
      }

      // ----------------------------------------------------------
      // DATE
      // ----------------------------------------------------------
      else if (h === 'date') {

        updatedRow[index] =
          formData.date
            ? new Date(formData.date)
            : new Date();

      }

      // ----------------------------------------------------------
      // SUPPLIER
      // ----------------------------------------------------------
      else if (
        h === 'supplier' ||
        h === 'supplier name' ||
        h === 'party'
      ) {

        updatedRow[index] =
          formData.party || '';

      }

      // ----------------------------------------------------------
      // DR NUMBER
      // ----------------------------------------------------------
      else if (
        h === 'dr #' ||
        h === 'dr no' ||
        h === 'dr number' ||
        h === 'dr' ||
        h.includes('delivery receipt')
      ) {

        updatedRow[index] =
          formData.drNumber || '';

      }

      // ----------------------------------------------------------
      // MATERIAL
      // ----------------------------------------------------------
      else {

        const material = MATERIAL_COLUMNS.find(mat => {

          return (
            mat.key.toLowerCase().trim() === h ||
            mat.name.toLowerCase().trim() === h
          );

        });

        if (material) {

          const qty =
            formData.items &&
            formData.items[material.key]
              ? Number(formData.items[material.key])
              : 0;

          updatedRow[index] = qty;
        }
      }
    });

    // ============================================================
    // SAVE
    // ============================================================
    sheet
      .getRange(
        rowIndex,
        1,
        1,
        headers.length
      )
      .setValues([updatedRow]);

    return {
      success: true
    };

  } catch (err) {

    console.error(
      'updateIncoming error:',
      err
    );

    return {
      success: false,
      error: err.message
    };
  }
}
