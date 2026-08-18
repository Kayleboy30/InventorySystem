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
    if (!sheet) return { totals: {}, monthlyIncoming: {}, transactions: [], pastRecords: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { totals: {}, monthlyIncoming: {}, transactions: [], pastRecords: [] };

    const totals = {};
    MATERIAL_COLUMNS.forEach(m => totals[m.key] = 0);

    const monthlyIncoming = {};
    const transactions = [];
    const pastRecords = [];

    // Col 0: ID, Col 1: Date, Col 2: Supplier, Cols 3-15: 13 Materials, Col 16: DR #
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawDate = row[1];
      if (!rawDate) continue;

      const dateObj = new Date(rawDate);
      const formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const monthKey = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'MMM yy');

      if (!monthlyIncoming[monthKey]) {
        monthlyIncoming[monthKey] = {};
        MATERIAL_COLUMNS.forEach(m => monthlyIncoming[monthKey][m.key] = 0);
      }

      const supplierName = String(row[2] || '-').trim();
      const drNum = String(row[16] || row[row.length - 1] || '-').trim();

      const recordItems = {};

      MATERIAL_COLUMNS.forEach((mat, idx) => {
        const val = Number(row[3 + idx]) || 0;
        totals[mat.key] += val;
        monthlyIncoming[monthKey][mat.key] += val;
        recordItems[mat.key] = val;
      });

      transactions.push({
        type: 'Incoming',
        date: formattedDate,
        rawDate: dateObj.getTime(),
        party: supplierName,
        ref: drNum !== '-' ? drNum : 'N/A'
      });

      pastRecords.push({
        id: i + 1,
        recordId: row[0] || ('INC-' + (i + 1)),
        date: formattedDate,
        supplier: supplierName,
        drNumber: drNum !== '-' ? drNum : '',
        items: recordItems
      });
    }

    pastRecords.reverse();

    return {
      totals: totals,
      monthlyIncoming: monthlyIncoming,
      transactions: transactions,
      pastRecords: pastRecords
    };
  } catch (err) {
    return { totals: {}, monthlyIncoming: {}, transactions: [], pastRecords: [], error: err.message };
  }
}

function recordIncoming(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.INCOMING);
    if (!sheet) throw new Error('Sheet ' + SHEET_NAMES.INCOMING + ' not found.');

    const newId = 'INC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    const dateVal = formData.date ? new Date(formData.date) : new Date();

    const rowData = [
      newId,
      dateVal,
      formData.party || ''
    ];

    MATERIAL_COLUMNS.forEach(mat => {
      const qty = (formData.items && formData.items[mat.key]) ? Number(formData.items[mat.key]) : 0;
      rowData.push(qty);
    });

    rowData.push(formData.drNumber || '');

    sheet.appendRow(rowData);
    return { success: true, id: newId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateIncoming(rowIndex, formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.INCOMING);
    if (!sheet) throw new Error('Sheet ' + SHEET_NAMES.INCOMING + ' not found.');

    const existingId = sheet.getRange(rowIndex, 1).getValue() || ('INC-' + rowIndex);
    const dateVal = formData.date ? new Date(formData.date) : new Date();

    const rowData = [
      existingId,
      dateVal,
      formData.party || ''
    ];

    MATERIAL_COLUMNS.forEach(mat => {
      const qty = (formData.items && formData.items[mat.key]) ? Number(formData.items[mat.key]) : 0;
      rowData.push(qty);
    });

    rowData.push(formData.drNumber || '');

    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, id: existingId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}