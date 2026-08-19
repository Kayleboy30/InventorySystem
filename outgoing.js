/**
 * ==========================================
 * OUTGOING MODULE (Dispatches / Gate Pass)
 * ==========================================
 */

/**
 * Reads Cluster, Cluster Head, Contact, Base Station, and Complete Address from 'Database' sheet
 */
function fetchRegionDirectory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Database') || 
              ss.getSheetByName('database') || 
              ss.getSheetByName('DB') ||
              ss.getSheetByName('Address / Region');

  const directory = {};
  const regionList = [];

  if (!sheet || sheet.getLastRow() <= 1) {
    return { directory, regionList };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());

  let clusterColIdx = -1;
  let headColIdx = -1;
  let contactColIdx = -1;
  let baseColIdx = -1;
  let addressColIdx = -1;

  headers.forEach((h, idx) => {
    if (h.includes('complete address') || h === 'address' || h.includes('delivery address')) {
      addressColIdx = idx;
    } else if (h.includes('contact') || h.includes('phone') || h.includes('mobile')) {
      contactColIdx = idx;
    } else if (h.includes('cluster head') || h.includes('head name') || h === 'head') {
      headColIdx = idx;
    } else if (h.includes('base station') || h.includes('station') || h.includes('base')) {
      baseColIdx = idx;
    } else if (h.includes('cluster') || h.includes('region')) {
      if (clusterColIdx === -1) clusterColIdx = idx;
    }
  });

  // Fallback defaults
  if (clusterColIdx === -1) clusterColIdx = 19;
  if (headColIdx === -1) headColIdx = 20;
  if (contactColIdx === -1) contactColIdx = 21;
  if (baseColIdx === -1) baseColIdx = 22;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const rawCluster     = (row[clusterColIdx] !== undefined) ? String(row[clusterColIdx]).trim() : '';
    const clusterHead    = (row[headColIdx] !== undefined) ? String(row[headColIdx]).trim() : '';
    const clusterContact = (row[contactColIdx] !== undefined) ? String(row[contactColIdx]).trim() : '';
    const baseStation    = (row[baseColIdx] !== undefined) ? String(row[baseColIdx]).trim() : '';
    const address        = (addressColIdx !== -1 && row[addressColIdx] !== undefined) ? String(row[addressColIdx]).trim() : '';

    if (rawCluster) {
      const dataObj = {
        displayName: rawCluster,
        clusterHead: clusterHead,
        clusterHeadContact: clusterContact,
        baseStation: baseStation,
        completeAddress: address
      };

      directory[rawCluster] = dataObj;
      directory[rawCluster.toLowerCase()] = dataObj;
      directory[rawCluster.toLowerCase().replace(/\s+/g, '')] = dataObj;

      const numMatch = rawCluster.match(/\d+/);
      if (numMatch) {
        directory[numMatch[0]] = dataObj;
      }

      if (!regionList.includes(rawCluster)) {
        regionList.push(rawCluster);
      }
    }
  }

  return { directory, regionList: regionList.sort() };
}

/**
 * Reads AVP, Division, Operation, Destination, and Complete Address from 'Database' sheet
 */
function fetchAvpDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Database') || 
              ss.getSheetByName('database') || 
              ss.getSheetByName('DB');

  const avpDirectory = {};
  const avpList = [];

  if (!sheet || sheet.getLastRow() <= 1) {
    return { avpDirectory, avpList };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());

  const avpIdx     = headers.findIndex(h => h.includes('avp'));
  const divIdx     = headers.findIndex(h => h.includes('division'));
  const opIdx      = headers.findIndex(h => h.includes('operation'));
  const destIdx    = headers.findIndex(h => h.includes('destination'));
  const addressIdx = headers.findIndex(h => h.includes('complete address') || h === 'address');

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const avpName     = (avpIdx !== -1 && row[avpIdx] !== undefined) ? String(row[avpIdx]).trim() : '';
    const division    = (divIdx !== -1 && row[divIdx] !== undefined) ? String(row[divIdx]).trim() : '';
    const operation   = (opIdx !== -1 && row[opIdx] !== undefined) ? String(row[opIdx]).trim() : '';
    const destination = (destIdx !== -1 && row[destIdx] !== undefined) ? String(row[destIdx]).trim() : '';
    const address     = (addressIdx !== -1 && row[addressIdx] !== undefined) ? String(row[addressIdx]).trim() : '';

    if (avpName) {
      const info = {
        division: division,
        operation: operation,
        destination: destination,
        completeAddress: address
      };

      if (!avpDirectory[avpName]) {
        avpDirectory[avpName] = info;
        avpDirectory[avpName.toLowerCase()] = info;
        avpList.push(avpName);
      }
    }
  }

  return { avpDirectory, avpList: avpList.sort() };
}

function fetchOutgoingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outSheet = ss.getSheetByName(SHEET_NAMES.OUTGOING);

  const totals = {};
  const monthlyOutgoing = {};
  const transactions = [];
  const pastRecords = [];

  MATERIAL_COLUMNS.forEach(mat => {
    totals[mat.key] = 0;
  });

  if (!outSheet || outSheet.getLastRow() <= 1) {
    return {
      totals,
      monthlyOutgoing,
      transactions,
      pastRecords
    };
  }

  const values = outSheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const dateIdx = headers.findIndex(
    h => h.toLowerCase() === 'date'
  );

  const avpIdx = headers.findIndex(
    h =>
      h.toLowerCase() === 'avp name' ||
      h.toLowerCase() === 'avp'
  );

  const divIdx = headers.findIndex(
    h => h.toLowerCase() === 'division'
  );

  const destIdx = headers.findIndex(
    h => h.toLowerCase() === 'destination'
  );

  const controlIdx = headers.findIndex(
    h => h.toLowerCase().includes('control')
  );

  const regIdx = headers.findIndex(
    h =>
      h.toLowerCase().includes('region') ||
      (
        h.toLowerCase().includes('cluster') &&
        !h.toLowerCase().includes('head')
      )
  );

  const opIdx = headers.findIndex(
    h => h.toLowerCase() === 'operation'
  );

  const headIdx = headers.findIndex(
    h => h.toLowerCase() === 'cluster head'
  );

  const contactIdx = headers.findIndex(
    h =>
      h.toLowerCase().includes('head contact') ||
      h.toLowerCase().includes('contact')
  );

  const baseIdx = headers.findIndex(
    h =>
      h.toLowerCase() === 'base station' ||
      h.toLowerCase().includes('base')
  );

  const notesIdx = headers.findIndex(
    h => h.toLowerCase() === 'notes'
  );


  // ============================================================
  // DATE PARSER
  // Handles Google Sheets Date objects and date strings
  // ============================================================
  function parseTransactionDate(value) {

    if (!value) {
      return null;
    }

    // Google Sheets Date object
    if (Object.prototype.toString.call(value) === '[object Date]') {
      if (!isNaN(value.getTime())) {
        return value;
      }
    }

    // String date
    const str = String(value).trim();

    if (!str) {
      return null;
    }

    // MM/DD/YYYY
    const match = str.match(
      /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/
    );

    if (match) {
      const month = Number(match[1]);
      const day = Number(match[2]);
      const year = Number(match[3]);

      const date = new Date(
        year,
        month - 1,
        day
      );

      if (!isNaN(date.getTime())) {
        return date;
      }
    }

    // Fallback
    const parsed = new Date(str);

    if (!isNaN(parsed.getTime())) {
      return parsed;
    }

    return null;
  }


  // ============================================================
  // READ ALL OUTGOING RECORDS
  // ============================================================

  for (let r = 1; r < values.length; r++) {

    const row = values[r];

    const dateVal =
      dateIdx !== -1
        ? row[dateIdx]
        : null;

    const transactionDate =
      parseTransactionDate(dateVal);

    const rawDate =
      transactionDate
        ? transactionDate.getTime()
        : 0;


    const monthStr =
      transactionDate
        ? Utilities.formatDate(
            transactionDate,
            Session.getScriptTimeZone(),
            'MMM yy'
          )
        : '';


    const avpVal =
      avpIdx !== -1 && row[avpIdx]
        ? String(row[avpIdx]).trim()
        : '-';

    const divVal =
      divIdx !== -1 && row[divIdx]
        ? String(row[divIdx]).trim()
        : '-';

    const destVal =
      destIdx !== -1 && row[destIdx]
        ? String(row[destIdx]).trim()
        : '-';

    const controlVal =
      controlIdx !== -1 && row[controlIdx]
        ? String(row[controlIdx]).trim()
        : '-';

    const regVal =
      regIdx !== -1 && row[regIdx]
        ? String(row[regIdx]).trim()
        : '-';

    const opVal =
      opIdx !== -1 && row[opIdx]
        ? String(row[opIdx]).trim()
        : '';

    const headVal =
      headIdx !== -1 && row[headIdx]
        ? String(row[headIdx]).trim()
        : '';

    const contactVal =
      contactIdx !== -1 && row[contactIdx]
        ? String(row[contactIdx]).trim()
        : '';

    const baseVal =
      baseIdx !== -1 && row[baseIdx]
        ? String(row[baseIdx]).trim()
        : '';

    const notesVal =
      notesIdx !== -1 && row[notesIdx]
        ? String(row[notesIdx]).trim()
        : '';


    // ============================================================
    // MATERIAL QUANTITIES
    // ============================================================

    const itemQuantities = {};
    let rowHasData = false;

    MATERIAL_COLUMNS.forEach(mat => {

      const colIdx =
        headers.findIndex(
          h =>
            h.toLowerCase() ===
            mat.key.toLowerCase()
        );

      if (colIdx !== -1) {

        const qty =
          parseFloat(row[colIdx]) || 0;

        if (qty > 0) {

          rowHasData = true;

          itemQuantities[mat.key] = qty;

          totals[mat.key] += qty;


          if (monthStr) {

            if (!monthlyOutgoing[monthStr]) {
              monthlyOutgoing[monthStr] = {};
            }

            if (!monthlyOutgoing[monthStr][mat.key]) {
              monthlyOutgoing[monthStr][mat.key] = 0;
            }

            monthlyOutgoing[monthStr][mat.key] += qty;
          }
        }
      }
    });


    // ============================================================
    // ADD TRANSACTION
    // ============================================================

    if (rowHasData) {

      const dateString =
        transactionDate
          ? Utilities.formatDate(
              transactionDate,
              Session.getScriptTimeZone(),
              'M/d/yyyy'
            )
          : 'N/A';


      transactions.push({
          type: 'Outgoing',
          date: dateString,
          party: `${destVal} (${avpVal})`.trim(),
          ref: controlVal,

          // Date shown in dashboard
          rawDate: rawDate,

          // Actual encoding order
          rowIndex: r + 1
      });


      pastRecords.push({

        id: r,

        date: dateString,

        controlNo: controlVal,

        avpName: avpVal,

        division: divVal,

        destination: destVal,

        cluster: regVal,

        operation: opVal,

        clusterHead: headVal,

        clusterHeadContact: contactVal,

        baseStation: baseVal,

        notes: notesVal,

        items: itemQuantities,

        rawDate: rawDate,

        rowIndex: r + 1
      });
    }
  }


  // ============================================================
  // SORT PAST OUTGOING RECORDS
  // LATEST DATE FIRST
  // SAME DATE = LATEST ENCODED ROW FIRST
  // ============================================================

  pastRecords.sort((a, b) => {

    if (b.rawDate !== a.rawDate) {
      return b.rawDate - a.rawDate;
    }

    return b.rowIndex - a.rowIndex;
  });


  // ============================================================
  // SORT TRANSACTIONS
  // LATEST DATE FIRST
  // SAME DATE = LATEST ENCODED ROW FIRST
  // ============================================================

  transactions.sort((a, b) => {

    if (b.rawDate !== a.rawDate) {
      return b.rawDate - a.rawDate;
    }

    return b.rowIndex - a.rowIndex;
  });


  return {
    totals,
    monthlyOutgoing,
    transactions,
    pastRecords
  };
}



/**
 * Appends a NEW outgoing record to the 'Outgoing' sheet
 */
function recordOutgoing(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.OUTGOING);

  if (!sheet) throw new Error(`Sheet '${SHEET_NAMES.OUTGOING}' not found.`);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const newRow = new Array(headers.length).fill('');
  const nextId = sheet.getLastRow();

  headers.forEach((header, idx) => {
    const hLower = header.toLowerCase();
    if      (hLower === 'id')            newRow[idx] = nextId;
    else if (hLower === 'date')          newRow[idx] = formData.date;
    else if (hLower === 'avp name')      newRow[idx] = formData.avpName || '';
    else if (hLower === 'division')      newRow[idx] = formData.division || '';
    else if (hLower.includes('region') || (hLower.includes('cluster') && !hLower.includes('head')))
                                         newRow[idx] = formData.regionCluster || '';
    else if (hLower === 'destination')   newRow[idx] = formData.destination || '';
    else if (hLower.includes('control')) newRow[idx] = formData.controlNo || '';
    else if (hLower === 'operation')     newRow[idx] = formData.operation || '';
    else if (hLower.includes('head contact') || hLower.includes('contact'))
                                         newRow[idx] = formData.clusterHeadContact || '';
    else if (hLower.includes('cluster head') || hLower === 'head')
                                         newRow[idx] = formData.clusterHead || '';
    else if (hLower === 'base station' || hLower.includes('base') || hLower.includes('station'))
                                         newRow[idx] = formData.baseStation || '';
    else if (hLower === 'notes')         newRow[idx] = formData.notes || '';
    else {
      const mat = MATERIAL_COLUMNS.find(m => m.key.toLowerCase() === hLower);
      if (mat && formData.items && formData.items[mat.key]) {
        newRow[idx] = parseFloat(formData.items[mat.key]) || 0;
      }
    }
  });

  sheet.appendRow(newRow);
  return { success: true };
}

/**
 * UPDATES an EXISTING outgoing record in the 'Outgoing' sheet by row index
 */
function updateOutgoing(rowIndex, formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.OUTGOING);

  if (!sheet) throw new Error(`Sheet '${SHEET_NAMES.OUTGOING}' not found.`);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const existingRow = sheet.getRange(rowIndex + 1, 1, 1, headers.length).getValues()[0];
  const updatedRow = existingRow.slice();

  headers.forEach((header, idx) => {
    const hLower = header.toLowerCase();
    if      (hLower === 'id')            { /* Keep ID unchanged */ }
    else if (hLower === 'date')          updatedRow[idx] = formData.date;
    else if (hLower === 'avp name')      updatedRow[idx] = formData.avpName || '';
    else if (hLower === 'division')      updatedRow[idx] = formData.division || '';
    else if (hLower.includes('region') || (hLower.includes('cluster') && !hLower.includes('head')))
                                         updatedRow[idx] = formData.regionCluster || '';
    else if (hLower === 'destination')   updatedRow[idx] = formData.destination || '';
    else if (hLower.includes('control')) updatedRow[idx] = formData.controlNo || '';
    else if (hLower === 'operation')     updatedRow[idx] = formData.operation || '';
    else if (hLower.includes('head contact') || hLower.includes('contact'))
                                         updatedRow[idx] = formData.clusterHeadContact || '';
    else if (hLower.includes('cluster head') || hLower === 'head')
                                         updatedRow[idx] = formData.clusterHead || '';
    else if (hLower === 'base station' || hLower.includes('base') || hLower.includes('station'))
                                         updatedRow[idx] = formData.baseStation || '';
    else if (hLower === 'notes')         updatedRow[idx] = formData.notes || '';
    else {
      const mat = MATERIAL_COLUMNS.find(m => m.key.toLowerCase() === hLower);
      if (mat) {
        updatedRow[idx] = (formData.items && formData.items[mat.key]) ? parseFloat(formData.items[mat.key]) || 0 : 0;
      }
    }
  });

  sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
  return { success: true };
}
