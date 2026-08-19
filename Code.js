/**
 * Main Controller & Multi-View Web App UI
 */

function doGet() {
  return HtmlService.createHtmlOutput(getHtmlContent())
    .setTitle('Logistics Inventory Dashboard & Delivery Receipts - ASA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  const incoming       = fetchIncomingData();
  const outgoing       = fetchOutgoingData();
  const regionData     = fetchRegionDirectory();
  const avpData        = fetchAvpDatabase();
  const supplierData   = fetchSupplierDatabase();
  const divisionBudget = fetchDivisionBudgetData();

  const monthlyData = {};
  const allMonths = new Set([...Object.keys(incoming.monthlyIncoming), ...Object.keys(outgoing.monthlyOutgoing)]);

  allMonths.forEach(m => {
    monthlyData[m] = {};
    MATERIAL_COLUMNS.forEach(mat => {
      const inVal  = (incoming.monthlyIncoming[m] && incoming.monthlyIncoming[m][mat.key]) || 0;
      const outVal = (outgoing.monthlyOutgoing[m] && outgoing.monthlyOutgoing[m][mat.key]) || 0;
      monthlyData[m][mat.key] = { in: inVal, out: outVal };
    });
  });

  const inventoryList = MATERIAL_COLUMNS.map(mat => {
    const inQty  = incoming.totals[mat.key] || 0;
    const outQty = outgoing.totals[mat.key] || 0;
    const stock  = inQty - outQty;
    return { key: mat.key, name: mat.name, inQty, outQty, stock, isLow: stock <= 2000 };
  });

  // Combine Incoming + Outgoing transactions
    const allTransactions = [
    ...incoming.transactions,
    ...outgoing.transactions
  ];

  // Sort by:
  // 1. Latest date first
  // 2. If same date, latest encoded row first
  allTransactions.sort((a, b) => {
    return (b.rowIndex || 0) - (a.rowIndex || 0);
  });

  return {
    materials: MATERIAL_COLUMNS,
    inventoryList,
    monthlyData,
    recentTransactions: allTransactions.slice(0, 10),
    pastIncomingRecords: incoming.pastRecords || [],
    pastOutgoingRecords: outgoing.pastRecords || [],
    regionDirectory: regionData.directory,
    avpDirectory: avpData.avpDirectory,
    avpList: avpData.avpList,
    supplierList: supplierData.supplierList || [],
    divisionBudgetData: divisionBudget
  };
}

function recordTransaction(type, formData) {
  return type === 'INCOMING' ? recordIncoming(formData) : recordOutgoing(formData);
}

function updateTransaction(type, rowIndex, formData) {
  return type === 'INCOMING' ? updateIncoming(rowIndex, formData) : updateOutgoing(rowIndex, formData);
}

// ─── DIVISION BUDGET BACKEND HANDLERS ───────────────────────────────────────

function fetchDivisionBudgetData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Division Budget');
    if (!sheet) {
      sheet = ss.insertSheet('Division Budget');
      sheet.appendRow([
        'Row ID', 'AVP NAME', 'Division',
        'GTR Op Request', 'GTR 5%', 'GTR Delivered Old', 'GTR Delivered New', 'GTR Balance',
        'FAF Op Request', 'FAF 5%', 'FAF With Less 5%', 'FAF Delivered', 'FAF Balance',
        'PB Op Request', 'PB 5%', 'PB With Less 5%', 'PB Delivered Old', 'PB Delivered New', 'PB Balance',
        'Notes'
      ]);
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { rows: [], divisions: [] };

    const rows = [];
    const divisionsSet = new Set();
    const startRow = (String(data[0][1]).toUpperCase().includes('AVP') || String(data[0][0]).toUpperCase().includes('AVP')) ? 1 : 2;

    for (let i = startRow; i < data.length; i++) {
      const r = data[i];
      const avp = String(r[1] || r[0] || '').trim();
      const div = String(r[2] || r[1] || '').trim();
      if (!avp && !div) continue;

      if (div) divisionsSet.add(div);

      const gtrOp     = Number(r[3]) || 0;
      const gtr5      = Number(r[4]) || (gtrOp * 0.05);
      const gtrOld    = Number(r[5]) || 0;
      const gtrNew    = Number(r[6]) || 0;
      const gtrBal    = Number(r[7]) || (gtrOp - (gtrOld + gtrNew));

      const fafOp     = Number(r[8]) || 0;
      const faf5      = Number(r[9]) || (fafOp * 0.05);
      const fafLess5  = Number(r[10]) || (fafOp - faf5);
      const fafDel    = Number(r[11]) || 0;
      const fafBal    = Number(r[12]) || (fafLess5 - fafDel);

      const pbOp      = Number(r[13]) || 0;
      const pb5       = Number(r[14]) || (pbOp * 0.05);
      const pbLess5   = Number(r[15]) || (pbOp - pb5);
      const pbOld     = Number(r[16]) || 0;
      const pbNew     = Number(r[17]) || 0;
      const pbBal     = Number(r[18]) || (pbLess5 - (pbOld + pbNew));

      const notes     = String(r[19] || '');

      rows.push({
        rowIndex: i + 1,
        avpName: avp,
        division: div,
        gtr: { opRequest: gtrOp, fivePercent: gtr5, deliveredOld: gtrOld, deliveredNew: gtrNew, balance: gtrBal },
        faf: { opRequest: fafOp, fivePercent: faf5, withLess5: fafLess5, delivered: fafDel, balance: fafBal },
        pb:  { opRequest: pbOp, fivePercent: pb5, withLess5: pbLess5, deliveredOld: pbOld, deliveredNew: pbNew, balance: pbBal },
        notes: notes
      });
    }

    return {
      rows: rows,
      divisions: Array.from(divisionsSet)
    };
  } catch (err) {
    return { rows: [], divisions: [], error: err.message };
  }
}

function saveDivisionBudgetRow(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Division Budget');
    if (!sheet) sheet = ss.insertSheet('Division Budget');

    const gtrOp    = Number(formData.gtrOpRequest) || 0;
    const gtr5     = gtrOp * 0.05;
    const gtrOld   = Number(formData.gtrDeliveredOld) || 0;
    const gtrNew   = Number(formData.gtrDeliveredNew) || 0;
    const gtrBal   = gtrOp - (gtrOld + gtrNew);

    const fafOp    = Number(formData.fafOpRequest) || 0;
    const faf5     = fafOp * 0.05;
    const fafLess5 = fafOp - faf5;
    const fafDel   = Number(formData.fafDelivered) || 0;
    const fafBal   = fafLess5 - fafDel;

    const pbOp     = Number(formData.pbOpRequest) || 0;
    const pb5      = pbOp * 0.05;
    const pbLess5  = pbOp - pb5;
    const pbOld    = Number(formData.pbDeliveredOld) || 0;
    const pbNew    = Number(formData.pbDeliveredNew) || 0;
    const pbBal    = pbLess5 - (pbOld + pbNew);

    const rowData = [
      formData.rowIndex || '',
      formData.avpName || '',
      formData.division || '',
      gtrOp, gtr5, gtrOld, gtrNew, gtrBal,
      fafOp, faf5, fafLess5, fafDel, fafBal,
      pbOp, pb5, pbLess5, pbOld, pbNew, pbBal,
      formData.notes || ''
    ];

    if (formData.rowIndex && Number(formData.rowIndex) > 1) {
      sheet.getRange(Number(formData.rowIndex), 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ─── CLIENT SPA INTERFACE ──────────────────────────────────────────────────

function getHtmlContent() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>ASA Logistics Inventory, Receipts & Division Budget</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    :root { --asa-orange: #d96800; --asa-sidebar-active: #b55300; }
    .bg-asa-orange { background-color: var(--asa-orange); }
    .bg-asa-active { background-color: var(--asa-sidebar-active); }
    body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f8fafc; }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #f1f5f9; }
    ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }

    /* GATE PASS STYLES */
    #printableGatePassArea { font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 8pt; line-height: 1.15; }
    .gp-table, .gp-table th, .gp-table td { border: 1.2px solid #000 !important; border-collapse: collapse; box-sizing: border-box; }
    .gp-lbl { background-color: #dce4ec !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; font-weight: bold; font-size: 8pt; color: #000 !important; padding: 2px 4px !important; }
    .gp-val { font-size: 8pt; padding: 2px 4px !important; color: #000 !important; }
    .gp-input { width: 100%; outline: none; background: transparent; font-size: 8pt; font-family: Calibri, 'Segoe UI', Arial, sans-serif; color: #000 !important; padding: 0 !important; margin: 0 !important; border: none !important; }

    /* DIVISION BUDGET STYLES */
    .db-header-gtr { background-color: #7e22ce; color: #fff; }
    .db-sub-gtr    { background-color: #f3e8ff; color: #581c87; font-weight: 600; }
    .db-header-faf { background-color: #0284c7; color: #fff; }
    .db-sub-faf    { background-color: #e0f2fe; color: #0369a1; font-weight: 600; }
    .db-header-pb  { background-color: #d97706; color: #fff; }
    .db-sub-pb     { background-color: #fef3c7; color: #92400e; font-weight: 600; }

    /* A4 PRINT SETTINGS */
    @page { size: A4 portrait; margin: 0mm !important; }
    @media print {
      html, body { width: 210mm !important; height: 297mm !important; max-height: 297mm !important; margin: 0 !important; padding: 0 !important; background: #fff !important; overflow: hidden !important; }
      body * { visibility: hidden !important; }
      #printableGatePassArea, #printableGatePassArea * { visibility: visible !important; }
      #printableGatePassArea { position: absolute !important; left: 0 !important; top: 0 !important; width: 210mm !important; height: 297mm !important; max-height: 297mm !important; box-sizing: border-box !important; padding: 5mm 8mm !important; margin: 0 !important; border: none !important; box-shadow: none !important; background: #fff !important; display: flex !important; flex-direction: column !important; justify-content: space-between !important; overflow: hidden !important; }
      .gp-copy-block { height: 138mm !important; max-height: 138mm !important; display: flex !important; flex-direction: column !important; justify-content: space-between !important; box-sizing: border-box !important; overflow: hidden !important; }
      .gp-cut-line { height: 6mm !important; max-height: 6mm !important; display: flex !important; align-items: center !important; justify-content: center !important; margin: 0 !important; }
      .no-print { display: none !important; }
      .gp-input::placeholder { color: transparent !important; }
    }
  </style>
</head>
<body class="flex h-screen overflow-hidden text-slate-800">

  <!-- SIDEBAR -->
  <aside class="w-64 bg-asa-orange text-white flex flex-col flex-shrink-0 shadow-xl justify-between select-none no-print">
    <div>
      <div class="p-5 flex flex-col items-center border-b border-orange-400/30">
        <img src="https://lh3.googleusercontent.com/d/1xpNqF3k1eHr6m_4sx-bawegdBsLMvaT9" alt="ASA Philippines Foundation" class="h-14 w-auto object-contain bg-white rounded-xl p-1.5 shadow-md mb-2">
        <div class="text-center font-black tracking-wider text-xs uppercase leading-tight">LOGISTICS INVENTORY</div>
      </div>

      <!-- MAIN NAVIGATION VIEWS -->
      <div class="px-3 pt-4">
        <div class="px-3 pb-2 text-[10px] font-bold uppercase tracking-wider text-orange-200/80">Navigation Views</div>
        <nav class="space-y-1">
          <button type="button" id="navDashboard" onclick="switchView('dashboard')" class="w-full flex items-center gap-3 px-4 py-2.5 rounded-xl bg-asa-active font-semibold shadow-sm text-sm text-left transition cursor-pointer">
            <i class="fa-solid fa-gauge-high w-4"></i> Dashboard
          </button>
          <button type="button" id="navInventory" onclick="switchView('inventory')" class="w-full flex items-center gap-3 px-4 py-2.5 rounded-xl hover:bg-orange-600/50 font-medium text-orange-100 transition text-sm text-left cursor-pointer">
            <i class="fa-solid fa-boxes-stacked w-4"></i> Inventory
          </button>
          <button type="button" id="navGatePass" onclick="switchView('gatepass')" class="w-full flex items-center gap-3 px-4 py-2.5 rounded-xl hover:bg-orange-600/50 font-medium text-orange-100 transition text-sm text-left cursor-pointer">
            <i class="fa-solid fa-print w-4"></i> Gate Pass
          </button>
          <button type="button" id="navDivisionBudget" onclick="switchView('divisionBudget')" class="w-full flex items-center gap-3 px-4 py-2.5 rounded-xl hover:bg-orange-600/50 font-medium text-orange-100 transition text-sm text-left cursor-pointer">
            <i class="fa-solid fa-folder-open w-4"></i> Division Budget
          </button>
        </nav>
      </div>

      <!-- QUICK TRANSACTION ACTIONS -->
      <div class="px-3 pt-5 mt-4 border-t border-orange-400/30">
        <div class="px-3 pb-2 text-[10px] font-bold uppercase tracking-wider text-orange-200/80">Quick Actions</div>
        <div class="space-y-1.5">
          <button type="button" onclick="openModal('INCOMING')" class="w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl bg-white/10 hover:bg-emerald-600/40 text-orange-50 font-semibold text-xs transition cursor-pointer border border-white/10 shadow-xs">
            <span class="flex items-center gap-2.5"><i class="fa-solid fa-circle-plus text-emerald-300"></i> New Delivery</span>
            <span class="text-[9.5pt] bg-emerald-500/30 text-emerald-200 px-1.5 py-0.5 rounded font-mono font-bold">IN</span>
          </button>
          <button type="button" onclick="openModal('OUTGOING')" class="w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl bg-white/10 hover:bg-blue-600/40 text-orange-50 font-semibold text-xs transition cursor-pointer border border-white/10 shadow-xs">
            <span class="flex items-center gap-2.5"><i class="fa-solid fa-paper-plane text-blue-300"></i> New Dispatch</span>
            <span class="text-[9.5pt] bg-blue-500/30 text-blue-200 px-1.5 py-0.5 rounded font-mono font-bold">OUT</span>
          </button>
        </div>
      </div>
    </div>

    <!-- FOOTER -->
    <div class="p-4">
      <div class="bg-white/20 hover:bg-white/30 text-white rounded-xl py-2 px-3 text-center text-xs font-semibold flex items-center justify-center gap-2 cursor-pointer transition shadow-xs">
        <i class="fa-solid fa-circle-user"></i> Active Session
      </div>
    </div>
  </aside>

  <!-- MAIN CONTENT AREA -->
  <main class="flex-1 flex flex-col overflow-y-auto bg-slate-50">

    <!-- VIEW 1: DASHBOARD -->
    <div id="viewDashboard" class="flex-1 flex flex-col">
      <header class="bg-white border-b px-8 py-5 flex items-center justify-between sticky top-0 z-10 no-print">
        <div>
          <h1 class="text-2xl font-bold text-slate-800">Inventory Dashboard</h1>
          <p class="text-xs text-slate-500 mt-0.5">Live physical stock balance and transaction overview</p>
        </div>
      </header>

      <div class="p-8 space-y-6 max-w-7xl w-full">
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div class="lg:col-span-2 bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm">
            <div class="flex items-center justify-between mb-4">
              <h2 class="text-base font-bold text-slate-800">Monthly Transactions</h2>
              <select id="materialFilter" onchange="filterChart()" class="border border-slate-300 rounded-lg px-3 py-1.5 text-xs text-slate-700 font-semibold outline-none focus:ring-2 focus:ring-blue-500">
                <option value="ALL">All Materials</option>
              </select>
            </div>
            <div class="relative h-72"><canvas id="transactionChart"></canvas></div>
          </div>
          <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm flex flex-col">
            <div class="flex items-center justify-between mb-4">
              <h2 class="text-base font-bold text-slate-800">Inventory Summary</h2>
              <button type="button" onclick="switchView('inventory')" class="text-xs font-semibold text-blue-600 hover:underline cursor-pointer">View All</button>
            </div>
            <div class="grid grid-cols-2 text-xs font-semibold text-slate-400 uppercase tracking-wider pb-2 border-b">
              <span>Material</span><span class="text-right">Stock</span>
            </div>
            <div id="summaryList" class="divide-y divide-slate-100 text-xs max-h-72 overflow-y-auto pr-1">
              <div class="py-4 text-center text-slate-400">Loading...</div>
            </div>
          </div>
        </div>

        <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm">
          <div class="flex items-center justify-between mb-4">
            <h3 class="text-base font-bold text-slate-800">Recent Deliveries / Dispatches</h3>
            <div class="flex items-center gap-3">
              <button type="button" onclick="openSearchPastIncomingModal()" class="text-xs font-semibold text-emerald-600 hover:underline cursor-pointer">View All Incoming</button>
              <span class="text-slate-300">|</span>
              <button type="button" onclick="openSearchPastModal()" class="text-xs font-semibold text-blue-600 hover:underline cursor-pointer">View All Outgoing</button>
            </div>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs text-slate-600">
              <thead class="bg-slate-50 text-slate-400 uppercase font-semibold">
                <tr>
                  <th class="p-3">Type</th><th class="p-3">Date</th>
                  <th class="p-3">Supplier / Destination</th><th class="p-3">DR / Control No</th>
                </tr>
              </thead>
              <tbody id="transactionsTable" class="divide-y divide-slate-100">
                <tr><td colspan="4" class="text-center py-4 text-slate-400">Loading...</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <!-- VIEW 2: FULL MATERIAL INVENTORY -->
    <div id="viewInventory" class="hidden flex-1 flex flex-col">
      <header class="bg-white border-b px-8 py-5 flex items-center justify-between sticky top-0 z-10 no-print">
        <div>
          <h1 class="text-2xl font-bold text-slate-900">Full Material Inventory</h1>
          <p class="text-xs text-slate-500 mt-0.5">Live stock levels for all 13 tracked items</p>
        </div>
      </header>
      <div class="p-8 space-y-4 max-w-7xl w-full">
        <div class="relative w-full max-w-lg">
          <i class="fa-solid fa-magnifying-glass absolute left-3.5 top-3 text-slate-400 text-xs"></i>
          <input type="text" id="inventorySearchInput" oninput="renderFullInventoryTable()" placeholder="Search material name..." class="w-full pl-9 pr-4 py-2 text-xs border border-slate-300 rounded-lg outline-none focus:ring-2 focus:ring-blue-500 bg-white">
        </div>
        <div class="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs">
              <thead class="bg-slate-50/70 text-slate-400 font-semibold border-b">
                <tr>
                  <th class="py-3.5 px-6 font-medium">Material Name</th>
                  <th class="py-3.5 px-6 text-right font-medium">Total Received</th>
                  <th class="py-3.5 px-6 text-right font-medium">Total Released</th>
                  <th class="py-3.5 px-6 text-right font-medium">Available Stock</th>
                  <th class="py-3.5 px-6 text-center font-medium">Status</th>
                </tr>
              </thead>
              <tbody id="fullInventoryTableBody" class="divide-y divide-slate-100 text-slate-700">
                <tr><td colspan="5" class="py-8 text-center text-slate-400">Loading...</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <!-- VIEW 3: GATE PASS -->
    <div id="viewGatePass" class="hidden flex-1 flex flex-col">
      <header class="bg-white border-b px-8 py-4 flex items-center justify-between sticky top-0 z-10 no-print shadow-sm">
        <div class="flex items-center gap-4">
          <div>
            <h1 class="text-xl font-bold text-slate-800">Gate Pass / Delivery Receipts</h1>
            <p class="text-xs text-slate-500">Official 2-part A4 printable receipts</p>
          </div>
          <div class="flex items-center gap-2 bg-slate-100 px-3 py-1.5 rounded-xl border border-slate-200 ml-4">
            <span class="text-xs font-semibold text-slate-600">Select Shipment:</span>
            <select id="gpShipmentSelect" onchange="onSelectGatePassShipment()" class="text-xs font-bold bg-white border border-slate-300 rounded-lg px-2.5 py-1 outline-none text-blue-700">
              <option value="">-- Choose Control No / Record --</option>
            </select>
          </div>
        </div>
        <div class="flex items-center gap-3">
          <button type="button" onclick="clearGatePassForm()" class="px-3 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-xl border">
            <i class="fa-solid fa-rotate-left"></i> Reset Form
          </button>
          <button type="button" onclick="window.print()" class="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-xl text-sm font-semibold shadow-md transition cursor-pointer">
            <i class="fa-solid fa-print"></i> Print Gate Pass (A4)
          </button>
        </div>
      </header>

      <div class="p-4 flex justify-center bg-slate-200/60 overflow-y-auto">
        <div id="printableGatePassArea" class="bg-white shadow-xl p-3 w-full max-w-[210mm] border border-slate-300 text-slate-900 select-text">
          <!-- TOP HALF: LOGISTICS COPY -->
          <div class="gp-copy-block">
            <div class="flex items-center justify-between pb-0.5">
              <div class="flex items-center"><img src="https://lh3.googleusercontent.com/d/1xpNqF3k1eHr6m_4sx-bawegdBsLMvaT9" alt="ASA Philippines Foundation" class="h-8 w-auto object-contain"></div>
              <h2 class="text-[11pt] font-black tracking-wide text-slate-900 uppercase">ASA LOGISTIC DELIVERY RECEIPTS</h2>
              <div class="text-[7.5pt] font-bold text-slate-800 tracking-wider">Logistics Copy</div>
            </div>
            <table class="w-full gp-table text-left my-0.5">
              <tr>
                <td class="gp-lbl" style="width: 11%;">To :</td>
                <td class="gp-val" style="width: 39%;"><input type="text" id="gpTo" oninput="syncGatePassCopies('gpTo')" placeholder="AVP / Consignee Name" class="gp-input font-bold"></td>
                <td class="gp-lbl" style="width: 9%;">Contact</td>
                <td class="gp-val" style="width: 19%;"><input type="text" id="gpContact" oninput="syncGatePassCopies('gpContact')" placeholder="Contact No." class="gp-input"></td>
                <td class="gp-lbl" style="width: 8%;">Date</td>
                <td class="gp-val" style="width: 14%;"><input type="text" id="gpDate" oninput="syncGatePassCopies('gpDate')" placeholder="YYYY-MM-DD" class="gp-input font-medium"></td>
              </tr>
              <tr>
                <td class="gp-lbl">C/O R.A</td>
                <td class="gp-val"><input type="text" id="gpCoRa" oninput="syncGatePassCopies('gpCoRa')" placeholder="Regional Assistant / C/O Name" class="gp-input"></td>
                <td class="gp-lbl">Contact</td>
                <td class="gp-val"><input type="text" id="gpRaContact" oninput="syncGatePassCopies('gpRaContact')" placeholder="R.A Contact No." class="gp-input"></td>
                <td class="gp-val font-bold text-center" colspan="2">
                  <select id="gpCourier" onchange="syncGatePassCopies('gpCourier')" class="gp-input font-bold text-slate-800 bg-transparent text-center">
                    <option value="Bus Cargo">Bus Cargo</option>
                    <option value="In-House Delivery">In-House Delivery</option>
                    <option value="LBC Express">LBC Express</option>
                    <option value="Van / Truck Cargo">Van / Truck Cargo</option>
                    <option value="Other Carrier">Other Carrier</option>
                  </select>
                </td>
              </tr>
              <tr>
                <td class="gp-lbl">Address:</td>
                <td class="gp-val" colspan="3"><input type="text" id="gpAddress" oninput="syncGatePassCopies('gpAddress')" placeholder="Branch / Delivery Address" class="gp-input"></td>
                <td class="gp-lbl">Cntrl No.</td>
                <td class="gp-val font-bold"><input type="text" id="gpControlNo" oninput="syncGatePassCopies('gpControlNo')" placeholder="Control #" class="gp-input font-bold text-blue-700"></td>
              </tr>
              <tr>
                <td class="gp-lbl">Branch Code :</td>
                <td class="gp-val"><input type="text" id="gpBranchCode" oninput="syncGatePassCopies('gpBranchCode')" placeholder="Branch Code" class="gp-input"></td>
                <td class="gp-lbl">Cluster</td>
                <td class="gp-val"><input type="text" id="gpCluster" oninput="syncGatePassCopies('gpCluster')" placeholder="Cluster" class="gp-input font-bold"></td>
                <td class="gp-lbl">Division</td>
                <td class="gp-val"><input type="text" id="gpDivision" oninput="syncGatePassCopies('gpDivision')" placeholder="Division" class="gp-input font-medium"></td>
              </tr>
            </table>
            <table class="w-full gp-table text-center my-0.5">
              <tr class="gp-lbl">
                <th rowspan="2" style="width: 14%;" class="font-bold">DESCRIPTION</th>
                <th colspan="2" class="border">Financing Agreement</th>
                <th colspan="2" class="border">Passbook</th>
                <th colspan="3" class="border">Group Treasurer Register</th>
                <th colspan="2" class="border">Calendar</th>
                <th rowspan="2" style="width: 9%;" class="font-bold border">Guide Book</th>
              </tr>
              <tr class="gp-lbl">
                <th style="width: 8%;" class="border font-medium text-[7pt]">Regular</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Barmm</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Regular</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Barmm</th>
                <th style="width: 7%;" class="border font-medium text-[7pt]">New</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Regular</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Barmm</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Desk</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Wall</th>
              </tr>
              <tr>
                <td class="gp-lbl text-left px-1 font-bold">QUANTITY :</td>
                <td><input type="number" id="gpFaf" oninput="syncGatePassCopies('gpFaf');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpFafBarmm" oninput="syncGatePassCopies('gpFafBarmm');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpPassbook" oninput="syncGatePassCopies('gpPassbook');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpPassbookBarmm" oninput="syncGatePassCopies('gpPassbookBarmm');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpGtrNew" oninput="syncGatePassCopies('gpGtrNew');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpGtr" oninput="syncGatePassCopies('gpGtr');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpGtrBarmm" oninput="syncGatePassCopies('gpGtrBarmm');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpDeskCal" oninput="syncGatePassCopies('gpDeskCal');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpWallCal" oninput="syncGatePassCopies('gpWallCal');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpGuideBook" oninput="syncGatePassCopies('gpGuideBook');" placeholder="0" class="gp-input text-center font-bold"></td>
              </tr>
              <tr class="gp-lbl">
                <th class="text-left px-1 font-bold">DESCRIPTION</th>
                <th class="border font-medium text-[7pt]">Insurance</th>
                <th class="border font-medium text-[7pt]">Coverage</th>
                <th colspan="2" class="border font-medium text-[7pt]">Poster Acrylic</th>
                <th colspan="2" class="border font-medium text-[7pt]">Survey Form</th>
                <th class="border font-medium text-[7pt]"></th>
                <th class="border font-medium text-[7pt]"></th>
                <th class="border font-medium text-[7pt]"></th>
                <th class="border font-medium text-[7pt]"></th>
              </tr>
              <tr>
                <td class="gp-lbl text-left px-1 font-bold">QUANTITY :</td>
                <td><input type="number" id="gpEnrolment" oninput="syncGatePassCopies('gpEnrolment');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td><input type="number" id="gpCoverage" oninput="syncGatePassCopies('gpCoverage');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td colspan="2"><input type="number" id="gpPoster" oninput="syncGatePassCopies('gpPoster');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td colspan="2"><input type="number" id="gpSurveyForm" oninput="syncGatePassCopies('gpSurveyForm');" placeholder="0" class="gp-input text-center font-bold"></td>
                <td>0</td><td>0</td><td>0</td><td>0</td>
              </tr>
              <tr>
                <td class="gp-lbl text-left px-1 font-bold">NOTE :</td>
                <td colspan="10" class="text-left px-1"><input type="text" id="gpNote" oninput="syncGatePassCopies('gpNote')" placeholder="Remarks / Notes regarding shipment" class="gp-input"></td>
              </tr>
            </table>
            <div class="grid grid-cols-12 items-end pt-1 pb-0 text-center">
              <div class="col-span-3 space-y-0.5">
                <div class="text-left font-bold text-slate-800 mb-4 text-[7.5pt]">Released By :</div>
                <div class="border-b-2 border-black w-11/12 mx-auto"></div>
                <div class="font-bold text-slate-900 text-[8pt]">Sheren Ponteres</div>
              </div>
              <div class="col-span-3 flex justify-center">
                <div class="border-2 border-black flex overflow-hidden w-32 shadow-xs bg-white">
                  <div class="w-12 bg-white flex items-center justify-center font-bold text-[8.5pt] border-r-2 border-black py-2.5">Qty</div>
                  <div class="w-20 bg-white flex flex-col items-center justify-center font-bold text-[7pt] leading-tight py-2">
                    <span class="text-[6.5pt] text-slate-900 uppercase font-bold text-center leading-tight">BUNDLES /<br>PCS</span>
                  </div>
                </div>
              </div>
              <div class="col-span-3 space-y-0.5">
                <div class="text-left font-bold text-slate-800 mb-4 text-[7.5pt]">Approved by:</div>
                <div class="border-b-2 border-black w-11/12 mx-auto"></div>
                <div class="font-bold text-slate-900 text-[8pt]">Efren Camacan</div>
              </div>
              <div class="col-span-3 space-y-0.5">
                <div class="text-left font-bold text-slate-800 mb-4 text-[7.5pt]">Received by:</div>
                <div class="border-b-2 border-black w-11/12 mx-auto"></div>
                <div class="font-bold text-slate-700 text-[6.5pt] uppercase tracking-wider">SIGNATURE OVER PRINTED NAME</div>
              </div>
            </div>
          </div>

          <!-- PERFORATED CUT LINE -->
          <div class="gp-cut-line relative flex items-center justify-center border-t-2 border-dashed border-slate-400">
            <span class="bg-white px-4 text-[7pt] font-black text-slate-600 tracking-wider uppercase"><i class="fa-solid fa-scissors mr-1.5"></i> CUT HERE</span>
          </div>

          <!-- BOTTOM HALF: CONSIGNEE COPY -->
          <div class="gp-copy-block">
            <div class="flex items-center justify-between pb-0.5">
              <div class="flex items-center"><img src="https://lh3.googleusercontent.com/d/1xpNqF3k1eHr6m_4sx-bawegdBsLMvaT9" alt="ASA Philippines Foundation" class="h-8 w-auto object-contain"></div>
              <h2 class="text-[11pt] font-black tracking-wide text-slate-900 uppercase">ASA LOGISTIC DELIVERY RECEIPTS</h2>
              <div class="text-[7.5pt] font-bold text-slate-800 tracking-wider">Consignee Copy</div>
            </div>
            <table class="w-full gp-table text-left my-0.5">
              <tr>
                <td class="gp-lbl" style="width: 11%;">To :</td>
                <td class="gp-val" style="width: 39%;"><input type="text" id="gpTo_c" readonly class="gp-input font-bold"></td>
                <td class="gp-lbl" style="width: 9%;">Contact</td>
                <td class="gp-val" style="width: 19%;"><input type="text" id="gpContact_c" readonly class="gp-input"></td>
                <td class="gp-lbl" style="width: 8%;">Date</td>
                <td class="gp-val" style="width: 14%;"><input type="text" id="gpDate_c" readonly class="gp-input font-medium"></td>
              </tr>
              <tr>
                <td class="gp-lbl">C/O R.A</td>
                <td class="gp-val"><input type="text" id="gpCoRa_c" readonly class="gp-input"></td>
                <td class="gp-lbl">Contact</td>
                <td class="gp-val"><input type="text" id="gpRaContact_c" readonly class="gp-input"></td>
                <td class="gp-val font-bold text-center" colspan="2"><input type="text" id="gpCourier_c" readonly value="Bus Cargo" class="gp-input font-bold text-center"></td>
              </tr>
              <tr>
                <td class="gp-lbl">Address:</td>
                <td class="gp-val" colspan="3"><input type="text" id="gpAddress_c" readonly class="gp-input"></td>
                <td class="gp-lbl">Cntrl No.</td>
                <td class="gp-val font-bold"><input type="text" id="gpControlNo_c" readonly class="gp-input font-bold text-blue-700"></td>
              </tr>
              <tr>
                <td class="gp-lbl">Branch Code :</td>
                <td class="gp-val"><input type="text" id="gpBranchCode_c" readonly class="gp-input"></td>
                <td class="gp-lbl">Cluster</td>
                <td class="gp-val"><input type="text" id="gpCluster_c" readonly class="gp-input font-bold"></td>
                <td class="gp-lbl">Division</td>
                <td class="gp-val"><input type="text" id="gpDivision_c" readonly class="gp-input font-medium"></td>
              </tr>
            </table>
            <table class="w-full gp-table text-center my-0.5">
              <tr class="gp-lbl">
                <th rowspan="2" style="width: 14%;" class="font-bold">DESCRIPTION</th>
                <th colspan="2" class="border">Financing Agreement</th>
                <th colspan="2" class="border">Passbook</th>
                <th colspan="3" class="border">Group Treasurer Register</th>
                <th colspan="2" class="border">Calendar</th>
                <th rowspan="2" style="width: 9%;" class="font-bold border">Guide Book</th>
              </tr>
              <tr class="gp-lbl">
                <th style="width: 8%;" class="border font-medium text-[7pt]">Regular</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Barmm</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Regular</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Barmm</th>
                <th style="width: 7%;" class="border font-medium text-[7pt]">New</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Regular</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Barmm</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Desk</th>
                <th style="width: 8%;" class="border font-medium text-[7pt]">Wall</th>
              </tr>
              <tr>
                <td class="gp-lbl text-left px-1 font-bold">QUANTITY :</td>
                <td><input type="text" id="gpFaf_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpFafBarmm_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpPassbook_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpPassbookBarmm_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpGtrNew_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpGtr_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpGtrBarmm_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpDeskCal_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpWallCal_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpGuideBook_c" readonly class="gp-input text-center font-bold"></td>
              </tr>
              <tr class="gp-lbl">
                <th class="text-left px-1 font-bold">DESCRIPTION</th>
                <th class="border font-medium text-[7pt]">Insurance</th>
                <th class="border font-medium text-[7pt]">Coverage</th>
                <th colspan="2" class="border font-medium text-[7pt]">Poster Acrylic</th>
                <th colspan="2" class="border font-medium text-[7pt]">Survey Form</th>
                <th class="border font-medium text-[7pt]"></th>
                <th class="border font-medium text-[7pt]"></th>
                <th class="border font-medium text-[7pt]"></th>
                <th class="border font-medium text-[7pt]"></th>
              </tr>
              <tr>
                <td class="gp-lbl text-left px-1 font-bold">QUANTITY :</td>
                <td><input type="text" id="gpEnrolment_c" readonly class="gp-input text-center font-bold"></td>
                <td><input type="text" id="gpCoverage_c" readonly class="gp-input text-center font-bold"></td>
                <td colspan="2"><input type="text" id="gpPoster_c" readonly class="gp-input text-center font-bold"></td>
                <td colspan="2"><input type="text" id="gpSurveyForm_c" readonly class="gp-input text-center font-bold"></td>
                <td>0</td><td>0</td><td>0</td><td>0</td>
              </tr>
              <tr>
                <td class="gp-lbl text-left px-1 font-bold">NOTE :</td>
                <td colspan="10" class="text-left px-1"><input type="text" id="gpNote_c" readonly class="gp-input"></td>
              </tr>
            </table>
            <div class="grid grid-cols-12 items-end pt-1 pb-0 text-center">
              <div class="col-span-3 space-y-0.5">
                <div class="text-left font-bold text-slate-800 mb-4 text-[7.5pt]">Released By :</div>
                <div class="border-b-2 border-black w-11/12 mx-auto"></div>
                <div class="font-bold text-slate-900 text-[8pt]">Sheren Ponteres</div>
              </div>
              <div class="col-span-3 flex justify-center">
                <div class="border-2 border-black flex overflow-hidden w-32 shadow-xs bg-white">
                  <div class="w-12 bg-white flex items-center justify-center font-bold text-[8.5pt] border-r-2 border-black py-2.5">Qty</div>
                  <div class="w-20 bg-white flex flex-col items-center justify-center font-bold text-[7pt] leading-tight py-2">
                    <span class="text-[6.5pt] text-slate-900 uppercase font-bold text-center leading-tight">BUNDLES /<br>PCS</span>
                  </div>
                </div>
              </div>
              <div class="col-span-3 space-y-0.5">
                <div class="text-left font-bold text-slate-800 mb-4 text-[7.5pt]">Approved by:</div>
                <div class="border-b-2 border-black w-11/12 mx-auto"></div>
                <div class="font-bold text-slate-900 text-[8pt]">Efren Camacan</div>
              </div>
              <div class="col-span-3 space-y-0.5">
                <div class="text-left font-bold text-slate-800 mb-4 text-[7.5pt]">Received by:</div>
                <div class="border-b-2 border-black w-11/12 mx-auto"></div>
                <div class="font-bold text-slate-700 text-[6.5pt] uppercase tracking-wider">SIGNATURE OVER PRINTED NAME</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- VIEW 4: DIVISION BUDGET -->
    <div id="viewDivisionBudget" class="hidden flex-1 flex flex-col">
      <header class="bg-white border-b px-8 py-5 flex items-center justify-between sticky top-0 z-10 no-print shadow-sm">
        <div>
          <div class="flex items-center gap-2">
            <span class="bg-purple-100 text-purple-700 font-bold px-2 py-0.5 rounded text-[11px] uppercase tracking-wide">Budget Monitoring</span>
            <h1 class="text-2xl font-bold text-slate-900">Division Budget</h1>
          </div>
          <p class="text-xs text-slate-500 mt-0.5 font-medium">GTR · FAF · PASSBOOK — Operation Requests vs Actual Delivery & Balance Tracking</p>
        </div>
        <div class="flex items-center gap-3">
          <button type="button" onclick="loadDashboard()" class="flex items-center gap-2 border border-slate-300 hover:bg-slate-100 text-slate-700 px-3.5 py-2 rounded-xl text-xs font-semibold transition cursor-pointer">
            <i class="fa-solid fa-arrows-rotate text-blue-600"></i> Refresh Data
          </button>
          <button type="button" onclick="openDivisionBudgetModal()" class="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl text-xs font-bold transition shadow-sm cursor-pointer">
            <i class="fa-solid fa-plus"></i> Add Budget Row
          </button>
        </div>
      </header>

      <div class="p-8 space-y-6 max-w-full w-full">
        <!-- KPI METRICS SUMMARY -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
          <div class="bg-white p-5 rounded-2xl border border-purple-200/80 shadow-sm relative overflow-hidden">
            <div class="absolute right-3 top-3 text-purple-200 text-5xl opacity-40"><i class="fa-solid fa-book-bookmark"></i></div>
            <div class="text-xs font-bold text-purple-700 uppercase tracking-wider mb-1">🟣 GTR Budget Total</div>
            <div class="text-2xl font-black text-slate-900" id="kpiGtrReq">0</div>
            <div class="text-[11px] text-slate-500 mt-2 flex justify-between border-t pt-2 border-purple-50">
              <span>Delivered: <b id="kpiGtrDel" class="text-slate-800 font-bold">0</b></span>
              <span>Balance: <b id="kpiGtrBal" class="text-purple-700 font-bold">0</b></span>
            </div>
          </div>

          <div class="bg-white p-5 rounded-2xl border border-sky-200/80 shadow-sm relative overflow-hidden">
            <div class="absolute right-3 top-3 text-sky-200 text-5xl opacity-40"><i class="fa-solid fa-file-invoice-dollar"></i></div>
            <div class="text-xs font-bold text-sky-700 uppercase tracking-wider mb-1">🔵 FAF Budget Total</div>
            <div class="text-2xl font-black text-slate-900" id="kpiFafReq">0</div>
            <div class="text-[11px] text-slate-500 mt-2 flex justify-between border-t pt-2 border-sky-50">
              <span>Delivered: <b id="kpiFafDel" class="text-slate-800 font-bold">0</b></span>
              <span>Balance: <b id="kpiFafBal" class="text-sky-700 font-bold">0</b></span>
            </div>
          </div>

          <div class="bg-white p-5 rounded-2xl border border-amber-200/80 shadow-sm relative overflow-hidden">
            <div class="absolute right-3 top-3 text-amber-200 text-5xl opacity-40"><i class="fa-solid fa-address-book"></i></div>
            <div class="text-xs font-bold text-amber-700 uppercase tracking-wider mb-1">🟡 Passbook Budget Total</div>
            <div class="text-2xl font-black text-slate-900" id="kpiPbReq">0</div>
            <div class="text-[11px] text-slate-500 mt-2 flex justify-between border-t pt-2 border-amber-50">
              <span>Delivered: <b id="kpiPbDel" class="text-slate-800 font-bold">0</b></span>
              <span>Balance: <b id="kpiPbBal" class="text-amber-700 font-bold">0</b></span>
            </div>
          </div>
        </div>

        <!-- FILTERS & SEARCH -->
        <div class="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-wrap items-center justify-between gap-4">
          <div class="flex items-center gap-3 flex-1 min-w-[280px]">
            <div class="relative flex-1 max-w-sm">
              <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
              <input type="text" id="dbSearchInput" oninput="renderDivisionBudgetTable()" placeholder="Search by AVP Name..." class="w-full pl-8 pr-3 py-1.5 text-xs border border-slate-300 rounded-lg outline-none focus:ring-2 focus:ring-blue-500">
            </div>
            <select id="dbDivisionFilter" onchange="renderDivisionBudgetTable()" class="border border-slate-300 rounded-lg px-3 py-1.5 text-xs text-slate-700 font-semibold outline-none focus:ring-2 focus:ring-blue-500 bg-white">
              <option value="ALL">All Divisions</option>
            </select>
          </div>
          <div class="flex items-center gap-2 text-xs font-semibold text-slate-500">
            <span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-purple-600 inline-block"></span> GTR</span>
            <span class="flex items-center gap-1.5 ml-2"><span class="w-2.5 h-2.5 rounded-full bg-sky-500 inline-block"></span> FAF</span>
            <span class="flex items-center gap-1.5 ml-2"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span> PASSBOOK</span>
          </div>
        </div>

        <!-- MAIN BUDGET TABLE -->
        <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs border-collapse">
              <thead>
                <tr class="text-center font-bold text-xs uppercase tracking-wider">
                  <th rowspan="2" class="p-3 bg-slate-800 text-white border-r border-slate-700 min-w-[140px] text-left">AVP NAME</th>
                  <th rowspan="2" class="p-3 bg-slate-800 text-white border-r border-slate-700 min-w-[90px]">Division</th>
                  <th colspan="5" class="p-2 db-header-gtr border-r border-purple-800">GTR</th>
                  <th colspan="5" class="p-2 db-header-faf border-r border-sky-800">FAF</th>
                  <th colspan="6" class="p-2 db-header-pb border-r border-amber-800">PASSBOOK</th>
                  <th rowspan="2" class="p-3 bg-slate-800 text-white min-w-[120px]">Notes</th>
                  <th rowspan="2" class="p-3 bg-slate-800 text-white text-center">Action</th>
                </tr>
                <tr class="text-center text-[10.5px] border-b border-slate-200">
                  <th class="p-2 db-sub-gtr border-r border-purple-200">Operation Request</th>
                  <th class="p-2 db-sub-gtr border-r border-purple-200">5% of Request</th>
                  <th class="p-2 db-sub-gtr border-r border-purple-200">Actual Delivered / Old</th>
                  <th class="p-2 db-sub-gtr border-r border-purple-200">Actual Delivered / New</th>
                  <th class="p-2 db-sub-gtr border-r border-purple-300 font-black">Balance</th>
                  <th class="p-2 db-sub-faf border-r border-sky-200">Operation Request</th>
                  <th class="p-2 db-sub-faf border-r border-sky-200">5% of Request</th>
                  <th class="p-2 db-sub-faf border-r border-sky-200">With Less 5%</th>
                  <th class="p-2 db-sub-faf border-r border-sky-200">Actual Delivered</th>
                  <th class="p-2 db-sub-faf border-r border-sky-300 font-black">Balance</th>
                  <th class="p-2 db-sub-pb border-r border-amber-200">Operation Request</th>
                  <th class="p-2 db-sub-pb border-r border-amber-200">5% of Request</th>
                  <th class="p-2 db-sub-pb border-r border-amber-200">With Less 5%</th>
                  <th class="p-2 db-sub-pb border-r border-amber-200">Actual Delivered / Old</th>
                  <th class="p-2 db-sub-pb border-r border-amber-200">Actual Delivered / New</th>
                  <th class="p-2 db-sub-pb border-r border-amber-300 font-black">Balance</th>
                </tr>
              </thead>
              <tbody id="dbTableBody" class="divide-y divide-slate-100 font-medium text-slate-700">
                <tr><td colspan="19" class="py-12 text-center text-slate-400">Loading Division Budget data...</td></tr>
              </tbody>
              <tfoot id="dbTableFoot" class="bg-slate-900 text-white font-bold text-center border-t-2 border-slate-800"></tfoot>
            </table>
          </div>
        </div>
      </div>
    </div>

  </main>

  <!-- MODAL: SEARCH PAST OUTGOING -->
  <div id="searchPastModal" class="fixed inset-0 bg-slate-900/40 backdrop-blur-sm hidden items-center justify-center z-50 p-4 no-print">
    <div class="bg-white w-full max-w-5xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[85vh]">
      <div class="px-6 py-4 border-b flex justify-between items-center bg-slate-50">
        <div class="flex items-center gap-2 font-bold text-base">
          <i class="fa-solid fa-clock-rotate-left text-blue-600"></i>
          <span class="text-slate-800">Search Past Outgoing Shipments</span>
        </div>
        <button type="button" onclick="closeSearchPastModal()" class="text-slate-400 hover:text-slate-600 cursor-pointer text-lg"><i class="fa-solid fa-xmark"></i></button>
      </div>
      <div class="p-4 border-b bg-white">
        <div class="relative w-full">
          <i class="fa-solid fa-magnifying-glass absolute left-3.5 top-3 text-slate-400 text-xs"></i>
          <input type="text" id="pastOutgoingSearch" oninput="renderPastOutgoingTable()" placeholder="Search by Control No, AVP Name, Division, Destination, Cluster, or Date..." class="w-full pl-9 pr-4 py-2 text-xs border border-slate-300 rounded-lg outline-none focus:ring-2 focus:ring-blue-500 bg-white">
        </div>
      </div>
      <div class="overflow-y-auto p-4 flex-1">
        <table class="w-full text-left text-xs border border-slate-200 rounded-lg overflow-hidden">
          <thead class="bg-slate-50 text-slate-700 font-bold border-b border-slate-200">
            <tr>
              <th class="p-3 border-r">Date</th>
              <th class="p-3 border-r">Control No</th>
              <th class="p-3 border-r">AVP Name</th>
              <th class="p-3 border-r">Division</th>
              <th class="p-3 border-r">Destination</th>
              <th class="p-3 border-r">Cluster</th>
              <th class="p-3 text-center">Action</th>
            </tr>
          </thead>
          <tbody id="pastOutgoingTableBody" class="divide-y divide-slate-100">
            <tr><td colspan="7" class="py-8 text-center text-slate-400">Loading past records...</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- MODAL: SEARCH PAST INCOMING -->
  <div id="searchPastIncomingModal" class="fixed inset-0 bg-slate-900/40 backdrop-blur-sm hidden items-center justify-center z-50 p-4 no-print">
    <div class="bg-white w-full max-w-5xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[85vh]">
      <div class="px-6 py-4 border-b flex justify-between items-center bg-slate-50">
        <div class="flex items-center gap-2 font-bold text-base">
          <i class="fa-solid fa-truck-ramp-box text-emerald-600"></i>
          <span class="text-slate-800">Search Past Incoming Deliveries</span>
        </div>
        <button type="button" onclick="closeSearchPastIncomingModal()" class="text-slate-400 hover:text-slate-600 cursor-pointer text-lg"><i class="fa-solid fa-xmark"></i></button>
      </div>
      <div class="p-4 border-b bg-white">
        <div class="relative w-full">
          <i class="fa-solid fa-magnifying-glass absolute left-3.5 top-3 text-slate-400 text-xs"></i>
          <input type="text" id="pastIncomingSearch" oninput="renderPastIncomingTable()" placeholder="Search by DR #, Supplier Name, or Date..." class="w-full pl-9 pr-4 py-2 text-xs border border-slate-300 rounded-lg outline-none focus:ring-2 focus:ring-emerald-500 bg-white">
        </div>
      </div>
      <div class="overflow-y-auto p-4 flex-1">
        <table class="w-full text-left text-xs border border-slate-200 rounded-lg overflow-hidden">
          <thead class="bg-slate-50 text-slate-700 font-bold border-b border-slate-200">
            <tr>
              <th class="p-3 border-r">Date</th>
              <th class="p-3 border-r">DR #</th>
              <th class="p-3 border-r">Supplier Name</th>
              <th class="p-3 border-r">Items Delivered (Summary)</th>
              <th class="p-3 text-center">Action</th>
            </tr>
          </thead>
          <tbody id="pastIncomingTableBody" class="divide-y divide-slate-100">
            <tr><td colspan="5" class="py-8 text-center text-slate-400">Loading incoming records...</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- MODAL: ADD / EDIT DIVISION BUDGET ROW -->
  <div id="divisionBudgetModal" class="fixed inset-0 bg-slate-900/40 backdrop-blur-sm hidden items-center justify-center z-50 p-4 no-print">
    <div class="bg-white w-full max-w-4xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
      <div class="px-6 py-4 border-b flex justify-between items-center bg-slate-50">
        <div class="flex items-center gap-2 font-bold text-sm">
          <i class="fa-solid fa-folder-open text-purple-600"></i>
          <span id="dbModalTitle" class="text-slate-800">Add / Edit Division Budget Entry</span>
        </div>
        <button type="button" onclick="closeDivisionBudgetModal()" class="text-slate-400 hover:text-slate-600 cursor-pointer"><i class="fa-solid fa-xmark text-lg"></i></button>
      </div>

      <form id="dbForm" onsubmit="submitDivisionBudget(event)" class="p-6 overflow-y-auto space-y-4 text-xs">
        <input type="hidden" id="dbRowIndex" value="">

        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 pb-2 border-b">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">AVP Name *</label>
            <input type="text" id="dbFormAvp" required placeholder="e.g. John Doe" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-purple-500 bg-purple-50/20 font-bold">
          </div>
          <div>
            <label class="block font-semibold text-slate-700 mb-1">Division *</label>
            <input type="text" id="dbFormDivision" required placeholder="e.g. AVP 1 / Luzon" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-purple-500 bg-purple-50/20 font-bold">
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 pt-1">
          <div class="bg-purple-50/40 p-3 rounded-xl border border-purple-200 space-y-2">
            <h4 class="font-bold text-purple-800 uppercase tracking-wide text-[11px] flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-purple-600"></span> GTR Details</h4>
            <div>
              <label class="text-[10.5px] text-slate-600">Op Request:</label>
              <input type="number" id="dbGtrOp" step="any" placeholder="0" class="w-full border rounded p-1.5 bg-white text-xs">
            </div>
            <div>
              <label class="text-[10.5px] text-slate-600">Delivered (Old):</label>
              <input type="number" id="dbGtrOld" step="any" placeholder="0" class="w-full border rounded p-1.5 bg-white text-xs">
            </div>
            <div>
              <label class="text-[10.5px] text-slate-600">Delivered (New):</label>
              <input type="number" id="dbGtrNew" step="any" placeholder="0" class="w-full border rounded p-1.5 bg-white text-xs">
            </div>
          </div>

          <div class="bg-sky-50/40 p-3 rounded-xl border border-sky-200 space-y-2">
            <h4 class="font-bold text-sky-800 uppercase tracking-wide text-[11px] flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-sky-600"></span> FAF Details</h4>
            <div>
              <label class="text-[10.5px] text-slate-600">Op Request:</label>
              <input type="number" id="dbFafOp" step="any" placeholder="0" class="w-full border rounded p-1.5 bg-white text-xs">
            </div>
            <div>
              <label class="text-[10.5px] text-slate-600">Delivered:</label>
              <input type="number" id="dbFafDel" step="any" placeholder="0" class="w-full border rounded p-1.5 bg-white text-xs">
            </div>
          </div>

          <div class="bg-amber-50/40 p-3 rounded-xl border border-amber-200 space-y-2">
            <h4 class="font-bold text-amber-800 uppercase tracking-wide text-[11px] flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-amber-600"></span> Passbook Details</h4>
            <div>
              <label class="text-[10.5px] text-slate-600">Op Request:</label>
              <input type="number" id="dbPbOp" step="any" placeholder="0" class="w-full border rounded p-1.5 bg-white text-xs">
            </div>
            <div>
              <label class="text-[10.5px] text-slate-600">Delivered (Old):</label>
              <input type="number" id="dbPbOld" step="any" placeholder="0" class="w-full border rounded p-1.5 bg-white text-xs">
            </div>
            <div>
              <label class="text-[10.5px] text-slate-600">Delivered (New):</label>
              <input type="number" id="dbPbNew" step="any" placeholder="0" class="w-full border rounded p-1.5 bg-white text-xs">
            </div>
          </div>
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">Notes / Remarks</label>
          <input type="text" id="dbFormNotes" placeholder="Optional notes regarding this budget allocation" class="w-full border rounded-lg p-2 outline-none">
        </div>

        <div class="flex justify-end gap-3 pt-4 border-t">
          <button type="button" onclick="closeDivisionBudgetModal()" class="px-4 py-2 border rounded-xl font-semibold text-slate-600 hover:bg-slate-100 cursor-pointer">Cancel</button>
          <button type="submit" id="dbSubmitBtn" class="px-5 py-2 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl shadow cursor-pointer">Save Budget Entry</button>
        </div>
      </form>
    </div>
  </div>

  <!-- MODAL: ADD / EDIT OUTGOING OR INCOMING -->
  <div id="transactionModal" class="fixed inset-0 bg-slate-900/40 backdrop-blur-sm hidden items-center justify-center z-50 p-4 no-print">
    <div class="bg-white w-full max-w-3xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
      <div class="px-6 py-4 border-b flex justify-between items-center bg-slate-50">
        <div class="flex items-center gap-3">
          <h3 id="modalTitle" class="font-bold text-slate-800 text-sm">Add Entry</h3>
          <button type="button" id="searchPastOutgoingLink" onclick="openSearchPastModal()" class="text-[11px] font-semibold text-blue-600 hover:underline flex items-center gap-1 cursor-pointer">
            <i class="fa-solid fa-clock-rotate-left"></i> Search Past Outgoing Record?
          </button>
          <button type="button" id="searchPastIncomingLink" onclick="openSearchPastIncomingModal()" class="text-[11px] font-semibold text-emerald-600 hover:underline flex items-center gap-1 cursor-pointer hidden">
            <i class="fa-solid fa-truck-ramp-box"></i> Search Past Incoming Record?
          </button>
        </div>
        <button type="button" onclick="closeModal()" class="text-slate-400 hover:text-slate-600 cursor-pointer"><i class="fa-solid fa-xmark text-lg"></i></button>
      </div>

      <form id="transForm" onsubmit="submitTransaction(event)" class="p-6 overflow-y-auto space-y-4 text-xs">
        <div id="incomingFields" class="grid grid-cols-1 md:grid-cols-3 gap-3">
          <div>
            <label class="block font-semibold text-slate-600 mb-1">Date *</label>
            <input type="date" id="incDate" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500">
          </div>
          <div>
            <label class="block font-semibold text-slate-600 mb-1">Supplier (Listbox) *</label>
            <select id="incSupplier" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500 bg-amber-50/50 font-medium">
              <option value="">-- Select Supplier --</option>
            </select>
          </div>
          <div>
            <label class="block font-semibold text-slate-600 mb-1">DR #</label>
            <input type="text" id="incDr" placeholder="e.g. DR-001" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500">
          </div>
        </div>

        <div id="outgoingFields" class="hidden space-y-3">
          <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Date *</label>
              <input type="date" id="outDate" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">AVP Name *</label>
              <select id="outAvp" onchange="onAvpSelect()" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500 bg-amber-50/50 font-medium">
                <option value="">-- Select AVP --</option>
              </select>
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Division (Auto)</label>
              <input type="text" id="outDivision" readonly placeholder="Auto-filled from AVP" class="w-full border rounded-lg p-2 outline-none bg-slate-100 font-medium text-slate-700">
            </div>
          </div>
          <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Region / Cluster *</label>
              <input type="text" id="outRegion" oninput="onRegionType()" placeholder="e.g. CLUSTER 35" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500 bg-amber-50/50 font-medium">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Destination (Auto from AVP)</label>
              <input type="text" id="outDestination" placeholder="Branch / Destination" class="w-full border rounded-lg p-2 outline-none bg-slate-100 font-medium text-slate-700">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Control No</label>
              <input type="text" id="outControl" placeholder="Control #" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500">
            </div>
          </div>
          <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Cluster Head (Auto)</label>
              <input type="text" id="outClusterHead" placeholder="Auto-filled from Cluster" class="w-full border rounded-lg p-2 outline-none bg-slate-100 font-medium text-slate-700">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Cluster Head Contact (Auto)</label>
              <input type="text" id="outClusterContact" placeholder="Auto-filled from Cluster" class="w-full border rounded-lg p-2 outline-none bg-slate-100 font-medium text-slate-700">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Base Station (Auto)</label>
              <input type="text" id="outBaseStation" placeholder="Auto-filled from Cluster" class="w-full border rounded-lg p-2 outline-none bg-slate-100 font-medium text-slate-700">
            </div>
          </div>
          <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Operation (Auto from AVP)</label>
              <input type="text" id="outOperation" placeholder="Operation / Purpose" class="w-full border rounded-lg p-2 outline-none bg-slate-100 font-medium text-slate-700">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Notes</label>
              <input type="text" id="outNotes" placeholder="Additional notes" class="w-full border rounded-lg p-2 outline-none focus:ring-2 focus:ring-blue-500">
            </div>
          </div>
        </div>

        <div class="pt-2">
          <label class="block font-bold text-slate-700 mb-2 border-b pb-1">Enter Material Quantities</label>
          <div id="materialInputsGrid" class="grid grid-cols-2 sm:grid-cols-3 gap-3"></div>
        </div>

        <div class="flex justify-end gap-3 pt-4 border-t">
          <button type="button" onclick="closeModal()" class="px-4 py-2 border rounded-xl font-semibold text-slate-600 hover:bg-slate-100 cursor-pointer">Cancel</button>
          <button type="submit" id="submitBtn" class="px-5 py-2 bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-xl shadow cursor-pointer">Save Entry</button>
        </div>
      </form>
    </div>
  </div>

  <script>
    var globalData = null;
    var chartInstance = null;
    var currentModalType = 'INCOMING';
    var editingRowIndex = null;

    // VIEW SWITCHING
    window.switchView = function(view) {
      var isDash = (view === 'dashboard');
      var isInv  = (view === 'inventory');
      var isGp   = (view === 'gatepass');
      var isDb   = (view === 'divisionBudget');

      document.getElementById('viewDashboard').classList.toggle('hidden', !isDash);
      document.getElementById('viewInventory').classList.toggle('hidden', !isInv);
      document.getElementById('viewGatePass').classList.toggle('hidden', !isGp);
      document.getElementById('viewDivisionBudget').classList.toggle('hidden', !isDb);

      var navDash = document.getElementById('navDashboard');
      var navInv  = document.getElementById('navInventory');
      var navGp   = document.getElementById('navGatePass');
      var navDb   = document.getElementById('navDivisionBudget');

      var activeClass = 'w-full flex items-center gap-3 px-4 py-2.5 rounded-xl bg-asa-active font-semibold shadow-sm text-sm text-left transition cursor-pointer';
      var inactiveClass = 'w-full flex items-center gap-3 px-4 py-2.5 rounded-xl hover:bg-orange-600/50 font-medium text-orange-100 transition text-sm text-left cursor-pointer';

      navDash.className = isDash ? activeClass : inactiveClass;
      navInv.className  = isInv  ? activeClass : inactiveClass;
      navGp.className   = isGp   ? activeClass : inactiveClass;
      navDb.className   = isDb   ? activeClass : inactiveClass;

      if (isInv) renderFullInventoryTable();
      if (isDb)  renderDivisionBudgetTable();
    };

    // INIT
    window.addEventListener('DOMContentLoaded', function() {
      var today = new Date().toISOString().split('T')[0];
      document.getElementById('incDate').value = today;
      document.getElementById('outDate').value = today;
      document.getElementById('gpDate').value = today;
      syncGatePassCopies('gpDate');
      loadDashboard();
    });

    function loadDashboard() {
      google.script.run
        .withSuccessHandler(renderDashboard)
        .withFailureHandler(function(err) { alert('Failed to load: ' + err.message); })
        .getDashboardData();
    }

    function renderDashboard(data) {
      globalData = data;

      var sel = document.getElementById('materialFilter');
      sel.innerHTML = '<option value="ALL">All Materials</option>';
      data.materials.forEach(function(mat) {
        sel.innerHTML += '<option value="' + mat.key + '">' + mat.name + '</option>';
      });

      var avpSel = document.getElementById('outAvp');
      var avpHtml = '<option value="">-- Select AVP --</option>';
      (data.avpList || []).forEach(function(avp) {
        avpHtml += '<option value="' + avp + '">' + avp + '</option>';
      });
      avpSel.innerHTML = avpHtml;

      var supSel = document.getElementById('incSupplier');
      var supHtml = '<option value="">-- Select Supplier --</option>';
      (data.supplierList || []).forEach(function(sup) {
        supHtml += '<option value="' + sup + '">' + sup + '</option>';
      });
      supSel.innerHTML = supHtml;

      var gpSel = document.getElementById('gpShipmentSelect');
      var gpHtml = '<option value="">-- Choose Control No / Record --</option>';
      (data.pastOutgoingRecords || []).forEach(function(r) {
        gpHtml += '<option value="' + r.id + '">' + (r.controlNo || 'N/A') + ' | ' + r.avpName + ' (' + r.destination + ') - ' + r.date + '</option>';
      });
      gpSel.innerHTML = gpHtml;

      var dbDivSel = document.getElementById('dbDivisionFilter');
      dbDivSel.innerHTML = '<option value="ALL">All Divisions</option>';
      if (data.divisionBudgetData && data.divisionBudgetData.divisions) {
        data.divisionBudgetData.divisions.forEach(function(d) {
          dbDivSel.innerHTML += '<option value="' + d + '">' + d + '</option>';
        });
      }

      var grid = document.getElementById('materialInputsGrid');
      grid.innerHTML = data.materials.map(function(mat) {
        return '<div class="bg-slate-50 p-2 rounded-lg border border-slate-200">' +
          '<label class="block text-[11px] font-semibold text-slate-700 truncate" title="' + mat.name + '">' + mat.key + '</label>' +
          '<input type="number" min="0" step="any" data-mat="' + mat.key + '" placeholder="0" class="w-full mt-1 border rounded p-1.5 text-xs bg-white focus:ring-1 focus:ring-blue-500 outline-none">' +
          '</div>';
      }).join('');

      document.getElementById('summaryList').innerHTML = data.inventoryList.map(function(item) {
        var badge = item.isLow ? '<span class="bg-amber-100 text-amber-600 font-bold px-1.5 py-0.5 rounded text-[10px]">Low</span>' : '';
        var cls = item.stock < 0 ? 'text-red-500' : 'text-slate-800';
        return '<div class="py-2.5 flex items-center justify-between">' +
          '<div class="flex items-center gap-2"><span class="font-medium text-slate-700">' + item.name + '</span>' + badge + '</div>' +
          '<div class="font-bold ' + cls + '">' + Number(item.stock).toLocaleString() + '</div></div>';
      }).join('');

      var txHtml = data.recentTransactions.length === 0
        ? '<tr><td colspan="4" class="text-center py-4 text-slate-400">No records found.</td></tr>'
        : data.recentTransactions.map(function(t) {
            var cls = t.type === 'Incoming' ? 'text-emerald-600' : 'text-blue-600';
            return '<tr class="hover:bg-slate-50"><td class="p-3 font-semibold ' + cls + '">' + t.type + '</td>' +
              '<td class="p-3 text-slate-500">' + t.date + '</td>' +
              '<td class="p-3 font-medium text-slate-800">' + t.party + '</td>' +
              '<td class="p-3 text-slate-500">' + t.ref + '</td></tr>';
          }).join('');
      document.getElementById('transactionsTable').innerHTML = txHtml;

      renderFullInventoryTable();
      renderDivisionBudgetTable();
      buildChart('ALL');
    }

    // ── PAST INCOMING MODAL & SELECTION ─────────────────────────────────────

    window.openSearchPastIncomingModal = function() {
      document.getElementById('searchPastIncomingModal').classList.remove('hidden');
      document.getElementById('searchPastIncomingModal').classList.add('flex');
      document.getElementById('pastIncomingSearch').value = '';
      renderPastIncomingTable();
    };

    window.closeSearchPastIncomingModal = function() {
      document.getElementById('searchPastIncomingModal').classList.add('hidden');
      document.getElementById('searchPastIncomingModal').classList.remove('flex');
    };

    window.renderPastIncomingTable = function() {
      if (!globalData || !globalData.pastIncomingRecords) return;
      var query = (document.getElementById('pastIncomingSearch').value || '').toLowerCase().trim();
      var records = globalData.pastIncomingRecords.filter(function(r) {
        return [r.drNumber, r.supplier, r.date]
          .some(function(v) { return v && v.toLowerCase().indexOf(query) !== -1; });
      });

      if (records.length === 0) {
        document.getElementById('pastIncomingTableBody').innerHTML = '<tr><td colspan="5" class="py-6 text-center text-slate-400">No matching incoming records found.</td></tr>';
        return;
      }

      document.getElementById('pastIncomingTableBody').innerHTML = records.map(function(r) {
        var summary = [];
        if (r.items) {
          Object.keys(r.items).forEach(function(k) {
            if (r.items[k] > 0) summary.push(k + ': ' + Number(r.items[k]).toLocaleString());
          });
        }
        var summaryText = summary.length > 0 ? summary.slice(0, 3).join(', ') + (summary.length > 3 ? '...' : '') : 'No items';

        return '<tr class="hover:bg-slate-50 transition border-b">' +
          '<td class="p-3 font-medium text-slate-600 border-r">' + r.date + '</td>' +
          '<td class="p-3 font-bold text-emerald-600 border-r">' + (r.drNumber || 'N/A') + '</td>' +
          '<td class="p-3 font-medium text-slate-800 border-r">' + r.supplier + '</td>' +
          '<td class="p-3 text-slate-600 border-r truncate max-w-xs">' + summaryText + '</td>' +
          '<td class="p-3 text-center">' +
          '<button type="button" onclick="loadPastIncomingRecordIntoForm(' + r.id + ')" class="px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 font-semibold rounded text-[11px] transition cursor-pointer">Edit</button>' +
          '</td></tr>';
      }).join('');
    };

    window.loadPastIncomingRecordIntoForm = function(recordId) {
      if (!globalData || !globalData.pastIncomingRecords) return;
      var record = globalData.pastIncomingRecords.find(function(r) { return r.id === recordId; });
      if (!record) return;

      editingRowIndex = record.id;
      closeSearchPastIncomingModal();
      openModal('INCOMING');

      var btn = document.getElementById('submitBtn');
      btn.innerText = 'Update Incoming Entry';
      btn.classList.remove('bg-blue-600', 'hover:bg-blue-700');
      btn.classList.add('bg-emerald-600', 'hover:bg-emerald-700');

      if (record.date) document.getElementById('incDate').value = record.date;
      if (record.drNumber && record.drNumber !== '-') document.getElementById('incDr').value = record.drNumber;

      if (record.supplier && record.supplier !== '-') {
        var supSel = document.getElementById('incSupplier');
        var targetSup = record.supplier.trim().toLowerCase();
        var matchIndex = -1;

        for (var i = 0; i < supSel.options.length; i++) {
          if (supSel.options[i].value.trim().toLowerCase() === targetSup) {
            matchIndex = i;
            break;
          }
        }

        if (matchIndex >= 0) {
          supSel.selectedIndex = matchIndex;
        } else {
          var opt = document.createElement('option');
          opt.value = record.supplier.trim();
          opt.text = record.supplier.trim();
          opt.selected = true;
          supSel.appendChild(opt);
        }
      }

      document.querySelectorAll('#materialInputsGrid input').forEach(function(input) {
        var qty = record.items && record.items[input.dataset.mat];
        input.value = (qty && qty > 0) ? qty : '';
      });
    };

    // ── DIVISION BUDGET RENDERING ──────────────────────────────────────────

    window.renderDivisionBudgetTable = function() {
      if (!globalData || !globalData.divisionBudgetData) return;

      var rows = globalData.divisionBudgetData.rows || [];
      var query = (document.getElementById('dbSearchInput').value || '').toLowerCase().trim();
      var divFilter = document.getElementById('dbDivisionFilter').value;

      var filtered = rows.filter(function(r) {
        var matchName = r.avpName.toLowerCase().indexOf(query) !== -1;
        var matchDiv  = (divFilter === 'ALL' || r.division === divFilter);
        return matchName && matchDiv;
      });

      var formatBal = function(val) {
        var num = Number(val) || 0;
        var formatted = (num >= 0 ? '+' : '') + num.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
        var cls = num >= 0 ? 'text-emerald-600 bg-emerald-50/50' : 'text-red-500 bg-red-50/50 font-bold';
        return '<td class="p-2 text-right border-r font-semibold ' + cls + '">' + formatted + '</td>';
      };

      var formatNum = function(val, borderCls) {
        var num = Number(val) || 0;
        return '<td class="p-2 text-right border-r ' + (borderCls || '') + '">' + (num > 0 ? num.toLocaleString() : '-') + '</td>';
      };

      if (filtered.length === 0) {
        document.getElementById('dbTableBody').innerHTML = '<tr><td colspan="19" class="py-10 text-center text-slate-400">No Division Budget rows found. Click "+ Add Budget Row" to create one.</td></tr>';
        document.getElementById('dbTableFoot').innerHTML = '';
        return;
      }

      var t = {
        gtrOp: 0, gtr5: 0, gtrOld: 0, gtrNew: 0, gtrBal: 0,
        fafOp: 0, faf5: 0, fafLess5: 0, fafDel: 0, fafBal: 0,
        pbOp: 0, pb5: 0, pbLess5: 0, pbOld: 0, pbNew: 0, pbBal: 0
      };

      var html = filtered.map(function(r) {
        t.gtrOp += r.gtr.opRequest; t.gtr5 += r.gtr.fivePercent; t.gtrOld += r.gtr.deliveredOld; t.gtrNew += r.gtr.deliveredNew; t.gtrBal += r.gtr.balance;
        t.fafOp += r.faf.opRequest; t.faf5 += r.faf.fivePercent; t.fafLess5 += r.faf.withLess5; t.fafDel += r.faf.delivered; t.fafBal += r.faf.balance;
        t.pbOp  += r.pb.opRequest;  t.pb5  += r.pb.fivePercent;  t.pbLess5  += r.pb.withLess5;  t.pbOld  += r.pb.deliveredOld;  t.pbNew  += r.pb.deliveredNew; t.pbBal += r.pb.balance;

        return '<tr class="hover:bg-slate-50 transition border-b border-slate-100">' +
          '<td class="p-2.5 font-bold text-slate-900 border-r">' + r.avpName + '</td>' +
          '<td class="p-2.5 text-center font-semibold text-slate-600 border-r bg-slate-50/50">' + r.division + '</td>' +
          formatNum(r.gtr.opRequest) + formatNum(r.gtr.fivePercent) + formatNum(r.gtr.deliveredOld) + formatNum(r.gtr.deliveredNew) + formatBal(r.gtr.balance) +
          formatNum(r.faf.opRequest) + formatNum(r.faf.fivePercent) + formatNum(r.faf.withLess5) + formatNum(r.faf.delivered) + formatBal(r.faf.balance) +
          formatNum(r.pb.opRequest) + formatNum(r.pb.fivePercent) + formatNum(r.pb.withLess5) + formatNum(r.pb.deliveredOld) + formatNum(r.pb.deliveredNew) + formatBal(r.pb.balance) +
          '<td class="p-2 text-slate-500 truncate max-w-[120px] text-left">' + (r.notes || '-') + '</td>' +
          '<td class="p-2 text-center">' +
            '<button type="button" onclick="editDivisionBudgetRow(' + r.rowIndex + ')" class="px-2 py-1 bg-purple-50 hover:bg-purple-100 text-purple-700 font-bold rounded text-[11px] transition">Edit</button>' +
          '</td>' +
          '</tr>';
      }).join('');

      document.getElementById('dbTableBody').innerHTML = html;

      document.getElementById('dbTableFoot').innerHTML = '<tr>' +
        '<td colspan="2" class="p-3 text-left font-black tracking-wider uppercase">Total Summary</td>' +
        '<td class="p-2 text-right">' + t.gtrOp.toLocaleString() + '</td><td class="p-2 text-right">' + t.gtr5.toLocaleString() + '</td><td class="p-2 text-right">' + t.gtrOld.toLocaleString() + '</td><td class="p-2 text-right">' + t.gtrNew.toLocaleString() + '</td><td class="p-2 text-right ' + (t.gtrBal >= 0 ? 'text-emerald-400' : 'text-red-400') + '">' + (t.gtrBal >= 0 ? '+' : '') + t.gtrBal.toLocaleString() + '</td>' +
        '<td class="p-2 text-right">' + t.fafOp.toLocaleString() + '</td><td class="p-2 text-right">' + t.faf5.toLocaleString() + '</td><td class="p-2 text-right">' + t.fafLess5.toLocaleString() + '</td><td class="p-2 text-right">' + t.fafDel.toLocaleString() + '</td><td class="p-2 text-right ' + (t.fafBal >= 0 ? 'text-emerald-400' : 'text-red-400') + '">' + (t.fafBal >= 0 ? '+' : '') + t.fafBal.toLocaleString() + '</td>' +
        '<td class="p-2 text-right">' + t.pbOp.toLocaleString() + '</td><td class="p-2 text-right">' + t.pb5.toLocaleString() + '</td><td class="p-2 text-right">' + t.pbLess5.toLocaleString() + '</td><td class="p-2 text-right">' + t.pbOld.toLocaleString() + '</td><td class="p-2 text-right">' + t.pbNew.toLocaleString() + '</td><td class="p-2 text-right ' + (t.pbBal >= 0 ? 'text-emerald-400' : 'text-red-400') + '">' + (t.pbBal >= 0 ? '+' : '') + t.pbBal.toLocaleString() + '</td>' +
        '<td colspan="2"></td>' +
        '</tr>';

      document.getElementById('kpiGtrReq').innerText = Number(t.gtrOp).toLocaleString();
      document.getElementById('kpiGtrDel').innerText = Number(t.gtrOld + t.gtrNew).toLocaleString();
      document.getElementById('kpiGtrBal').innerText = (t.gtrBal >= 0 ? '+' : '') + Number(t.gtrBal).toLocaleString();

      document.getElementById('kpiFafReq').innerText = Number(t.fafOp).toLocaleString();
      document.getElementById('kpiFafDel').innerText = Number(t.fafDel).toLocaleString();
      document.getElementById('kpiFafBal').innerText = (t.fafBal >= 0 ? '+' : '') + Number(t.fafBal).toLocaleString();

      document.getElementById('kpiPbReq').innerText  = Number(t.pbOp).toLocaleString();
      document.getElementById('kpiPbDel').innerText  = Number(t.pbOld + t.pbNew).toLocaleString();
      document.getElementById('kpiPbBal').innerText  = (t.pbBal >= 0 ? '+' : '') + Number(t.pbBal).toLocaleString();
    };

    window.openDivisionBudgetModal = function() {
      document.getElementById('dbRowIndex').value = '';
      document.getElementById('dbForm').reset();
      document.getElementById('dbModalTitle').innerText = 'Add Division Budget Entry';
      document.getElementById('dbSubmitBtn').innerText = 'Save Budget Entry';
      document.getElementById('divisionBudgetModal').classList.remove('hidden');
      document.getElementById('divisionBudgetModal').classList.add('flex');
    };

    window.closeDivisionBudgetModal = function() {
      document.getElementById('divisionBudgetModal').classList.add('hidden');
      document.getElementById('divisionBudgetModal').classList.remove('flex');
    };

    window.editDivisionBudgetRow = function(rowIndex) {
      if (!globalData || !globalData.divisionBudgetData) return;
      var record = globalData.divisionBudgetData.rows.find(function(r) { return r.rowIndex === rowIndex; });
      if (!record) return;

      document.getElementById('dbRowIndex').value     = record.rowIndex;
      document.getElementById('dbFormAvp').value      = record.avpName;
      document.getElementById('dbFormDivision').value = record.division;
      document.getElementById('dbGtrOp').value        = record.gtr.opRequest || '';
      document.getElementById('dbGtrOld').value       = record.gtr.deliveredOld || '';
      document.getElementById('dbGtrNew').value       = record.gtr.deliveredNew || '';
      document.getElementById('dbFafOp').value        = record.faf.opRequest || '';
      document.getElementById('dbFafDel').value       = record.faf.delivered || '';
      document.getElementById('dbPbOp').value         = record.pb.opRequest || '';
      document.getElementById('dbPbOld').value        = record.pb.deliveredOld || '';
      document.getElementById('dbPbNew').value        = record.pb.deliveredNew || '';
      document.getElementById('dbFormNotes').value    = record.notes || '';

      document.getElementById('dbModalTitle').innerText = 'Edit Division Budget Entry';
      document.getElementById('dbSubmitBtn').innerText = 'Update Budget Entry';
      document.getElementById('divisionBudgetModal').classList.remove('hidden');
      document.getElementById('divisionBudgetModal').classList.add('flex');
    };

    window.submitDivisionBudget = function(e) {
      e.preventDefault();
      var btn = document.getElementById('dbSubmitBtn');
      btn.disabled = true;
      btn.innerText = 'Saving...';

      var formData = {
        rowIndex: document.getElementById('dbRowIndex').value,
        avpName: document.getElementById('dbFormAvp').value.trim(),
        division: document.getElementById('dbFormDivision').value.trim(),
        gtrOpRequest: document.getElementById('dbGtrOp').value,
        gtrDeliveredOld: document.getElementById('dbGtrOld').value,
        gtrDeliveredNew: document.getElementById('dbGtrNew').value,
        fafOpRequest: document.getElementById('dbFafOp').value,
        fafDelivered: document.getElementById('dbFafDel').value,
        pbOpRequest: document.getElementById('dbPbOp').value,
        pbDeliveredOld: document.getElementById('dbPbOld').value,
        pbDeliveredNew: document.getElementById('dbPbNew').value,
        notes: document.getElementById('dbFormNotes').value.trim()
      };

      google.script.run
        .withSuccessHandler(function(res) {
          btn.disabled = false;
          btn.innerText = 'Save Budget Entry';
          closeDivisionBudgetModal();
          loadDashboard();
        })
        .withFailureHandler(function(err) {
          alert('Failed to save Division Budget: ' + err.message);
          btn.disabled = false;
          btn.innerText = 'Save Budget Entry';
        })
        .saveDivisionBudgetRow(formData);
    };

    // ── AVP & CLUSTER AUTOFILL ──────────────────────────────────────────────

    window.onAvpSelect = function() {
      if (!globalData || !globalData.avpDirectory) return;
      var avp = document.getElementById('outAvp').value.trim();
      var info = globalData.avpDirectory[avp] || globalData.avpDirectory[avp.toLowerCase()];
      document.getElementById('outDivision').value    = info ? (info.division    || '') : '';
      document.getElementById('outOperation').value   = info ? (info.operation   || '') : '';
      document.getElementById('outDestination').value = info ? (info.destination || '') : '';
    };

    window.onRegionType = function() {
      if (!globalData || !globalData.regionDirectory) return;
      var raw = document.getElementById('outRegion').value.trim();
      if (!raw) {
        document.getElementById('outClusterHead').value = '';
        document.getElementById('outClusterContact').value = '';
        document.getElementById('outBaseStation').value = '';
        return;
      }
      var info = globalData.regionDirectory[raw] ||
                 globalData.regionDirectory[raw.toLowerCase()] ||
                 globalData.regionDirectory[raw.toLowerCase().replace(/\\s+/g, '')];
      if (!info && !isNaN(raw)) {
        info = globalData.regionDirectory['cluster ' + raw] || globalData.regionDirectory['cluster' + raw];
      }
      if (info) {
        document.getElementById('outClusterHead').value    = info.clusterHead        || '';
        document.getElementById('outClusterContact').value = info.clusterHeadContact || '';
        document.getElementById('outBaseStation').value    = info.baseStation        || '';
      }
    };

    // ── GATE PASS LOGIC ─────────────────────────────────────────────────────

    window.syncGatePassCopies = function(fieldId) {
      var val = document.getElementById(fieldId).value;
      var mirror = document.getElementById(fieldId + '_c');
      if (mirror) mirror.value = val;
    };

    window.onSelectGatePassShipment = function() {
      var val = document.getElementById('gpShipmentSelect').value;
      if (!val || !globalData || !globalData.pastOutgoingRecords) return;
      var record = globalData.pastOutgoingRecords.find(function(r) { return String(r.id) === String(val); });
      if (!record) return;

      document.getElementById('gpTo').value = record.avpName !== '-' ? record.avpName : '';
      document.getElementById('gpContact').value = record.clusterHeadContact || '';
      document.getElementById('gpDate').value = record.date || '';
      document.getElementById('gpCoRa').value = record.clusterHead || '';
      document.getElementById('gpRaContact').value = record.clusterHeadContact || '';
      document.getElementById('gpControlNo').value = record.controlNo !== '-' ? record.controlNo : '';
      document.getElementById('gpBranchCode').value = record.baseStation || '';
      document.getElementById('gpCluster').value = record.cluster !== '-' ? record.cluster : '';
      document.getElementById('gpDivision').value = record.division !== '-' ? record.division : '';
      document.getElementById('gpSurveyForm').value = '';
      document.getElementById('gpNote').value = record.notes || '';

      var completeAddr = '';
      if (record.cluster && globalData.regionDirectory) {
        var cInfo = globalData.regionDirectory[record.cluster] || globalData.regionDirectory[record.cluster.toLowerCase()];
        if (cInfo && cInfo.completeAddress) completeAddr = cInfo.completeAddress;
      }
      if (!completeAddr && record.avpName && globalData.avpDirectory) {
        var aInfo = globalData.avpDirectory[record.avpName] || globalData.avpDirectory[record.avpName.toLowerCase()];
        if (aInfo && aInfo.completeAddress) completeAddr = aInfo.completeAddress;
      }
      document.getElementById('gpAddress').value = completeAddr || (record.destination !== '-' ? record.destination : '');

      var items = record.items || {};
      document.getElementById('gpFaf').value = items['FAF'] || '';
      document.getElementById('gpFafBarmm').value = items['FAF Barmm'] || '';
      document.getElementById('gpPassbook').value = items['Passbook'] || '';
      document.getElementById('gpPassbookBarmm').value = items['Passbook Barmm'] || '';
      document.getElementById('gpGtrNew').value = items['GTR New'] || '';
      document.getElementById('gpGtr').value = items['GTR'] || '';
      document.getElementById('gpGtrBarmm').value = items['GTR Barmm'] || '';
      document.getElementById('gpDeskCal').value = items['Desk Calendar'] || '';
      document.getElementById('gpWallCal').value = items['Wall Calendar'] || '';
      document.getElementById('gpGuideBook').value = items['Guide Book'] || '';
      document.getElementById('gpEnrolment').value = items['Enrolment'] || '';
      document.getElementById('gpCoverage').value = items['Coverage'] || '';
      document.getElementById('gpPoster').value = items['Poster'] || '';

      var allFields = ['gpTo','gpContact','gpDate','gpCoRa','gpRaContact','gpAddress','gpControlNo','gpBranchCode','gpCluster','gpDivision','gpSurveyForm','gpNote','gpFaf','gpFafBarmm','gpPassbook','gpPassbookBarmm','gpGtrNew','gpGtr','gpGtrBarmm','gpDeskCal','gpWallCal','gpGuideBook','gpEnrolment','gpCoverage','gpPoster'];
      allFields.forEach(function(f) { syncGatePassCopies(f); });
    };

    window.clearGatePassForm = function() {
      var allInputs = document.querySelectorAll('#printableGatePassArea input');
      allInputs.forEach(function(inp) { inp.value = ''; });
      var today = new Date().toISOString().split('T')[0];
      document.getElementById('gpDate').value = today;
      syncGatePassCopies('gpDate');
    };

    // ── PAST OUTGOING MODAL ─────────────────────────────────────────────────

    window.openSearchPastModal = function() {
      document.getElementById('searchPastModal').classList.remove('hidden');
      document.getElementById('searchPastModal').classList.add('flex');
      document.getElementById('pastOutgoingSearch').value = '';
      renderPastOutgoingTable();
    };

    window.closeSearchPastModal = function() {
      document.getElementById('searchPastModal').classList.add('hidden');
      document.getElementById('searchPastModal').classList.remove('flex');
    };

    window.renderPastOutgoingTable = function() {
      if (!globalData || !globalData.pastOutgoingRecords) return;
      var query = (document.getElementById('pastOutgoingSearch').value || '').toLowerCase().trim();
      var records = globalData.pastOutgoingRecords.filter(function(r) {
        return [r.controlNo, r.avpName, r.division, r.destination, r.cluster, r.date]
          .some(function(v) { return v && v.toLowerCase().indexOf(query) !== -1; });
      });

      if (records.length === 0) {
        document.getElementById('pastOutgoingTableBody').innerHTML = '<tr><td colspan="7" class="py-6 text-center text-slate-400">No matching records found.</td></tr>';
        return;
      }

      document.getElementById('pastOutgoingTableBody').innerHTML = records.map(function(r) {
        return '<tr class="hover:bg-slate-50 transition border-b">' +
          '<td class="p-3 font-medium text-slate-600 border-r">' + r.date + '</td>' +
          '<td class="p-3 font-bold text-blue-600 border-r">' + r.controlNo + '</td>' +
          '<td class="p-3 font-medium text-slate-800 border-r">' + r.avpName + '</td>' +
          '<td class="p-3 text-slate-600 border-r">' + r.division + '</td>' +
          '<td class="p-3 text-slate-600 border-r">' + r.destination + '</td>' +
          '<td class="p-3 font-semibold text-slate-700 border-r">' + r.cluster + '</td>' +
          '<td class="p-3 text-center">' +
          '<button type="button" onclick="loadPastRecordIntoForm(' + r.id + ')" class="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-700 font-semibold rounded text-[11px] transition cursor-pointer">Edit</button>' +
          '</td></tr>';
      }).join('');
    };

    window.loadPastRecordIntoForm = function(recordId) {
      if (!globalData || !globalData.pastOutgoingRecords) return;
      var record = globalData.pastOutgoingRecords.find(function(r) { return r.id === recordId; });
      if (!record) return;

      editingRowIndex = record.id;
      closeSearchPastModal();
      openModal('OUTGOING');

      var btn = document.getElementById('submitBtn');
      btn.innerText = 'Update Outgoing Entry';
      btn.classList.remove('bg-blue-600', 'hover:bg-blue-700');
      btn.classList.add('bg-amber-500', 'hover:bg-amber-600');

      if (record.avpName && record.avpName !== '-') { document.getElementById('outAvp').value = record.avpName; window.onAvpSelect(); }
      if (record.cluster && record.cluster !== '-') { document.getElementById('outRegion').value = record.cluster; window.onRegionType(); }
      if (record.destination && record.destination !== '-') document.getElementById('outDestination').value = record.destination;
      if (record.controlNo  && record.controlNo  !== '-') document.getElementById('outControl').value = record.controlNo;
      if (record.operation)    document.getElementById('outOperation').value    = record.operation;
      if (record.clusterHead)  document.getElementById('outClusterHead').value  = record.clusterHead;
      if (record.clusterHeadContact) document.getElementById('outClusterContact').value = record.clusterHeadContact;
      if (record.baseStation)  document.getElementById('outBaseStation').value  = record.baseStation;
      if (record.notes)        document.getElementById('outNotes').value        = record.notes;

      document.querySelectorAll('#materialInputsGrid input').forEach(function(input) {
        var qty = record.items && record.items[input.dataset.mat];
        input.value = (qty && qty > 0) ? qty : '';
      });
    };

    // ── TRANSACTIONS MODAL ──────────────────────────────────────────────────

    window.openModal = function(type) {
      currentModalType = type;
      var isInc = (type === 'INCOMING');
      document.getElementById('modalTitle').innerText = isInc ? '+ Add Incoming Entry (Delivery)' : '↗ Add Outgoing Entry (Gate Pass / Dispatch)';
      
      document.getElementById('searchPastOutgoingLink').classList.toggle('hidden', isInc);
      document.getElementById('searchPastIncomingLink').classList.toggle('hidden', !isInc);
      
      document.getElementById('incomingFields').classList.toggle('hidden', !isInc);
      document.getElementById('outgoingFields').classList.toggle('hidden', isInc);
      document.getElementById('transactionModal').classList.remove('hidden');
      document.getElementById('transactionModal').classList.add('flex');
    };

    window.closeModal = function() {
      document.getElementById('transactionModal').classList.add('hidden');
      document.getElementById('transactionModal').classList.remove('flex');
      document.getElementById('transForm').reset();

      editingRowIndex = null;
      var btn = document.getElementById('submitBtn');
      btn.innerText = 'Save Entry';
      btn.classList.remove('bg-amber-500', 'hover:bg-amber-600', 'bg-emerald-600', 'hover:bg-emerald-700');
      btn.classList.add('bg-blue-600', 'hover:bg-blue-700');
    };

    window.submitTransaction = function(e) {
      e.preventDefault();
      var btn = document.getElementById('submitBtn');
      var isUpdate = (editingRowIndex !== null);
      btn.disabled = true;
      btn.innerText = isUpdate ? 'Updating...' : 'Saving...';

      var items = {};
      document.querySelectorAll('#materialInputsGrid input').forEach(function(input) {
        var val = parseFloat(input.value);
        if (!isNaN(val) && val > 0) items[input.dataset.mat] = val;
      });

      var formData = currentModalType === 'INCOMING'
        ? {
            date: document.getElementById('incDate').value,
            party: document.getElementById('incSupplier').value,
            drNumber: document.getElementById('incDr').value,
            items: items
          }
        : {
            date: document.getElementById('outDate').value,
            avpName: document.getElementById('outAvp').value,
            division: document.getElementById('outDivision').value,
            regionCluster: document.getElementById('outRegion').value,
            destination: document.getElementById('outDestination').value,
            controlNo: document.getElementById('outControl').value,
            operation: document.getElementById('outOperation').value,
            clusterHead: document.getElementById('outClusterHead').value,
            clusterHeadContact: document.getElementById('outClusterContact').value,
            baseStation: document.getElementById('outBaseStation').value,
            notes: document.getElementById('outNotes').value,
            items: items
          };

      var onDone = function() {
        btn.disabled = false;
        btn.innerText = isUpdate ? 'Update Entry' : 'Save Entry';
        closeModal();
        loadDashboard();
      };
      var onFail = function(err) {
        alert('Error: ' + err.message);
        btn.disabled = false;
        btn.innerText = isUpdate ? 'Update Entry' : 'Save Entry';
      };

      if (isUpdate) {
        google.script.run.withSuccessHandler(onDone).withFailureHandler(onFail).updateTransaction(currentModalType, editingRowIndex, formData);
      } else {
        google.script.run.withSuccessHandler(onDone).withFailureHandler(onFail).recordTransaction(currentModalType, formData);
      }
    };

    // ── FULL INVENTORY TABLE ────────────────────────────────────────────────

    window.renderFullInventoryTable = function() {
      if (!globalData) return;
      var query = (document.getElementById('inventorySearchInput').value || '').toLowerCase().trim();
      var filtered = globalData.inventoryList.filter(function(item) {
        return item.name.toLowerCase().indexOf(query) !== -1 || item.key.toLowerCase().indexOf(query) !== -1;
      });

      document.getElementById('fullInventoryTableBody').innerHTML = filtered.length === 0
        ? '<tr><td colspan="5" class="py-8 text-center text-slate-400">No matching materials found.</td></tr>'
        : filtered.map(function(item) {
            var isLow = item.stock <= 2000;
            var sc = isLow ? 'bg-[#fee2e2] text-[#ef4444]' : 'bg-[#dcfce7] text-[#16a34a]';
            return '<tr class="hover:bg-slate-50 transition border-b border-slate-100">' +
              '<td class="py-4 px-6 font-medium text-slate-800">' + item.name + '</td>' +
              '<td class="py-4 px-6 text-right font-medium text-slate-600">' + Number(item.inQty).toLocaleString() + '</td>' +
              '<td class="py-4 px-6 text-right font-medium text-slate-600">' + Number(item.outQty).toLocaleString() + '</td>' +
              '<td class="py-4 px-6 text-right font-bold text-slate-800">' + Number(item.stock).toLocaleString() + '</td>' +
              '<td class="py-4 px-6 text-center"><span class="inline-block px-3 py-1 text-[11px] font-semibold rounded-full ' + sc + '">' +
              (isLow ? 'Low Stock' : 'In Stock') + '</span></td></tr>';
          }).join('');
    };

    // ── CHART ───────────────────────────────────────────────────────────────

    function buildChart(filterMaterial) {
      var ctx = document.getElementById('transactionChart').getContext('2d');
      var monthlyData = globalData.monthlyData || {};
      var months = ['Jan 26','Feb 26','Mar 26','Apr 26','May 26','Jun 26','Jul 26','Aug 26'];

      var inArr = [], outArr = [];
      months.forEach(function(m) {
        var inT = 0, outT = 0;
        if (monthlyData[m]) {
          Object.keys(monthlyData[m]).forEach(function(k) {
            if (filterMaterial === 'ALL' || filterMaterial === k) {
              inT  += monthlyData[m][k].in  || 0;
              outT += monthlyData[m][k].out || 0;
            }
          });
        }
        inArr.push(inT); outArr.push(outT);
      });

      if (chartInstance) chartInstance.destroy();
      chartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: months,
          datasets: [
            { label: 'Incoming Units', data: inArr,  backgroundColor: '#10b981', borderRadius: 4 },
            { label: 'Outgoing Units', data: outArr, backgroundColor: '#2563eb', borderRadius: 4 }
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend: { position: 'top', labels: { boxWidth: 12, usePointStyle: true } } },
          scales: {
            y: { beginAtZero: true, grid: { color: '#f1f5f9' }, ticks: { callback: function(v) { return Number(v).toLocaleString(); } } },
            x: { grid: { display: false } }
          }
        }
      });
    }

    window.filterChart = function() {
      buildChart(document.getElementById('materialFilter').value);
    };
  </script>
</body>
</html>`;
}
