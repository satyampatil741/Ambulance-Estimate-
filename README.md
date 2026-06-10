<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>108 Fleet – Repair Estimate Builder</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
<style>
  :root {
    --bg: #0d1117;
    --surface: #161b22;
    --surface2: #1c2333;
    --border: #30363d;
    --accent: #f97316;
    --accent2: #fb923c;
    --green: #22c55e;
    --red: #ef4444;
    --blue: #3b82f6;
    --text: #e6edf3;
    --text2: #8b949e;
    --text3: #6e7681;
    --radius: 8px;
    --font: 'Segoe UI', system-ui, sans-serif;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    background: var(--bg);
    color: var(--text);
    font-family: var(--font);
    min-height: 100vh;
    font-size: 14px;
  }

  header {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 14px 24px;
    display: flex;
    align-items: center;
    gap: 14px;
    position: sticky;
    top: 0;
    z-index: 100;
  }
  .logo-badge {
    background: var(--accent);
    color: #fff;
    font-weight: 800;
    font-size: 13px;
    padding: 5px 10px;
    border-radius: 6px;
    letter-spacing: 0.5px;
  }
  header h1 { font-size: 15px; font-weight: 600; color: var(--text); }
  header p { font-size: 12px; color: var(--text2); margin-top: 1px; }
  .header-text { flex: 1; }

  .container {
    max-width: 1100px;
    margin: 0 auto;
    padding: 24px 20px 60px;
  }

  .card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    margin-bottom: 20px;
    overflow: hidden;
  }
  .card-header {
    background: var(--surface2);
    border-bottom: 1px solid var(--border);
    padding: 12px 18px;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .card-header .icon {
    width: 28px; height: 28px;
    background: var(--accent);
    border-radius: 6px;
    display: flex; align-items: center; justify-content: center;
    font-size: 14px;
  }
  .card-header h2 {
    font-size: 13px;
    font-weight: 600;
    color: var(--text);
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  .card-body { padding: 18px; }

  .form-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
    gap: 14px;
  }
  .form-grid.three { grid-template-columns: repeat(3, 1fr); }
  .form-grid.two   { grid-template-columns: repeat(2, 1fr); }

  .field { display: flex; flex-direction: column; gap: 5px; }
  .field.span2 { grid-column: span 2; }
  .field.span3 { grid-column: span 3; }

  label {
    font-size: 11px;
    font-weight: 600;
    color: var(--text2);
    text-transform: uppercase;
    letter-spacing: 0.4px;
  }
  label .req { color: var(--accent); }

  input, select, textarea {
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: 6px;
    color: var(--text);
    font-family: var(--font);
    font-size: 13px;
    padding: 8px 10px;
    transition: border-color 0.15s;
    width: 100%;
  }
  input:focus, select:focus, textarea:focus {
    outline: none;
    border-color: var(--accent);
    box-shadow: 0 0 0 3px rgba(249,115,22,0.12);
  }
  input::placeholder { color: var(--text3); }
  select option { background: var(--surface2); }

  /* spare vehicle row hidden by default */
  #spareVehicleRow { display: none; }

  .line-table-wrap { overflow-x: auto; }
  .line-table {
    width: 100%;
    border-collapse: collapse;
    min-width: 960px;
  }
  .line-table th {
    background: var(--surface2);
    color: var(--text2);
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    padding: 9px 10px;
    text-align: left;
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
  }
  .line-table td {
    padding: 6px 6px;
    border-bottom: 1px solid rgba(48,54,61,0.5);
    vertical-align: middle;
  }
  .line-table tr:hover td { background: rgba(255,255,255,0.02); }
  .line-table input, .line-table select {
    padding: 6px 8px;
    font-size: 12px;
    min-width: 60px;
  }
  .line-table .num-input { width: 90px; text-align: right; }
  .line-table .desc-input { min-width: 150px; }
  .line-table .partno-input { width: 110px; }
  .line-table .qty-input { width: 60px; text-align: center; }

  .computed {
    font-size: 12px;
    font-weight: 600;
    color: var(--accent2);
    text-align: right;
    padding: 0 8px;
    white-space: nowrap;
  }
  .row-num {
    color: var(--text3);
    font-size: 11px;
    font-weight: 600;
    text-align: center;
    width: 30px;
  }

  .btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 8px 16px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    border: none;
    transition: all 0.15s;
    font-family: var(--font);
    letter-spacing: 0.3px;
  }
  .btn-ghost { background: transparent; border: 1px dashed var(--border); color: var(--text2); }
  .btn-ghost:hover { border-color: var(--accent); color: var(--accent); }
  .btn-danger { background: rgba(239,68,68,0.1); color: var(--red); border: 1px solid rgba(239,68,68,0.3); }
  .btn-danger:hover { background: rgba(239,68,68,0.2); }
  .btn-primary { background: var(--accent); color: #fff; }
  .btn-primary:hover { background: var(--accent2); }
  .btn-success { background: var(--green); color: #fff; }
  .btn-success:hover { background: #16a34a; }
  .btn-blue { background: var(--blue); color: #fff; }
  .btn-blue:hover { background: #2563eb; }

  .table-actions {
    padding: 14px 18px;
    border-top: 1px solid var(--border);
    display: flex;
    gap: 10px;
    align-items: center;
  }

  .summary-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
  }
  .summary-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px 14px;
    border-radius: 6px;
    background: var(--surface2);
    border: 1px solid var(--border);
  }
  .summary-row .lbl { font-size: 12px; color: var(--text2); }
  .summary-row .val {
    font-size: 15px;
    font-weight: 700;
    color: var(--text);
    font-variant-numeric: tabular-nums;
  }
  .summary-row.total {
    background: rgba(249,115,22,0.08);
    border-color: var(--accent);
  }
  .summary-row.total .lbl { color: var(--accent); font-weight: 600; }
  .summary-row.total .val { color: var(--accent); font-size: 18px; }

  .gst-row {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 14px;
    border-radius: 6px;
    background: var(--surface2);
    border: 1px solid var(--border);
  }
  .gst-row label { text-transform: none; font-size: 12px; color: var(--text2); white-space: nowrap; }
  .gst-row input { width: 80px; }

  .export-bar {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 16px 20px;
    display: flex;
    align-items: center;
    gap: 14px;
    flex-wrap: wrap;
  }
  .export-bar .exp-label {
    font-size: 12px;
    color: var(--text2);
    flex: 1;
    min-width: 200px;
  }
  .export-bar .exp-label strong { color: var(--text); display: block; font-size: 13px; margin-bottom: 2px; }

  #toast {
    position: fixed;
    bottom: 24px;
    right: 24px;
    background: var(--green);
    color: #fff;
    padding: 12px 20px;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    opacity: 0;
    transform: translateY(10px);
    transition: all 0.3s;
    z-index: 9999;
    pointer-events: none;
  }
  #toast.show { opacity: 1; transform: translateY(0); }

  @media (max-width: 700px) {
    .form-grid.three { grid-template-columns: 1fr 1fr; }
    .summary-grid { grid-template-columns: 1fr; }
    .field.span2, .field.span3 { grid-column: span 1; }
  }
</style>
</head>
<body>

<header>
  <div class="logo-badge">108</div>
  <div class="header-text">
    <h1>Repair Estimate Builder</h1>
    <p>Maharashtra Medical Emergency Services – Fleet Management</p>
  </div>
</header>

<div class="container">

  <!-- JOB / VEHICLE DETAILS -->
  <div class="card">
    <div class="card-header">
      <div class="icon">📋</div>
      <h2>Job / Vehicle Details</h2>
    </div>
    <div class="card-body">
      <div class="form-grid three" style="row-gap:14px;">

        <div class="field">
          <label>Vehicle No. <span class="req">*</span></label>
          <input id="vehNo" type="text" placeholder="e.g. MH-12 AB 1234">
        </div>
        <div class="field">
          <label>Make <span class="req">*</span></label>
          <select id="vehMake">
            <option value="">– Select –</option>
            <option>Force Traveler</option>
            <option>AL Dost</option>
            <option>Other</option>
          </select>
        </div>
        <div class="field">
          <label>Model / Variant</label>
          <select id="vehModel">
            <option value="">– Select –</option>
            <option>BS-3</option>
            <option>BS-4</option>
            <option>BS-6</option>
            <option>BS-6+</option>
          </select>
        </div>

        <div class="field">
          <label>Chassis No.</label>
          <input id="chassisNo" type="text" placeholder="Optional">
        </div>
        <div class="field">
          <label>Cluster No. <span class="req">*</span></label>
          <select id="cluster">
            <option value="">– Select –</option>
            <option>C1</option><option>C2</option><option>C3</option><option>C4</option><option>C5</option>
            <option>C6</option><option>C7</option><option>C8</option><option>C9</option><option>C10</option>
            <option>C11</option><option>C12</option><option>C13</option><option>C14</option><option>C15</option>
            <option>C16</option><option>C17</option><option>C18</option><option>C19</option><option>C20</option>
            <option>C21</option><option>C22</option><option>C23</option><option>C24</option><option>C25</option>
            <option>C26</option><option>C27</option><option>C28</option><option>C29</option><option>C30</option>
            <option>C31</option><option>C32</option><option>C33</option><option>C34</option><option>C35</option>
            <option>C36</option><option>C37</option><option>C38</option><option>C39</option><option>C40</option>
            <option>C41</option><option>C42</option><option>C43</option><option>C44</option><option>C45</option>
          </select>
        </div>
        <div class="field">
          <label>Cluster Name</label>
          <input id="clusterName" type="text" placeholder="e.g. Pune Urban">
        </div>

        <div class="field">
          <label>Base Location</label>
          <input id="baseLocation" type="text" placeholder="e.g. Pune Station">
        </div>
        <div class="field">
          <label>Vehicle Status <span class="req">*</span></label>
          <select id="vehicleStatus" onchange="toggleSpare()">
            <option value="">– Select –</option>
            <option value="On Road">On Road</option>
            <option value="Off Road">Off Road</option>
          </select>
        </div>
        <div class="field" id="spareVehicleRow">
          <label>Spare Amb Vehicle No.</label>
          <input id="spareVehNo" type="text" placeholder="Spare vehicle no.">
        </div>

        <div class="field">
          <label>Odometer (km)</label>
          <input id="odometer" type="number" placeholder="e.g. 45000">
        </div>
        <div class="field">
          <label>Vendor / Workshop</label>
          <input id="vendor" type="text" placeholder="Workshop name">
        </div>
        <div class="field">
          <label>Estimate Date <span class="req">*</span></label>
          <input id="estDate" type="date">
        </div>

        <div class="field">
          <label>Estimate Ref. No.</label>
          <input id="estRef" type="text" placeholder="e.g. EST-2025-001">
        </div>
        <div class="field span2">
          <label>Nature of Repair / Job Description</label>
          <input id="jobDesc" type="text" placeholder="e.g. Scheduled PM Service + Brake System Overhaul">
        </div>

      </div>
    </div>
  </div>

  <!-- PARTS & LABOUR LINE ITEMS -->
  <div class="card">
    <div class="card-header">
      <div class="icon">🔧</div>
      <h2>Parts &amp; Labour Line Items</h2>
    </div>
    <div class="card-body" style="padding: 0;">
      <div class="line-table-wrap">
        <table class="line-table">
          <thead>
            <tr>
              <th style="width:34px;">#</th>
              <th style="min-width:160px;">Job / Activity Description <span style="color:var(--accent)">*</span></th>
              <th style="min-width:130px;">Part Description</th>
              <th style="width:110px;">Part No.</th>
              <th style="width:60px;">Qty</th>
              <th style="width:100px;">Part Rate (₹)</th>
              <th style="width:110px;">Part Amt (₹)</th>
              <th style="width:100px;">Labour (₹)</th>
              <th style="width:100px;">Part Tax (₹)</th>
              <th style="width:100px;">Labour Tax (₹)</th>
              <th style="width:110px;">Line Total (₹)</th>
              <th style="width:40px;"></th>
            </tr>
          </thead>
          <tbody id="lineBody"></tbody>
        </table>
      </div>
    </div>
    <div class="table-actions">
      <button class="btn btn-ghost" onclick="addRow()">＋ Add Row</button>
      <button class="btn btn-ghost" onclick="addRows(5)">＋ Add 5 Rows</button>
      <span style="color:var(--text3);font-size:11px;margin-left:8px;">Part Amt = Qty × Rate · Line Total = Part Amt + Labour + Part Tax + Labour Tax</span>
    </div>
  </div>

  <!-- LUBRICANTS & CONSUMABLES -->
  <div class="card">
    <div class="card-header">
      <div class="icon">🛢️</div>
      <h2>Lubricants &amp; Consumables</h2>
    </div>
    <div class="card-body" style="padding:0;">
      <div class="line-table-wrap">
        <table class="line-table">
          <thead>
            <tr>
              <th style="width:34px;">#</th>
              <th style="min-width:160px;">Item Description <span style="color:var(--accent)">*</span></th>
              <th style="width:80px;">Grade / Spec</th>
              <th style="width:70px;">Qty</th>
              <th style="width:80px;">Unit</th>
              <th style="width:110px;">Rate / Unit (₹)</th>
              <th style="width:120px;">Amount (₹)</th>
              <th style="width:40px;"></th>
            </tr>
          </thead>
          <tbody id="lubeBody"></tbody>
        </table>
      </div>
    </div>
    <div class="table-actions">
      <button class="btn btn-ghost" onclick="addLubeRow()">＋ Add Lube / Consumable</button>
    </div>
  </div>

  <!-- SUMMARY -->
  <div class="card">
    <div class="card-header">
      <div class="icon">💰</div>
      <h2>Estimate Summary</h2>
    </div>
    <div class="card-body">
      <div class="summary-grid">
        <div class="summary-row">
          <span class="lbl">Total Parts Value</span>
          <span class="val" id="sumParts">₹ 0.00</span>
        </div>
        <div class="summary-row">
          <span class="lbl">Total Labour Charges</span>
          <span class="val" id="sumLabour">₹ 0.00</span>
        </div>
        <div class="summary-row">
          <span class="lbl">Part Tax</span>
          <span class="val" id="sumPartTax">₹ 0.00</span>
        </div>
        <div class="summary-row">
          <span class="lbl">Labour Tax</span>
          <span class="val" id="sumLabourTax">₹ 0.00</span>
        </div>
        <div class="summary-row">
          <span class="lbl">Lubricants / Consumables</span>
          <span class="val" id="sumLube">₹ 0.00</span>
        </div>
        <div class="summary-row">
          <span class="lbl">Sub-Total (before GST)</span>
          <span class="val" id="sumSub">₹ 0.00</span>
        </div>
        <div class="gst-row">
          <label>GST (%)</label>
          <input id="gstPct" type="number" value="18" min="0" max="100" oninput="recalc()">
          <span style="color:var(--text2);font-size:12px;">applicable rate</span>
        </div>
        <div class="summary-row">
          <span class="lbl">GST Amount</span>
          <span class="val" id="sumGst">₹ 0.00</span>
        </div>
        <div class="summary-row total" style="grid-column: span 2;">
          <span class="lbl">GRAND TOTAL</span>
          <span class="val" id="sumGrand">₹ 0.00</span>
        </div>
      </div>
    </div>
  </div>

  <!-- EXPORT -->
  <div class="export-bar">
    <div class="exp-label">
      <strong>Ready to export?</strong>
      Download the estimate as a formatted Excel (.xlsx) — attach directly to your repair approval mail.
    </div>
    <button class="btn btn-blue" onclick="clearAll()">🗑 Clear All</button>
    <button class="btn btn-success" onclick="exportXLSX()">⬇ Download Estimate (.xlsx)</button>
  </div>

</div>

<div id="toast">✓ Estimate exported successfully!</div>

<script>
let rows = [];
let lubeRows = [];
let rowCounter = 0;
let lubeCounter = 0;

document.getElementById('estDate').value = new Date().toISOString().split('T')[0];
for (let i = 0; i < 5; i++) addRow();
addLubeRow();

function toggleSpare() {
  const val = document.getElementById('vehicleStatus').value;
  document.getElementById('spareVehicleRow').style.display = (val === 'Off Road') ? 'flex' : 'none';
}

function addRow() {
  rowCounter++;
  const id = rowCounter;
  rows.push(id);
  const tbody = document.getElementById('lineBody');
  const tr = document.createElement('tr');
  tr.id = 'row-' + id;
  tr.innerHTML = `
    <td class="row-num">${rows.length}</td>
    <td><input class="desc-input" type="text" placeholder="e.g. Engine Oil Change" oninput="recalc()"></td>
    <td><input class="desc-input" type="text" placeholder="e.g. Engine Oil 15W40" oninput="recalc()"></td>
    <td><input class="partno-input" type="text" placeholder="OEM / Ref No." oninput="recalc()"></td>
    <td><input class="qty-input" type="number" value="" min="0" placeholder="0" oninput="recalcRow(${id})"></td>
    <td><input class="num-input" type="number" value="" min="0" placeholder="0.00" oninput="recalcRow(${id})"></td>
    <td class="computed" id="partamt-${id}">–</td>
    <td><input class="num-input" type="number" value="" min="0" placeholder="0.00" oninput="recalcRow(${id})"></td>
    <td><input class="num-input" type="number" value="" min="0" placeholder="0.00" oninput="recalcRow(${id})"></td>
    <td><input class="num-input" type="number" value="" min="0" placeholder="0.00" oninput="recalcRow(${id})"></td>
    <td class="computed" id="linetotal-${id}">–</td>
    <td><button class="btn btn-danger" style="padding:4px 8px;font-size:11px;" onclick="removeRow(${id})">✕</button></td>
  `;
  tbody.appendChild(tr);
  renumberRows();
  recalc();
}

function addRows(n) { for (let i = 0; i < n; i++) addRow(); }

function removeRow(id) {
  const idx = rows.indexOf(id);
  if (idx > -1) rows.splice(idx, 1);
  const el = document.getElementById('row-' + id);
  if (el) el.remove();
  renumberRows();
  recalc();
}

function renumberRows() {
  document.querySelectorAll('#lineBody tr').forEach((tr, i) => {
    const c = tr.querySelector('.row-num');
    if (c) c.textContent = i + 1;
  });
}

function addLubeRow() {
  lubeCounter++;
  const id = lubeCounter;
  lubeRows.push(id);
  const tbody = document.getElementById('lubeBody');
  const tr = document.createElement('tr');
  tr.id = 'luberow-' + id;
  tr.innerHTML = `
    <td class="row-num">${lubeRows.length}</td>
    <td><input class="desc-input" type="text" placeholder="e.g. Engine Oil" oninput="recalc()"></td>
    <td><input style="width:90px;" type="text" placeholder="e.g. 15W40" oninput="recalc()"></td>
    <td><input class="qty-input" type="number" min="0" placeholder="0" oninput="recalcLube(${id})"></td>
    <td>
      <select style="width:80px;" onchange="recalc()">
        <option>Ltr</option><option>ml</option><option>Kg</option><option>No.</option><option>Set</option>
      </select>
    </td>
    <td><input class="num-input" type="number" min="0" placeholder="0.00" oninput="recalcLube(${id})"></td>
    <td class="computed" id="lubeamt-${id}">–</td>
    <td><button class="btn btn-danger" style="padding:4px 8px;font-size:11px;" onclick="removeLubeRow(${id})">✕</button></td>
  `;
  tbody.appendChild(tr);
  renumberLubeRows();
  recalc();
}

function removeLubeRow(id) {
  const idx = lubeRows.indexOf(id);
  if (idx > -1) lubeRows.splice(idx, 1);
  const el = document.getElementById('luberow-' + id);
  if (el) el.remove();
  renumberLubeRows();
  recalc();
}

function renumberLubeRows() {
  document.querySelectorAll('#lubeBody tr').forEach((tr, i) => {
    const c = tr.querySelector('.row-num');
    if (c) c.textContent = i + 1;
  });
}

function recalcRow(id) {
  const tr = document.getElementById('row-' + id);
  if (!tr) return;
  const inputs = tr.querySelectorAll('input[type="number"]');
  const qty      = parseFloat(inputs[0].value) || 0;
  const rate     = parseFloat(inputs[1].value) || 0;
  const labour   = parseFloat(inputs[2].value) || 0;
  const partTax  = parseFloat(inputs[3].value) || 0;
  const labTax   = parseFloat(inputs[4].value) || 0;
  const partAmt  = qty * rate;
  const lineTotal = partAmt + labour + partTax + labTax;
  document.getElementById('partamt-' + id).textContent   = partAmt > 0 ? '₹ ' + fmt(partAmt) : '–';
  document.getElementById('linetotal-' + id).textContent = lineTotal > 0 ? '₹ ' + fmt(lineTotal) : '–';
  recalcSummary();
}

function recalcLube(id) {
  const tr = document.getElementById('luberow-' + id);
  if (!tr) return;
  const nums = tr.querySelectorAll('input[type="number"]');
  const qty  = parseFloat(nums[0].value) || 0;
  const rate = parseFloat(nums[1].value) || 0;
  const amt  = qty * rate;
  document.getElementById('lubeamt-' + id).textContent = amt > 0 ? '₹ ' + fmt(amt) : '–';
  recalcSummary();
}

function recalc() {
  rows.forEach(id => recalcRow(id));
  lubeRows.forEach(id => recalcLube(id));
  recalcSummary();
}

function recalcSummary() {
  let totalParts = 0, totalLabour = 0, totalPartTax = 0, totalLabTax = 0, totalLube = 0;

  rows.forEach(id => {
    const tr = document.getElementById('row-' + id);
    if (!tr) return;
    const inputs = tr.querySelectorAll('input[type="number"]');
    const qty     = parseFloat(inputs[0].value) || 0;
    const rate    = parseFloat(inputs[1].value) || 0;
    const labour  = parseFloat(inputs[2].value) || 0;
    const partTax = parseFloat(inputs[3].value) || 0;
    const labTax  = parseFloat(inputs[4].value) || 0;
    totalParts   += qty * rate;
    totalLabour  += labour;
    totalPartTax += partTax;
    totalLabTax  += labTax;
  });

  lubeRows.forEach(id => {
    const tr = document.getElementById('luberow-' + id);
    if (!tr) return;
    const nums = tr.querySelectorAll('input[type="number"]');
    totalLube += (parseFloat(nums[0].value) || 0) * (parseFloat(nums[1].value) || 0);
  });

  const sub    = totalParts + totalLabour + totalPartTax + totalLabTax + totalLube;
  const gstPct = parseFloat(document.getElementById('gstPct').value) || 0;
  const gstAmt = sub * gstPct / 100;
  const grand  = sub + gstAmt;

  setText('sumParts',    '₹ ' + fmt(totalParts));
  setText('sumLabour',   '₹ ' + fmt(totalLabour));
  setText('sumPartTax',  '₹ ' + fmt(totalPartTax));
  setText('sumLabourTax','₹ ' + fmt(totalLabTax));
  setText('sumLube',     '₹ ' + fmt(totalLube));
  setText('sumSub',      '₹ ' + fmt(sub));
  setText('sumGst',      '₹ ' + fmt(gstAmt));
  setText('sumGrand',    '₹ ' + fmt(grand));
}

function setText(id, v) { document.getElementById(id).textContent = v; }
function fmt(n) { return n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

function clearAll() {
  if (!confirm('Clear all line items and reset the form?')) return;
  document.getElementById('lineBody').innerHTML = '';
  document.getElementById('lubeBody').innerHTML = '';
  rows = []; lubeRows = [];
  for (let i = 0; i < 5; i++) addRow();
  addLubeRow();
  ['vehNo','vehMake','vehModel','chassisNo','cluster','clusterName','baseLocation',
   'vehicleStatus','spareVehNo','odometer','vendor','estDate','estRef','jobDesc'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    if (el.tagName === 'INPUT') el.value = id === 'estDate' ? new Date().toISOString().split('T')[0] : '';
    else el.value = '';
  });
  document.getElementById('spareVehicleRow').style.display = 'none';
}

// ── EXPORT XLSX matching sample_estimate.xlsx format ──────────────────────────
// ── PROFESSIONAL STYLED EXPORT ────────────────────────────────────────────────
function exportXLSX() {

  /* ── 1. Collect form values ─────────────────────────────────────── */
  const vehNo       = document.getElementById('vehNo').value       || '';
  const make        = document.getElementById('vehMake').value      || '';
  const model       = document.getElementById('vehModel').value     || '';
  const chassis     = document.getElementById('chassisNo').value    || '';
  const clusterNo   = document.getElementById('cluster').value      || '';
  const clusterName = document.getElementById('clusterName').value  || '';
  const baseLoc     = document.getElementById('baseLocation').value || '';
  const vehStatus   = document.getElementById('vehicleStatus').value|| '';
  const spareVeh    = document.getElementById('spareVehNo').value   || '';
  const odo         = document.getElementById('odometer').value     || '';
  const vendor      = document.getElementById('vendor').value       || '';
  const estDate     = document.getElementById('estDate').value      || '';
  const estRef      = document.getElementById('estRef').value       || '';
  const job         = document.getElementById('jobDesc').value      || '';
  const gstPct      = parseFloat(document.getElementById('gstPct').value) || 0;

  const clusterDisplay = clusterNo + (clusterName ? ' \u2013 ' + clusterName : '');
  const makeModel      = make + (model ? ' ' + model : '');
  const statusDisplay  = vehStatus + (spareVeh ? ' | Spare Amb: ' + spareVeh : '');

  /* ── 2. Collect line data ───────────────────────────────────────── */
  let totParts = 0, totLabour = 0, totPTax = 0, totLTax = 0;
  let hasPTax = false, hasLTax = false;
  const pRows = [];

  rows.forEach(id => {
    const tr = document.getElementById('row-' + id); if (!tr) return;
    const inp = tr.querySelectorAll('input');
    const num = tr.querySelectorAll('input[type="number"]');
    const jobAct=inp[0].value||'', partDesc=inp[1].value||'', partNo=inp[2].value||'';
    const qty=parseFloat(num[0].value)||0, rate=parseFloat(num[1].value)||0;
    const labour=parseFloat(num[2].value)||0;
    const pTax=parseFloat(num[3].value)||0, lTax=parseFloat(num[4].value)||0;
    if (!jobAct && !partDesc && qty===0 && rate===0 && labour===0) return;
    if (pTax > 0) hasPTax = true;
    if (lTax > 0) hasLTax = true;
    const pAmt=qty*rate, lt=pAmt+labour+pTax+lTax;
    totParts+=pAmt; totLabour+=labour; totPTax+=pTax; totLTax+=lTax;
    pRows.push({n:pRows.length+1, jobAct, partDesc, partNo, qty, rate, pAmt, labour, pTax, lTax, lt});
  });

  let totLube = 0;
  const lRows = [];
  lubeRows.forEach(id => {
    const tr = document.getElementById('luberow-' + id); if (!tr) return;
    const inp=tr.querySelectorAll('input'), sel=tr.querySelectorAll('select');
    const desc=inp[0].value||'', grade=inp[1].value||'';
    const qty=parseFloat(inp[2].value)||0, unit=sel[0]?sel[0].value:'', rate=parseFloat(inp[3].value)||0;
    if (!desc && qty===0) return;
    const amt=qty*rate; totLube+=amt;
    lRows.push({n:lRows.length+1, desc, grade, qty, unit, rate, amt});
  });

  const sub    = totParts+totLabour+totPTax+totLTax+totLube;
  const gstAmt = sub*gstPct/100;
  const grand  = sub+gstAmt;

  /* ── 3. Style palette ───────────────────────────────────────────── */
  // Dynamic columns: always A-H base, + Part Tax col if hasPTax, + Labour Tax col if hasLTax, then Line Total
  // Col indices: 0=Sr, 1=JobAct, 2=PartDesc, 3=PartNo, 4=Qty, 5=Rate, 6=PartAmt, 7=Labour, [8=PTax?], [9=LTax?], last=LineTotal
  const extraCols = (hasPTax ? 1 : 0) + (hasLTax ? 1 : 0);
  const NC = 9 + extraCols; // total columns

  function ptaxCol() { return hasPTax ? 8 : -1; }
  function ltaxCol() { return hasLTax ? 8 + (hasPTax?1:0) : -1; }
  function ltotalCol() { return 8 + extraCols; }

  const NAVY   = 'FF1B3A5C';
  const NAVY2  = 'FF22527A';
  const ORG    = 'FFD9600A';
  const ORGL   = 'FFFFF3E6';
  const WHITE  = 'FFFFFFFF';
  const GREY1  = 'FFF2F5F8';
  const GREY2  = 'FFE2E8EE';
  const GHDR   = 'FFD4DCE6';
  const GREENL = 'FFE6F4EA';
  const GREEND = 'FF1E6E2E';
  const BORD   = 'FFAABBCC';
  const TXTD   = 'FF111111';
  const TXTN   = 'FF1B3A5C';
  const TXTG   = 'FF666666';
  const TXTW   = 'FFFFFFFF';

  const bT = (c) => ({ style:'thin',   color:{ rgb: c||BORD } });
  const bM = (c) => ({ style:'medium', color:{ rgb: c||NAVY } });
  const bT2 = { top:bT(), bottom:bT(), left:bT(), right:bT() };
  const bM2 = { top:bM(), bottom:bM(), left:bM(), right:bM() };
  const bSect= { top:bM(ORG), bottom:bM(ORG), left:bM(), right:bM() };

  /* ── Cell constructors ── */
  function C(v, s) {
    const t = (typeof v === 'number') ? 'n' : 's';
    return { v, t, s };
  }
  function Cn(v, s) { return { v: v||0, t:'n', s }; }

  function fill(rgb)  { return { fgColor:{ rgb } }; }
  function font(opts) { return { name:'Arial', ...opts }; }
  function al(h, wrap) { return { horizontal: h||'left', vertical:'center', wrapText: !!wrap }; }

  function empty(bg) {
    return C('', { fill:fill(bg||WHITE), border:bT2 });
  }
  function blankLine() {
    return Array.from({length:NC}, () => C('', { fill:fill(WHITE) }));
  }

  /* Title bar */
  function cTitle(v) {
    return C(v, { font:font({bold:true, sz:14, color:{rgb:TXTW}}), fill:fill(NAVY), alignment:al('center'), border:bM2 });
  }
  function cTitleFill() { return C('', { fill:fill(NAVY), border:bM2 }); }

  /* Orange subtitle bar */
  function cSub(v) {
    return C(v, { font:font({bold:true, sz:11, color:{rgb:TXTW}}), fill:fill(ORG), alignment:al('center'), border:bM2 });
  }
  function cSubFill() { return C('', { fill:fill(ORG), border:bM2 }); }

  /* Info block: label cell */
  function cLbl(v) {
    return C(v, { font:font({bold:true, sz:9, color:{rgb:TXTN}}), fill:fill(GHDR), alignment:al('left',true), border:bT2 });
  }
  /* Info block: value cell */
  function cVal(v) {
    return C(v, { font:font({sz:9, color:{rgb:TXTD}}), fill:fill(WHITE), alignment:al('left',true), border:bT2 });
  }
  function cValFill() { return C('', { fill:fill(WHITE), border:bT2 }); }
  function cLblFill() { return C('', { fill:fill(GHDR),  border:bT2 }); }

  /* Section header */
  function cSect(v) {
    return C(v, { font:font({bold:true, sz:10, color:{rgb:TXTN}}), fill:fill(ORGL), alignment:al('left'), border:bSect });
  }
  function cSectFill() { return C('', { fill:fill(ORGL), border:bSect }); }

  /* Column header */
  function cHdr(v) {
    return C(v, { font:font({bold:true, sz:9, color:{rgb:TXTW}}), fill:fill(NAVY2), alignment:{horizontal:'center',vertical:'center',wrapText:true}, border:bT2 });
  }

  /* Data: text */
  function cD(v, shade, h) {
    return C(v, { font:font({sz:9, color:{rgb:TXTD}}), fill:fill(shade?GREY1:WHITE), alignment:al(h||'left',true), border:bT2 });
  }
  /* Data: number */
  function cN(v, shade) {
    const hasVal = v && v !== 0;
    return Cn(v, {
      font:font({sz:9, color:{rgb: hasVal ? TXTD : 'FFCCCCCC'}}),
      fill:fill(shade?GREY1:WHITE),
      alignment:{horizontal:'right',vertical:'center'},
      border:bT2,
      numFmt:'#,##0.00'
    });
  }
  /* Data: serial number */
  function cSr(v, shade) {
    return C(v, { font:font({sz:9, color:{rgb:TXTG}}), fill:fill(shade?GREY2:GREY1), alignment:al('center'), border:bT2 });
  }

  /* Summary label */
  function cSumLbl(v, grand) {
    return C(v, {
      font:font({bold:true, sz: grand?11:9, color:{rgb: grand?TXTW:TXTN}}),
      fill:fill(grand?GREEND:GREENL),
      alignment:{horizontal:'right',vertical:'center'},
      border: grand?bM2:bT2
    });
  }
  function cSumLblFill(grand) {
    return C('', { fill:fill(grand?GREEND:GREENL), border: grand?bM2:bT2 });
  }
  /* Summary value */
  function cSumVal(v, grand) {
    return Cn(v, {
      font:font({bold:true, sz: grand?11:9, color:{rgb: grand?TXTW:TXTN}}),
      fill:fill(grand?GREEND:GREENL),
      alignment:{horizontal:'right',vertical:'center'},
      border: grand?bM2:bT2,
      numFmt:'#,##0.00'
    });
  }

  /* Terms */
  function cTermsHdr(v) {
    return C(v, { font:font({bold:true, sz:9, color:{rgb:TXTW}}), fill:fill(NAVY), alignment:al('left'), border:bT2 });
  }
  function cTermsHdrFill() { return C('', { fill:fill(NAVY), border:bT2 }); }
  function cTermsLine(v) {
    return C(v, { font:font({sz:8, italic:true, color:{rgb:TXTG}}), fill:fill('FFF9FAFB'), alignment:al('left',true), border:bT2 });
  }
  function cTermsFill() { return C('', { fill:fill('FFF9FAFB'), border:bT2 }); }

  /* ── 4. Build rows ──────────────────────────────────────────────── */
  const wsData = [];
  const merges = [];
  let R = 0;

  function mFull()         { merges.push({ s:{r:R,c:0}, e:{r:R,c:NC-1} }); }
  function mRange(c1, c2)  { merges.push({ s:{r:R,c:c1}, e:{r:R,c:c2} }); }

  /* ROW 0: Main Title */
  wsData.push([cTitle('108 \u2013 MAHARASHTRA MEDICAL EMERGENCY SERVICES'), ...Array(NC-1).fill(cTitleFill())]);
  mFull(); R++;

  /* ROW 1: Subtitle */
  wsData.push([cSub('VEHICLE REPAIR ESTIMATE'), ...Array(NC-1).fill(cSubFill())]);
  mFull(); R++;

  /* ROW 2: Spacer */
  wsData.push(blankLine()); R++;

  /* ROWS 3-7: Info block (two-pane layout) */
  // Left pane: cols 0-1 label, 2-3 value | gap col 4 | Right pane: cols 5-6 label, 7-(NC-1) value
  const infoRows = [
    ['Vehicle No.',    vehNo||'N/A',         'Estimate Ref.',     estRef||'N/A'],
    ['Make / Model',   makeModel||'N/A',     'Date',              estDate||'N/A'],
    ['Chassis No.',    chassis||'N/A',       'Cluster No.',       clusterDisplay||'N/A'],
    ['Odometer (km)',  odo||'N/A',           'Vendor / Workshop', vendor||'N/A'],
    ['Base Location',  baseLoc||'N/A',       'Vehicle Status',    statusDisplay||'N/A'],
  ];
  infoRows.forEach(([l1,v1,l2,v2]) => {
    const row = [];
    // left label: 0-1
    row.push(cLbl(l1)); row.push(cLblFill());
    // left value: 2-3
    row.push(cVal(v1)); row.push(cValFill());
    // gap: 4
    row.push(empty(WHITE));
    // right label: 5-6
    row.push(cLbl(l2)); row.push(cLblFill());
    // right value: 7 to NC-1
    row.push(cVal(v2));
    for (let x=8; x<NC; x++) row.push(cValFill());
    wsData.push(row);
    mRange(0,1); mRange(2,3); mRange(5,6); mRange(7,NC-1);
    R++;
  });

  /* ROW 8: Job Description */
  wsData.push([
    cLbl('Job Description'), cLblFill(),
    cVal(job||'N/A'), ...Array(NC-3).fill(cValFill()),
    cValFill()
  ]);
  mRange(0,1); mRange(2,NC-1); R++;

  /* ROW 9: Spacer */
  wsData.push(blankLine()); R++;

  /* ROW 10: Parts section header */
  wsData.push([cSect('PARTS & LABOUR LINE ITEMS'), ...Array(NC-1).fill(cSectFill())]);
  mFull(); R++;

  /* ROW 11: Column headers — dynamic */
  const hdrCells = [
    cHdr('Sr.'),
    cHdr('Job / Activity\nDescription'),
    cHdr('Part\nDescription'),
    cHdr('Part No.'),
    cHdr('Qty'),
    cHdr('Part Rate\n(\u20B9)'),
    cHdr('Part Amt\n(\u20B9)'),
    cHdr('Labour\n(\u20B9)'),
  ];
  if (hasPTax) hdrCells.push(cHdr('Part Tax\n(\u20B9)'));
  if (hasLTax) hdrCells.push(cHdr('Labour Tax\n(\u20B9)'));
  hdrCells.push(cHdr('Line Total\n(\u20B9)'));
  wsData.push(hdrCells); R++;

  /* Parts data rows */
  const minP = Math.max(pRows.length, 10);
  for (let i=0; i<minP; i++) {
    const d=pRows[i], sh=i%2===1;
    const row = [];
    if (d) {
      row.push(cSr(d.n, sh));
      row.push(cD(d.jobAct, sh));
      row.push(cD(d.partDesc, sh));
      row.push(cD(d.partNo, sh, 'center'));
      row.push(cN(d.qty, sh));
      row.push(cN(d.rate, sh));
      row.push(cN(d.pAmt, sh));
      row.push(cN(d.labour, sh));
      if (hasPTax) row.push(cN(d.pTax, sh));
      if (hasLTax) row.push(cN(d.lTax, sh));
      row.push(cN(d.lt, sh));
    } else {
      row.push(cSr(i+1, sh));
      for (let x=1; x<NC; x++) row.push(empty(sh?GREY1:WHITE));
    }
    wsData.push(row); R++;
  }

  /* Spacer */
  wsData.push(blankLine()); R++;

  /* Lube section header */
  wsData.push([cSect('LUBRICANTS & CONSUMABLES'), ...Array(NC-1).fill(cSectFill())]);
  mFull(); R++;

  /* Lube column headers */
  const lubeHdrCells = [
    cHdr('Sr.'), cHdr('Item Description'), cHdr('Grade /\nSpec'),
    cHdr('Qty'), cHdr('Unit'), cHdr('Rate / Unit\n(\u20B9)'), cHdr('Amount\n(\u20B9)'),
  ];
  for (let x=7; x<NC; x++) lubeHdrCells.push(empty(NAVY2));
  wsData.push(lubeHdrCells); R++;

  /* Lube data rows */
  const minL = Math.max(lRows.length, 4);
  for (let i=0; i<minL; i++) {
    const d=lRows[i], sh=i%2===1;
    const row = [];
    if (d) {
      row.push(cSr(d.n, sh));
      row.push(cD(d.desc, sh));
      row.push(cD(d.grade, sh, 'center'));
      row.push(cN(d.qty, sh));
      row.push(cD(d.unit, sh, 'center'));
      row.push(cN(d.rate, sh));
      row.push(cN(d.amt, sh));
    } else {
      row.push(cSr(i+1, sh));
      for (let x=1; x<7; x++) row.push(empty(sh?GREY1:WHITE));
    }
    for (let x=7; x<NC; x++) row.push(empty(sh?GREY1:WHITE));
    wsData.push(row); R++;
  }

  /* Spacer */
  wsData.push(blankLine()); R++;

  /* ── 5. Summary block ──────────────────────────────────────────── */
  // Summary uses right half of sheet: label spans left side, value in last col
  const sumSplit = Math.floor(NC * 0.55); // label from col sumSplit to NC-2, value = NC-1

  function pushSummaryRow(label, value, isGrand) {
    const row = [];
    // left side empty
    for (let x=0; x<sumSplit; x++) row.push(empty(WHITE));
    // label cell
    row.push(cSumLbl(label, isGrand));
    // label fill cells
    for (let x=sumSplit+1; x<NC-1; x++) row.push(cSumLblFill(isGrand));
    // value cell
    row.push(cSumVal(value, isGrand));
    wsData.push(row);
    mRange(sumSplit, NC-2);
    R++;
  }

  pushSummaryRow('Parts Value (\u20B9)',              totParts,  false);
  pushSummaryRow('Labour Charges (\u20B9)',            totLabour, false);
  if (hasPTax) pushSummaryRow('Part Tax (\u20B9)',    totPTax,   false);
  if (hasLTax) pushSummaryRow('Labour Tax (\u20B9)',  totLTax,   false);
  pushSummaryRow('Lubricants / Consumables (\u20B9)', totLube,   false);
  pushSummaryRow('Sub-Total (\u20B9)',                 sub,       false);
  if (gstPct > 0) pushSummaryRow('GST @ '+gstPct+'% (\u20B9)', gstAmt, false);
  pushSummaryRow('GRAND TOTAL (\u20B9)',               grand,     true);

  /* Spacer */
  wsData.push(blankLine()); R++;

  /* ── 6. Terms ──────────────────────────────────────────────────── */
  wsData.push([cTermsHdr('TERMS & CONDITIONS'), ...Array(NC-1).fill(cTermsHdrFill())]);
  mFull(); R++;

  const termsLines = [
    '1.  Part costs and labour charges are shown separately and shall not be clubbed under a single description.',
    '2.  All lubricants / consumables are listed separately with quantity, unit, and rate details.',
    '3.  Vendor is requested to provide revised quotations strictly in this format for all future repair estimates.',
    '4.  Lube prices and brake-pad labour / part split-up to be verified against OEM Part Catalogue before finalising.',
  ];
  termsLines.forEach(t => {
    wsData.push([cTermsLine(t), ...Array(NC-1).fill(cTermsFill())]);
    mFull(); R++;
  });

  /* ── 7. Assemble workbook ───────────────────────────────────────── */
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!merges'] = merges;

  // Column widths — base 9 cols + optional tax cols
  const colWidths = [
    {wch:5},   // Sr
    {wch:28},  // Job desc
    {wch:20},  // Part desc
    {wch:12},  // Part No
    {wch:7},   // Qty
    {wch:12},  // Rate
    {wch:12},  // Part Amt
    {wch:12},  // Labour
  ];
  if (hasPTax) colWidths.push({wch:11}); // Part Tax
  if (hasLTax) colWidths.push({wch:11}); // Labour Tax
  colWidths.push({wch:13}); // Line Total
  ws['!cols'] = colWidths;

  // Row heights
  const rh = {};
  rh[0] = {hpt:32}; // title
  rh[1] = {hpt:22}; // subtitle
  rh[11]= {hpt:36}; // col headers (wrapped text)
  ws['!rows'] = rh;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Estimate');

  const fname = 'RepairEstimate_' + (vehNo||'Vehicle').replace(/\s/g,'') + '_' + (estDate||'Date') + '.xlsx';
  XLSX.writeFile(wb, fname);

  const t = document.getElementById('toast');
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3000);
}


</script>
</body>
</html>
