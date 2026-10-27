// ============================================================
// Workshop Tools - Apps Script (v18)
// v18: families. Drill bits, sanding paper and similar items, each on its own tab named after the family.
// Accessories go out and come back. Consumables are used up. The Log now records quantity (I), unit cost (J) and total (K).
// v17: a small "ping" request the scanner sends every few minutes while staff are signed in, to keep the server responsive.
// v16: tool lookups return dueDate (yyyy-MM-dd), used by the scanner's Undo.
// v15: staff pages header renamed to Workshop Tools. Otherwise the same as v14.
// v14: students can see the workshop tools signed out to them (student number only).
// v12: new QR labels contain only the tool ID (e.g. TL-042), which makes them much simpler to scan.
// Old labels with the full web-app link still work.
// Students sheet: A = StudentNo, B = Name, C = Surname, D = Year.
// One complete file. Replace everything in Extensions > Apps Script with this.
// Contains: v7 staff pages and QR links (unchanged) + JSON API for CPUT-Workshop-Scanner
// (sign-in, lookup, check out, check in, and tools out to a student).
// Then: Deploy > Manage deployments > edit existing > New version.
// ============================================================

const TOOLS = 'Tools', STUDENTS = 'Students', LOG = 'Log', STAFF = 'Staff';
const TEAL = '#00897b', TEAL_DARK = '#00473f', TEAL_LIGHT = '#b2dfdb',
      TEAL_MID = '#80cbc4', TEAL_PALE = '#e0f2f1';
const BATCH_SECONDS = 1800;  // a student session lasts 30 minutes (v7 pages)
const SESSION_SECONDS = 600; // scanner sign-in lasts 10 minutes of inactivity

// ---------- Web entry point ----------

function doGet(e) {
  if (e && e.parameter && e.parameter.api) return apiRespond(e);
  const p = (e && e.parameter) || {};
  let body;
  if (p.view === 'students') body = listPage();
  else if (p.page === 'start') body = startPage();
  else if (p.page === 'add') body = addPage();
  else if (p.tool) body = staffPage(p.tool);
  else body = '<p>Scan a tool QR code, or open the students list.</p>';
  return HtmlService.createHtmlOutput(
    '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<body style="font-family:sans-serif;padding:12px;font-size:18px;background:' + TEAL_PALE +
    ';color:#000;margin:0">' +
    '<div style="background:' + TEAL + ';color:#fff;padding:10px 14px;border-radius:8px;margin-bottom:8px;' +
    'font-weight:bold">Workshop Tools</div>' + body + '</body>');
}

// ---------- Start a student (batch) page ----------

function startPage() {
  return '<h2 style="color:' + TEAL_DARK + ';margin:12px 0 0">Start a student</h2>' +
    '<p>Enter the student once. Every tool you add in the next 30 minutes goes to this student.</p>' +
    field('pin', 'Your PIN', 'password', 'inputmode="numeric"') +
    field('stu', 'Student number', 'text', 'inputmode="numeric"') +
    dateField() +
    button('Set student', "run('startBatch')", TEAL) +
    button('End student session', "run('endBatch')", TEAL_DARK) +
    startScript();
}

function startScript() {
  return '<script>function run(fn){' +
    'var v=function(i){var el=document.getElementById(i);return el?el.value:""};' +
    'if(!v("pin")){alert("Enter your PIN.");return;}' +
    'if(fn==="startBatch"&&(!v("stu")||!v("ret"))){alert("Enter the student number and the return date.");return;}' +
    'google.script.run.withSuccessHandler(function(m){document.body.innerHTML=m})' +
    '.withFailureHandler(function(err){alert(err.message)})' +
    '[fn](v("pin"),v("stu"),v("ret"));}</script>';
}

// ---------- Add several tools by category (during a student session) ----------

function addPage() {
  return '<h2 style="color:' + TEAL_DARK + ';margin:12px 0 0">Add tools for this student</h2>' +
    '<p>Pick a category, then tick the tools you have fetched.</p>' +
    field('pin', 'Your PIN', 'password', 'inputmode="numeric"') +
    '<label for="cat" style="display:block;margin:16px 0 6px;font-size:18px;font-weight:bold;color:' +
    TEAL_DARK + '">Category</label>' +
    '<select id="cat" style="display:block;width:100%;height:52px;font-size:18px;padding:8px;box-sizing:border-box;' +
    'color:#000;background:#fff;border:2px solid ' + TEAL + ';border-radius:6px">' + categoryOptions() + '</select>' +
    button('Show available tools', 'showList()', TEAL) +
    '<p style="margin-top:20px">' + navLink('start', 'Start a new student instead') + '</p>' +
    addScript();
}

function addScript() {
  return '<script>var PIN_SAVED="";' +
    'function showList(){var pin=document.getElementById("pin").value;' +
    'var cat=document.getElementById("cat").value;' +
    'if(!pin){alert("Enter your PIN.");return;}PIN_SAVED=pin;' +
    'google.script.run.withSuccessHandler(function(m){document.body.innerHTML=m})' +
    '.withFailureHandler(function(err){alert(err.message)}).availableList(pin,cat);}' +
    'function checkoutSelected(){var boxes=document.querySelectorAll("input.pick:checked");' +
    'var ids=Array.prototype.map.call(boxes,function(c){return c.value});' +
    'if(!ids.length){alert("Tick at least one tool.");return;}' +
    'google.script.run.withSuccessHandler(function(m){document.body.innerHTML=m})' +
    '.withFailureHandler(function(err){alert(err.message)}).checkoutBatch(PIN_SAVED,ids);}</script>';
}

function categoryOptions() {
  const rows = SpreadsheetApp.getActive().getSheetByName(TOOLS).getDataRange().getValues();
  const cats = [...new Set(rows.slice(1).map(r => String(r[2]).trim()).filter(c => c))].sort();
  return cats.map(c => '<option value="' + esc(c) + '">' + esc(c) + '</option>').join('');
}

// Called from the browser: lists the available tools in one category.
function availableList(pin, category) {
  const staff = findStaff(pin);
  if (!staff) throw new Error('Wrong PIN');
  const b = getBatch(staff.id);
  if (!b) throw new Error('No student set. Start a new student first.');

  const rows = SpreadsheetApp.getActive().getSheetByName(TOOLS).getDataRange().getValues();
  const avail = rows.slice(1).filter(r =>
    String(r[0]).trim() !== '' && r[4] !== 'Out' && String(r[2]).trim() === String(category).trim());

  const back = '<p style="margin-top:20px">' + navLink('add', 'Choose another category') + '</p>';

  if (!avail.length) {
    return '<h2 style="color:' + TEAL_DARK + '">No ' + esc(category) + ' available</h2>' + back;
  }

  let html = '<h2 style="color:' + TEAL_DARK + ';margin:12px 0 0">' + esc(category) + ' available</h2>' +
    '<p>For <b>' + esc(b.holder) + '</b>. Tick the tools you have fetched.</p>';
  html += avail.map(r =>
    '<label style="display:block;padding:14px;margin:8px 0;background:#fff;border:2px solid ' + TEAL +
    ';border-radius:8px;font-size:20px">' +
    '<input type="checkbox" class="pick" value="' + esc(r[0]) + '" style="transform:scale(1.6);margin-right:14px">' +
    esc(r[0]) + ' &middot; ' + esc(r[1]) + '</label>').join('');
  html += button('Check out selected', 'checkoutSelected()', TEAL) + back;
  return html;
}

// Called from the browser: checks out every ticked tool to the current student.
function checkoutBatch(pin, toolIds) {
  const staff = findStaff(pin);
  if (!staff) throw new Error('Wrong PIN');
  const b = getBatch(staff.id);
  if (!b) throw new Error('Student session has ended. Start a new student.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive();
    const ts = ss.getSheetByName(TOOLS);
    const log = ss.getSheetByName(LOG);
    const rows = ts.getDataRange().getValues();
    const now = new Date();
    const due = new Date(b.returnDate + 'T23:59:00');
    const done = [], taken = [];

    toolIds.forEach(id => {
      const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
      if (idx < 0 || rows[idx][4] === 'Out') { taken.push(id); return; }
      const sheetRow = idx + 1;
      ts.getRange(sheetRow, 5, 1, 5).setValues([['Out', b.holder, now, due, staff.name]]);
      rows[idx][4] = 'Out'; // keep the in-memory copy in sync
      log.appendRow([now, id, 'Out', b.studentNo, b.holder, staff.id, staff.name, 'Batch']);
      done.push(id);
    });

    let html = '<h2 style="color:' + TEAL_DARK + '">Checked out</h2>' +
      '<p><b>' + esc(b.holder) + '</b>, return by <b>' + fmt(due) + '</b></p>' +
      '<ul style="font-size:20px">' + done.map(id => '<li>' + esc(id) + '</li>').join('') + '</ul>';
    if (taken.length) {
      html += '<p style="color:#c00"><b>Not checked out (already taken or unknown):</b> ' +
        esc(taken.join(', ')) + '</p>';
    }
    html += '<p style="margin-top:20px">' + navLink('add', 'Add more tools for this student') + '</p>';
    return html;
  } finally {
    lock.releaseLock();
  }
}

// ---------- Staff page (QR code target) ----------

function staffPage(toolId) {
  const t = findTool(toolId);
  if (!t) return '<p>Tool not found: ' + esc(toolId) + '</p>';
  const title = '<h2 style="color:' + TEAL_DARK + ';margin:12px 0 0">' + esc(t.name) + '</h2>' +
    '<p style="color:#333;margin:4px 0 0">Tool ID: ' + esc(t.id) + '</p>';
  const pin = field('pin', 'Your PIN', 'password', 'inputmode="numeric"');
  const links = '<p style="margin-top:20px">' + navLink('start', 'Start a new student') + '</p>' +
    '<p>' + navLink('add', 'Add tools for this student') + '</p>';

  if (t.status === 'Out') {
    const due = t.due ? new Date(t.due) : null;
    const overdue = due && due < new Date();
    return title +
      banner('#ffcdd2', '#000',
        '<div style="font-size:14px;font-weight:bold">CHECKED OUT TO</div>' +
        '<div style="font-size:20px;font-weight:bold">' + esc(t.holder) + '</div>' +
        '<div style="font-size:14px;font-weight:bold;margin-top:8px">RETURN BY</div>' +
        '<div style="font-size:20px;font-weight:bold">' + (due ? fmt(due) : 'No date set') +
        (overdue ? ' <span style="color:#c00">OVERDUE</span>' : '') + '</div>' +
        '<div style="font-size:14px;font-weight:bold;margin-top:8px">SIGNED OUT BY</div>' +
        '<div style="font-size:18px">' + esc(t.signedBy || '-') + '</div>') +
      pin +
      button('Check in tool', "act('in')", TEAL) + script(t.id) + links;
  }

  return title +
    banner(TEAL_MID, TEAL_DARK, '<div style="font-size:20px;font-weight:bold">AVAILABLE</div>') +
    pin +
    '<p style="color:#333;font-size:15px;margin:14px 0 0">A student set with "Start a new student" ' +
    'takes this tool automatically. Leave the student number blank in that case.</p>' +
    field('stu', 'Student number (direct checkout)', 'text', 'inputmode="numeric"') +
    dateField() +
    button('Check out tool', "act('out')", TEAL) + script(t.id) + links;
}

// ---------- Student list (read-only) ----------

function listPage() {
  const rows = SpreadsheetApp.getActive().getSheetByName(TOOLS).getDataRange().getValues();
  let html = '<h2 style="color:' + TEAL_DARK + '">Tool availability</h2>' +
    '<table border="1" cellpadding="8" style="border-collapse:collapse;width:100%;border-color:' + TEAL_MID + '">' +
    '<tr style="background:' + TEAL + ';color:#fff"><th align="left">Tool</th><th align="left">Category</th>' +
    '<th align="left">Status</th><th align="left">Back by</th></tr>';
  rows.slice(1).forEach(r => {
    if (!r[0]) return;
    const out = r[4] === 'Out';
    const due = r[7] ? new Date(r[7]) : null;
    const overdue = out && due && due < new Date();
    const status = out ? '<b>Checked out</b>' : '<b style="color:' + TEAL_DARK + '">Available</b>';
    const back = out ? (due ? fmt(due) : '-') + (overdue ? ' <b style="color:#c00">(overdue)</b>' : '') : '-';
    html += '<tr style="background:' + (out ? '#ffebee' : TEAL_LIGHT) + ';color:#000">' +
      '<td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td><td>' + status + '</td><td>' + back + '</td></tr>';
  });
  return html + '</table>';
}

// ---------- Client-side script (tool page buttons) ----------

function script(id) {
  return '<script>function act(a){' +
    'var v=function(i){var el=document.getElementById(i);return el?el.value:""};' +
    'if(!v("pin")){alert("Enter your PIN.");return;}' +
    'google.script.run.withSuccessHandler(function(m){document.body.innerHTML=m})' +
    '.withFailureHandler(function(err){alert(err.message)})' +
    '.recordMove("' + id + '",a,v("pin"),v("stu"),v("ret"));}</script>';
}

// ---------- Server-side: check out / check in a single tool ----------

function recordMove(toolId, action, pin, studentNo, returnDate) {
  const staff = findStaff(pin);
  if (!staff) throw new Error('Wrong PIN');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive();
    const ts = ss.getSheetByName(TOOLS);
    const rows = ts.getDataRange().getValues();
    const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === toolId);
    if (idx < 0) throw new Error('Unknown tool');
    const row = rows[idx], sheetRow = idx + 1, now = new Date();
    const log = ss.getSheetByName(LOG);

    if (action === 'out') {
      if (row[4] === 'Out') throw new Error('Tool is already out');
      let holder, stuNo, due, mode;

      if (studentNo) {
        // Direct checkout: student number typed on the tool page
        if (!returnDate) throw new Error('Return date is required');
        holder = studentLookup(studentNo);
        if (!holder) throw new Error('Student number not found');
        stuNo = String(studentNo);
        due = new Date(returnDate + 'T23:59:00');
        mode = 'Direct';
      } else {
        // Batch checkout: student set with "Start a new student"
        const b = getBatch(staff.id);
        if (!b) throw new Error('No student set. Start a new student first, or enter a student number.');
        holder = b.holder;
        stuNo = b.studentNo;
        due = new Date(b.returnDate + 'T23:59:00');
        mode = 'Batch';
      }

      ts.getRange(sheetRow, 5, 1, 5).setValues([['Out', holder, now, due, staff.name]]);
      log.appendRow([now, toolId, 'Out', stuNo, holder, staff.id, staff.name, mode]);
      const next = mode === 'Batch' ? '<p>Scan the next tool, or use Add tools for this student.</p>' : '';
      return '<h2 style="color:' + TEAL_DARK + '">Checked out</h2><p>' + esc(toolId) + ' to ' +
        esc(holder) + '. Return by ' + fmt(due) + '.</p>' + next;
    }

    if (row[4] !== 'Out') throw new Error('Tool is already in');
    log.appendRow([now, toolId, 'In', '', row[5], staff.id, staff.name, '']);
    ts.getRange(sheetRow, 5, 1, 5).setValues([['In', '', '', '', '']]);
    return '<h2 style="color:' + TEAL_DARK + '">Checked in</h2><p>' + esc(toolId) +
      ' is available again.</p>';
  } finally {
    lock.releaseLock();
  }
}

// ---------- Server-side: student sessions (batch) ----------

function startBatch(pin, studentNo, returnDate) {
  const staff = findStaff(pin);
  if (!staff) throw new Error('Wrong PIN');
  const holder = studentLookup(studentNo);
  if (!holder) throw new Error('Student number not found');
  if (!returnDate) throw new Error('Return date is required');

  CacheService.getScriptCache().put('batch:' + staff.id,
    JSON.stringify({ studentNo: String(studentNo), holder: holder, returnDate: returnDate }),
    BATCH_SECONDS);

  return '<h2 style="color:' + TEAL_DARK + '">Student set</h2>' +
    '<p style="font-size:20px"><b>' + esc(holder) + '</b></p>' +
    '<p>Return by <b>' + fmt(new Date(returnDate + 'T23:59:00')) + '</b></p>' +
    '<p>Scan a tool, or use Add tools for this student. This lasts 30 minutes.</p>';
}

function endBatch(pin) {
  const staff = findStaff(pin);
  if (!staff) throw new Error('Wrong PIN');
  CacheService.getScriptCache().remove('batch:' + staff.id);
  return '<h2 style="color:' + TEAL_DARK + '">Session ended</h2><p>No student is set for you now.</p>';
}

function getBatch(staffId) {
  const s = CacheService.getScriptCache().get('batch:' + staffId);
  return s ? JSON.parse(s) : null;
}

// ---------- Lookups ----------

function findStaff(pin) {
  const p = String(pin || '').trim();
  if (!p) return null;
  const rows = SpreadsheetApp.getActive().getSheetByName(STAFF).getDataRange().getValues();
  const hits = rows.slice(1).filter(r => String(r[0]).trim() !== '' && String(r[2]).trim() === p);
  if (hits.length !== 1) return null;
  return { id: String(hits[0][0]).trim(), name: String(hits[0][1]).trim() };
}

function findTool(id) {
  const rows = SpreadsheetApp.getActive().getSheetByName(TOOLS).getDataRange().getValues();
  const r = rows.slice(1).find(r => String(r[0]) === id);
  return r ? { id: r[0], name: r[1], status: r[4], holder: r[5], due: r[7], signedBy: r[8] } : null;
}

// Students sheet: A = StudentNo, B = Name, C = Year (1st year, 2nd year, 3rd year, Advanced Diploma, Staff)
function studentRow(no) {
  const rows = SpreadsheetApp.getActive().getSheetByName(STUDENTS).getDataRange().getValues();
  return rows.slice(1).find(r => String(r[0]) === String(no)) || null;
}

function studentLookup(no) {
  const r = studentRow(no);
  return r ? r[1] + ' (' + no + ')' : null;
}

// Student row for a holder label like "Name (123)". Null if not found.
function studentRowForHolder(holder) {
  const m = /\(([^()]*)\)\s*$/.exec(String(holder || ''));
  return m ? studentRow(m[1]) : null;
}

// Year for a holder label like "Name (123)". Empty if not found.
function yearForHolder(holder) {
  const r = studentRowForHolder(holder);
  return r && r[3] ? String(r[3]).trim() : '';
}

// Full name for a holder label, e.g. "Thandi Mkhize (123)". Falls back to the label.
function fullNameForHolder(holder) {
  const m = /\(([^()]*)\)\s*$/.exec(String(holder || ''));
  const r = m ? studentRow(m[1]) : null;
  if (!r) return holder;
  return (String(r[1]) + ' ' + String(r[2] || '')).trim() + ' (' + m[1] + ')';
}

// ---------- HTML helpers ----------

// Links must open at the top of the window, otherwise Apps Script shows a blank screen.
function navLink(page, label) {
  return '<a href="' + ScriptApp.getService().getUrl() + '?page=' + page + '" target="_top" ' +
    'style="color:' + TEAL_DARK + ';font-weight:bold">' + label + '</a>';
}

function banner(bg, textColor, html) {
  return '<div style="background:' + bg + ';color:' + textColor +
    ';padding:14px;border-radius:8px;margin:12px 0;border-left:6px solid ' + TEAL_DARK + '">' + html + '</div>';
}

function field(id, label, type, extra) {
  return '<label for="' + id + '" style="display:block;margin:16px 0 6px;font-size:18px;font-weight:bold;color:' +
    TEAL_DARK + '">' + label + '</label>' +
    '<input id="' + id + '" type="' + type + '" ' + extra +
    ' style="display:block;width:100%;height:52px;font-size:18px;padding:8px 10px;box-sizing:border-box;' +
    'color:#000;background:#fff;border:2px solid ' + TEAL + ';border-radius:6px">';
}

// Return date, pre-filled with today on the server (no inline JavaScript needed).
function dateField() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return '<label for="ret" style="display:block;margin:16px 0 2px;font-size:20px;font-weight:bold;color:' +
    TEAL_DARK + '">RETURN DATE</label>' +
    '<div style="display:block;color:#000;font-size:15px;margin:0 0 6px">Pre-filled with today. Returning today is fine.</div>' +
    '<input id="ret" type="date" value="' + today + '" min="' + today + '" style="display:block;width:100%;' +
    'height:52px;font-size:18px;padding:8px 10px;box-sizing:border-box;color:#000;background:#fff;' +
    'border:2px solid ' + TEAL + ';border-radius:6px">';
}

function button(label, onclick, color) {
  return '<button onclick="' + onclick + '" style="display:block;width:100%;padding:16px;font-size:20px;margin-top:18px;' +
    'background:' + color + ';color:#fff;border:0;border-radius:8px;font-weight:bold">' + label + '</button>';
}

function fmt(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEE d MMM yyyy');
}

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g,
    c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ---------- QR code export (run manually) ----------
// Each QR code holds only the tool ID, so it is a small, simple code that phones read easily.

function saveQRCodes() {
  const rows = SpreadsheetApp.getActive().getSheetByName(TOOLS).getDataRange().getValues();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
  const folder = DriveApp.createFolder('Tool QR codes ' + stamp);

  rows.slice(1).forEach(r => {
    const id = String(r[0]);
    if (!id) return;
    const url = 'https://api.qrserver.com/v1/create-qr-code/?size=600x600&ecc=L&margin=1&data=' +
      encodeURIComponent(id);
    const blob = UrlFetchApp.fetch(url).getBlob().setName(id + '.png');
    folder.createFile(blob);
  });
}

// ============================================================
// JSON API for CPUT-Workshop-Scanner
// ============================================================

function apiRespond(e) {
  const p = e.parameter;
  const cb = String(p.callback || 'cb').replace(/[^\w.$]/g, '');
  let out;
  try {
    out = { ok: true, data: apiRoute(p) };
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function apiRoute(p) {
  const ids = String(p.ids || '').split(',').map(s => s.trim()).filter(Boolean);
  switch (p.api) {
    case 'signin':   return apiSignin(p.pin);
    case 'student':  return apiStudent(p.token, p.no);
    case 'tool':     return apiTool(p.token, p.id);
    case 'holding':  return apiHolding(p.token, p.no);
    case 'ping':     return { time: new Date().toISOString() };
    case 'mytools':  return apiMyTools(p.no);
    case 'checkout': return apiCheckout(p.token, ids, p.no, p.ret);
    case 'checkin':  return apiCheckin(p.token, ids);
    case 'family':   return apiFamily(p.token, p.qr);
    case 'take':     return apiTake(p.token, p.no, p.lines);
    case 'return':   return apiReturnItems(p.token, p.no, p.lines);
    default: throw new Error('Unknown request');
  }
}

// PIN is used once, at sign-in. Returns a token that expires after 10 minutes of inactivity.
function apiSignin(pin) {
  const staff = findStaff(pin);
  if (!staff) throw new Error('Wrong PIN');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('tok:' + token, JSON.stringify(staff), SESSION_SECONDS);
  return { token: token, name: staff.name, seconds: SESSION_SECONDS };
}

// Every other call uses the token. Each use refreshes the 10-minute timer.
function staffFromToken(token) {
  const cache = CacheService.getScriptCache();
  const s = token ? cache.get('tok:' + token) : null;
  if (!s) throw new Error('Session expired. Sign in again.');
  cache.put('tok:' + token, s, SESSION_SECONDS);
  return JSON.parse(s);
}

function apiStudent(token, no) {
  staffFromToken(token);
  const r = studentRow(no);
  if (!r) throw new Error('Student number not found');
  const full = (String(r[1]) + ' ' + String(r[2] || '')).trim();
  return { holder: full + ' (' + no + ')', year: r[3] ? String(r[3]).trim() : '' };
}

function apiTool(token, id) {
  staffFromToken(token);
  const t = findTool(id);
  if (!t) throw new Error('Tool not found: ' + id);
  const out = t.status === 'Out';
  const due = t.due ? new Date(t.due) : null;
  return {
    id: t.id, name: t.name, out: out,
    holder: out ? t.holder : '',
    due: due ? fmt(due) : '',
    dueDate: due ? Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    overdue: !!(out && due && due < new Date()),
    borrower: out ? fullNameForHolder(t.holder) : '',
    year: out ? yearForHolder(t.holder) : ''
  };
}

// Tools currently out to one student, earliest return first. Overdue flagged for red display.
function apiHolding(token, no) {
  staffFromToken(token);
  const holder = studentLookup(no);
  if (!holder) throw new Error('Student number not found');
  const rows = SpreadsheetApp.getActive().getSheetByName(TOOLS).getDataRange().getValues();
  const now = new Date();
  const items = rows.slice(1)
    .filter(r => String(r[0]).trim() !== '' && r[4] === 'Out' && String(r[5]) === holder)
    .map(r => {
      const due = r[7] ? new Date(r[7]) : null;
      return {
        id: String(r[0]),
        name: String(r[1]),
        due: due ? fmt(due) : 'No date set',
        dueDate: due ? Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        dueTime: due ? due.getTime() : Infinity,
        overdue: !!(due && due < now)
      };
    })
    .sort((a, b) => a.dueTime - b.dueTime)
    .map(({ dueTime, ...rest }) => rest);
  return { items: items, accessories: accessoriesHeld(no) };
}

// Student self-service: the workshop tools signed out to this student, with return dates.
function apiMyTools(no) {
  const r = studentRow(no);
  if (!r) throw new Error('Student number not found.');
  const holder = studentLookup(no);
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const now = new Date();
  const rows = SpreadsheetApp.getActive().getSheetByName(TOOLS).getDataRange().getValues();
  const items = rows.slice(1)
    .filter(t => String(t[0]).trim() !== '' && t[4] === 'Out' && String(t[5]) === holder)
    .map(t => {
      const due = t[7] ? new Date(t[7]) : null;
      return {
        id: String(t[0]),
        name: String(t[1]),
        due: due ? fmt(due) : 'No date set',
        dueTime: due ? due.getTime() : Infinity,
        dueToday: !!due && Utilities.formatDate(due, tz, 'yyyy-MM-dd') === today,
        overdue: !!due && due < now
      };
    })
    .sort((a, b) => a.dueTime - b.dueTime)
    .map(({ dueTime, ...rest }) => rest);
  return { name: String(r[1]).trim(), items: items, accessories: accessoriesHeld(no) };
}

// Student and return date are passed on every call, so no 30-minute batch timer is involved.
function apiCheckout(token, ids, no, ret) {
  const staff = staffFromToken(token);
  if (!ids.length) throw new Error('No tools scanned');
  const holder = studentLookup(no);
  if (!holder) throw new Error('Student number not found');
  if (!ret) throw new Error('Return date is required');
  const stuNo = String(no), due = new Date(ret + 'T23:59:00');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive();
    const ts = ss.getSheetByName(TOOLS), log = ss.getSheetByName(LOG);
    const rows = ts.getDataRange().getValues();
    const now = new Date(), done = [], taken = [];

    ids.forEach(id => {
      const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === id);
      if (idx < 0 || rows[idx][4] === 'Out') { taken.push(id); return; }
      ts.getRange(idx + 1, 5, 1, 5).setValues([['Out', holder, now, due, staff.name]]);
      rows[idx][4] = 'Out';
      log.appendRow([now, id, 'Out', stuNo, holder, staff.id, staff.name, 'Scanner']);
      done.push(id);
    });
    return { holder: holder, returnBy: fmt(due), done: done, taken: taken };
  } finally {
    lock.releaseLock();
  }
}

function apiCheckin(token, ids) {
  const staff = staffFromToken(token);
  if (!ids.length) throw new Error('No tools scanned');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive();
    const ts = ss.getSheetByName(TOOLS), log = ss.getSheetByName(LOG);
    const rows = ts.getDataRange().getValues();
    const now = new Date(), done = [], notOut = [];

    ids.forEach(id => {
      const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === id);
      if (idx < 0 || rows[idx][4] !== 'Out') { notOut.push(id); return; }
      log.appendRow([now, id, 'In', '', rows[idx][5], staff.id, staff.name, 'Scanner']);
      ts.getRange(idx + 1, 5, 1, 5).setValues([['In', '', '', '', '']]);
      done.push(id);
    });
    return { done: done, notOut: notOut };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// v18: families (drill bits, sanding paper and similar items)
// Each family has its own tab, named exactly as the Family column on the Families tab.
// Families tab columns: Family | Group | QR Code.
// Item tab columns: Code | Size or Grit | Stock (and Unit cost, Minimum for consumables).
// Columns are found by their heading, so their order does not matter.
// ============================================================

const FAMILIES = 'Families';
const LOG_QTY = 8;    // Log column I (zero-based index)
const LOG_COST = 9;   // Log column J
const LOG_TOTAL = 10; // Log column K

function familyRows() {
  const rows = SpreadsheetApp.getActive().getSheetByName(FAMILIES).getDataRange().getValues();
  const head = rows[0].map(h => String(h).trim().toLowerCase());
  const iF = head.indexOf('family'), iG = head.indexOf('group'), iQ = head.indexOf('qr code');
  if (iF < 0 || iQ < 0) throw new Error('Families tab needs Family and QR Code headings');
  return rows.slice(1)
    .filter(r => String(r[iF]).trim() !== '')
    .map(r => ({
      family: String(r[iF]).trim(),
      group: iG >= 0 ? String(r[iG]).trim() : '',
      qr: String(r[iQ]).trim().toUpperCase()
    }));
}

function isConsumable(group) {
  return String(group || '').trim().toLowerCase() === 'consumables';
}

// Reads one family's tab. Columns are found by heading.
function familyTab(family) {
  const sh = SpreadsheetApp.getActive().getSheetByName(family);
  if (!sh) throw new Error('Tab missing for ' + family);
  const rows = sh.getDataRange().getValues();
  const head = rows[0].map(h => String(h).trim().toLowerCase());
  const col = {
    code: head.indexOf('code'), name: head.indexOf('name'), stock: head.indexOf('stock'),
    cost: head.indexOf('unit cost'), min: head.indexOf('minimum')
  };
  if (col.code < 0 || col.stock < 0) throw new Error(family + ' tab needs Code and Stock headings');
  return { sheet: sh, rows: rows, col: col, family: family };
}

// What people see: the Name column if it is filled in, otherwise the family and its size or grit.
function displayName(tab, r) {
  if (tab.col.name >= 0 && String(r[tab.col.name]).trim()) return String(r[tab.col.name]).trim();
  return tab.family + ' ' + String(r[1]).trim();
}

function itemFrom(tab, r) {
  const stock = Number(r[tab.col.stock]) || 0;
  const min = tab.col.min >= 0 ? Number(r[tab.col.min]) || 0 : 0;
  return {
    code: String(r[tab.col.code]).trim().toUpperCase(),
    label: displayName(tab, r),
    stock: stock,
    unitCost: tab.col.cost >= 0 ? Number(r[tab.col.cost]) || 0 : 0,
    minimum: min,
    low: min > 0 && stock <= min
  };
}

// Item codes begin with their family's QR code, so DB-5.0 belongs to the family with QR code DB.
function familyForItem(code, fams) {
  const prefix = String(code).split('-')[0].trim().toUpperCase();
  return fams.find(f => f.qr === prefix) || null;
}

// Parses "DB-5.0:2,SP-120:3" into [{code, qty}].
function parseLines(text) {
  return String(text || '').split(',').map(s => s.trim()).filter(Boolean).map(s => {
    const m = /^(.+):(\d+)$/.exec(s);
    if (!m) throw new Error('Bad line: ' + s);
    return { code: m[1].trim().toUpperCase(), qty: parseInt(m[2], 10) };
  });
}

// Scanning a family's QR code: the family name and all its items with stock.
function apiFamily(token, qr) {
  staffFromToken(token);
  const f = familyRows().find(x => x.qr === String(qr || '').trim().toUpperCase());
  if (!f) throw new Error('Family not found');
  const tab = familyTab(f.family);
  const items = tab.rows.slice(1)
    .filter(r => String(r[tab.col.code]).trim() !== '')
    .map(r => itemFrom(tab, r));
  return { family: f.family, group: f.group, items: items };
}

// Accessories still out to a student: Out minus In, from the Log's quantity column.
// Tool lines have no quantity, so they are skipped here.
function accessoriesHeld(no) {
  const log = SpreadsheetApp.getActive().getSheetByName(LOG).getDataRange().getValues();
  const net = new Map(), since = new Map();
  log.slice(1).forEach(r => {
    if (String(r[3]).trim() !== String(no).trim()) return;
    const qty = Number(r[LOG_QTY]);
    if (!qty) return;
    const code = String(r[1]).trim().toUpperCase();
    const dir = r[2] === 'Out' ? 1 : r[2] === 'In' ? -1 : 0;
    net.set(code, (net.get(code) || 0) + dir * qty);
    if (dir === 1 && !since.has(code)) since.set(code, r[0]); // first time it was taken
  });
  const fams = familyRows();
  const out = [];
  net.forEach((qty, code) => {
    if (qty <= 0) return;
    const f = familyForItem(code, fams);
    if (!f || isConsumable(f.group)) return;
    const tab = familyTab(f.family);
    const row = tab.rows.slice(1).find(r => String(r[tab.col.code]).trim().toUpperCase() === code);
    const taken = since.get(code);
    out.push({ code: code, qty: qty, label: row ? displayName(tab, row) : '', family: f.family,
      taken: taken ? Utilities.formatDate(new Date(taken), Session.getScriptTimeZone(), 'd MMM yyyy') : '' });
  });
  return out;
}

// Takes items out to the current student. Accessories are expected back; consumables are used up.
// Every line is checked before anything is written, so a bad line takes nothing.
function apiTake(token, no, text) {
  const staff = staffFromToken(token);
  const holder = studentLookup(no);
  if (!holder) throw new Error('Student number not found');
  const lines = parseLines(text);
  if (!lines.length) throw new Error('Nothing to take');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive(), log = ss.getSheetByName(LOG);
    const fams = familyRows();
    const plan = lines.map(l => {
      const f = familyForItem(l.code, fams);
      if (!f) throw new Error('Unknown item: ' + l.code);
      const tab = familyTab(f.family);
      const idx = tab.rows.findIndex((r, i) => i > 0 && String(r[tab.col.code]).trim().toUpperCase() === l.code);
      if (idx < 0) throw new Error('Unknown item: ' + l.code);
      const item = itemFrom(tab, tab.rows[idx]);
      if (!(l.qty >= 1)) throw new Error('Quantity must be at least 1');
      if (l.qty > item.stock) throw new Error('Only ' + item.stock + ' of ' + item.label + ' in stock');
      return { l: l, f: f, tab: tab, row: idx + 1, item: item };
    });

    const now = new Date();
    plan.forEach(p => {
      const used = isConsumable(p.f.group);
      p.tab.sheet.getRange(p.row, p.tab.col.stock + 1).setValue(p.item.stock - p.l.qty);
      const cost = used ? p.item.unitCost : '';
      const total = used ? Math.round(p.item.unitCost * p.l.qty * 100) / 100 : '';
      log.appendRow([now, p.l.code, used ? 'Used' : 'Out', String(no), holder, staff.id, staff.name, 'Scanner',
        p.l.qty, cost, total]);
    });
    return { holder: holder, done: plan.map(p => ({ code: p.l.code, qty: p.l.qty, label: p.item.label,
      total: isConsumable(p.f.group) ? Math.round(p.item.unitCost * p.l.qty * 100) / 100 : 0 })) };
  } finally {
    lock.releaseLock();
  }
}

// Returns accessories. Only what this student actually has out can come back.
function apiReturnItems(token, no, text) {
  const staff = staffFromToken(token);
  const holder = studentLookup(no);
  if (!holder) throw new Error('Student number not found');
  const lines = parseLines(text);
  if (!lines.length) throw new Error('Nothing to return');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive(), log = ss.getSheetByName(LOG);
    const fams = familyRows();
    const held = new Map(accessoriesHeld(no).map(a => [a.code, a.qty]));
    const plan = lines.map(l => {
      const f = familyForItem(l.code, fams);
      if (!f || isConsumable(f.group)) throw new Error(l.code + ' is not a returnable item');
      const tab = familyTab(f.family);
      const idx = tab.rows.findIndex((r, i) => i > 0 && String(r[tab.col.code]).trim().toUpperCase() === l.code);
      if (idx < 0) throw new Error('Unknown item: ' + l.code);
      const have = held.get(l.code) || 0;
      if (!(l.qty >= 1) || l.qty > have) throw new Error(holder + ' has ' + have + ' of ' + itemFrom(tab, tab.rows[idx]).label + ' out');
      return { l: l, tab: tab, row: idx + 1, item: itemFrom(tab, tab.rows[idx]) };
    });

    const now = new Date();
    plan.forEach(p => {
      p.tab.sheet.getRange(p.row, p.tab.col.stock + 1).setValue(p.item.stock + p.l.qty);
      log.appendRow([now, p.l.code, 'In', String(no), holder, staff.id, staff.name, 'Scanner', p.l.qty, '', '']);
    });
    return { holder: holder, done: plan.map(p => ({ code: p.l.code, qty: p.l.qty, label: p.item.label })) };
  } finally {
    lock.releaseLock();
  }
}
