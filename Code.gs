/**
 * Wedding Planner — Google backend (Apps Script)
 * Lives inside YOUR "Wedding Master Tracker" Google Sheet.
 * The planner web page reads and writes the tabs "Master Tracker" and "Vendors & Payments".
 * Follow GUIDE-PLANNER.md. The only thing to edit here is the PIN below.
 */

var PIN = '1202';                       // <-- CHANGE THIS to your own family PIN, then re-deploy
var TASK_SHEET = 'Master Tracker';
var VENDOR_SHEET = 'Vendors & Payments';
var PAY_SHEET = 'Payments';
var VENDOR_EXTRA = ['ID', 'Split', 'Shrijeet Share %'];
var PAY_HEADERS = ['Pay ID', 'Vendor ID', 'Date', 'Amount', 'Paid By', 'Note', 'Added By', 'Added On'];
var TZ = 'Asia/Kolkata';
var EXTRA_COLS = ['ID', 'Owner', 'Due Date', 'Function', 'Completed On', 'Last Updated By', 'Last Updated'];
var FUNCTIONS = ['Before the wedding', 'Faldaan', 'Mehndi', 'Engagement & Sangeet', 'Haldi', 'Varmala & Shaadi', 'After the wedding'];
var STATUSES = ['Pending', 'In Progress', 'Completed'];
var PRIORITIES = ['P0', 'P1', 'P2'];

function guessFunction_(t) {
  t = String(t).toLowerCase();
  if (!t.trim()) return '';
  if (/faldaan/.test(t)) return 'Faldaan';
  if (/haldi/.test(t)) return 'Haldi';
  if (/mehndi|mehendi/.test(t)) return 'Mehndi';
  if (/sangeet|\bdj\b|entry song|dance|performance/.test(t)) return 'Engagement & Sangeet';
  if (/village reception/.test(t)) return 'After the wedding';
  if (/barat|baraat|dhol|shehnai|pandit|ritual|pheras|varmala|mandap/.test(t)) return 'Varmala & Shaadi';
  return 'Before the wedding';
}

function setup() { ensureColumns_(); }   // optional: run once from the editor

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (String(p.pin || '') !== PIN) return json_({ success: false, code: 'PIN', message: 'Wrong PIN' });
    if (p.action === 'all') return json_(readAll_());
    return json_({ success: true, message: 'Planner backend is running.' });
  } catch (err) { return json_({ success: false, message: 'Could not load: ' + err }); }
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    if (String(b.pin || '') !== PIN) return json_({ success: false, code: 'PIN', message: 'Wrong PIN' });
    if (b.action === 'update') return json_(updateTask_(b));
    if (b.action === 'add') return json_(addTask_(b));
    if (b.action === 'vendorAdd') return json_(vendorAdd_(b));
    if (b.action === 'vendorUpdate') return json_(vendorUpdate_(b));
    if (b.action === 'vendorDelete') return json_(vendorDelete_(b));
    if (b.action === 'payAdd') return json_(payAdd_(b));
    if (b.action === 'payUpdate') return json_(payUpdate_(b));
    if (b.action === 'payDelete') return json_(payDelete_(b));
    return json_({ success: false, message: 'Unknown action' });
  } catch (err) { return json_({ success: false, message: 'Could not save: ' + err }); }
}

/* ---------------- columns ---------------- */
function norm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }

function headerMap_(sheet) {
  var last = Math.max(sheet.getLastColumn(), 1);
  var hdr = sheet.getRange(1, 1, 1, last).getValues()[0];
  var map = {};
  for (var i = 0; i < hdr.length; i++) { var k = norm_(hdr[i]); if (k && map[k] === undefined) map[k] = i + 1; }
  return map;
}

function ensureColumns_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(TASK_SHEET);
  if (!sheet) throw new Error('Tab "' + TASK_SHEET + '" not found');
  var map = headerMap_(sheet);
  var col = sheet.getLastColumn(), newFunction = false;
  EXTRA_COLS.forEach(function (name) {
    if (!map[norm_(name)]) { col++; sheet.getRange(1, col).setValue(name).setFontWeight('bold'); if (name === 'Function') newFunction = true; }
  });
  map = headerMap_(sheet);
  var taskCol = map.task, idCol = map.id, rows = sheet.getLastRow();
  if (rows < 2) return map;
  var data = sheet.getRange(2, 1, rows - 1, sheet.getLastColumn()).getValues();
  var max = 0;
  data.forEach(function (r) { var m = /^T(\d+)$/.exec(String(r[idCol - 1])); if (m) max = Math.max(max, Number(m[1])); });
  var ids = data.map(function (r) {
    var cur = String(r[idCol - 1] || '');
    if (cur) return [cur];
    if (!String(r[taskCol - 1] || '').trim()) return [''];
    max++; return ['T' + ('000' + max).slice(-3)];
  });
  sheet.getRange(2, idCol, ids.length, 1).setValues(ids);
  if (newFunction) {   // first time only: guess which function each task belongs to (you can change it any time)
    var fns = data.map(function (r) { return [guessFunction_(String(r[taskCol - 1]))]; });
    sheet.getRange(2, map.function, fns.length, 1).setValues(fns);
  }
  sheet.getRange(2, map.duedate, rows - 1, 1).setNumberFormat('yyyy-mm-dd');
  return map;
}

/* ---------------- read ---------------- */
function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}
function status_(v) {
  var s = norm_(v);
  if (s.indexOf('complete') === 0 || s === 'done') return 'Completed';
  if (s.indexOf('progress') >= 0) return 'In Progress';
  return 'Pending';
}

function readAll_() {
  var cache = null;
  try { cache = CacheService.getScriptCache(); var hit = cache.get('ALL'); if (hit) return JSON.parse(hit); } catch (e) { cache = null; }
  var out = readAllFresh_();
  try { if (cache) cache.put('ALL', JSON.stringify(out), 25); } catch (e) {}
  return out;
}
function clearCache_() { try { CacheService.getScriptCache().remove('ALL'); } catch (e) {} }

function setupNeeded_(sheet) {
  var map = headerMap_(sheet);
  for (var i = 0; i < EXTRA_COLS.length; i++) if (!map[norm_(EXTRA_COLS[i])]) return true;
  return false;
}

function readAllFresh_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(TASK_SHEET);
  if (!sheet) throw new Error('Tab "' + TASK_SHEET + '" not found');
  if (setupNeeded_(sheet)) ensureColumns_();
  seedOnce_(); seedAfterWedding_();
  var map = headerMap_(sheet);
  var rows = sheet.getLastRow(), tasks = [];
  if (rows >= 2) {
    var data = sheet.getRange(2, 1, rows - 1, sheet.getLastColumn()).getValues();
    var missingId = data.some(function (r) { return String(r[map.task - 1] || '').trim() && !String(r[map.id - 1] || ''); });
    if (missingId) { ensureColumns_(); data = sheet.getRange(2, 1, rows - 1, sheet.getLastColumn()).getValues(); }
    var g = function (r, key) { return map[key] ? r[map[key] - 1] : ''; };
    data.forEach(function (r) {
      var id = str_(g(r, 'id'));
      if (!id || !str_(g(r, 'task')).trim()) return;
      tasks.push({
        id: id, category: str_(g(r, 'category')).trim(), task: str_(g(r, 'task')).trim(),
        status: status_(g(r, 'status')), priority: str_(g(r, 'priority')).trim().toUpperCase() || 'P1',
        vendor: str_(g(r, 'vendorcontact')), cost: str_(g(r, 'cost')),
        remarks: str_(g(r, 'nextactionremarks')), planAhead: str_(g(r, 'planahead')),
        owner: str_(g(r, 'owner')), due: str_(g(r, 'duedate')), fn: str_(g(r, 'function')), completedOn: str_(g(r, 'completedon')),
        updatedBy: str_(g(r, 'lastupdatedby')), updatedAt: str_(g(r, 'lastupdated'))
      });
    });
  }
  var vs = ss.getSheetByName(VENDOR_SHEET);
  if (vs) vendorSetup_(vs);
  var payments = readPayments_(), paid = {};
  payments.forEach(function (p) { if (p.vid !== 'SETTLE') paid[p.vid] = (paid[p.vid] || 0) + p.amount; });
  var vendors = [];
  if (vs && vs.getLastRow() >= 2) {
    var vm = headerMap_(vs);
    var vd = vs.getRange(2, 1, vs.getLastRow() - 1, vs.getLastColumn()).getValues();
    var vg = function (r, key) { return vm[key] ? r[vm[key] - 1] : ''; };
    vd.forEach(function (r) {
      var cat = str_(vg(r, 'category')).trim(), id = str_(vg(r, 'id'));
      if (!id || (!cat && !str_(vg(r, 'vendor')).trim())) return;
      var quoted = num_(vg(r, 'quotedamount')), pd = paid[id] || 0, pct = vg(r, 'shrijeetshare');
      vendors.push({
        id: id, category: cat, vendor: str_(vg(r, 'vendor')).trim(), contact: str_(vg(r, 'contact')).trim(),
        deliverables: str_(vg(r, 'deliverables')).trim(), quoted: quoted, advance: pd, balance: quoted - pd,
        status: str_(vg(r, 'status')).trim(), remarks: str_(vg(r, 'remarks')).trim(),
        split: str_(vg(r, 'split')).trim() === 'Personal' ? 'Personal' : 'Shared', pct: pct === '' || pct == null ? 50 : num_(pct)
      });
    });
  }
  return { success: true, tasks: tasks, vendors: vendors, payments: payments, serverTime: new Date().getTime() };
}
function num_(v) { var n = Number(v); return isFinite(n) && v !== '' && v !== null ? n : 0; }

/* ---------------- write ---------------- */
function safe_(v, max) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function who_(b) { return safe_(b.by, 40) || 'Someone'; }

function findRowById_(sheet, map, id) {
  var rows = sheet.getLastRow();
  if (rows < 2) return 0;
  var ids = sheet.getRange(2, map.id, rows - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

var FIELD_COL = { category: 'category', task: 'task', status: 'status', priority: 'priority', vendor: 'vendorcontact',
  cost: 'cost', remarks: 'nextactionremarks', planAhead: 'planahead', owner: 'owner', due: 'duedate', fn: 'function' };

function updateTask_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(TASK_SHEET);
    var map = headerMap_(sheet);
    if (!map.id || !map.completedon) map = ensureColumns_();
    var row = findRowById_(sheet, map, b.id);
    if (!row) return { success: false, message: 'Task not found (was it deleted in the sheet?)' };
    var f = b.fields || {}, changed = 0;
    Object.keys(f).forEach(function (k) {
      var col = map[FIELD_COL[k]]; if (!col) return;
      var v = f[k];
      if (k === 'status') {
        if (STATUSES.indexOf(v) < 0) return;
        if (map.completedon) sheet.getRange(row, map.completedon).setNumberFormat('yyyy-mm-dd').setValue(v === 'Completed' ? Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') : '');
      }
      else if (k === 'priority') { if (PRIORITIES.indexOf(v) < 0) return; }
      else if (k === 'due') {
        if (v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) return;
        var cell = sheet.getRange(row, col);
        cell.setNumberFormat('yyyy-mm-dd');
        if (v) { var p = v.split('-'); cell.setValue(new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0)); } else cell.clearContent();
        changed++; return;
      }
      else if (k === 'fn') { if (v && FUNCTIONS.indexOf(v) < 0) return; }
      else v = safe_(v, k === 'remarks' || k === 'planAhead' ? 1000 : 200);
      sheet.getRange(row, col).setValue(v); changed++;
    });
    if (changed) {
      sheet.getRange(row, map.lastupdatedby).setValue(who_(b));
      sheet.getRange(row, map.lastupdated).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
    }
    clearCache_();
    return { success: true };
  } finally { lock.releaseLock(); }
}

function addTask_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(TASK_SHEET);
    var map = ensureColumns_();
    var f = b.fields || {};
    var task = safe_(f.task, 200);
    if (!task) return { success: false, message: 'Please type the task.' };
    var rows = sheet.getLastRow(), max = 0;
    if (rows >= 2) sheet.getRange(2, map.id, rows - 1, 1).getValues().forEach(function (r) { var m = /^T(\d+)$/.exec(String(r[0])); if (m) max = Math.max(max, Number(m[1])); });
    var id = 'T' + ('000' + (max + 1)).slice(-3);
    var row = new Array(sheet.getLastColumn()).fill('');
    row[map.id - 1] = id; row[map.task - 1] = task;
    row[map.category - 1] = safe_(f.category, 80) || 'Other';
    row[map.status - 1] = 'Pending';
    row[map.priority - 1] = PRIORITIES.indexOf(f.priority) >= 0 ? f.priority : 'P1';
    row[map.function - 1] = FUNCTIONS.indexOf(f.fn) >= 0 ? f.fn : 'Before the wedding';
    if (map.owner) row[map.owner - 1] = safe_(f.owner, 60);
    if (map.nextactionremarks) row[map.nextactionremarks - 1] = safe_(f.remarks, 1000);
    row[map.lastupdatedby - 1] = who_(b);
    row[map.lastupdated - 1] = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    sheet.appendRow(row);
    var r = sheet.getLastRow();
    if (f.due && /^\d{4}-\d{2}-\d{2}$/.test(f.due)) { var p = f.due.split('-'); sheet.getRange(r, map.duedate).setNumberFormat('yyyy-mm-dd').setValue(new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0)); }
    clearCache_();
    return { success: true, id: id };
  } finally { lock.releaseLock(); }
}

var SUGGESTED = [
  ["Ceremonies & Religious Arrangements", "Faldaan items ready: tilak thali, gifts, sweets, dry fruits", "P0", "Faldaan", "Check the list with Pandit Ji and the elders a day before"],
  ["Event Planning & Execution", "Faldaan: seating, photographer and arrival of guests on 30 Nov", "P1", "Faldaan", "Who receives guests, where they sit, tea and snacks"],
  ["Makeup & Mehndi", "Mehndi: seating for artists, chargers, snacks and drinks for the ladies", "P1", "Mehndi", "Easy to forget; assign one person"],
  ["Ceremonies & Religious Arrangements", "Engagement: rings, ring tray and who exchanges", "P0", "Engagement & Sangeet", "Rings, tray, flowers and the order of the ceremony"],
  ["Music, Sangeet & Entertainment", "Sangeet: stage, mic, lights and printed order of performances", "P0", "Engagement & Sangeet", "Give a copy to the emcee, DJ and photographer"],
  ["Ceremonies & Religious Arrangements", "Haldi items: haldi paste, flowers, towels, old / yellow clothes", "P1", "Haldi", "Keep extra towels and a change of clothes for everyone"],
  ["Ceremonies & Religious Arrangements", "Barat welcome (milni): garlands, order and who greets whom", "P0", "Varmala & Shaadi", "Plan names, order and photographer position"],
["Payments & Vendor Management", "Master wedding budget", "P0", "Before the wedding", "Track estimated vs actual spending and keep a contingency amount"], ["Payments & Vendor Management", "Written vendor scope / contracts", "P0", "Before the wedding", "Get inclusions and deliverables in writing from every vendor to avoid disputes"], ["Event Planning & Execution", "Family responsibility matrix", "P0", "Before the wedding", "Give one owner each for guests, rooms, payments, rituals and transport"], ["Event Planning & Execution", "Vendor arrival / setup schedule", "P0", "Before the wedding", "Exact time each vendor reaches and sets up, function by function"], ["Guests, Travel & Transport", "Guest pickup / drop coordinator", "P1", "Before the wedding", "One person for airport / station / venue transfers"], ["Venue & Accommodation", "Room key / check-in desk", "P1", "Before the wedding", "Desk with room list, keys and a guest register for smooth check-in"], ["Guests, Travel & Transport", "Welcome kits for outstation guests", "P1", "Before the wedding", "Itinerary, contacts, room and venue information, water and snacks"], ["Event Planning & Execution", "Function-wise guest count", "P0", "Before the wedding", "Avoid catering and seating mismatch"], ["Event Planning & Execution", "Stage / seating plan", "P1", "Varmala & Shaadi", "Avoid crowding and confusion during ceremonies; seats for elders"], ["Venue & Accommodation", "Power backup / generator confirmation", "P1", "Before the wedding", "Critical for lights, DJ, décor and photography"], ["Event Planning & Execution", "Sound / mic plan", "P1", "Before the wedding", "Separate needs for rituals, speeches and performances; sound check before each function"], ["Event Planning & Execution", "Rain / weather contingency", "P1", "Before the wedding", "Backup plan for outdoor functions; seaside wind and humidity"], ["Event Planning & Execution", "Gift / envelope receiving desk", "P1", "Varmala & Shaadi", "One trusted person, a locked box and a written tally"], ["Event Planning & Execution", "Jewellery / valuables safekeeping", "P1", "Before the wedding", "Who keeps what, where, and the key; list of items"], ["Makeup & Mehndi", "Bride and groom emergency kits", "P2", "Before the wedding", "Safety pins, sewing kit, double-sided tape, pain relief, bandages, wet wipes, snacks, water, comb, extra makeup"], ["Photography, Video & Memories", "Photographer shot list", "P1", "Before the wedding", "Must-have moments and who to capture, shared with the photographer"], ["Photography, Video & Memories", "Family group-photo list", "P1", "Before the wedding", "Group combinations in order, with one person calling families"], ["Payments & Vendor Management", "Vendor master contact sheet", "P0", "Before the wedding", "Names, numbers, arrival times and balance payments in one printout"], ["Event Planning & Execution", "Final 48-hour checklist", "P0", "Before the wedding", "Walk through every function, vendor and room two days before"], ["Venue & Accommodation", "Entry passes for guests staying in naval premises", "P0", "Before the wedding", "Gate passes, ID proof list and vehicle passes for NOI, Command Mess and HSL stays"], ["Venue & Accommodation", "Resort rules: DJ cut-off time, crackers, outside decorator / food", "P0", "Before the wedding", "Confirm in writing with the resort what is allowed and until what time"], ["Venue & Accommodation", "Loudspeaker / late-night permission for Shaadi (2:30 AM) and Barat", "P1", "Before the wedding", "Check what permission or timing the resort and local authorities require"], ["Venue & Accommodation", "Early check-in / late check-out arrangements", "P1", "Before the wedding", "Resort check-in 10 AM, check-out 11 AM; Hotel FantaSea check-out 3 Dec 11 AM"], ["Venue & Accommodation", "Rooms ready: extra mattresses, blankets, hot water, AC working", "P2", "Before the wedding", "Inspect every stay one day before guests arrive"], ["Food & Catering", "Food tasting before finalising the menu", "P1", "Before the wedding", "Taste with family; confirm per-plate price and live counters"], ["Food & Catering", "Meals for vendors, drivers and helpers", "P1", "Before the wedding", "Photographers, decorators, DJ, drivers need food all day"], ["Food & Catering", "Special food: Jain, diabetic, children, elders", "P2", "Before the wedding", "Tell the caterer in advance"], ["Food & Catering", "Early morning tea / breakfast for guests every day", "P1", "Before the wedding", "Decide timing, place and who serves"], ["Food & Catering", "Late-night tea and snacks around the Shaadi (2:30 AM)", "P1", "Varmala & Shaadi", "Hot tea, coffee and light snacks for guests who stay awake"], ["Food & Catering", "Drinking water and dustbins at every venue", "P2", "Before the wedding", "Water bottles, dustbins and cleaning staff for each function"], ["Décor & Flowers", "Décor site visit and final look approval", "P1", "Before the wedding", "Check mandap, stage, entry and lights with the décor team a day before"], ["Ceremonies & Religious Arrangements", "Hawan kund, samagri and mandap items — who provides?", "P0", "Varmala & Shaadi", "Clarify with Pandit Ji and the decorator so nothing is missed"], ["Ceremonies & Religious Arrangements", "Muhurat timings for every ritual in writing", "P0", "Before the wedding", "Written from Pandit Ji, shared with family, photographer and caterer"], ["Ceremonies & Religious Arrangements", "Sindoor, mangalsutra, garlands and ring items ready", "P0", "Varmala & Shaadi", "Keep in one labelled box with a responsible person"], ["Ceremonies & Religious Arrangements", "Moli / kalira for the bride's bangles", "P2", "Before the wedding", "Small item, easy to forget"], ["Ceremonies & Religious Arrangements", "Family roles in rituals (kanyadaan, pheras, vidaai)", "P1", "Before the wedding", "Tell each person their role and when to be ready"], ["Ceremonies & Religious Arrangements", "Vidaai arrangements", "P1", "Varmala & Shaadi", "Car, rice, flowers, tissues and who stays with the bride"], ["Ceremonies & Religious Arrangements", "Ghar pravesh arrangements", "P1", "After the wedding", "Aarti thali, rangoli, kalash and photographer timing"], ["Ceremonies & Religious Arrangements", "Ghodi / chariot booking for Barat (if needed)", "P2", "Before the wedding", "Small vendors need early booking and advance"], ["Music, Sangeet & Entertainment", "Backup copies of entry songs and videos", "P1", "Before the wedding", "Pendrive plus phone plus laptop; give the DJ one copy"], ["Music, Sangeet & Entertainment", "Person to confirm songs and cues with the DJ", "P1", "Before the wedding", "Share a written song list for each function"], ["Makeup & Mehndi", "Bride makeup and hair trial", "P1", "Before the wedding", "Do it at least two weeks before"], ["Makeup & Mehndi", "Makeup / hair for family women", "P2", "Before the wedding", "Book slots so the schedule does not slip"], ["Makeup & Mehndi", "Bridal jewellery packing and transport to makeup venue", "P1", "Varmala & Shaadi", "A trusted person carries it and keeps it safe"], ["Makeup & Mehndi", "Backup footwear for bride and groom", "P2", "Before the wedding", "Comfortable pair for dancing and long ceremonies"], ["Shopping & Personal Arrangements", "Alterations and trials for all outfits", "P0", "Before the wedding", "Tailor visits at least two weeks before; keep a list per person"], ["Shopping & Personal Arrangements", "Accessories: footwear, dupatta, safa, sehra, kalgi", "P1", "Before the wedding", "Check each outfit has all its accessories"], ["Shopping & Personal Arrangements", "Lena-dena list: gifts and envelopes for relatives, Pandit Ji, helpers", "P1", "Before the wedding", "Prepare envelopes in advance with names"], ["Shopping & Personal Arrangements", "Tips / envelopes for band, staff, drivers, resort staff", "P2", "Before the wedding", "Prepare small notes and decide who hands them over"], ["Guests, Travel & Transport", "Re-check train / flight tickets for outstation guests", "P0", "Before the wedding", "Confirm waitlist or RAC status and fix alternatives"], ["Guests, Travel & Transport", "Return tickets for guests and family", "P1", "Before the wedding", "Book early; tickets for 3 Dec onwards get full"], ["Guests, Travel & Transport", "Shuttle schedule between hotels, bungalows and resort", "P0", "Before the wedding", "Guests are staying in many places; fix timings and drivers"], ["Guests, Travel & Transport", "Driver names and numbers shared with guests", "P1", "Before the wedding", "Send on the WhatsApp group"], ["Guests, Travel & Transport", "Plan for elderly guests: ground-floor rooms, wheelchair, helper", "P1", "Before the wedding", "Ask elders about needs in advance"], ["Guests, Travel & Transport", "Nearest hospital, doctor and first-aid at venue", "P1", "Before the wedding", "Save numbers; keep a basic first-aid box with the coordinator"], ["Guests, Travel & Transport", "Sightseeing plan for free time in Vizag", "P2", "Before the wedding", "RK Beach, Kailasagiri, museums; share on the Explore page"], ["Event Planning & Execution", "Rehearsal of Varmala and Shaadi sequence", "P1", "Haldi", "Walk-through with family, photographer, DJ and Pandit Ji"], ["Event Planning & Execution", "Phone chargers and power strips at venue", "P2", "Before the wedding", "Small things guests always ask for"], ["Event Planning & Execution", "Live stream for relatives who cannot come", "P2", "Before the wedding", "Check internet at the resort and who runs it"], ["Invitations & Communication", "Personal calls to elders and key relatives", "P1", "Before the wedding", "A call counts more than a card"], ["Invitations & Communication", "Share wedding website link on family WhatsApp groups", "P0", "Before the wedding", "RSVP and photo upload need the link and QR"], ["Invitations & Communication", "Test the website RSVP and photo upload on 2–3 phones", "P0", "Before the wedding", "Do it before sharing widely"], ["Invitations & Communication", "Check the RSVP sheet weekly and follow up non-responders", "P1", "Before the wedding", "Call or message those who have not replied"], ["Photography, Video & Memories", "Print the Photo QR cards and place them at venues", "P1", "Before the wedding", "Entrance, tables and reception desk; scan once to test"], ["Payments & Vendor Management", "Final vendor payments and receipts", "P0", "After the wedding", "Settle balances and collect written receipts"], ["Venue & Accommodation", "Room checkout inspection and return of keys / rented items", "P1", "After the wedding", "Check each stay and note damages or missing items"], ["Event Planning & Execution", "Tally of gifts and envelopes", "P1", "After the wedding", "Record who gave what in a notebook or sheet"], ["Invitations & Communication", "Thank-you messages with the photo gallery link", "P2", "After the wedding", "Send on WhatsApp once photos are ready"], ["Photography, Video & Memories", "Photo and video delivery, album selection follow-up", "P2", "After the wedding", "Get a delivery date in writing"], ["Photography, Video & Memories", "Download RSVP sheet and guest photos backup", "P1", "After the wedding", "Save a copy from Google Sheet and Drive"], ["Food & Catering", "Leftover food distribution and lost-and-found", "P2", "After the wedding", "Assign one person"]];

/* Adds the suggested tasks ONE time (a flag stops repeats) and skips any task you already have. */

var SUGGESTED_AFTER = [
  ["Legal & Documents", "Take 6–8 certified copies of the marriage certificate", "P0", "After the wedding", "Every office below asks for one; keep scans in Drive too"],
  ["Legal & Documents", "Update name / surname on Aadhaar (if Shivangi is changing it)", "P1", "After the wedding", "Needs the marriage certificate; check the current UIDAI process"],
  ["Legal & Documents", "Update name on PAN card and link it", "P1", "After the wedding", "Do this after Aadhaar so the details match"],
  ["Legal & Documents", "Passport: change name and add spouse details", "P1", "After the wedding", "Check passport validity before honeymoon travel"],
  ["Legal & Documents", "Update name / address on bank accounts, cards, driving licence, voter ID", "P2", "After the wedding", "One list; tick each as it is done"],
  ["Navy & Service Records", "Update marital status and next of kin in service records (unit / pay office)", "P0", "After the wedding", "Confirm the exact forms and the proof needed with your unit and pay office"],
  ["Navy & Service Records", "Spouse dependent ID card and medical entitlement for Shivangi", "P1", "After the wedding", "Ask the unit / Station HQ what documents and photos they want"],
  ["Navy & Service Records", "Check married accommodation, HRA and allowances after marriage", "P2", "After the wedding", "Ask the pay office what applies and from which date"],
  ["Navy & Service Records", "Leave and joining date: confirm with unit and plan travel", "P1", "After the wedding", "Keep the leave sanction copy with you"],
  ["Finance & Insurance", "Add Shivangi to health and life insurance; update nominees", "P1", "After the wedding", "Also provident fund, mutual funds, bank accounts and demat"],
  ["Finance & Insurance", "Open a joint account or add Shivangi as nominee (if you plan to)", "P2", "After the wedding", "Discuss and decide together"],
  ["Payments & Vendor Management", "Collect refundable deposits from the resort, decorator and others", "P1", "After the wedding", "Note the deposit amounts and the dates they should come back"],
  ["Payments & Vendor Management", "Keep soft copies of all vendor bills and receipts in one folder", "P1", "After the wedding", "Needed for the expense split and any disputes"],
  ["Payments & Vendor Management", "Settle shared expenses with Shivangi (Hisaab tab in this planner)", "P1", "After the wedding", "Mark personal items, add every payment, then settle up"],
  ["Shopping & Personal Arrangements", "Store jewellery safely or move it to a bank locker", "P1", "After the wedding", "Make a list with photos and weights"],
  ["Shopping & Personal Arrangements", "Dry-clean and store wedding outfits properly", "P2", "After the wedding", "Wrap in cotton; keep the safa, sehra and dupattas with each outfit"],
  ["Ceremonies & Religious Arrangements", "Post-wedding rituals and visits as per family custom", "P2", "After the wedding", "Confirm the days and the visits with the elders"],
  ["Invitations & Communication", "Thank-you calls to elders, helpers and vendors", "P2", "After the wedding", "A short personal call is remembered"],
  ["Travel & Honeymoon", "Plan and book honeymoon (tickets, stay, visas if needed)", "P1", "After the wedding", "Check leave dates, passport validity and cancellation rules"]
];

function seedOnce_() { seedList_('SEEDED_V2', SUGGESTED); }
function seedAfterWedding_() { seedList_('SEEDED_V3', SUGGESTED_AFTER); }
function seedList_(KEY, LIST) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(KEY)) return;
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    if (props.getProperty(KEY)) return;
    var sheet = SpreadsheetApp.getActive().getSheetByName(TASK_SHEET);
    var map = ensureColumns_(), rows = sheet.getLastRow(), have = {}, max = 0;
    if (rows >= 2) sheet.getRange(2, 1, rows - 1, sheet.getLastColumn()).getValues().forEach(function (r) {
      have[norm_(r[map.task - 1])] = 1;
      var m = /^T(\d+)$/.exec(String(r[map.id - 1])); if (m) max = Math.max(max, Number(m[1]));
    });
    var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), width = sheet.getLastColumn(), out = [];
    LIST.forEach(function (s) {
      if (have[norm_(s[1])]) return;
      max++;
      var row = new Array(width).fill('');
      row[map.id - 1] = 'T' + ('000' + max).slice(-3);
      row[map.category - 1] = s[0]; row[map.task - 1] = s[1]; row[map.status - 1] = 'Pending'; row[map.priority - 1] = s[2];
      row[map.function - 1] = s[3]; row[map.nextactionremarks - 1] = s[4];
      row[map.lastupdatedby - 1] = 'Suggested'; row[map.lastupdated - 1] = stamp;
      out.push(row);
    });
    if (out.length) sheet.getRange(sheet.getLastRow() + 1, 1, out.length, width).setValues(out);
    props.setProperty(KEY, '1');
  } finally { lock.releaseLock(); }
}


/* ---------------- vendors & payments ---------------- */
function guessSplit_(cat, name) {
  var t = String(cat + ' ' + name).toLowerCase();
  return /jewel|cloth|outfit|lehenga|sherwani|saree|suit|dress|gift|shopping|personal/.test(t) ? 'Personal' : 'Shared';
}
function vendorSetup_(vs, held) {
  var map = headerMap_(vs);
  var missing = VENDOR_EXTRA.some(function (n) { return !map[norm_(n)]; });
  var idCol = map.id, rows = vs.getLastRow(), missingId = false;
  if (!missing && idCol && rows >= 2) {
    var data0 = vs.getRange(2, 1, rows - 1, vs.getLastColumn()).getValues();
    missingId = data0.some(function (r) { return (String(r[(map.category || 1) - 1] || '').trim() || String(r[(map.vendor || 1) - 1] || '').trim()) && !String(r[idCol - 1] || ''); });
  }
  var props = PropertiesService.getScriptProperties();
  if (!missing && !missingId && props.getProperty('PAYMIG_V1')) return map;
  var lock = LockService.getScriptLock(); if (!held) lock.waitLock(25000);
  try {
    map = headerMap_(vs);
    var col = vs.getLastColumn();
    VENDOR_EXTRA.forEach(function (n) { if (!map[norm_(n)]) { col++; vs.getRange(1, col).setValue(n).setFontWeight('bold'); } });
    map = headerMap_(vs);
    rows = vs.getLastRow();
    if (rows >= 2) {
      var data = vs.getRange(2, 1, rows - 1, vs.getLastColumn()).getValues(), max = 0;
      data.forEach(function (r) { var m = /^V(\d+)$/.exec(String(r[map.id - 1])); if (m) max = Math.max(max, Number(m[1])); });
      data.forEach(function (r, i) {
        var cat = String(r[(map.category || 1) - 1] || '').trim(), nm = String(r[(map.vendor || 1) - 1] || '').trim();
        if (!cat && !nm) return;
        var row = i + 2;
        if (!String(r[map.id - 1] || '')) { max++; vs.getRange(row, map.id).setValue('V' + ('000' + max).slice(-3)); }
        if (!String(r[map.split - 1] || '')) vs.getRange(row, map.split).setValue(guessSplit_(cat, nm));
        if (String(r[map.shrijeetshare - 1] || '') === '') vs.getRange(row, map.shrijeetshare).setValue(50);
      });
    }
    if (!props.getProperty('PAYMIG_V1')) {
      var ps = paySheet_(), vals = vs.getLastRow() >= 2 ? vs.getRange(2, 1, vs.getLastRow() - 1, vs.getLastColumn()).getValues() : [], out = [], n = ps.getLastRow() - 1, stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
      vals.forEach(function (r) {
        var id = String(r[map.id - 1] || ''), adv = num_(r[(map.advancepaid || 1) - 1]);
        if (!id || adv <= 0 || !map.advancepaid) return;
        var pb = String(r[(map.paidby || 1) - 1] || '').toLowerCase();
        var by = /shivangi/.test(pb) ? 'Shivangi' : /family|papa|mummy|father|mother|parent/.test(pb) ? 'Family' : 'Shrijeet';
        n++; out.push(['P' + ('0000' + n).slice(-4), id, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), adv, by, 'Advance already paid (from your sheet)', 'Planner', stamp]);
      });
      if (out.length) ps.getRange(ps.getLastRow() + 1, 1, out.length, PAY_HEADERS.length).setValues(out);
      props.setProperty('PAYMIG_V1', '1');
    }
    return headerMap_(vs);
  } finally { if (!held) lock.releaseLock(); }
}
function nextPayId_(ps) {
  var max = 0;
  if (ps.getLastRow() >= 2) ps.getRange(2, 1, ps.getLastRow() - 1, 1).getValues().forEach(function (r) { var m = /^P(\d+)$/.exec(String(r[0])); if (m) max = Math.max(max, Number(m[1])); });
  return 'P' + ('0000' + (max + 1)).slice(-4);
}
function paySheet_() {
  var ss = SpreadsheetApp.getActive(), ps = ss.getSheetByName(PAY_SHEET);
  if (!ps) { ps = ss.insertSheet(PAY_SHEET); ps.getRange(1, 1, 1, PAY_HEADERS.length).setValues([PAY_HEADERS]).setFontWeight('bold'); ps.setFrozenRows(1); }
  return ps;
}
function readPayments_() {
  var ps = SpreadsheetApp.getActive().getSheetByName(PAY_SHEET), out = [];
  if (!ps || ps.getLastRow() < 2) return out;
  ps.getRange(2, 1, ps.getLastRow() - 1, PAY_HEADERS.length).getValues().forEach(function (r) {
    if (!String(r[0] || '')) return;
    out.push({ id: String(r[0]), vid: String(r[1]), date: str_(r[2]), amount: num_(r[3]), by: String(r[4] || ''), note: String(r[5] || ''), addedBy: String(r[6] || '') });
  });
  return out;
}
function syncVendor_(vs, map, vid) {
  var row = findRowById_(vs, map, vid); if (!row) return;
  var sum = 0; readPayments_().forEach(function (p) { if (p.vid === vid) sum += p.amount; });
  if (map.advancepaid) vs.getRange(row, map.advancepaid).setValue(sum);
  if (map.balance) { var c = vs.getRange(row, map.balance); if (!c.getFormula()) c.setValue(num_(vs.getRange(row, map.quotedamount).getValue()) - sum); }
}
var VFIELD = { category: 'category', vendor: 'vendor', contact: 'contact', deliverables: 'deliverables', status: 'status', remarks: 'remarks' };
function applyVendorFields_(vs, map, row, f) {
  Object.keys(f).forEach(function (k) {
    if (VFIELD[k] && map[VFIELD[k]]) vs.getRange(row, map[VFIELD[k]]).setValue(safe_(f[k], k === 'remarks' || k === 'deliverables' ? 500 : 120));
    else if (k === 'quoted' && map.quotedamount) vs.getRange(row, map.quotedamount).setValue(Math.max(0, num_(f[k])));
    else if (k === 'split') vs.getRange(row, map.split).setValue(f[k] === 'Personal' ? 'Personal' : 'Shared');
    else if (k === 'pct') vs.getRange(row, map.shrijeetshare).setValue(Math.min(100, Math.max(0, num_(f[k]))));
  });
}
function vendorAdd_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var vs = SpreadsheetApp.getActive().getSheetByName(VENDOR_SHEET); if (!vs) return { success: false, message: 'Tab "' + VENDOR_SHEET + '" not found' };
    var map = vendorSetup_(vs, true), f = b.fields || {};
    if (!safe_(f.vendor, 120) && !safe_(f.category, 120)) return { success: false, message: 'Please type a vendor or item name.' };
    var rows = vs.getLastRow(), max = 0;
    if (rows >= 2) vs.getRange(2, map.id, rows - 1, 1).getValues().forEach(function (r) { var m = /^V(\d+)$/.exec(String(r[0])); if (m) max = Math.max(max, Number(m[1])); });
    var id = 'V' + ('000' + (max + 1)).slice(-3), row = new Array(vs.getLastColumn()).fill('');
    row[map.id - 1] = id; row[map.split - 1] = 'Shared'; row[map.shrijeetshare - 1] = 50;
    vs.appendRow(row);
    var r = vs.getLastRow();
    applyVendorFields_(vs, map, r, f);
    syncVendor_(vs, map, id);
    clearCache_();
    return { success: true, id: id };
  } finally { lock.releaseLock(); }
}
function vendorUpdate_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var vs = SpreadsheetApp.getActive().getSheetByName(VENDOR_SHEET), map = vendorSetup_(vs, true), row = findRowById_(vs, map, b.id);
    if (!row) return { success: false, message: 'Vendor not found (was it deleted in the sheet?)' };
    applyVendorFields_(vs, map, row, b.fields || {});
    syncVendor_(vs, map, b.id); clearCache_();
    return { success: true };
  } finally { lock.releaseLock(); }
}
function vendorDelete_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var vs = SpreadsheetApp.getActive().getSheetByName(VENDOR_SHEET), map = vendorSetup_(vs, true), row = findRowById_(vs, map, b.id);
    if (!row) return { success: true };
    var ps = paySheet_();
    if (ps.getLastRow() >= 2) {
      var v = ps.getRange(2, 1, ps.getLastRow() - 1, 2).getValues();
      for (var i = v.length - 1; i >= 0; i--) if (String(v[i][1]) === String(b.id)) ps.deleteRow(i + 2);
    }
    vs.deleteRow(row); clearCache_();
    return { success: true };
  } finally { lock.releaseLock(); }
}
function cleanBy_(v) { v = String(v || ''); return ['Shrijeet', 'Shivangi', 'Family'].indexOf(v) >= 0 ? v : 'Shrijeet'; }
function cleanDate_(v) { return /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) ? String(v) : Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function payAdd_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var f = b.fields || {}, amt = Number(f.amount);
    if (!isFinite(amt) || amt === 0) return { success: false, message: 'Please enter the amount.' };
    var vs = SpreadsheetApp.getActive().getSheetByName(VENDOR_SHEET), map = vendorSetup_(vs, true), vid = String(f.vid || '');
    if (vid !== 'SETTLE' && !findRowById_(vs, map, vid)) return { success: false, message: 'Vendor not found.' };
    var ps = paySheet_(), id = nextPayId_(ps);
    ps.appendRow([id, vid, cleanDate_(f.date), amt, cleanBy_(f.by), safe_(f.note, 200), who_(b), Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')]);
    ps.getRange(ps.getLastRow(), 3).setNumberFormat('@').setValue(cleanDate_(f.date));
    if (vid !== 'SETTLE') syncVendor_(vs, map, vid);
    clearCache_();
    return { success: true, id: id };
  } finally { lock.releaseLock(); }
}
function findPayRow_(ps, id) {
  if (ps.getLastRow() < 2) return 0;
  var v = ps.getRange(2, 1, ps.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(id)) return i + 2;
  return 0;
}
function payUpdate_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var ps = paySheet_(), row = findPayRow_(ps, b.id); if (!row) return { success: false, message: 'Payment not found.' };
    var f = b.fields || {}, vid = String(ps.getRange(row, 2).getValue());
    if (f.amount !== undefined) { var a = Number(f.amount); if (!isFinite(a) || a === 0) return { success: false, message: 'Please enter the amount.' }; ps.getRange(row, 4).setValue(a); }
    if (f.date !== undefined) ps.getRange(row, 3).setNumberFormat('@').setValue(cleanDate_(f.date));
    if (f.by !== undefined) ps.getRange(row, 5).setValue(cleanBy_(f.by));
    if (f.note !== undefined) ps.getRange(row, 6).setValue(safe_(f.note, 200));
    ps.getRange(row, 7).setValue(who_(b));
    if (vid !== 'SETTLE') { var vs = SpreadsheetApp.getActive().getSheetByName(VENDOR_SHEET); syncVendor_(vs, vendorSetup_(vs, true), vid); }
    clearCache_();
    return { success: true };
  } finally { lock.releaseLock(); }
}
function payDelete_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var ps = paySheet_(), row = findPayRow_(ps, b.id); if (!row) return { success: true };
    var vid = String(ps.getRange(row, 2).getValue()); ps.deleteRow(row);
    if (vid !== 'SETTLE') { var vs = SpreadsheetApp.getActive().getSheetByName(VENDOR_SHEET); syncVendor_(vs, vendorSetup_(vs, true), vid); }
    clearCache_();
    return { success: true };
  } finally { lock.releaseLock(); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
