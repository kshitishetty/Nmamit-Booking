// ===== EDIT THESE (1 or 2 fixed admins) =====
const ADMINS = { principal: 'ChangeMe-1!', coordinator: 'ChangeMe-2!' };
// ============================================
const DOMAINS = ['nmamit.in', 'nitte.in'];
const DEFAULT_VENUES = ['Sambhram','Shambhavi','Sadananda','Amphitheatre','Phalguni','Abhyuday Hall','Conference Room','Vista Lab'];
const OPEN = '08:00', CLOSE = '20:00';
const HEAD = {
  Venues: ['id','name','capacity'],
  Bookings: ['id','venue','date','start','end','title','org','att','phone','notes','status','email','created','grp'],
  Blocks: ['id','venue','date','reason']
};
const _s = {};
const uid = () => Utilities.getUuid().slice(0, 8);
const today = () => Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');

function authorize() { MailApp.getRemainingDailyQuota(); SpreadsheetApp.getActive(); } // run once to grant permissions

function sh(n) {
  if (_s[n]) return _s[n];
  const h = HEAD[n], ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.getRange(2, 1, 2000, h.length).setNumberFormat('@'); s.setFrozenRows(1); }
  if (s.getRange(1, 1, 1, h.length).getValues()[0].join('|') !== h.join('|'))
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  if (n === 'Venues' && s.getLastRow() < 2) DEFAULT_VENUES.forEach(v => s.appendRow([uid(), v, '']));
  return (_s[n] = s);
}
function rows(n) {
  const v = sh(n).getDataRange().getDisplayValues(), h = v.shift();
  return v.filter(r => r[0]).map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function rowOf(n, id) {
  const s = sh(n), last = s.getLastRow(); if (last < 2) return 0;
  const i = s.getRange(2, 1, last - 1, 1).getValues().flat().indexOf(id); return i < 0 ? 0 : i + 2;
}
const put = (n, o) => sh(n).appendRow(HEAD[n].map(k => o[k] ?? ''));
const setRow = (n, r, o) => sh(n).getRange(r, 1, 1, HEAD[n].length).setValues([HEAD[n].map(k => o[k] ?? '')]);
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function doGet() { return out({ ok: true }); }

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try { return out(Object.assign({ ok: true }, route(JSON.parse(e.postData.contents)))); }
  catch (err) { return out({ ok: false, error: err.message || String(err) }); }
  finally { lock.releaseLock(); }
}
const fail = m => { throw new Error(m); };

function route(d) {
  const a = d.action, C = CacheService.getScriptCache();
  if (a === 'sendotp') {
    const em = String(d.email || '').trim().toLowerCase();
    if (!DOMAINS.some(x => em.endsWith('@' + x))) fail('Use your @nmamit.in or @nitte.in email.');
    if (C.get('rl_' + em)) fail('Please wait a minute before asking for another code.');
    const code = String(Math.floor(100000 + Math.random() * 900000));
    C.put('otp_' + em, code, 600); C.put('rl_' + em, '1', 60); C.remove('try_' + em);
    MailApp.sendEmail({ to: em, name: 'NMAMIT Bookings', subject: 'Your NMAMIT Bookings code: ' + code,
      body: 'Your sign-in code is ' + code + '. It is valid for 10 minutes.\nIf you did not ask for it, ignore this email.' });
    return {};
  }
  if (a === 'verify') {
    const em = String(d.email || '').trim().toLowerCase(), tries = +(C.get('try_' + em) || 0);
    if (tries >= 5) fail('Too many wrong codes. Request a new code.');
    if (!C.get('otp_' + em) || C.get('otp_' + em) !== String(d.code || '').trim()) { C.put('try_' + em, String(tries + 1), 600); fail('Wrong or expired code.'); }
    C.remove('otp_' + em);
    const t = Utilities.getUuid(); C.put('t_' + t, JSON.stringify({ email: em, admin: false }), 21600);
    return { token: t };
  }
  if (a === 'adminlogin') {
    const u = String(d.username || '').trim().toLowerCase();
    if (!ADMINS[u] || ADMINS[u] !== d.password) fail('Wrong username or password.');
    const t = Utilities.getUuid(); C.put('t_' + t, JSON.stringify({ email: u, admin: true }), 21600);
    return { token: t };
  }
  const raw = C.get('t_' + d.token); if (!raw) fail('Session expired. Please sign in again.');
  const u = JSON.parse(raw);
  if (a === 'signout') { C.remove('t_' + d.token); return {}; }
  if (a === 'list') return listFor(u);

  if (a === 'create') {
    const f = clean(d); chk(f, false);
    const n = Math.min(12, Math.max(1, +d.weeks || 1)), dates = [];
    for (let i = 0; i < n; i++) dates.push(addDays(f.date, 7 * i));
    dates.forEach(dt => { const m = clash(Object.assign({}, f, { date: dt }), ''); if (m) fail(m); });
    const g = uid();
    dates.forEach(dt => put('Bookings', Object.assign({}, f, { id: uid(), date: dt, status: u.admin ? 'approved' : 'pending', email: u.email, created: new Date().toISOString(), grp: g })));
    return {};
  }
  if (a === 'update' || a === 'cancel') {
    const b = rows('Bookings').find(x => x.id === d.id); if (!b) fail('Booking not found.');
    if (!u.admin && b.email !== u.email) fail('Not allowed.');
    const r = rowOf('Bookings', d.id);
    if (a === 'cancel') { sh('Bookings').deleteRow(r); return {}; }
    const f = clean(d); chk(f, f.date === b.date); const m = clash(f, d.id); if (m) fail(m);
    setRow('Bookings', r, Object.assign({}, b, f, { status: u.admin ? b.status : 'pending' }));
    return {};
  }
  if (!u.admin) fail('Admin only.');
  if (a === 'decide') {
    const b = rows('Bookings').find(x => x.id === d.id); if (!b) fail('Booking not found.');
    if (d.status === 'approved') { const m = clash(b, b.id, true); if (m) fail(m); }
    setRow('Bookings', rowOf('Bookings', d.id), Object.assign({}, b, { status: d.status === 'approved' ? 'approved' : 'rejected' }));
    return {};
  }
  if (a === 'block') { if (!/^\d{4}-\d\d-\d\d$/.test(d.date || '')) fail('Pick a date.'); put('Blocks', { id: uid(), venue: String(d.venue || 'All'), date: d.date, reason: String(d.reason || '').slice(0, 50) }); return {}; }
  if (a === 'unblock') { const r = rowOf('Blocks', d.id); if (r) sh('Blocks').deleteRow(r); return {}; }
  if (a === 'vsave') {
    const name = String(d.name || '').trim().slice(0, 40), cap = String(d.capacity || '').slice(0, 8);
    if (!name) fail('Enter a venue name.');
    const vs = rows('Venues'), ex = d.id && vs.find(v => v.id === d.id);
    if (vs.some(v => v.name.toLowerCase() === name.toLowerCase() && v.id !== d.id)) fail('That venue already exists.');
    if (ex) {
      setRow('Venues', rowOf('Venues', d.id), { id: d.id, name: name, capacity: cap });
      if (ex.name !== name) {
        const bs = sh('Bookings'), n = bs.getLastRow();
        if (n > 1) { const rg = bs.getRange(2, 2, n - 1, 1); rg.setValues(rg.getValues().map(x => [x[0] === ex.name ? name : x[0]])); }
        const ks = sh('Blocks'), m = ks.getLastRow();
        if (m > 1) { const rg = ks.getRange(2, 2, m - 1, 1); rg.setValues(rg.getValues().map(x => [x[0] === ex.name ? name : x[0]])); }
      }
    } else put('Venues', { id: uid(), name: name, capacity: cap });
    return {};
  }
  if (a === 'vdel') { const r = rowOf('Venues', d.id); if (r) sh('Venues').deleteRow(r); return {}; }
  fail('Unknown action.');
}

function listFor(u) {
  const bookings = rows('Bookings').filter(b => u.admin || b.email === u.email || b.status !== 'rejected').map(b => {
    const own = b.email === u.email;
    if (u.admin || own) return Object.assign({}, b, { own: own });
    return { id: b.id, venue: b.venue, date: b.date, start: b.start, end: b.end, title: b.title, org: b.org, status: b.status, own: false };
  });
  return { admin: u.admin, venues: rows('Venues'), bookings: bookings, blocks: rows('Blocks') };
}
function clean(d) {
  const s = (x, n) => String(x == null ? '' : x).trim().slice(0, n);
  return { venue: s(d.venue, 40), title: s(d.title, 80), org: s(d.org, 60), date: s(d.date, 10), start: s(d.start, 5), end: s(d.end, 5), att: s(d.att, 6), phone: s(d.phone, 15), notes: s(d.notes, 120) };
}
function chk(f, skipPast) {
  if (!f.venue || !f.title || !/^\d{4}-\d\d-\d\d$/.test(f.date) || !/^\d\d:\d\d$/.test(f.start) || !/^\d\d:\d\d$/.test(f.end)) fail('Fill in venue, title, date and times.');
  if (f.end <= f.start) fail('End time must be after start time.');
  if (f.start < OPEN || f.end > CLOSE) fail('Venues are open 8 AM to 8 PM.');
  if (!skipPast && f.date < today()) fail('Pick today or a future date.');
  if (!rows('Venues').some(v => v.name === f.venue)) fail('Unknown venue.');
}
function clash(f, skipId, ignoreBlocks) {
  const k = !ignoreBlocks && rows('Blocks').find(x => x.date === f.date && (x.venue === 'All' || x.venue === f.venue));
  if (k) return f.date + ' is blocked' + (k.reason ? ': ' + k.reason : '') + '.';
  if (rows('Bookings').some(b => b.id !== skipId && b.venue === f.venue && b.date === f.date && b.status !== 'rejected' && f.start < b.end && f.end > b.start))
    return f.venue + ' is already booked at that time on ' + f.date + '.';
  return '';
}
function addDays(s, n) { const d = new Date(s + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
