/* Paramotor engine logbook — offline-first, all data stays on the device.
   Excel (.xlsx) is the backup format, compatible with the original "Paramotor Flights.xlsx". */
'use strict';

const STORE_KEY = 'ppg-logbook-v1';
const SHEET_LOG = 'LOG BOOK HOURS';
const SHEET_MT = 'Maintenance Tracker';
const SHEET_ML = 'Maintenance Log';
const SHEET_SITES = 'Sites';
const XL_EPOCH = Date.UTC(1899, 11, 30);

// Manufacturer intervals for the Vittorazi Moster 185 MY19 (taken from the user's sheet).
// [name, replace every (h), or every (months), notes, check every (h)]
const DEFAULT_COMPONENTS = [
  ['Spark Plug', 25, 0, '', 0],
  ['Spark Plug Connector', 150, 0, 'Check every 50h', 50],
  ['Carburettor', 100, 0, 'Only to clean, check it every 25h', 25],
  ['Carburettor Membranes', 100, 12, 'Replace (or after 1 year)', 0],
  ['Airbox Snaplock', 25, 0, 'Check it every 10h', 10],
  ['Airbox Sponge & Sleeve', 100, 12, 'Check and Clean every 25h; replace at 100h or after 1 year', 25],
  ['Reed Valve Petals', 150, 0, 'Check it every 100h', 100],
  ['Pull Starter System', 100, 0, 'Rope, spring & hooks (or full system)', 0],
  ['Kit Exhaust Bushing (1st joint)', 50, 0, 'Check every 25h', 25],
  ['Exhaust Bushing Springs (2nd joint)', 50, 0, 'Check every 25h ,at 50h clean and replace springs', 25],
  ['Exhaust manifold with springs', 150, 0, '', 0],
  ['Soundproofing Material – Silencer', 150, 0, 'Check for deterioration', 0],
  ['Silencer Rubber', 150, 12, 'Replace if cracked or after 1 year', 0],
  ['Gaskets (cylinder,carburettor, reed valve, exhaust, silencer)', 100, 0, 'Each time the component is disassembled', 0],
  ['Rubber Mountings (Engine & Exhaust)', 150, 12, 'Check for cracks; replace after 1 year if needed', 0],
  ['Piston', 150, 0, 'at 100h clean and measure, at 150h replace', 100],
  ['Piston Roller Bearing', 150, 0, '', 0],
  ['Head & Cylinder', 100, 0, 'Only to measure and clean exhaust port, decompressor hole, head', 0],
  ['O-ring head', 100, 0, 'Each time the component is disassembled', 0],
  ['Belt', 100, 0, 'Re-tension belt every 25h', 25],
  ['Reduction Bearings', 150, 0, '', 0],
  ['Centrifugal Clutch', 150, 0, '', 0],
  ['Clutch Bell', 150, 0, 'Only to measure, inspect for wear/glazing', 0],
  ['Oil Seal Carter Case', 150, 0, '', 0],
  ['Crankshaft Bearings', 150, 0, 'Replace both main bearings', 0],
  ['Crankshaft', 150, 0, 'Only to measure', 0],
];

/* ---------------- utils ---------------- */
const $ = (s, el = document) => el.querySelector(s);
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = n => String(n).padStart(2, '0');
const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const fmtDate = iso => { if (!iso) return ''; const [y, m, d] = iso.split('-'); return `${d}/${m}/${y.slice(2)}`; };
const HE_MONTHS = ['ינו׳', 'פבר׳', 'מרץ', 'אפר׳', 'מאי', 'יוני', 'יולי', 'אוג׳', 'ספט׳', 'אוק׳', 'נוב׳', 'דצמ׳'];

/** hours (float) -> "h:mm" */
function hm(hours) {
  const neg = hours < 0;
  let mins = Math.round(Math.abs(hours) * 60);
  const h = Math.floor(mins / 60), m = mins % 60;
  return (neg ? '-' : '') + h + ':' + pad(m);
}
/** "h:mm" | "h.mm decimal" | number -> hours (float) or NaN */
function parseHM(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').trim();
  if (!s) return NaN;
  const m = s.match(/^(-?\d+)\s*[:：]\s*(\d{1,2})$/);
  if (m) return (m[1].startsWith('-') ? -1 : 1) * (Math.abs(+m[1]) + (+m[2]) / 60);
  const n = Number(s.replace(',', '.'));
  return Number.isFinite(n) ? n : NaN;
}
function monthsBetween(isoA, isoB) {
  const a = new Date(isoA), b = new Date(isoB);
  return (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth()) + (b.getDate() - a.getDate()) / 30;
}
function addMonthsISO(iso, months) {
  const d = new Date(iso + 'T00:00:00');
  d.setMonth(d.getMonth() + months);
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function daysUntil(iso) { return Math.round((new Date(iso + 'T00:00:00') - new Date(todayISO() + 'T00:00:00')) / 86400000); }
function haversineKm(a, b) {
  const R = 6371, rad = Math.PI / 180;
  const dLat = (b.lat - a.lat) * rad, dLon = (b.lon - a.lon) * rad;
  const x = Math.sin(dLat / 2) ** 2 + Math.cos(a.lat * rad) * Math.cos(b.lat * rad) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(x));
}

/* ---------------- state ---------------- */
function freshState() {
  return {
    version: 1,
    engineName: 'Vittorazi Moster 185 MY19',
    flights: [],       // {id, date:'YYYY-MM-DD', min:int, site, notes}
    components: DEFAULT_COMPONENTS.map(([name, intervalH, intervalMonths, notes, checkH]) => ({ id: uid(), name, intervalH, intervalMonths, lastResetH: 0, lastResetDate: '', notes, checkH, lastCheckH: 0 })),
    maint: [],         // {id, date, hours, items:[componentId replaced], checks:[componentId checked], note}
    sites: [],         // {name, lat, lon}
    settings: { reserveRepack: '', reserveMonths: 6, autoBackup: true },
    meta: { lastBackup: null, changesSinceBackup: 0 },
  };
}
let S = load();
function load() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (raw) { const s = JSON.parse(raw); return Object.assign(freshState(), s, { settings: Object.assign(freshState().settings, s.settings), meta: Object.assign(freshState().meta, s.meta) }); }
  } catch (e) { console.error(e); }
  return freshState();
}
function save(changed = true) {
  if (changed) S.meta.changesSinceBackup = (S.meta.changesSinceBackup || 0) + 1;
  localStorage.setItem(STORE_KEY, JSON.stringify(S));
  updateHeader();
  if (changed) scheduleLinkedWrite();
}
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});

const totalHours = () => S.flights.reduce((a, f) => a + (f.min || 0), 0) / 60;
const sortedFlights = () => [...S.flights].sort((a, b) => (b.date || '').localeCompare(a.date || '') || 0);
const knownSites = () => {
  const count = {};
  for (const f of S.flights) { const s = (f.site || '').trim(); if (s) count[s] = (count[s] || 0) + 1; }
  return Object.entries(count).sort((a, b) => b[1] - a[1]).map(e => e[0]);
};

function compStatus(c, total = totalHours()) {
  const since = total - (c.lastResetH || 0);
  const remaining = (c.intervalH || 0) - since;
  let dueDate = null, daysLeft = null;
  if (c.intervalMonths && c.lastResetDate) { dueDate = addMonthsISO(c.lastResetDate, c.intervalMonths); daysLeft = daysUntil(dueDate); }
  let level = remaining < 10 ? 'bad' : remaining <= 25 ? 'warn' : 'ok';
  if (daysLeft !== null) { if (daysLeft < 0) level = 'bad'; else if (daysLeft <= 30 && level === 'ok') level = 'warn'; }
  const pct = c.intervalH ? Math.max(0, Math.min(1, since / c.intervalH)) : 0;
  // periodic check (inspection / re-tension) — a replacement also counts as a check
  let checkRemaining = null, checkLevel = null;
  if (c.checkH > 0) {
    const lastCheck = Math.max(c.lastCheckH || 0, c.lastResetH || 0);
    checkRemaining = c.checkH - (total - lastCheck);
    checkLevel = checkRemaining < 0 ? 'bad' : checkRemaining <= Math.max(2, c.checkH * 0.2) ? 'warn' : 'ok';
  }
  const rank = { ok: 0, warn: 1, bad: 2 };
  const worst = checkLevel && rank[checkLevel] > rank[level] ? checkLevel : level;
  // sort key: hours until the next thing (check or replacement) is due
  const nextDue = checkRemaining !== null ? Math.min(remaining, checkRemaining) : remaining;
  return { since, remaining, dueDate, daysLeft, level, pct, checkRemaining, checkLevel, worst, nextDue, overdue: remaining < 0 || (daysLeft !== null && daysLeft < 0) };
}
function reserveStatus() {
  const r = S.settings.reserveRepack; if (!r) return null;
  const last = r.length === 7 ? r + '-01' : r;
  const due = addMonthsISO(last, S.settings.reserveMonths || 6);
  const days = daysUntil(due);
  return { last, due, days, level: days < 0 ? 'bad' : days <= 30 ? 'warn' : 'ok' };
}

/* ---------------- UI shell ---------------- */
let tab = 'dash';
const view = $('#view');
document.querySelectorAll('#tabbar button').forEach(b => b.addEventListener('click', () => go(b.dataset.tab)));
function go(t, arg) {
  tab = t;
  document.querySelectorAll('#tabbar button').forEach(b => b.classList.toggle('active', b.dataset.tab === t));
  window.scrollTo(0, 0);
  ({ dash: renderDash, flights: renderFlights, add: () => renderFlightForm(arg), maint: renderMaint, settings: renderSettings })[t]();
}
function toast(msg, ms = 2600) {
  const t = $('#toast'); t.textContent = msg; t.classList.remove('hidden');
  clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.add('hidden'), ms);
}
function openModal(html, onMount) {
  $('#modalCard').innerHTML = html; $('#modal').classList.remove('hidden');
  onMount && onMount($('#modalCard'));
}
function closeModal() { $('#modal').classList.add('hidden'); $('#modalCard').innerHTML = ''; }
$('#modal').addEventListener('click', e => { if (e.target.id === 'modal') closeModal(); });
function confirmBox(msg, okText = 'אישור', danger = false) {
  return new Promise(res => {
    openModal(`<h3>${esc(msg)}</h3><div class="btns"><button class="btn" data-a="no">ביטול</button><button class="btn ${danger ? 'danger' : 'primary'}" data-a="yes">${esc(okText)}</button></div>`, card => {
      card.querySelectorAll('[data-a]').forEach(b => b.onclick = () => { closeModal(); res(b.dataset.a === 'yes'); });
    });
  });
}

function updateHeader() {
  $('#engineName').textContent = S.engineName || 'יומן מנוע';
  const el = $('#backupStatus');
  const n = S.meta.changesSinceBackup || 0;
  if (linked.handle && linked.ok) { el.textContent = 'מסונכרן לקובץ אקסל'; el.classList.remove('stale'); }
  else if (!S.meta.lastBackup) { el.textContent = S.flights.length ? 'עדיין לא נעשה גיבוי לאקסל' : ''; el.classList.toggle('stale', !!S.flights.length); }
  else {
    const d = new Date(S.meta.lastBackup);
    el.textContent = n ? `${n} שינויים מאז הגיבוי (${fmtDate(d.toISOString().slice(0, 10))})` : `גובה לאקסל ${fmtDate(d.toISOString().slice(0, 10))}`;
    el.classList.toggle('stale', n > 0);
  }
}
$('#quickBackup').addEventListener('click', () => exportExcel());

/* ---------------- Dashboard ---------------- */
function renderDash() {
  const total = totalHours();
  if (!S.flights.length) {
    view.innerHTML = `
      <div class="card hero"><div class="label">סה״כ שעות מנוע</div><div class="big num">0:00</div></div>
      <div class="card"><h2>מתחילים</h2>
        <p class="muted small">יש לך כבר קובץ אקסל עם היסטוריית טיסות? אפשר לייבא אותו ישירות, כולל גיליון התחזוקה.</p>
        <div class="btns"><button class="btn primary" id="impX">ייבוא מקובץ אקסל</button><button class="btn" id="addF">טיסה ראשונה</button></div>
      </div>`;
    $('#impX').onclick = () => $('#fileExcel').click();
    $('#addF').onclick = () => go('add');
    return;
  }
  const comps = S.components.map(c => ({ c, s: compStatus(c, total) })).sort((a, b) => a.s.nextDue - b.s.nextDue);
  const due = comps.filter(x => x.s.worst !== 'ok');
  const yearNow = String(new Date().getFullYear());
  const yearFlights = S.flights.filter(f => (f.date || '').startsWith(yearNow));
  const yearH = yearFlights.reduce((a, f) => a + f.min, 0) / 60;
  const last = sortedFlights()[0];
  const rs = reserveStatus();
  const alerts = [];
  if ((S.meta.changesSinceBackup || 0) >= 1 && !(linked.handle && linked.ok))
    alerts.push(`<div class="alert info"><span>💾</span><span>יש ${S.meta.changesSinceBackup} שינויים שעוד לא נשמרו לקובץ האקסל.</span><button class="btn sm" id="bk">גבה עכשיו</button></div>`);
  if (rs && rs.level !== 'ok')
    alerts.push(`<div class="alert ${rs.level}"><span>🪂</span><span>${rs.days < 0 ? `קיפול רזרבי באיחור של ${-rs.days} ימים` : `קיפול רזרבי בעוד ${rs.days} ימים`} (אחרון: ${fmtDate(rs.last).slice(3)})</span></div>`);

  view.innerHTML = `
    <div class="card hero">
      <div class="label">סה״כ שעות מנוע</div>
      <div class="big num">${hm(total)}</div>
      <div class="stats">
        <div class="stat"><b class="num">${S.flights.length}</b><span>טיסות</span></div>
        <div class="stat"><b class="num">${hm(yearH)}</b><span>שעות ב-${yearNow}</span></div>
        <div class="stat"><b class="num">${last ? fmtDate(last.date) : '—'}</b><span>טיסה אחרונה</span></div>
      </div>
    </div>
    ${alerts.join('')}
    <div class="card">
      <div class="row"><h2>${due.length ? `דורש תשומת לב (${due.length})` : 'כל החלקים בסדר ✓'}</h2><button class="btn sm" id="toMaint">כל החלקים</button></div>
      ${(due.length ? due : comps.slice(0, 3)).map(x => compRow(x.c, x.s)).join('')}
    </div>
    ${rs && rs.level === 'ok' ? `<div class="card"><div class="row"><span>🪂 קיפול רזרבי הבא</span><b class="num">${fmtDate(rs.due)}</b></div></div>` : ''}
    <button class="btn primary block" id="addF">＋ הוסף טיסה</button>`;
  $('#toMaint').onclick = () => go('maint');
  $('#addF').onclick = () => go('add');
  const bk = $('#bk'); if (bk) bk.onclick = () => exportExcel();
  view.querySelectorAll('.comp').forEach(el => el.onclick = () => openComponent(el.dataset.id));
}
function compRow(c, s) {
  const remTxt = s.remaining < 0 ? `באיחור ${hm(-s.remaining)}` : `נותרו ${hm(s.remaining)}`;
  const dateTxt = s.dueDate ? ` · עד ${fmtDate(s.dueDate)}${s.daysLeft < 0 ? ' (עבר)' : ''}` : '';
  return `<div class="comp" data-id="${c.id}">
    <div class="row"><div class="name">${esc(c.name)}</div><div class="rem s-${s.level}"><span class="num">${remTxt}</span></div></div>
    <div class="meta">כל ${c.intervalH} ש׳${c.intervalMonths ? ` או ${c.intervalMonths} ח׳` : ''} · הוחלף ב-<span class="num">${hm(c.lastResetH || 0)}</span>${c.lastResetDate ? ` (${fmtDate(c.lastResetDate)})` : ''} · הבא ב-<span class="num">${hm((c.lastResetH || 0) + c.intervalH)}</span>${dateTxt}</div>
    ${s.checkRemaining !== null ? `<div class="meta">🔍 בדיקה כל ${c.checkH} ש׳ · <span class="s-${s.checkLevel} num">${s.checkRemaining < 0 ? `בדיקה באיחור ${hm(-s.checkRemaining)}` : `בדיקה בעוד ${hm(s.checkRemaining)}`}</span></div>` : ''}
    <div class="bar"><i class="b-${s.level}" style="width:${Math.round(s.pct * 100)}%"></i></div>
  </div>`;
}

/* ---------------- Flights ---------------- */
let flightQuery = '';
function renderFlights() {
  const q = flightQuery.trim();
  const list = sortedFlights().filter(f => !q || (f.site || '').includes(q) || (f.notes || '').includes(q) || fmtDate(f.date).includes(q));
  const byYear = {};
  for (const f of list) (byYear[(f.date || '????').slice(0, 4)] ||= []).push(f);
  view.innerHTML = `
    <input type="text" class="search" id="q" placeholder="חיפוש לפי מקום, הערה או תאריך" value="${esc(flightQuery)}">
    ${list.length ? '' : '<div class="empty">אין טיסות להצגה</div>'}
    ${Object.keys(byYear).sort().reverse().map(y => {
      const fl = byYear[y]; const h = fl.reduce((a, f) => a + f.min, 0) / 60;
      return `<div class="year"><b>${y}</b><span>${fl.length} טיסות · <span class="num">${hm(h)}</span> ש׳</span></div>
      <div class="card list">${fl.map(f => {
        const [, m, d] = (f.date || '--').split('-');
        return `<div class="flight" data-id="${f.id}">
          <div class="d"><b class="num">${+d || '?'}</b><span>${HE_MONTHS[+m - 1] || ''}</span></div>
          <div><div class="site">${esc(f.site || '—')}</div>${f.notes ? `<div class="note">${esc(f.notes)}</div>` : ''}</div>
          <div class="t num">${hm(f.min / 60)}</div></div>`;
      }).join('')}</div>`;
    }).join('')}`;
  const qi = $('#q');
  qi.oninput = () => { flightQuery = qi.value; const pos = qi.selectionStart; renderFlights(); const n = $('#q'); n.focus(); n.setSelectionRange(pos, pos); };
  view.querySelectorAll('.flight').forEach(el => el.onclick = () => go('add', el.dataset.id));
}

let pendingTrack = null; // filled by IGC/GPX import
function renderFlightForm(editId) {
  const f = editId ? S.flights.find(x => x.id === editId) : null;
  const t = !f && pendingTrack ? pendingTrack : null;
  const min = f ? f.min : t ? t.min : 0;
  const sites = knownSites();
  view.innerHTML = `
    <div class="card">
      <div class="row"><h2>${f ? 'עריכת טיסה' : 'טיסה חדשה'}</h2>
        ${f ? '' : '<button class="btn sm" id="impTrack">📥 ייבוא מ-Gaggle (IGC/GPX)</button>'}</div>
      ${t ? `<div class="track-info">📍 נטען מקובץ <b>${esc(t.file)}</b><br>זמן טיסה מזוהה: <b class="num">${hm(t.min / 60)}</b> (המראה ${t.takeoff} → נחיתה ${t.landing})${t.totalMin !== t.min ? ` · כל ההקלטה: <span class="num">${hm(t.totalMin / 60)}</span>` : ''}${t.siteMatch ? `<br>אתר מזוהה: <b>${esc(t.siteMatch)}</b>` : t.lat ? '<br>אתר חדש – רשום את שמו, המיקום יישמר לזיהוי אוטומטי בפעם הבאה' : ''}</div>` : ''}
      <label class="f"><span>תאריך</span><input type="date" id="fDate" value="${f ? f.date : t ? t.date : todayISO()}"></label>
      <label class="f"><span>זמן טיסה (שעות : דקות)</span>
        <div class="dur"><input type="number" inputmode="numeric" id="fH" min="0" max="12" value="${Math.floor(min / 60)}"><span class="colon">:</span><input type="number" inputmode="numeric" id="fM" min="0" max="59" value="${pad(min % 60)}"></div>
        <div class="dur-labels"><span>שעות</span><span>דקות</span></div>
      </label>
      <label class="f"><span>מקום המראה</span><input type="text" id="fSite" list="siteList" value="${esc(f ? f.site : t ? (t.siteMatch || '') : (sites[0] || ''))}" autocomplete="off">
        <datalist id="siteList">${sites.map(s => `<option value="${esc(s)}">`).join('')}</datalist>
        <div class="chips">${sites.slice(0, 6).map(s => `<span class="chip" data-site="${esc(s)}">${esc(s)}</span>`).join('')}</div>
      </label>
      <label class="f"><span>הערות (טיפול, תקלה וכו׳)</span><textarea id="fNotes">${esc(f ? f.notes : '')}</textarea></label>
      ${f ? '' : `<label class="f checks" style="margin-bottom:14px"><label style="border:0;padding:0"><input type="checkbox" id="fMaint"> בוצע טיפול / הוחלף חלק בטיסה הזו</label></label>`}
      <div class="btns">
        ${f ? '<button class="btn danger" id="del">מחק</button>' : ''}
        <button class="btn primary" id="saveF">${f ? 'שמור שינויים' : 'שמור טיסה'}</button>
      </div>
    </div>`;
  view.querySelectorAll('.chip').forEach(c => c.onclick = () => { $('#fSite').value = c.dataset.site; });
  const it = $('#impTrack'); if (it) it.onclick = () => $('#fileTrack').click();
  $('#saveF').onclick = () => {
    const h = parseInt($('#fH').value || '0', 10), m = parseInt($('#fM').value || '0', 10);
    const date = $('#fDate').value, site = $('#fSite').value.trim(), notes = $('#fNotes').value.trim();
    const total = (h || 0) * 60 + (m || 0);
    if (!date) return toast('חסר תאריך');
    if (total <= 0) return toast('זמן טיסה חייב להיות גדול מ-0');
    if (m > 59) return toast('דקות: 0–59');
    if (f) Object.assign(f, { date, min: total, site, notes });
    else S.flights.push({ id: uid(), date, min: total, site, notes });
    if (t && t.lat && site && !S.sites.some(s => s.name === site)) S.sites.push({ name: site, lat: t.lat, lon: t.lon });
    const wantMaint = !f && $('#fMaint') && $('#fMaint').checked;
    pendingTrack = null;
    save();
    toast(f ? 'הטיסה עודכנה' : `נשמר ✓ סה״כ ${hm(totalHours())} שעות`);
    if (wantMaint) { go('maint'); openMaintDialog(null, date, notes); }
    else if (f) go('flights');
    else { go('dash'); afterChangeBackup(); }
  };
  const del = $('#del');
  if (del) del.onclick = async () => {
    if (!(await confirmBox(`למחוק את הטיסה מ-${fmtDate(f.date)} (${hm(f.min / 60)})?`, 'מחק', true))) return;
    S.flights = S.flights.filter(x => x.id !== f.id); save(); toast('הטיסה נמחקה'); go('flights');
  };
}

/* ---------------- IGC / GPX import (Gaggle export) ---------------- */
$('#fileTrack').addEventListener('change', async e => {
  const files = [...e.target.files]; e.target.value = '';
  await handleTrackFiles(await Promise.all(files.map(async f => ({ name: f.name, text: await f.text() }))));
});
async function handleTrackFiles(files) {
  if (!files.length) return;
  const parsed = [];
  for (const file of files) {
    try { const p = parseTrack(file.text, file.name); if (p) parsed.push(p); else toast(`לא נמצאו נקודות בקובץ ${file.name}`); }
    catch (err) { console.error(err); toast(`שגיאה בקריאת ${file.name}`); }
  }
  if (!parsed.length) return;
  if (parsed.length === 1) { pendingTrack = parsed[0]; go('add'); return; }
  // several files: add them all at once (skip duplicates by date+duration)
  let added = 0;
  for (const p of parsed) {
    if (S.flights.some(f => f.date === p.date && Math.abs(f.min - p.min) <= 2)) continue;
    S.flights.push({ id: uid(), date: p.date, min: p.min, site: p.siteMatch || '', notes: '' }); added++;
  }
  save(); toast(`נוספו ${added} טיסות מתוך ${parsed.length} קבצים`); go('flights');
}
// Android: files shared from Gaggle ("Share" -> this app) are stashed by sw.js, picked up here
async function consumeSharedFiles() {
  if (!new URLSearchParams(location.search).has('shared') || !('caches' in window)) return;
  history.replaceState(null, '', location.pathname);
  const cache = await caches.open('ppg-share');
  const files = [];
  for (const req of await cache.keys()) {
    const res = await cache.match(req);
    files.push({ name: decodeURIComponent(res.headers.get('x-name') || 'track.igc'), text: await res.text() });
    await cache.delete(req);
  }
  if (files.length) await handleTrackFiles(files); else toast('לא התקבל קובץ מהשיתוף');
}

function parseTrack(text, fileName) {
  let pts = [];
  if (/^\s*</.test(text)) { // GPX
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    for (const p of doc.querySelectorAll('trkpt, rtept')) {
      const tm = p.querySelector('time'); if (!tm) continue;
      const t = Date.parse(tm.textContent.trim()); if (!Number.isFinite(t)) continue;
      pts.push({ t, lat: +p.getAttribute('lat'), lon: +p.getAttribute('lon') });
    }
  } else { // IGC
    let base = null, prevSec = -1, dayOffset = 0;
    for (const line of text.split(/\r?\n/)) {
      if (line.startsWith('HFDTE')) {
        const m = line.match(/(\d{2})(\d{2})(\d{2})/);
        if (m) base = Date.UTC(2000 + +m[3], +m[2] - 1, +m[1]);
      } else if (line[0] === 'B' && line.length >= 24 && base !== null) {
        const hh = +line.slice(1, 3), mm = +line.slice(3, 5), ss = +line.slice(5, 7);
        let sec = hh * 3600 + mm * 60 + ss;
        if (prevSec >= 0 && sec + 3600 < prevSec) dayOffset += 86400; // crossed UTC midnight
        prevSec = sec;
        const lat = (+line.slice(7, 9) + +line.slice(9, 14) / 60000) * (line[14] === 'S' ? -1 : 1);
        const lon = (+line.slice(15, 18) + +line.slice(18, 23) / 60000) * (line[23] === 'W' ? -1 : 1);
        pts.push({ t: base + (sec + dayOffset) * 1000, lat, lon });
      }
    }
  }
  pts = pts.filter(p => Number.isFinite(p.lat) && Number.isFinite(p.lon)).sort((a, b) => a.t - b.t);
  if (pts.length < 2) return null;

  // detect takeoff / landing by sustained ground speed (> ~18 km/h for 20 s)
  const moving = i => {
    let j = i; while (j < pts.length - 1 && pts[j].t - pts[i].t < 20000) j++;
    const dt = (pts[j].t - pts[i].t) / 1000; if (dt < 5) return false;
    return haversineKm(pts[i], pts[j]) * 1000 / dt > 5;
  };
  const winEnd = i => { let j = i; while (j < pts.length - 1 && pts[j].t - pts[i].t < 20000) j++; return j; };
  let a = 0; while (a < pts.length - 1 && !moving(a)) a++;
  let last = pts.length - 2; while (last > a && !moving(last)) last--;
  let b = winEnd(last);
  if (a >= pts.length - 1 || b <= a) { a = 0; b = pts.length - 1; }
  const totalMin = Math.round((pts[pts.length - 1].t - pts[0].t) / 60000);
  const min = Math.round((pts[b].t - pts[a].t) / 60000) || totalMin;
  const start = new Date(pts[a].t), end = new Date(pts[b].t);
  const date = `${start.getFullYear()}-${pad(start.getMonth() + 1)}-${pad(start.getDate())}`;
  const tk = { lat: pts[a].lat, lon: pts[a].lon };
  let siteMatch = '', best = 2.0;
  for (const s of S.sites) { const d = haversineKm(tk, s); if (d < best) { best = d; siteMatch = s.name; } }
  return { file: fileName, date, min, totalMin, lat: tk.lat, lon: tk.lon, siteMatch,
    takeoff: `${pad(start.getHours())}:${pad(start.getMinutes())}`, landing: `${pad(end.getHours())}:${pad(end.getMinutes())}` };
}

/* ---------------- Maintenance ---------------- */
function renderMaint() {
  const total = totalHours();
  const comps = S.components.map(c => ({ c, s: compStatus(c, total) })).sort((a, b) => a.s.nextDue - b.s.nextDue);
  const log = [...S.maint].sort((a, b) => b.date.localeCompare(a.date));
  const byId = Object.fromEntries(S.components.map(c => [c.id, c]));
  view.innerHTML = `
    <button class="btn primary block" id="doMaint" style="margin-bottom:12px">🔧 רישום טיפול / החלפת חלקים</button>
    <div class="card">
      <div class="row"><h2>חלקים (${comps.length}) · מנוע <span class="num">${hm(total)}</span></h2><button class="btn sm" id="addComp">＋ חלק</button></div>
      ${comps.map(x => compRow(x.c, x.s)).join('')}
    </div>
    <div class="card">
      <h2>היסטוריית טיפולים</h2>
      ${log.length ? log.map(m => `<div class="comp" data-mid="${m.id}">
        <div class="row"><div class="name">${fmtDate(m.date)}</div><div class="num muted">${hm(m.hours)} ש׳</div></div>
        ${m.items.length ? `<div class="meta">🔁 ${m.items.map(id => esc(byId[id]?.name || '(חלק שנמחק)')).join(' · ')}</div>` : ''}
        ${(m.checks || []).length ? `<div class="meta">🔍 ${m.checks.map(id => esc(byId[id]?.name || '(חלק שנמחק)')).join(' · ')}</div>` : ''}
        ${m.note ? `<div class="meta" style="color:var(--accent)">${esc(m.note)}</div>` : ''}
      </div>`).join('') : '<div class="muted small">עדיין לא נרשמו טיפולים באפליקציה. כל טיפול שתרשום יאפס את השעון של החלקים שנבחרו.</div>'}
    </div>`;
  $('#doMaint').onclick = () => openMaintDialog();
  $('#addComp').onclick = () => openComponent(null);
  view.querySelectorAll('.comp[data-id]').forEach(el => el.onclick = () => openComponent(el.dataset.id));
  view.querySelectorAll('.comp[data-mid]').forEach(el => el.onclick = () => openMaintEntry(el.dataset.mid));
}

function openMaintDialog(preselect, date, note) {
  const total = totalHours();
  const comps = S.components.map(c => ({ c, s: compStatus(c, total) })).sort((a, b) => a.s.nextDue - b.s.nextDue);
  openModal(`<h3>רישום טיפול</h3>
    <div class="row" style="gap:8px">
      <label class="f" style="flex:1"><span>תאריך</span><input type="date" id="mDate" value="${date || todayISO()}"></label>
      <label class="f" style="flex:1"><span>בשעת מנוע</span><input type="text" id="mHours" inputmode="decimal" class="num" value="${hm(total)}"></label>
    </div>
    <span class="muted small">סמן 🔁 לחלק שהוחלף (מאפס את שעון ההחלפה) או 🔍 לחלק שרק נבדק / טופל (מאפס את שעון הבדיקה)</span>
    <div class="card" style="padding:4px 10px;margin-top:6px">
      <div class="mrow mhead small muted"><span>חלק</span><span>🔁 הוחלף</span><span>🔍 נבדק</span></div>
      ${comps.map(({ c, s }) => `<div class="mrow">
        <span>${esc(c.name)}<br><span class="small num s-${s.level}">${hm(s.remaining)}</span>${s.checkRemaining !== null ? ` <span class="small num s-${s.checkLevel}">· 🔍 ${hm(s.checkRemaining)}</span>` : ''}</span>
        <label><input type="checkbox" class="rep" value="${c.id}" ${preselect === c.id ? 'checked' : ''}></label>
        <label>${c.checkH > 0 ? `<input type="checkbox" class="chk" value="${c.id}" ${preselect === c.id + ':check' ? 'checked' : ''}>` : ''}</label>
      </div>`).join('')}
    </div>
    <label class="f"><span>הערה</span><input type="text" id="mNote" value="${esc(note || '')}" placeholder="למשל: פלאג NGK BR10EG"></label>
    <div class="btns"><button class="btn" id="mCancel">ביטול</button><button class="btn primary" id="mSave">שמור טיפול</button></div>`, card => {
    $('#mCancel', card).onclick = closeModal;
    $('#mSave', card).onclick = () => {
      const items = [...card.querySelectorAll('input.rep:checked')].map(i => i.value);
      const checks = [...card.querySelectorAll('input.chk:checked')].map(i => i.value).filter(id => !items.includes(id));
      const hours = parseHM($('#mHours', card).value), d = $('#mDate', card).value;
      if (!items.length && !checks.length) return toast('בחר לפחות חלק אחד');
      if (!Number.isFinite(hours) || hours < 0) return toast('שעת מנוע לא תקינה (פורמט 123:45)');
      for (const id of items) { const c = S.components.find(x => x.id === id); c.lastResetH = hours; c.lastResetDate = d; c.lastCheckH = hours; }
      for (const id of checks) { const c = S.components.find(x => x.id === id); c.lastCheckH = hours; }
      S.maint.push({ id: uid(), date: d, hours, items, checks, note: $('#mNote', card).value.trim() });
      save(); closeModal(); toast(`נרשם: ${items.length} הוחלפו, ${checks.length} נבדקו`); renderMaint(); afterChangeBackup();
    };
  });
}
function openMaintEntry(id) {
  const m = S.maint.find(x => x.id === id); if (!m) return;
  const byId = Object.fromEntries(S.components.map(c => [c.id, c]));
  openModal(`<h3>טיפול ${fmtDate(m.date)} · <span class="num">${hm(m.hours)}</span></h3>
    <p>${m.items.map(i => '🔁 ' + esc(byId[i]?.name || '(חלק שנמחק)')).concat((m.checks || []).map(i => '🔍 ' + esc(byId[i]?.name || '(חלק שנמחק)'))).join('<br>')}</p>${m.note ? `<p class="muted">${esc(m.note)}</p>` : ''}
    <p class="muted small">מחיקת הרשומה לא משנה את שעון החלקים. לתיקון השעון ערוך את החלק עצמו.</p>
    <div class="btns"><button class="btn danger" id="mDel">מחק רשומה</button><button class="btn" id="mClose">סגור</button></div>`, card => {
    $('#mClose', card).onclick = closeModal;
    $('#mDel', card).onclick = () => { S.maint = S.maint.filter(x => x.id !== id); save(); closeModal(); renderMaint(); };
  });
}
function openComponent(id) {
  const c = id ? S.components.find(x => x.id === id) : { id: null, name: '', intervalH: 100, intervalMonths: 0, lastResetH: totalHours(), lastResetDate: todayISO(), notes: '' };
  if (!c) return;
  const s = id ? compStatus(c) : null;
  const hist = id ? S.maint.filter(m => m.items.includes(id) || (m.checks || []).includes(id)).sort((a, b) => b.date.localeCompare(a.date)) : [];
  openModal(`<h3>${id ? esc(c.name) : 'חלק חדש'}</h3>
    ${s ? `<div class="track-info">שעות מאז החלפה: <b class="num">${hm(s.since)}</b> · ${s.remaining < 0 ? 'באיחור' : 'נותרו'} <b class="num s-${s.level}">${hm(Math.abs(s.remaining))}</b>${s.dueDate ? `<br>תאריך יעד קלנדרי: ${fmtDate(s.dueDate)}` : ''}</div>` : ''}
    <label class="f"><span>שם החלק</span><input type="text" id="cName" value="${esc(c.name)}"></label>
    <div class="row" style="gap:8px">
      <label class="f" style="flex:1"><span>מרווח (שעות)</span><input type="number" inputmode="decimal" id="cInt" value="${c.intervalH}"></label>
      <label class="f" style="flex:1"><span>או כל (חודשים, 0=ללא)</span><input type="number" inputmode="numeric" id="cMon" value="${c.intervalMonths || 0}"></label>
    </div>
    <div class="row" style="gap:8px">
      <label class="f" style="flex:1"><span>הוחלף לאחרונה בשעה</span><input type="text" class="num" id="cLast" value="${hm(c.lastResetH || 0)}"></label>
      <label class="f" style="flex:1"><span>בתאריך</span><input type="date" id="cDate" value="${c.lastResetDate || ''}"></label>
    </div>
    <div class="row" style="gap:8px">
      <label class="f" style="flex:1"><span>🔍 בדיקה כל (שעות, 0=ללא)</span><input type="number" inputmode="decimal" id="cChk" value="${c.checkH || 0}"></label>
      <label class="f" style="flex:1"><span>נבדק לאחרונה בשעה</span><input type="text" class="num" id="cLastChk" value="${hm(Math.max(c.lastCheckH || 0, c.lastResetH || 0))}"></label>
    </div>
    <label class="f"><span>הערות יצרן</span><textarea id="cNotes">${esc(c.notes)}</textarea></label>
    ${hist.length ? `<div class="small muted" style="margin-bottom:12px">היסטוריה: ${hist.map(m => `${m.items.includes(id) ? '🔁' : '🔍'} ${fmtDate(m.date)} (<span class="num">${hm(m.hours)}</span>)`).join(' · ')}</div>` : ''}
    <div class="btns">
      ${id ? `<button class="btn danger" id="cDel">מחק</button>${c.checkH > 0 ? '<button class="btn" id="cChkNow">🔍 נבדק</button>' : ''}<button class="btn" id="cDone">🔁 הוחלף</button>` : ''}
      <button class="btn primary" id="cSave">שמור</button>
    </div>`, card => {
    $('#cSave', card).onclick = () => {
      const name = $('#cName', card).value.trim(), intervalH = parseFloat($('#cInt', card).value), last = parseHM($('#cLast', card).value);
      if (!name) return toast('חסר שם');
      if (!(intervalH > 0)) return toast('מרווח שעות לא תקין');
      if (!Number.isFinite(last)) return toast('שעת החלפה לא תקינה (פורמט 123:45)');
      const lastChk = parseHM($('#cLastChk', card).value);
      if (!Number.isFinite(lastChk)) return toast('שעת בדיקה לא תקינה (פורמט 123:45)');
      Object.assign(c, { name, intervalH, intervalMonths: parseInt($('#cMon', card).value || '0', 10) || 0, lastResetH: last, lastResetDate: $('#cDate', card).value, notes: $('#cNotes', card).value.trim(),
        checkH: parseFloat($('#cChk', card).value) || 0, lastCheckH: lastChk });
      if (!id) { c.id = uid(); S.components.push(c); }
      save(); closeModal(); go(tab);
    };
    const del = $('#cDel', card);
    if (del) del.onclick = async () => { if (!(await confirmBox(`למחוק את "${c.name}" מרשימת החלקים?`, 'מחק', true))) return; S.components = S.components.filter(x => x.id !== id); save(); go(tab); };
    const done = $('#cDone', card);
    if (done) done.onclick = () => { closeModal(); openMaintDialog(id); };
    const chkNow = $('#cChkNow', card);
    if (chkNow) chkNow.onclick = () => { closeModal(); openMaintDialog(id + ':check'); };
  });
}

/* ---------------- Settings / backup ---------------- */
function renderSettings() {
  const fsa = 'showSaveFilePicker' in window;
  const rs = reserveStatus();
  view.innerHTML = `
    <div class="card">
      <h2>גיבוי לאקסל</h2>
      <p class="small muted">הקובץ נשמר באותו מבנה של הקובץ המקורי שלך (גיליון טיסות + גיליון תחזוקה עם נוסחאות), ובנוסף גיליון היסטוריית טיפולים. אפשר לפתוח אותו באקסל גם בלי האפליקציה, ולייבא אותו בחזרה בכל זמן.</p>
      <div class="btns"><button class="btn primary" id="exp">⤓ שמור קובץ אקסל</button><button class="btn" id="imp">⤒ ייבוא מאקסל</button></div>
      ${fsa ? `<hr class="sep"><div class="row"><div><b>סנכרון אוטומטי לקובץ</b><div class="small muted">${linked.handle ? `מקושר ל-<b>${esc(linked.name)}</b>${linked.ok ? ' · כל שינוי נכתב מיד' : ' · נדרש אישור גישה'}` : 'בחר קובץ (למשל בתיקיית OneDrive) והאפליקציה תעדכן אותו אחרי כל שינוי'}</div></div>
        ${linked.handle ? `<div class="btns" style="flex:0 0 auto">${linked.ok ? '' : '<button class="btn sm primary" id="grant">אשר גישה</button>'}<button class="btn sm" id="unlink">נתק</button></div>` : '<button class="btn sm" id="link">קשר קובץ</button>'}</div>` : ''}
      <hr class="sep">
      <label class="checks"><label style="border:0;padding:0"><input type="checkbox" id="auto" ${S.settings.autoBackup ? 'checked' : ''}> להציע שמירת אקסל אחרי כל טיסה${fsa ? ' (כשאין קובץ מקושר)' : ''}</label></label>
    </div>
    <div class="card">
      <h2>🪂 מצנח רזרבי</h2>
      <div class="row" style="gap:8px">
        <label class="f" style="flex:1"><span>קיפול אחרון</span><input type="month" id="rDate" value="${(S.settings.reserveRepack || '').slice(0, 7)}"></label>
        <label class="f" style="flex:1"><span>כל (חודשים)</span><input type="number" id="rMon" value="${S.settings.reserveMonths || 6}"></label>
      </div>
      ${rs ? `<div class="small s-${rs.level}">קיפול הבא: ${fmtDate(rs.due)} (${rs.days < 0 ? `באיחור ${-rs.days} ימים` : `בעוד ${rs.days} ימים`})</div>` : ''}
    </div>
    <div class="card">
      <h2>כללי</h2>
      <label class="f"><span>שם המנוע</span><input type="text" id="eng" value="${esc(S.engineName)}"></label>
      <div class="small muted" style="margin-bottom:10px">אתרי המראה עם מיקום שמור (לזיהוי אוטומטי מקבצי Gaggle): ${S.sites.length ? S.sites.map(s => esc(s.name)).join(' · ') : 'אין עדיין'}</div>
      <div class="btns"><button class="btn" id="jexp">גיבוי JSON</button><button class="btn" id="jimp">שחזור JSON</button></div>
      <hr class="sep">
      <button class="btn danger block" id="wipe">מחק את כל הנתונים במכשיר</button>
    </div>
    <p class="small muted" style="text-align:center">הנתונים נשמרים מקומית במכשיר הזה בלבד. גבה לאקסל באופן קבוע.</p>`;
  $('#exp').onclick = () => exportExcel();
  $('#imp').onclick = () => $('#fileExcel').click();
  $('#jexp').onclick = () => saveBlob(new Blob([JSON.stringify(S, null, 1)], { type: 'application/json' }), `ppg-logbook-${todayISO()}.json`);
  $('#jimp').onclick = () => $('#fileJson').click();
  $('#auto').onchange = e => { S.settings.autoBackup = e.target.checked; save(false); };
  $('#rDate').onchange = e => { S.settings.reserveRepack = e.target.value; save(); renderSettings(); };
  $('#rMon').onchange = e => { S.settings.reserveMonths = parseInt(e.target.value, 10) || 6; save(); renderSettings(); };
  $('#eng').onchange = e => { S.engineName = e.target.value.trim() || 'יומן מנוע'; save(); };
  $('#wipe').onclick = async () => {
    if (!(await confirmBox('למחוק את כל הטיסות, החלקים וההגדרות מהמכשיר? (קובץ האקסל לא יימחק)', 'מחק הכל', true))) return;
    S = freshState(); save(false); go('dash');
  };
  const l = $('#link'); if (l) l.onclick = linkFile;
  const u = $('#unlink'); if (u) u.onclick = async () => { await idbSet('linked', null); linked = { handle: null }; updateHeader(); renderSettings(); };
  const g = $('#grant'); if (g) g.onclick = async () => { if (await ensurePermission(true)) { await writeLinked(); renderSettings(); } };
}

$('#fileJson').addEventListener('change', async e => {
  const file = e.target.files[0]; e.target.value = ''; if (!file) return;
  try {
    const data = JSON.parse(await file.text());
    if (!Array.isArray(data.flights)) throw new Error('bad');
    if (!(await confirmBox(`לשחזר ${data.flights.length} טיסות מהקובץ? הנתונים הנוכחיים יוחלפו.`, 'שחזר'))) return;
    S = Object.assign(freshState(), data); save(); go('dash'); toast('שוחזר');
  } catch { toast('קובץ JSON לא תקין'); }
});

/* ---------------- Excel export ---------------- */
const toXlDur = hours => hours / 24;                       // Excel duration = fraction of a day
const isoToXlDate = iso => { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };

async function buildWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PPG Engine Logbook';
  wb.calcProperties.fullCalcOnLoad = true;
  const total = totalHours();
  const flights = [...S.flights].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  const head = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } }, alignment: { horizontal: 'center', vertical: 'middle', wrapText: true } };
  const style = (cell, st) => Object.assign(cell, st);

  // Sheet 1 — same layout as the original logbook
  const ws = wb.addWorksheet(SHEET_LOG, { views: [{ state: 'frozen', ySplit: 2 }] });
  ws.columns = [{ width: 12 }, { width: 11 }, { width: 26 }, { width: 30 }, { width: 15 }, { width: 13 }];
  ws.getCell('E1').value = "סה''כ זמן טיסה";
  ws.getCell('F1').value = 'פתיחת רזרבי';
  ['A2', 'B2', 'C2', 'D2'].forEach((a, i) => { ws.getCell(a).value = ['תאריך', 'זמן טיסה', 'אזור', 'הערות'][i]; style(ws.getCell(a), head); });
  ['E1', 'F1'].forEach(a => style(ws.getCell(a), head));
  const lastRow = 2 + Math.max(flights.length, 1);
  ws.getCell('E2').value = { formula: `SUM(B3:B${Math.max(lastRow, 2000)})`, result: toXlDur(total) };
  ws.getCell('E2').numFmt = '[h]:mm';
  ws.getCell('E2').font = { bold: true, size: 13 };
  const r = S.settings.reserveRepack;
  ws.getCell('F2').value = r ? `${r.slice(5, 7)}/${r.slice(2, 4)}` : '';
  flights.forEach((f, i) => {
    const row = ws.getRow(3 + i);
    row.getCell(1).value = isoToXlDate(f.date); row.getCell(1).numFmt = 'dd/mm/yy';
    row.getCell(2).value = toXlDur(f.min / 60); row.getCell(2).numFmt = '[h]:mm';
    row.getCell(3).value = f.site || null;
    row.getCell(4).value = f.notes || null;
  });

  // Sheet 2 — maintenance tracker with live formulas
  const mt = wb.addWorksheet(SHEET_MT, { views: [{ state: 'frozen', ySplit: 3 }] });
  mt.columns = [{ width: 40 }, { width: 13 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 50 }, { width: 15 }, { width: 13 }, { width: 13 }, { width: 15 }, { width: 16 }];
  mt.getCell('A1').value = `🔧 ${S.engineName} — Maintenance Tracker`;
  mt.getCell('A1').font = { bold: true, size: 14 };
  mt.getCell('A2').value = '⏱ Total Engine Hours:';
  mt.getCell('A2').font = { bold: true };
  mt.getCell('E2').value = { formula: `'${SHEET_LOG}'!E2`, result: toXlDur(total) };
  mt.getCell('E2').numFmt = '[h]:mm'; mt.getCell('E2').font = { bold: true, size: 13 };
  const hdr = ['Component', 'Interval (hrs)', 'Last Reset at (hrs)', 'Hours Since Reset', 'Next Due at (hrs)', 'Hours Remaining', 'Notes', 'Last Reset Date', 'Interval (months)', 'Check Every (hrs)', 'Last Check at (hrs)', 'Check Remaining'];
  hdr.forEach((h, i) => { const c = mt.getRow(3).getCell(i + 1); c.value = h; style(c, head); });
  mt.getRow(3).height = 32;
  S.components.forEach((c, i) => {
    const n = 4 + i, row = mt.getRow(n), s = compStatus(c, total);
    row.getCell(1).value = c.name;
    row.getCell(2).value = toXlDur(c.intervalH);
    row.getCell(3).value = toXlDur(c.lastResetH || 0);
    row.getCell(4).value = { formula: `$E$2-C${n}`, result: toXlDur(s.since) };
    row.getCell(5).value = { formula: `C${n}+B${n}`, result: toXlDur((c.lastResetH || 0) + c.intervalH) };
    row.getCell(6).value = { formula: `E${n}-$E$2`, result: toXlDur(s.remaining) };
    for (let k = 2; k <= 6; k++) row.getCell(k).numFmt = '[h]:mm';
    row.getCell(7).value = c.notes || null;
    if (c.lastResetDate) { row.getCell(8).value = isoToXlDate(c.lastResetDate); row.getCell(8).numFmt = 'dd/mm/yy'; }
    row.getCell(9).value = c.intervalMonths || null;
    if (c.checkH > 0) {
      row.getCell(10).value = toXlDur(c.checkH);
      row.getCell(11).value = toXlDur(Math.max(c.lastCheckH || 0, c.lastResetH || 0));
      row.getCell(12).value = { formula: `MAX(K${n},C${n})+J${n}-$E$2`, result: toXlDur(s.checkRemaining) };
      for (let k = 10; k <= 12; k++) row.getCell(k).numFmt = '[h]:mm';
    }
  });
  const lastC = 3 + S.components.length;
  if (S.components.length) {
    const fill = argb => ({ fill: { type: 'pattern', pattern: 'solid', bgColor: { argb } } });
    mt.addConditionalFormatting({ ref: `F4:F${lastC}`, rules: [
      { type: 'cellIs', operator: 'lessThan', formulae: ['10/24'], priority: 1, style: fill('FFF8696B') },
      { type: 'cellIs', operator: 'between', formulae: ['10/24', '25/24'], priority: 2, style: fill('FFFFEB84') },
      { type: 'cellIs', operator: 'greaterThan', formulae: ['25/24'], priority: 3, style: fill('FF63BE7B') },
    ] });
  }
  const lg = lastC + 3;
  mt.getCell(`B${lg}`).value = 'Legend'; mt.getCell(`B${lg}`).font = { bold: true };
  [['Less than 10h', 'FFF8696B'], ['10h-25h', 'FFFFEB84'], ['Greater than 25h', 'FF63BE7B']].forEach(([t, c], i) => {
    const cell = mt.getRow(lg + 1).getCell(2 + i); cell.value = t;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: c } };
  });

  // Sheet 3 — maintenance history
  const ml = wb.addWorksheet(SHEET_ML, { views: [{ state: 'frozen', ySplit: 1 }] });
  ml.columns = [{ width: 12 }, { width: 14 }, { width: 55 }, { width: 40 }, { width: 55 }];
  ['Date', 'Engine Hours', 'Replaced', 'Notes', 'Checked'].forEach((h, i) => { const c = ml.getRow(1).getCell(i + 1); c.value = h; style(c, head); });
  const byId = Object.fromEntries(S.components.map(c => [c.id, c.name]));
  [...S.maint].sort((a, b) => a.date.localeCompare(b.date)).forEach((m, i) => {
    const row = ml.getRow(2 + i);
    row.getCell(1).value = isoToXlDate(m.date); row.getCell(1).numFmt = 'dd/mm/yy';
    row.getCell(2).value = toXlDur(m.hours); row.getCell(2).numFmt = '[h]:mm';
    row.getCell(3).value = m.items.map(id => byId[id] || '?').join(' | ');
    row.getCell(4).value = m.note || null;
    row.getCell(5).value = (m.checks || []).map(id => byId[id] || '?').join(' | ') || null;
  });

  // Sheet 4 — takeoff sites (for automatic matching of Gaggle tracks)
  const st = wb.addWorksheet(SHEET_SITES);
  st.columns = [{ width: 26 }, { width: 12 }, { width: 12 }];
  ['Site', 'Lat', 'Lon'].forEach((h, i) => { const c = st.getRow(1).getCell(i + 1); c.value = h; style(c, head); });
  S.sites.forEach((s, i) => { const row = st.getRow(2 + i); row.getCell(1).value = s.name; row.getCell(2).value = +s.lat.toFixed(6); row.getCell(3).value = +s.lon.toFixed(6); });
  st.getCell('E1').value = 'reserveMonths'; st.getCell('F1').value = S.settings.reserveMonths || 6;

  return wb;
}
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
async function workbookBlob() { const wb = await buildWorkbook(); return new Blob([await wb.xlsx.writeBuffer()], { type: XLSX_MIME }); }
function markBackedUp() { S.meta.lastBackup = Date.now(); S.meta.changesSinceBackup = 0; save(false); if (tab === 'dash') renderDash(); }

async function exportExcel() {
  try {
    if (linked.handle && (await ensurePermission(true))) { await writeLinked(); toast(`נשמר ל-${linked.name}`); return; }
    const blob = await workbookBlob();
    const ok = await saveBlob(blob, 'Paramotor Flights.xlsx');
    if (ok) { markBackedUp(); toast('קובץ האקסל נשמר ✓'); }
  } catch (e) { console.error(e); toast('שגיאה ביצירת האקסל'); }
}
async function saveBlob(blob, name) {
  const file = new File([blob], name, { type: blob.type });
  // Mobile: native share sheet lets you save straight to OneDrive / Drive / Files
  if (navigator.canShare && /Android|iPhone|iPad|iPod/i.test(navigator.userAgent) && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return true; }
    catch (e) { if (e.name === 'AbortError') return false; }
  }
  if ('showSaveFilePicker' in window) {
    try {
      const h = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: name.endsWith('.json') ? 'JSON' : 'Excel', accept: { [blob.type]: [name.slice(name.lastIndexOf('.'))] } }] });
      const w = await h.createWritable(); await w.write(blob); await w.close(); return true;
    } catch (e) { if (e.name === 'AbortError') return false; }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  return true;
}
function afterChangeBackup() {
  if (linked.handle && linked.ok) return; // already written automatically
  if (!S.settings.autoBackup) return;
  setTimeout(() => openModal(`<h3>לשמור גיבוי אקסל עכשיו?</h3><p class="muted small">כך תמיד יהיה לך קובץ עדכני גם בלי האפליקציה.${/Android|iPhone|iPad/i.test(navigator.userAgent) ? ' בחלון השיתוף בחר OneDrive / Drive / קבצים.' : ''}</p>
    <div class="btns"><button class="btn" id="later">אחר כך</button><button class="btn primary" id="now">⤓ שמור אקסל</button></div>`, card => {
    $('#later', card).onclick = closeModal;
    $('#now', card).onclick = () => { closeModal(); exportExcel(); };
  }), 300);
}

/* ---------------- Linked file (desktop Chrome/Edge: writes straight into the chosen .xlsx) ---------------- */
let linked = { handle: null, name: '', ok: false };
function idb() { return new Promise((res, rej) => { const r = indexedDB.open('ppg-logbook', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }
async function idbGet(k) { const db = await idb(); return new Promise(res => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(null); }); }
async function idbSet(k, v) { const db = await idb(); return new Promise(res => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => res(); t.onerror = () => res(); }); }
async function linkFile() {
  try {
    const h = await window.showSaveFilePicker({ suggestedName: 'Paramotor Flights.xlsx', types: [{ description: 'Excel', accept: { [XLSX_MIME]: ['.xlsx'] } }] });
    linked = { handle: h, name: h.name, ok: true }; await idbSet('linked', h);
    await writeLinked(); toast(`מקושר ל-${h.name}`); renderSettings();
  } catch (e) { if (e.name !== 'AbortError') toast('לא ניתן לקשר קובץ'); }
}
async function ensurePermission(ask) {
  if (!linked.handle) return false;
  let p = await linked.handle.queryPermission({ mode: 'readwrite' });
  if (p !== 'granted' && ask) p = await linked.handle.requestPermission({ mode: 'readwrite' });
  linked.ok = p === 'granted'; updateHeader(); return linked.ok;
}
async function writeLinked() {
  try {
    const blob = await workbookBlob();
    const w = await linked.handle.createWritable(); await w.write(blob); await w.close();
    markBackedUp();
  } catch (e) {
    console.error(e); linked.ok = false; updateHeader();
    toast('הכתיבה לקובץ נכשלה – אם הוא פתוח באקסל, סגור אותו ונסה שוב', 4500);
  }
}
let linkedTimer = null;
function scheduleLinkedWrite() {
  if (!linked.handle || !linked.ok) return;
  clearTimeout(linkedTimer); linkedTimer = setTimeout(writeLinked, 600);
}
(async () => {
  if (!('showSaveFilePicker' in window)) return;
  try { const h = await idbGet('linked'); if (h) { linked = { handle: h, name: h.name, ok: false }; await ensurePermission(false); if (tab === 'settings') renderSettings(); } } catch {}
})();

/* ---------------- Excel import ---------------- */
const xlToHours = v => {
  if (v == null || v === '') return NaN;
  if (typeof v === 'object' && !(v instanceof Date)) v = v.result ?? v.value ?? NaN; // formula / rich
  if (v instanceof Date) return (v.getTime() - XL_EPOCH) / 3600000;
  if (typeof v === 'number') return v * 24;
  return parseHM(v);
};
const xlToISO = v => {
  if (v instanceof Date) return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  if (typeof v === 'number') { const d = new Date(XL_EPOCH + v * 86400000); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
  const m = String(v || '').trim().match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${pad(m[2])}-${pad(m[1])}`;
  return '';
};
const cellText = v => { if (v == null) return ''; if (typeof v === 'object' && v.richText) return v.richText.map(r => r.text).join(''); if (typeof v === 'object' && 'result' in v) return String(v.result ?? ''); return String(v); };

async function importExcel(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  const ws = wb.getWorksheet(SHEET_LOG) || wb.worksheets[0];
  const flights = [], warnings = [];
  let prevDate = '';
  ws.eachRow((row, n) => {
    if (n < 3) return;
    const date = xlToISO(row.getCell(1).value), hours = xlToHours(row.getCell(2).value);
    if (!date && !Number.isFinite(hours)) return;
    if (!date || !Number.isFinite(hours) || hours <= 0) { warnings.push(`שורה ${n} דולגה (תאריך/זמן לא תקין)`); return; }
    if (prevDate && date < prevDate && prevDate.slice(0, 4) - date.slice(0, 4) >= 2) warnings.push(`שורה ${n}: התאריך ${fmtDate(date)} נראה כמו טעות הקלדה (הטיסה הקודמת ב-${fmtDate(prevDate)})`);
    prevDate = date;
    flights.push({ id: uid(), date, min: Math.round(hours * 60), site: cellText(row.getCell(3).value).trim(), notes: cellText(row.getCell(4).value).trim() });
  });
  let reserve = '';
  const f2 = ws.getCell('F2').value;
  if (f2 instanceof Date) reserve = `${f2.getUTCFullYear()}-${pad(f2.getUTCMonth() + 1)}`;
  else { const m = cellText(f2).match(/^(\d{1,2})\/(\d{2,4})$/); if (m) reserve = `${m[2].length === 2 ? '20' + m[2] : m[2]}-${pad(m[1])}`; }

  let components = null;
  const mt = wb.getWorksheet(SHEET_MT);
  if (mt) {
    components = [];
    for (let n = 4; n <= mt.rowCount; n++) {
      const row = mt.getRow(n);
      const name = cellText(row.getCell(1).value).replace(/\s*\n\s*/g, ' ').trim();
      const interval = xlToHours(row.getCell(2).value);
      if (!name || !Number.isFinite(interval) || interval <= 0) { if (!name) break; continue; }
      const def = DEFAULT_COMPONENTS.find(d => d[0].toLowerCase() === name.toLowerCase());
      const monthsCell = row.getCell(9).value;
      const chk = xlToHours(row.getCell(10).value), lastChk = xlToHours(row.getCell(11).value);
      const hasCheckCols = cellText(mt.getRow(3).getCell(10).value).startsWith('Check');
      components.push({
        checkH: hasCheckCols ? (Number.isFinite(chk) ? Math.round(chk * 100) / 100 : 0) : def ? def[4] : 0,
        lastCheckH: Number.isFinite(lastChk) ? Math.round(lastChk * 60) / 60 : null, // null = fill in after flights are known
        id: uid(), name, intervalH: Math.round(interval * 100) / 100,
        intervalMonths: typeof monthsCell === 'number' ? monthsCell : def ? def[2] : 0,
        lastResetH: Number.isFinite(xlToHours(row.getCell(3).value)) ? Math.round(xlToHours(row.getCell(3).value) * 60) / 60 : 0,
        lastResetDate: xlToISO(row.getCell(8).value) || '',
        notes: cellText(row.getCell(7).value).trim(),
      });
    }
  }
  const maint = [];
  const ml = wb.getWorksheet(SHEET_ML);
  if (ml && components) {
    ml.eachRow((row, n) => {
      if (n < 2) return;
      const date = xlToISO(row.getCell(1).value), hours = xlToHours(row.getCell(2).value);
      if (!date) return;
      const items = cellText(row.getCell(3).value).split('|').map(s => s.trim()).map(nm => components.find(c => c.name === nm)?.id).filter(Boolean);
      const checks = cellText(row.getCell(5).value).split('|').map(s => s.trim()).map(nm => components.find(c => c.name === nm)?.id).filter(Boolean);
      maint.push({ id: uid(), date, hours: Number.isFinite(hours) ? hours : 0, items, checks, note: cellText(row.getCell(4).value).trim() });
    });
  }
  const sites = []; let reserveMonths = null;
  const st = wb.getWorksheet(SHEET_SITES);
  if (st) {
    st.eachRow((row, n) => { if (n < 2) return; const name = cellText(row.getCell(1).value).trim(), lat = +row.getCell(2).value, lon = +row.getCell(3).value; if (name && Number.isFinite(lat) && Number.isFinite(lon)) sites.push({ name, lat, lon }); });
    if (typeof st.getCell('F1').value === 'number') reserveMonths = st.getCell('F1').value;
  }
  // The original sheet has no check history: assume checks are up to date as of today's engine hours.
  const totalImported = flights.reduce((a, f) => a + f.min, 0) / 60;
  let assumedChecks = 0;
  for (const c of components || []) if (c.lastCheckH === null) { c.lastCheckH = c.checkH > 0 ? totalImported : 0; if (c.checkH > 0) assumedChecks++; }
  return { flights, components, maint, sites, reserve, reserveMonths, warnings, assumedChecks };
}

$('#fileExcel').addEventListener('change', async e => {
  const file = e.target.files[0]; e.target.value = ''; if (!file) return;
  let r;
  try { r = await importExcel(file); } catch (err) { console.error(err); return toast('לא הצלחתי לקרוא את קובץ האקסל'); }
  const hours = r.flights.reduce((a, f) => a + f.min, 0) / 60;
  openModal(`<h3>ייבוא מ-${esc(file.name)}</h3>
    <div class="track-info">
      ✈ <b>${r.flights.length}</b> טיסות · סה״כ <b class="num">${hm(hours)}</b> שעות<br>
      🔧 ${r.components ? `<b>${r.components.length}</b> חלקים בגיליון התחזוקה` : 'לא נמצא גיליון תחזוקה – יישארו החלקים הנוכחיים'}<br>
      ${r.maint.length ? `📋 ${r.maint.length} רשומות טיפול<br>` : ''}${r.reserve ? `🪂 קיפול רזרבי: ${r.reserve.slice(5)}/${r.reserve.slice(2, 4)}` : ''}
    </div>
    ${r.assumedChecks ? `<div class="alert info"><span>🔍</span><span>לקובץ אין היסטוריית בדיקות, אז ל-${r.assumedChecks} חלקים עם בדיקה תקופתית הבדיקה האחרונה נרשמה כ"עכשיו". אפשר לשנות בכל חלק.</span></div>` : ''}
    ${r.warnings.length ? `<div class="alert warn"><span>⚠</span><span>${r.warnings.map(esc).join('<br>')}</span></div>` : ''}
    <p class="small muted">${S.flights.length ? `הנתונים הנוכחיים באפליקציה (${S.flights.length} טיסות) יוחלפו בתוכן הקובץ.` : ''}</p>
    <div class="btns"><button class="btn" id="ic">ביטול</button><button class="btn primary" id="io">ייבא</button></div>`, card => {
    $('#ic', card).onclick = closeModal;
    $('#io', card).onclick = () => {
      S.flights = r.flights;
      if (r.components) S.components = r.components;
      S.maint = r.maint;
      if (r.sites.length) S.sites = r.sites;
      if (r.reserve) S.settings.reserveRepack = r.reserve;
      if (r.reserveMonths) S.settings.reserveMonths = r.reserveMonths;
      S.meta.lastBackup = Date.now(); S.meta.changesSinceBackup = 0; // the source file is an up-to-date backup
      save(false); scheduleLinkedWrite(); closeModal(); toast(`יובאו ${r.flights.length} טיסות`); go('dash');
    };
  });
});

/* ---------------- boot ---------------- */
updateHeader();
go('dash');
consumeSharedFiles();
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('sw.js').catch(() => {});
// expose for testing
window.__ppg = { parseTrack, importExcel, buildWorkbook, get state() { return S; }, hm, parseHM, compStatus, totalHours };
