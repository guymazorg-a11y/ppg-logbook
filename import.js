/* Smart Excel import for logbooks in ANY layout (other pilots' files).
   Detects the date / flight-time / site / notes columns, lets the user correct the guess,
   and shows a live preview (count, total hours, sample rows) before anything is saved. */
'use strict';

const MAX_COLS = 26, MAX_ROWS = 5000;
const KW = {
  date: /תאריך|date|יום|day/i,
  dur: /זמן|משך|שעות|דקות|time|duration|hours?|mins?|minutes|airtime|flight\s*time/i,
  site: /מקום|אזור|אתר|מנחת|המראה|site|location|place|take\s*-?off|launch|field/i,
  notes: /הערות|הערה|notes?|remarks?|comments?|תיאור/i,
  minutes: /דק|min/i,
};

/** Is this a file in the app's own / the original logbook layout? (keeps full restore incl. maintenance) */
function isOwnFormat(wb) {
  const ws = wb.getWorksheet(SHEET_LOG);
  if (!ws || !wb.getWorksheet(SHEET_MT)) return false;
  return cellText(ws.getCell('A2').value).trim() === 'תאריך' && cellText(ws.getCell('B2').value).trim() === 'זמן טיסה';
}

/* ---------- cell parsers ---------- */
const unwrap = v => (v && typeof v === 'object' && !(v instanceof Date)) ? (v.richText ? v.richText.map(r => r.text).join('') : (v.result ?? v.text ?? v.value ?? null)) : v;
const isTimeFmt = fmt => !!fmt && /h|:/.test(fmt.replace(/"[^"]*"/g, '')) && !/y|d/.test(fmt.replace(/"[^"]*"/g, ''));
const isDateFmt = fmt => !!fmt && /d|y/.test(fmt.replace(/"[^"]*"/g, ''));

function cellDate(c) {
  const v = unwrap(c.v);
  let y, m, d;
  if (v instanceof Date) { y = v.getUTCFullYear(); m = v.getUTCMonth() + 1; d = v.getUTCDate(); }
  else if (typeof v === 'number' && v > 25000 && v < 60000 && isDateFmt(c.fmt)) { const t = new Date(XL_EPOCH + Math.round(v) * 86400000); y = t.getUTCFullYear(); m = t.getUTCMonth() + 1; d = t.getUTCDate(); }
  else if (typeof v === 'string') {
    const s = v.trim();
    let r = s.match(/^(\d{1,2})[./\-](\d{1,2})[./\-](\d{2}|\d{4})$/);
    if (r) { d = +r[1]; m = +r[2]; y = r[3].length === 2 ? 2000 + +r[3] : +r[3]; }
    else if ((r = s.match(/^(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})/))) { y = +r[1]; m = +r[2]; d = +r[3]; }
    else return '';
  } else return '';
  if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return '';
  return `${y}-${pad(m)}-${pad(d)}`;
}

/** duration in minutes. unit: 'auto-h' | 'auto-m' (how to read plain numbers) | 'hm' (h.mm written as 1.30 = 1:30) */
function cellMinutes(c, unit) {
  const v = unwrap(c.v);
  if (v == null || v === '') return NaN;
  if (v instanceof Date) { // Excel time / duration cell (epoch 1899-12-30)
    const mins = (v.getTime() - XL_EPOCH) / 60000;
    return mins > 0 && mins < 24 * 60 * 3 ? Math.round(mins) : NaN;
  }
  if (typeof v === 'number') {
    if (isTimeFmt(c.fmt)) return Math.round(v * 1440); // Excel time stored as fraction of a day
    return numToMinutes(v, unit);
  }
  const s = String(v).trim().replace(',', '.');
  let r = s.match(/^(\d{1,2})\s*[:：]\s*(\d{1,2})(?:\s*[:：]\s*(\d{1,2}))?$/);
  if (r) return +r[1] * 60 + +r[2] + Math.round((+r[3] || 0) / 60);
  r = s.match(/^(\d+(?:\.\d+)?)\s*(h|hr|hrs|ש(?:עות|׳|')?)\s*(?:(\d{1,2})\s*(m|min|ד(?:קות|׳|')?)?)?$/i);
  if (r) return Math.round(+r[1] * 60 + (+r[3] || 0));
  r = s.match(/^(\d+)\s*(m|min|mins|ד(?:קות|׳|')?)$/i);
  if (r) return +r[1];
  if (/^\d+(\.\d+)?$/.test(s)) return numToMinutes(+s, unit);
  return NaN;
}
function numToMinutes(n, unit) {
  if (!(n > 0)) return NaN;
  if (unit === 'auto-m') return Math.round(n);
  if (unit === 'hm') { const h = Math.floor(n), mm = Math.round((n - h) * 100); return mm < 60 ? h * 60 + mm : NaN; }
  return Math.round(n * 60); // hours (decimal)
}

/* ---------- read & detect ---------- */
function readGrid(ws) {
  const rows = [];
  const last = Math.min(ws.rowCount, MAX_ROWS);
  let maxCol = 0;
  for (let r = 1; r <= last; r++) {
    const row = ws.getRow(r), cells = [];
    for (let c = 1; c <= Math.min(ws.columnCount || MAX_COLS, MAX_COLS); c++) {
      const cell = row.getCell(c);
      const v = cell.value;
      cells.push({ v, fmt: cell.numFmt || '' });
      if (v != null && v !== '') maxCol = Math.max(maxCol, c);
    }
    rows.push(cells);
  }
  return { rows, cols: maxCol };
}
const colLetter = i => String.fromCharCode(65 + i);
const cellStr = c => { const v = unwrap(c && c.v); return v == null ? '' : v instanceof Date ? '' : String(v).trim(); };

function detect(grid) {
  const { rows, cols } = grid;
  // date column = most parseable dates
  let dateCol = -1, best = 0;
  for (let c = 0; c < cols; c++) { const n = rows.filter(r => cellDate(r[c])).length; if (n > best) { best = n; dateCol = c; } }
  if (dateCol < 0) return null;
  const start = rows.findIndex(r => cellDate(r[dateCol]));
  // header row = nearest row above data with text in it
  let headerRow = -1;
  for (let r = start - 1; r >= 0 && r >= start - 3; r--) if (rows[r].some(c => cellStr(c) && isNaN(+cellStr(c)))) { headerRow = r; break; }
  const head = c => headerRow >= 0 ? cellStr(rows[headerRow][c]) : '';
  const dataRows = rows.slice(start).filter(r => cellDate(r[dateCol]));
  // duration column: header keyword first, else the column with most plausible durations
  let durCol = -1; best = 0;
  for (let c = 0; c < cols; c++) {
    if (c === dateCol) continue;
    const n = dataRows.filter(r => { const m = cellMinutes(r[c], 'auto-h'); return m > 0 && m <= 12 * 60 * 60; }).length;
    const score = n + (KW.dur.test(head(c)) ? dataRows.length * 0.5 : 0);
    if (score > best) { best = score; durCol = c; }
  }
  // how to read plain numbers in that column
  let unit = 'auto-h';
  if (durCol >= 0) {
    const nums = dataRows.map(r => unwrap(r[durCol].v)).map(v => typeof v === 'string' ? +v.replace(',', '.') : v)
      .filter((v, i) => typeof v === 'number' && Number.isFinite(v) && !isTimeFmt(dataRows[i][durCol].fmt));
    if (nums.length) {
      const sorted = [...nums].sort((a, b) => a - b), median = sorted[Math.floor(sorted.length / 2)];
      if (KW.minutes.test(head(durCol)) || median > 8) unit = 'auto-m';
      else {
        // "1.30" style (hours.minutes): every fraction < .60 and some aren't quarter-hours like .25/.5/.75
        const fr = nums.map(n => Math.round((n - Math.floor(n)) * 100));
        if (fr.every(f => f < 60) && fr.some(f => ![0, 25, 50, 75].includes(f))) unit = 'hm';
      }
    }
  }
  // text columns
  const textScore = c => dataRows.filter(r => { const s = cellStr(r[c]); return s && isNaN(+s) && !cellDate(r[c]) && isNaN(cellMinutes(r[c], 'auto-h')); }).length;
  const pick = (kw, exclude) => {
    let col = -1;
    for (let c = 0; c < cols; c++) if (!exclude.includes(c) && kw.test(head(c))) { col = c; break; }
    return col;
  };
  let siteCol = pick(KW.site, [dateCol, durCol]);
  if (siteCol < 0) { best = 0; for (let c = 0; c < cols; c++) { if (c === dateCol || c === durCol) continue; const s = textScore(c); if (s > best) { best = s; siteCol = c; } } }
  let notesCol = pick(KW.notes, [dateCol, durCol, siteCol]);
  return { dateCol, durCol, siteCol, notesCol, unit, start, headerRow };
}

function buildFlights(grid, map) {
  const flights = [], skipped = [], warnings = [];
  const today = todayISO();
  grid.rows.forEach((r, i) => {
    if (i < map.start) return;
    const date = cellDate(r[map.dateCol]);
    const min = map.durCol >= 0 ? cellMinutes(r[map.durCol], map.unit) : NaN;
    const empty = r.every(c => cellStr(c) === '' && !(unwrap(c.v) instanceof Date));
    if (empty) return;
    if (!date || !(min > 0)) {
      // ignore summary rows ("סה״כ" etc.) silently, report the rest
      const txt = r.map(cellStr).join(' ');
      if (!/סה.?כ|total|sum/i.test(txt) && (date || min > 0)) skipped.push(i + 1);
      return;
    }
    if (min > 8 * 60) warnings.push(`שורה ${i + 1}: טיסה של ${hm(min / 60)} שעות – בדוק שהעמודה/יחידה נכונה`);
    if (date > today) warnings.push(`שורה ${i + 1}: תאריך עתידי ${fmtDate(date)}`);
    flights.push({ id: uid(), date, min, site: map.siteCol >= 0 ? cellStr(r[map.siteCol]) : '', notes: map.notesCol >= 0 ? cellStr(r[map.notesCol]) : '' });
  });
  return { flights, skipped, warnings };
}

/* ---------- UI ---------- */
function openMapper(wb, fileName) {
  const sheets = wb.worksheets.filter(ws => ws.rowCount > 0);
  let sheetIdx = 0, grid, map;
  const pickSheet = () => {
    grid = readGrid(sheets[sheetIdx]);
    map = detect(grid) || { dateCol: -1, durCol: -1, siteCol: -1, notesCol: -1, unit: 'auto-h', start: 0, headerRow: -1 };
  };
  // start on the sheet where detection finds the most dated rows
  let bestN = -1;
  sheets.forEach((ws, i) => { const g = readGrid(ws), m = detect(g); const n = m ? g.rows.filter(r => cellDate(r[m.dateCol])).length : 0; if (n > bestN) { bestN = n; sheetIdx = i; } });
  pickSheet();
  // existing data → default to "add" so nothing (incl. maintenance counters) is wiped by accident
  let mode = S.flights.length ? 'add' : 'replace';

  const colOptions = (sel, allowNone) => {
    let o = allowNone ? `<option value="-1" ${sel < 0 ? 'selected' : ''}>— אין —</option>` : '';
    for (let c = 0; c < grid.cols; c++) {
      const h = map.headerRow >= 0 ? cellStr(grid.rows[map.headerRow][c]) : '';
      o += `<option value="${c}" ${c === sel ? 'selected' : ''}>${colLetter(c)}${h ? ' · ' + esc(h.slice(0, 22)) : ''}</option>`;
    }
    return o;
  };

  const render = () => {
    const res = map.dateCol >= 0 && map.durCol >= 0 ? buildFlights(grid, map) : { flights: [], skipped: [], warnings: [] };
    const total = res.flights.reduce((a, f) => a + f.min, 0) / 60;
    const sample = res.flights.slice(0, 4).concat(res.flights.length > 6 ? [null] : [], res.flights.slice(Math.max(4, res.flights.length - 2)));
    const warn = [...res.warnings.slice(0, 4), res.warnings.length > 4 ? `ועוד ${res.warnings.length - 4} אזהרות` : '',
      res.skipped.length ? `${res.skipped.length} שורות דולגו (חסר תאריך או זמן): ${res.skipped.slice(0, 8).join(', ')}${res.skipped.length > 8 ? '…' : ''}` : ''].filter(Boolean);
    openModal(`<h3>ייבוא מ-${esc(fileName)}</h3>
      <p class="small muted" style="margin-top:-8px">זיהיתי את העמודות אוטומטית. בדוק שהתצוגה למטה נכונה, ואם לא – שנה את הבחירה.</p>
      ${sheets.length > 1 ? `<label class="f"><span>גיליון</span><select id="mpSheet">${sheets.map((s, i) => `<option value="${i}" ${i === sheetIdx ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select></label>` : ''}
      <div class="mapgrid">
        <label class="f"><span>📅 תאריך</span><select id="mpDate">${colOptions(map.dateCol, false)}</select></label>
        <label class="f"><span>⏱ זמן טיסה</span><select id="mpDur">${colOptions(map.durCol, false)}</select></label>
        <label class="f"><span>📍 מקום</span><select id="mpSite">${colOptions(map.siteCol, true)}</select></label>
        <label class="f"><span>📝 הערות</span><select id="mpNotes">${colOptions(map.notesCol, true)}</select></label>
      </div>
      <label class="f"><span>מספרים בעמודת הזמן הם…</span><select id="mpUnit">
        <option value="auto-h" ${map.unit === 'auto-h' ? 'selected' : ''}>שעות עשרוניות (1.5 = שעה וחצי)</option>
        <option value="auto-m" ${map.unit === 'auto-m' ? 'selected' : ''}>דקות (90 = שעה וחצי)</option>
        <option value="hm" ${map.unit === 'hm' ? 'selected' : ''}>שעות.דקות (1.30 = שעה וחצי)</option>
      </select><span class="small muted" style="display:block;margin-top:4px">ערכים כמו 1:30 ותאי זמן של אקסל מזוהים תמיד נכון.</span></label>

      <div class="track-info">
        ✈ <b>${res.flights.length}</b> טיסות · סה״כ <b class="num">${hm(total)}</b> שעות מנוע
        ${res.flights.length ? ` · ${fmtDate(res.flights.reduce((a, f) => f.date < a ? f.date : a, '9999'))} – ${fmtDate(res.flights.reduce((a, f) => f.date > a ? f.date : a, '0000'))}` : ''}
      </div>
      ${res.flights.length ? `<div class="card list" style="margin-bottom:12px">${sample.map(f => f ? `<div class="flight" style="cursor:default"><div class="d small num">${fmtDate(f.date)}</div><div><div class="site">${esc(f.site || '—')}</div>${f.notes ? `<div class="note">${esc(f.notes)}</div>` : ''}</div><div class="t num">${hm(f.min / 60)}</div></div>` : '<div class="flight small muted" style="cursor:default;display:block;text-align:center">⋯</div>').join('')}</div>` : '<div class="alert bad"><span>⚠</span><span>לא נמצאו טיסות עם הבחירה הנוכחית. נסה עמודה אחרת.</span></div>'}
      ${warn.length ? `<div class="alert warn"><span>⚠</span><span class="small">${warn.map(esc).join('<br>')}</span></div>` : ''}
      ${S.flights.length ? `<label class="f"><span>מה לעשות עם ${S.flights.length} הטיסות שכבר באפליקציה?</span><select id="mpMode">
        <option value="replace" ${mode === 'replace' ? 'selected' : ''}>להחליף בתוכן הקובץ</option>
        <option value="add" ${mode === 'add' ? 'selected' : ''}>להוסיף (טיסות כפולות ידולגו)</option></select></label>` : ''}
      ${mode === 'replace' ? `<div class="alert ${S.flights.length ? 'bad' : 'info'}"><span>🔧</span><span class="small">${S.flights.length ? '<b>שים לב:</b> הטיסות, היסטוריית הטיפולים ושעוני החלקים הנוכחיים יימחקו. ' : ''}הקובץ לא כולל נתוני תחזוקה, אז שעון כל החלקים יתחיל מ-0. אחרי הייבוא עבור ללשונית תחזוקה ועדכן לכל חלק מתי הוחלף לאחרונה.</span></div>` : ''}
      <div class="btns"><button class="btn" id="mpCancel">ביטול</button><button class="btn primary" id="mpGo" ${res.flights.length ? '' : 'disabled'}>ייבא ${res.flights.length} טיסות</button></div>`, card => {
      const num = id => parseInt($(id, card).value, 10);
      const sh = $('#mpSheet', card); if (sh) sh.onchange = () => { sheetIdx = num('#mpSheet'); pickSheet(); render(); };
      $('#mpDate', card).onchange = () => { map.dateCol = num('#mpDate'); const s = grid.rows.findIndex(r => cellDate(r[map.dateCol])); map.start = s < 0 ? 0 : s; render(); };
      $('#mpDur', card).onchange = () => { map.durCol = num('#mpDur'); render(); };
      $('#mpSite', card).onchange = () => { map.siteCol = num('#mpSite'); render(); };
      $('#mpNotes', card).onchange = () => { map.notesCol = num('#mpNotes'); render(); };
      $('#mpUnit', card).onchange = () => { map.unit = $('#mpUnit', card).value; render(); };
      const md = $('#mpMode', card); if (md) md.onchange = () => { mode = md.value; render(); };
      $('#mpCancel', card).onclick = closeModal;
      $('#mpGo', card).onclick = () => {
        let flights = res.flights, added = flights.length;
        if (mode === 'add') {
          const fresh = flights.filter(f => !S.flights.some(x => x.date === f.date && Math.abs(x.min - f.min) <= 1));
          added = fresh.length; S.flights = S.flights.concat(fresh);
        } else {
          S.flights = flights;
          // a new logbook: part counters start from zero unless the user updates them
          for (const c of S.components) { c.lastResetH = 0; c.lastResetDate = ''; c.lastCheckH = 0; }
          S.maint = [];
        }
        save(); closeModal(); toast(`יובאו ${added} טיסות`); go(mode === 'add' ? 'flights' : 'maint');
      };
    });
  };
  render();
}
