/* Google Drive sync: keeps ONE Excel file in the user's Drive up to date after every change.
   Uses Google Identity Services (token flow, no server) and the narrow "drive.file" scope:
   the app can only see files it created itself, nothing else in the Drive. */
'use strict';

// OAuth "Web application" client ID from Google Cloud Console (public value, not a secret).
const GOOGLE_CLIENT_ID = '1061826604285-903mab34rrg7av68hnc9ephco0ru8odk.apps.googleusercontent.com';
const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.file';
const DRIVE_FILE_NAME = 'Paramotor Logbook.xlsx';

const drive = {
  enabled: () => !!GOOGLE_CLIENT_ID,
  connected: () => !!(S.settings.drive && S.settings.drive.fileId),
  token: null, tokenExp: 0, client: null, gisLoaded: null, syncing: false, lastError: '',
};

function loadGis() {
  if (drive.gisLoaded) return drive.gisLoaded;
  drive.gisLoaded = new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = 'https://accounts.google.com/gsi/client'; s.async = true;
    s.onload = res; s.onerror = () => { drive.gisLoaded = null; rej(new Error('offline')); };
    document.head.appendChild(s);
  });
  return drive.gisLoaded;
}

/** Returns a valid access token. interactive=false never shows UI (fails instead). */
function driveToken(interactive) {
  try { const t = JSON.parse(sessionStorage.getItem('gd-token') || 'null'); if (t && t.exp > Date.now() + 60000) { drive.token = t.v; drive.tokenExp = t.exp; } } catch {}
  if (drive.token && drive.tokenExp > Date.now() + 60000) return Promise.resolve(drive.token);
  if (!interactive) return Promise.reject(new Error('need-auth'));
  if (!window.google || !google.accounts || !google.accounts.oauth2) return Promise.reject(new Error('offline'));
  return new Promise((res, rej) => {
    drive.client = google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID, scope: DRIVE_SCOPE,
      callback: r => {
        if (r.error) return rej(new Error(r.error));
        drive.token = r.access_token; drive.tokenExp = Date.now() + (r.expires_in || 3600) * 1000;
        sessionStorage.setItem('gd-token', JSON.stringify({ v: drive.token, exp: drive.tokenExp }));
        res(drive.token);
      },
      error_callback: e => rej(new Error(e.type || 'popup')),
    });
    // prompt:'' → no consent screen if already approved (just a quick account check)
    drive.client.requestAccessToken({ prompt: S.settings.drive && S.settings.drive.fileId ? '' : 'consent', hint: S.settings.drive?.email || undefined });
  });
}

async function driveFetch(url, opts = {}) {
  const r = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + drive.token } });
  if (r.status === 401) { drive.token = null; sessionStorage.removeItem('gd-token'); throw new Error('need-auth'); }
  return r;
}

async function driveUpload(blob) {
  const id = S.settings.drive && S.settings.drive.fileId;
  if (id) { // overwrite the SAME file (Drive keeps the old versions in its history)
    const r = await driveFetch(`https://www.googleapis.com/upload/drive/v3/files/${id}?uploadType=media&fields=id,name,modifiedTime`, { method: 'PATCH', headers: { 'Content-Type': XLSX_MIME }, body: blob });
    if (r.ok) return r.json();
    if (r.status !== 404) throw new Error('drive-' + r.status);
    // 404: the user deleted the file in Drive → create it again
  }
  const meta = { name: DRIVE_FILE_NAME, mimeType: XLSX_MIME };
  const form = new FormData();
  form.append('metadata', new Blob([JSON.stringify(meta)], { type: 'application/json' }));
  form.append('file', blob);
  const r = await driveFetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,modifiedTime', { method: 'POST', body: form });
  if (!r.ok) throw new Error('drive-' + r.status);
  const f = await r.json();
  S.settings.drive = Object.assign(S.settings.drive || {}, { fileId: f.id });
  return f;
}

/** Sync now. interactive: allowed to show Google's sign-in popup (only from a tap). */
async function driveSync(interactive) {
  if (!drive.enabled() || !drive.connected()) return false;
  if (drive.syncing) { drive.again = true; return false; }
  drive.syncing = true; drive.again = false; updateHeader();
  try {
    if (!navigator.onLine) throw new Error('offline');
    await loadGis();
    await driveToken(interactive);
    await driveUpload(await workbookBlob());
    drive.lastError = '';
    S.meta.drivePending = false;
    markBackedUp();
    return true;
  } catch (e) {
    console.warn('drive sync', e);
    S.meta.drivePending = true;
    drive.lastError = e.message;
    save(false);
    return false;
  } finally {
    drive.syncing = false; updateHeader();
    if (drive.again) driveSync(false);
  }
}

async function driveConnect() {
  try {
    await loadGis();
  } catch { return toast('צריך חיבור לאינטרנט כדי להתחבר ל-Google'); }
  try {
    S.settings.drive = { fileId: '' };
    await driveToken(true);
    const info = await (await driveFetch('https://www.googleapis.com/drive/v3/about?fields=user(emailAddress)')).json();
    S.settings.drive.email = info.user && info.user.emailAddress || '';
    await driveUpload(await workbookBlob());
    S.meta.drivePending = false; markBackedUp(); save(false);
    toast('מחובר ל-Google Drive ✓ הקובץ נוצר');
  } catch (e) {
    console.error(e); S.settings.drive = null; save(false);
    toast(e.message === 'popup_closed' ? 'ההתחברות בוטלה' : 'ההתחברות ל-Google נכשלה', 4000);
  }
  if (tab === 'settings') renderSettings();
}

function driveDisconnect() {
  if (drive.token && window.google && google.accounts) try { google.accounts.oauth2.revoke(drive.token); } catch {}
  drive.token = null; sessionStorage.removeItem('gd-token');
  S.settings.drive = null; S.meta.drivePending = false; save(false);
  toast('נותק מ-Google Drive (הקובץ נשאר ב-Drive)');
}

// called once from app.js after state is loaded:
// retry pending uploads when the connection comes back / the app is reopened
function driveInit() {
  if (!drive.enabled()) return;
  window.addEventListener('online', () => { if (S.meta.drivePending) driveSync(false); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && S.meta.drivePending) driveSync(false); });
  if (drive.connected()) loadGis().then(() => { if (S.meta.drivePending) driveSync(false); }).catch(() => {});
}
