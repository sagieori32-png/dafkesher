/**
 * דף קשר לכיתה — שרת קטן על Google Sheets.
 *
 * התקנה (פעם אחת):
 * 1. בגיליון Google Sheets חדש: Extensions (תוספים) → Apps Script.
 * 2. מדביקים את כל הקובץ הזה במקום הקוד הקיים ושומרים.
 * 3. משנים למטה את CLASS_CODE ו-ADMIN_PIN, בוחרים בפונקציה setup ולוחצים Run (ומאשרים הרשאות).
 * 4. Deploy → New deployment → סוג: Web app.
 *    Execute as: Me.  Who has access: Anyone.  → Deploy, ומעתיקים את ה-Web app URL.
 */

// ---- לשנות לפני ההרצה הראשונה של setup ----
const CLASS_CODE = 'g1-tashpaz';   // קוד הכיתה שמופיע בלינק להורים
const ADMIN_PIN  = '1234';          // קוד הניהול של המחנכת
// ---------------------------------------------

const KID_COLS = ['id','token','firstName','lastName','p1Name','p1Phone','p2Name','p2Phone','photo','source','updatedAt'];
const MAX_CELL = 49000;

function setup() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CLASS_CODE', CLASS_CODE);
  props.setProperty('ADMIN_PIN', ADMIN_PIN);
  const ss = SpreadsheetApp.getActive();
  let kids = ss.getSheetByName('kids');
  if (!kids) kids = ss.insertSheet('kids');
  kids.getRange(1, 1, kids.getMaxRows(), KID_COLS.length).setNumberFormat('@'); // טלפונים נשמרים כטקסט, עם ה-0 בהתחלה
  kids.getRange(1, 1, 1, KID_COLS.length).setValues([KID_COLS]);
  let cfg = ss.getSheetByName('config');
  if (!cfg) {
    cfg = ss.insertSheet('config');
    cfg.getRange(1, 1, 7, 2).setValues([
      ['classLabel', 'ג 1'], ['hebYear', 'תשפ״ז'], ['year', '2026-2027'],
      ['school', 'בית הספר הממלכתי הנריאטה סאלד'], ['teacherName', 'אלינור לוי'],
      ['teacherPhone', '052-8496724'], ['teacherPhoto', '']
    ]);
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function prop(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function checkCode(k) { if (!k || String(k) !== prop('CLASS_CODE')) throw new Error('bad_code'); }
function checkPin(p) { if (!p || String(p) !== prop('ADMIN_PIN')) throw new Error('bad_pin'); }

function kidsSheet() { return SpreadsheetApp.getActive().getSheetByName('kids'); }
function readKids() {
  const sh = kidsSheet(); const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, KID_COLS.length).getValues().map((r, i) => {
    const o = { _row: i + 2 }; KID_COLS.forEach((c, j) => o[c] = r[j] === null ? '' : String(r[j])); return o;
  }).filter(o => o.id);
}
function readConfig() {
  const sh = SpreadsheetApp.getActive().getSheetByName('config'); const o = {};
  const n = sh.getLastRow(); if (n < 1) return o;
  sh.getRange(1, 1, n, 2).getValues().forEach(r => { if (r[0]) o[String(r[0])] = String(r[1] ?? ''); });
  return o;
}
function clean(s, max) { return String(s == null ? '' : s).trim().slice(0, max || 80); }
function cleanPhoto(p) {
  p = String(p || '');
  if (!p) return '';
  if (!/^data:image\/(jpeg|png|webp);base64,/.test(p)) throw new Error('bad_photo');
  if (p.length > MAX_CELL) throw new Error('photo_too_big');
  return p;
}
function kidRow(id, token, k, source) {
  return [id, token, clean(k.firstName), clean(k.lastName), clean(k.p1Name), clean(k.p1Phone, 20),
          clean(k.p2Name), clean(k.p2Phone, 20), cleanPhoto(k.photo), source, new Date().toISOString()];
}
function publicKid(o) { const c = Object.assign({}, o); delete c.token; delete c._row; return c; }

function doGet(e) {
  try {
    checkCode(e.parameter.k);
    return out({ ok: true, cfg: readConfig(), kids: readKids().map(publicKid) });
  } catch (err) { return out({ ok: false, error: err.message }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    checkCode(b.k);
    lock.waitLock(20000);
    const sh = kidsSheet();

    if (b.action === 'save') {               // הורה: יצירה או עדכון של השורה שלו
      const kid = b.kid || {};
      if (!clean(kid.firstName) || !clean(kid.lastName)) throw new Error('missing_name');
      const existing = b.id ? readKids().find(o => o.id === b.id) : null;
      if (existing && existing.token === b.token) {
        sh.getRange(existing._row, 1, 1, KID_COLS.length).setValues([kidRow(existing.id, existing.token, kid, existing.source || 'parent')]);
        return out({ ok: true, id: existing.id, token: existing.token });
      }
      const id = Utilities.getUuid().slice(0, 8), token = Utilities.getUuid();
      sh.appendRow(kidRow(id, token, kid, 'parent'));
      return out({ ok: true, id: id, token: token });
    }

    checkPin(b.pin);                         // מכאן והלאה: ניהול בלבד
    if (b.action === 'checkPin') return out({ ok: true });

    if (b.action === 'adminSave') {
      const kid = b.kid || {};
      if (!clean(kid.firstName) || !clean(kid.lastName)) throw new Error('missing_name');
      const existing = b.id ? readKids().find(o => o.id === b.id) : null;
      if (existing) {
        sh.getRange(existing._row, 1, 1, KID_COLS.length).setValues([kidRow(existing.id, existing.token, kid, existing.source || 'admin')]);
        return out({ ok: true, id: existing.id });
      }
      const id = Utilities.getUuid().slice(0, 8);
      sh.appendRow(kidRow(id, Utilities.getUuid(), kid, 'admin'));
      return out({ ok: true, id: id });
    }

    if (b.action === 'adminDelete') {
      const existing = readKids().find(o => o.id === b.id);
      if (existing) sh.deleteRow(existing._row);
      return out({ ok: true });
    }

    if (b.action === 'saveConfig') {
      const c = b.cfg || {};
      const keys = ['classLabel','hebYear','year','school','teacherName','teacherPhone','teacherPhoto'];
      const rows = keys.map(key => [key, key === 'teacherPhoto' ? cleanPhoto(c[key]) : clean(c[key], 120)]);
      const csh = SpreadsheetApp.getActive().getSheetByName('config');
      csh.clearContents(); csh.getRange(1, 1, rows.length, 2).setValues(rows);
      return out({ ok: true });
    }

    throw new Error('unknown_action');
  } catch (err) {
    return out({ ok: false, error: err.message });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
