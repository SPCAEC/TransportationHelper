/**
 * Transportation Helper — Backend (Apps Script)
 * READ-ONLY on sheet. Clones Slides → per-pet PDFs (keep) → merged PDF (keep).
 */

const CFG = {
  TZ: 'America/New_York',

  SHEET_ID: '110OZsGAWmndDo07REdKQIrdR92XDBLwKgMvtfZ1oboU',
  SHEET_GID: 0,

  SLIDES_TEMPLATE_ID: '1eb_JRWgowvKckVGZ5JYc76Z0P-D1jE0MI4zqmz8JvbE',

  TEMP_FOLDER_ID: '1-JoMz-afUsCUYJu7hEx95NOpNGVa1Uht',
  INDIVIDUAL_PDFS_FOLDER_ID: '1z3XOvYJAcwpWlXddGHySMHy_BubYarlm',
  MERGED_PDFS_FOLDER_ID: '1e_NlS-TLwM4IuKmk3l6OVXXPRC7a43lc',

  MERGE_API_URL_FALLBACK: 'https://pdf-merge-service.onrender.com/merge',

  COLS: {
    DATE: 'Date',
    APPT_STATUS: 'Appointment Status',
    TRANSPORT_NEEDED: 'Transportation Needed',
    FIRST: 'First Name',
    LAST: 'Last Name',
    ADDRESS: 'Address',
    CITY: 'City',
    STATE: 'State',
    ZIP: 'Zip Code',
    PHONE: 'Phone Number',
    EMAIL: 'Email',
    PET_NAME: 'Pet Name',
    SPECIES: 'Species',
    BREED1: 'Breed One',
    BREED2: 'Breed Two',
    AGE: 'Age',
    SEX: 'Sex',
    COLOR: 'Color',
    APPT_TYPE: 'Appointment Type'
  },

  PLACEHOLDERS: {
    DATE: '{{Date}}',
    NAME: '{{Name}}',
    ADDRESS: '{{Address}}',
    ADDRESS2: '{{Address2}}',
    PHONE: '{{Phone}}',
    EMAIL: '{{Email}}',
    PET_NAME: '{{PetName}}',
    SPECIES_BREED: '{{Species_Breed}}',
    AGE_SEX_COLOR: '{{AgeSexColor}}',
    APPT_TYPE: '{{ApptType}}'
  }
};

/** Serve UI via template (so <?!= includes work) */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Transportation Helper')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);  // <— this line fixes the null issue
}

/** Script Properties */
function getScriptProp_(key, fallback) {
  try {
    const v = PropertiesService.getScriptProperties().getProperty(key);
    return (v !== null && v !== undefined && v !== '') ? v : fallback;
  } catch (e) {
    Logger.log('getScriptProp_ error for key %s: %s', key, e);
    return fallback;
  }
}
function getMainHubLink() { return getScriptProp_('MAIN_HUB_LINK', ''); }
function getMergeApiUrl_() { return getScriptProp_('MERGE_API_URL', CFG.MERGE_API_URL_FALLBACK); }

/** Utils */
function S_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function joinNonEmpty_(parts, sep) { return parts.filter(p => S_(p)).join(sep); }

/** Parse date from:
 *  - a real Date object, or
 *  - a text cell like "MM/DD/YYYY" or "M/D/YYYY"
 */
function parseSheetDate_(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val)) return val;

  const s = String(val).trim();
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const month = parseInt(m[1], 10) - 1; // 0-based
    const day   = parseInt(m[2], 10);
    const year  = parseInt(m[3], 10);
    const dt    = new Date(year, month, day);
    return isNaN(dt) ? null : dt;
  }

  // Fallback: native parse
  const dt = new Date(s);
  return isNaN(dt) ? null : dt;
}

/** Compare using TZ-safe yyyymmdd strings */
function isTodayOrTomorrow_(dt) {
  if (!(dt instanceof Date) || isNaN(dt)) return false;
  const tz = CFG.TZ;
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyyMMdd');
  const tomorrowStr = Utilities.formatDate(new Date(now.getTime() + 86400000), tz, 'yyyyMMdd');
  const dateStr = Utilities.formatDate(dt, tz, 'yyyyMMdd');
  return dateStr === todayStr || dateStr === tomorrowStr;
}

/** Strictly fetch the sheet by GID (no silent fallback) */
function getStrictSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();
  for (const sh of sheets) if (sh.getSheetId() === gid) return sh;
  throw new Error(`Sheet with GID ${gid} not found in ${CFG.SHEET_ID}`);
}

/** READ-ONLY: list appointments needing transport today/tomorrow */
function getTransportAppointments(targetDateStr) {
  const tz = CFG.TZ;
  const targetDate = targetDateStr ? parseSheetDate_(targetDateStr) : null;
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const sh = getStrictSheetByGid_(ss, CFG.SHEET_GID);
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const header = values[0].map(h => S_(h));
  const rows = values.slice(1);
  const colIndex = {};
  header.forEach((name, idx) => colIndex[name] = idx);

  const out = [];

  rows.forEach((r, i) => {
    const get = (colName) => r[colIndex[colName]];
    const apptStatus = S_(get(CFG.COLS.APPT_STATUS));
    const transNeeded = S_(get(CFG.COLS.TRANSPORT_NEEDED));
    const dateObj = parseSheetDate_(get(CFG.COLS.DATE));

    if (apptStatus !== 'Scheduled') return;
    if (transNeeded.toLowerCase() !== 'yes') return;
    if (!dateObj) return;

    // Date filter logic
    let include = false;
    if (targetDate) {
      const dateStr = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');
      const targetStr = Utilities.formatDate(targetDate, tz, 'yyyyMMdd');
      include = (dateStr === targetStr);
    } else {
      include = isTodayOrTomorrow_(dateObj);
    }
    if (!include) return;

    const first = S_(get(CFG.COLS.FIRST));
    const last  = S_(get(CFG.COLS.LAST));
    const name  = joinNonEmpty_([first, last], ' ');

    const address1 = S_(get(CFG.COLS.ADDRESS));
    const city  = S_(get(CFG.COLS.CITY));
    const state = S_(get(CFG.COLS.STATE));
    const zip   = S_(get(CFG.COLS.ZIP));
    const address2 = joinNonEmpty_([city, state, zip], ', ');

    const phone = S_(get(CFG.COLS.PHONE));
    const email = S_(get(CFG.COLS.EMAIL));
    const petName = S_(get(CFG.COLS.PET_NAME));
    const apptType = S_(get(CFG.COLS.APPT_TYPE));

    const species = S_(get(CFG.COLS.SPECIES));
    const b1 = S_(get(CFG.COLS.BREED1));
    const b2 = S_(get(CFG.COLS.BREED2));
    const speciesBreed = joinNonEmpty_([species, joinNonEmpty_([b1, b2], ' / ')], ' • ');

    const age = S_(get(CFG.COLS.AGE));
    const sex = S_(get(CFG.COLS.SEX));
    const color = S_(get(CFG.COLS.COLOR));
    const ageSexColor = joinNonEmpty_([age, sex, color], ' • ');

    const dateStr = Utilities.formatDate(dateObj, tz, 'MMMM d, yyyy');

    out.push({
      rowIndex: i + 2,
      dateRaw: dateObj,
      date: dateStr,
      name,
      address1,
      address2,
      phone,
      email,
      petName,
      apptType,
      speciesBreed,
      ageSexColor
    });
  });

  Logger.log('getTransportAppointments → %s rows', out.length);
  return JSON.parse(JSON.stringify(out));
}

/** Create Transportation Contracts (sheet is never modified) */
function createTransportContracts() {
  const tz = CFG.TZ;
  const appts = getTransportAppointments();
  if (!appts.length) {
    return { ok: false, message: 'No transport appointments for today/tomorrow.', individuals: [], merged: null };
  }

  const tempFolder   = DriveApp.getFolderById(CFG.TEMP_FOLDER_ID);
  const indivFolder  = DriveApp.getFolderById(CFG.INDIVIDUAL_PDFS_FOLDER_ID);
  const mergedFolder = DriveApp.getFolderById(CFG.MERGED_PDFS_FOLDER_ID);
  const templateFile = DriveApp.getFileById(CFG.SLIDES_TEMPLATE_ID);

  const individualPdfs = [];
  const tempClones = [];

  try {
    appts.forEach(a => {
      const cloneName = `TransportContract_${sanitizeName_(a.name)}_${Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss')}`;
      const clone = templateFile.makeCopy(cloneName, tempFolder);
      tempClones.push(clone);

      const pres = SlidesApp.openById(clone.getId());
      const map = {};
      map[CFG.PLACEHOLDERS.DATE]          = a.date || '';
      map[CFG.PLACEHOLDERS.NAME]          = a.name || '';
      map[CFG.PLACEHOLDERS.ADDRESS]       = a.address1 || '';
      map[CFG.PLACEHOLDERS.ADDRESS2]      = a.address2 || '';
      map[CFG.PLACEHOLDERS.PHONE]         = a.phone || '';
      map[CFG.PLACEHOLDERS.EMAIL]         = a.email || '';
      map[CFG.PLACEHOLDERS.PET_NAME]      = a.petName || '';
      map[CFG.PLACEHOLDERS.SPECIES_BREED] = a.speciesBreed || '';
      map[CFG.PLACEHOLDERS.AGE_SEX_COLOR] = a.ageSexColor || '';
      map[CFG.PLACEHOLDERS.APPT_TYPE]     = a.apptType || '';
      replaceInPresentation_(pres, map);

      const pdfBlob = DriveApp.getFileById(clone.getId()).getAs(MimeType.PDF);
      const pdfName = `${cloneName}.pdf`;
      const pdfFile = indivFolder.createFile(pdfBlob).setName(pdfName);
      individualPdfs.push({ id: pdfFile.getId(), name: pdfName, url: pdfFile.getUrl() });
    });

    tempClones.forEach(f => { try { f.setTrashed(true); } catch (e) { Logger.log('Temp clone trash fail %s → %s', f.getId(), e); } });

    const outputName = `Transportation_Contracts_${Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss')}.pdf`;
    const merged = mergePDFs_(individualPdfs, outputName);

    let mergedFileMeta = null;
    if (merged && merged.contentBase64) {
      const mergedBlob = Utilities.newBlob(Utilities.base64Decode(merged.contentBase64), MimeType.PDF, merged.fileName || outputName);
      const mergedFile = mergedFolder.createFile(mergedBlob).setName(merged.fileName || outputName);
      mergedFileMeta = { id: mergedFile.getId(), name: mergedFile.getName(), url: mergedFile.getUrl() };
    } else if (merged && merged.fileUrl) {
      mergedFileMeta = { id: null, name: outputName, url: merged.fileUrl };
    }

    return { ok: true, count: appts.length, individuals: individualPdfs, merged: mergedFileMeta };

  } catch (err) {
    Logger.log('createTransportContracts error: %s', err.stack || err);
    return { ok: false, message: 'Error creating transportation contracts. See logs.', error: String(err), individuals: individualPdfs, merged: null };
  }
}

/** Replace placeholders across all slides in a presentation (clone only) */
function replaceInPresentation_(presentation, map) {
  const slides = presentation.getSlides();
  Object.keys(map).forEach(needle => {
    const value = map[needle];
    slides.forEach(slide => {
      try { slide.replaceAllText(needle, value); }
      catch (e) { Logger.log('replaceAllText error on %s → %s', needle, e); }
    });
  });
  SlidesApp.flush();
}

/** Merge PDFs via Render service */
function mergePDFs_(pdfs, outputName) {
  const url = getMergeApiUrl_();
  Logger.log('Merging %s PDFs via %s', pdfs.length, url);

  const files = pdfs.map(p => {
    const blob = DriveApp.getFileById(p.id).getBlob();
    return { name: p.name || (p.id + '.pdf'), contentBase64: Utilities.base64Encode(blob.getBytes()) };
  });

  const payload = JSON.stringify({ outputName: outputName || 'merged.pdf', files });

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  Logger.log('Merge response %s: %s', code, text.substring(0, Math.min(500, text.length)));

  if (code >= 200 && code < 300) {
    try { return JSON.parse(text); }
    catch (_) { throw new Error('Invalid JSON from merge API'); }
  }
  throw new Error('Merge API error: ' + code + ' — ' + text);
}

/** Filename sanitizer */
function sanitizeName_(s) { return S_(s).replace(/[^\w\-. ]+/g, '_').slice(0, 80); }

function testTransport() {
  const data = getTransportAppointments();
  Logger.log(JSON.stringify(data, null, 2));
  return data;
}