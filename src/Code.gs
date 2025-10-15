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

  // Overridable via Script Property MERGE_API_URL
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
    .setSandboxMode(HtmlService.SandboxMode.IFRAME); // preserves prior null fix
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
function sanitizeName_(s) { return S_(s).replace(/[^\w\-. ]+/g, '_').slice(0, 80); }

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

/** Compare using TZ-safe yyyymmdd strings (compat fallback for no-date calls) */
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

/** Server-side helpers for accurate dates (fixes timezone drift on client) */
function getServerDate(offsetDays) {
  const tz = CFG.TZ;
  const base = new Date(new Date().getTime() + (offsetDays || 0) * 86400000);
  return Utilities.formatDate(base, tz, 'yyyy-MM-dd');
}
function debugServerNow() {
  return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm:ss z');
}

/**
 * READ-ONLY: list appointments needing transport for a **specific date** (YYYY-MM-DD).
 * If targetDateStr is falsy, preserves legacy behavior (today or tomorrow).
 */
function getTransportAppointments(targetDateStr) {
  const tz = CFG.TZ;
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const sh = getStrictSheetByGid_(ss, CFG.SHEET_GID);
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const header = values[0].map(h => S_(h));
  const rows = values.slice(1);
  const colIndex = {};
  header.forEach((name, idx) => colIndex[name] = idx);

  let selectedYMD = null;
  if (targetDateStr) {
    const [y, m, d] = targetDateStr.split('-').map(n => parseInt(n, 10));
    const target = new Date(y, m - 1, d);
    selectedYMD = Utilities.formatDate(target, tz, 'yyyyMMdd');
  }

  const out = [];

  rows.forEach((r, i) => {
    const get = (colName) => r[colIndex[colName]];
    const apptStatus = S_(get(CFG.COLS.APPT_STATUS));
    const transNeeded = S_(get(CFG.COLS.TRANSPORT_NEEDED));
    const dateObj = parseSheetDate_(get(CFG.COLS.DATE));
    if (apptStatus !== 'Scheduled') return;
    if (transNeeded.toLowerCase() !== 'yes') return;
    if (!dateObj) return;

    const rowYMD = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');
    let include = false;
    if (selectedYMD) {
      include = (rowYMD === selectedYMD);
    } else {
      // legacy: today or tomorrow
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

    const dateStr = Utilities.formatDate(dateObj, tz, 'MM/dd/yyyy'); // enhancement: mm/dd/yyyy

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

  Logger.log('getTransportAppointments(%s) → %s rows', targetDateStr || '(today|tomorrow)', out.length);
  return JSON.parse(JSON.stringify(out));
}

/**
 * Create Transportation Contracts for the **selected date only** (YYYY-MM-DD).
 * If called without a date, preserves legacy behavior (today/tomorrow).
 */
function createTransportContracts(targetDateStr) {
  const tz = CFG.TZ;
  const appts = getTransportAppointments(targetDateStr);
  if (!appts.length) {
    return { ok: false, message: 'No transport appointments for the selected date.', individuals: [], merged: null };
  }

  const tempFolder   = DriveApp.getFolderById(CFG.TEMP_FOLDER_ID);
  const indivFolder  = DriveApp.getFolderById(CFG.INDIVIDUAL_PDFS_FOLDER_ID);
  const mergedFolder = DriveApp.getFolderById(CFG.MERGED_PDFS_FOLDER_ID);
  const templateFile = DriveApp.getFileById(CFG.SLIDES_TEMPLATE_ID);

  const individualPdfs = [];
  const createdNames = [];

  try {
    appts.forEach((a, idx) => {
      const cloneName = `TransportContract_${sanitizeName_(a.name)}_${Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss')}_${idx+1}`;
      const clone = templateFile.makeCopy(cloneName, tempFolder);
      const pres = SlidesApp.openById(clone.getId());

      // Reliable presentation-level replacement
      const map = {};
      map[CFG.PLACEHOLDERS.DATE]           = a.date || '';
      map[CFG.PLACEHOLDERS.NAME]           = a.name || '';
      map[CFG.PLACEHOLDERS.ADDRESS]        = a.address1 || '';
      map[CFG.PLACEHOLDERS.ADDRESS2]       = a.address2 || '';
      map[CFG.PLACEHOLDERS.PHONE]          = a.phone || '';
      map[CFG.PLACEHOLDERS.EMAIL]          = a.email || '';
      map[CFG.PLACEHOLDERS.PET_NAME]       = a.petName || '';
      map[CFG.PLACEHOLDERS.SPECIES_BREED]  = a.speciesBreed || '';
      map[CFG.PLACEHOLDERS.AGE_SEX_COLOR]  = a.ageSexColor || '';
      map[CFG.PLACEHOLDERS.APPT_TYPE]      = a.apptType || '';

      replaceInPresentation_(pres, map);   // replaces + saveAndClose inside
      Utilities.sleep(400);                // give Drive a breath

      const pdfBlob = clone.getAs(MimeType.PDF); // export after saveAndClose
      const size = pdfBlob.getBytes().length;
      if (size < 1000) {
        Logger.log('WARNING: Skipping zero/very small PDF for %s (size=%s bytes)', cloneName, size);
        try { clone.setTrashed(true); } catch (e) {}
        return; // skip pushing invalid file
      }

      const pdfFile = indivFolder.createFile(pdfBlob).setName(`${cloneName}.pdf`);
      individualPdfs.push({ id: pdfFile.getId(), name: pdfFile.getName(), url: pdfFile.getUrl(), size });
      createdNames.push(pdfFile.getName());

      // Trash the temp Slides clone after successful export
      try { clone.setTrashed(true); } catch (e) { Logger.log('Temp clone trash fail %s → %s', clone.getId(), e); }
    });

    Logger.log('Created %s valid individual PDFs: %s', individualPdfs.length, createdNames.join(', '));

    if (!individualPdfs.length) {
      return { ok: false, message: 'All generated PDFs were empty or invalid. Please check the template placeholders.', individuals: [], merged: null };
    }

    // Merge PDFs via Render (wait a beat for Drive to finalize)
    Utilities.sleep(1200);
    const baseName = targetDateStr ? targetDateStr.replace(/-/g, '') : Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
    const outputName = `Transportation_Contracts_${baseName}.pdf`;
    const merged = mergePDFs_(individualPdfs, outputName);

    let mergedFileMeta = null;
    if (merged && merged.contentBase64) {
      const mergedBlob = Utilities.newBlob(Utilities.base64Decode(merged.contentBase64), MimeType.PDF, merged.fileName || outputName);
      const mergedFile = mergedFolder.createFile(mergedBlob).setName(merged.fileName || outputName);
      mergedFileMeta = { id: mergedFile.getId(), name: mergedFile.getName(), url: mergedFile.getUrl() };
    } else if (merged && merged.fileUrl) {
      mergedFileMeta = { id: null, name: outputName, url: merged.fileUrl };
    } else {
      return { ok: false, message: 'Merge service returned no file.', individuals: individualPdfs, merged: null };
    }

    return { ok: true, count: individualPdfs.length, individuals: individualPdfs, merged: mergedFileMeta };

  } catch (err) {
    Logger.log('createTransportContracts(%s) error: %s', targetDateStr, err.stack || err);
    return { ok: false, message: 'Error creating transportation contracts. See logs.', error: String(err), individuals: individualPdfs, merged: null };
  }
}

/** Replace placeholders across entire presentation then save (flushes edits) */
function replaceInPresentation_(presentation, map) {
  try {
    // presentation.replaceAllText is powerful but some templates need slide-level fallback
    Object.keys(map).forEach(needle => {
      const value = (map[needle] == null) ? '' : String(map[needle]);
      try { presentation.replaceAllText(needle, value); }
      catch (e) { Logger.log('presentation.replaceAllText error on %s → %s', needle, e); }
    });

    // As a safety net, also run per-slide replacement for any stubborn elements
    const slides = presentation.getSlides();
    slides.forEach(slide => {
      Object.keys(map).forEach(needle => {
        const value = (map[needle] == null) ? '' : String(map[needle]);
        try { slide.replaceAllText(needle, value); }
        catch (e) { /* ignore individual element errors */ }
      });
    });

  } catch (e) {
    Logger.log('replaceInPresentation_ global error: %s', e);
  } finally {
    // 🔑 Ensure edits are committed before export
    try { presentation.saveAndClose(); } catch (e) { Logger.log('saveAndClose error: %s', e); }
  }
}

/** Merge PDFs via Render service */
/** Merge PDFs via Render service */
function mergePDFs_(pdfs, outputName) {
  const url = getMergeApiUrl_();
  Logger.log('Merging %s PDFs via %s', pdfs.length, url);

  const files = pdfs.map(p => {
    try {
      const blob = DriveApp.getFileById(p.id).getBlob();
      const bytes = blob.getBytes();
      if (!bytes || !bytes.length) {
        Logger.log('⚠️ Skipping invalid or empty blob for %s', p.name);
        return null;
      }
      const base64 = Utilities.base64Encode(bytes);
      if (!base64) {
        Logger.log('⚠️ Skipping PDF with missing base64 data: %s', p.name);
        return null;
      }
      return { name: p.name || (p.id + '.pdf'), contentBase64: base64 };
    } catch (err) {
      Logger.log('⚠️ Error reading blob for %s: %s', p.name, err);
      return null;
    }
  }).filter(Boolean);

  if (!files.length) {
    throw new Error('No valid PDFs to merge.');
  }

  const payload = JSON.stringify({ outputName: outputName || 'merged.pdf', files });
  Logger.log('Payload prepared: %s files', files.length);

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

/** Test helper(s) */
function testTransport() {
  const today = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd');
  const data = getTransportAppointments(today);
  Logger.log(JSON.stringify(data, null, 2));
  return data;
}
function pingMergeService() {
  const url = 'https://pdf-merge-service.onrender.com/';
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    Logger.log('Pinged merge service → %s', code);
  } catch (err) {
    Logger.log('Ping error: %s', err);
  }
}