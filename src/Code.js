/**
 * Transportation Helper — Backend (Apps Script)
 * Generates transport contracts based on selected date (Today / Tomorrow / Custom).
 * - Reads sheet (read-only)
 * - Filters by chosen date + transport needed + scheduled
 * - Clones template per pet → replaces placeholders → exports to PDF → keeps individual PDFs
 * - Merges all PDFs via Render service → keeps merged file
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
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/** Script Properties helpers */
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

/** Utility functions */
function S_(v) { return (v == null) ? '' : String(v).trim(); }
function joinNonEmpty_(arr, sep) { return arr.filter(Boolean).join(sep); }
function sanitizeName_(s) { return S_(s).replace(/[^\w\-. ]+/g, '_').slice(0, 80); }

/** Parse a sheet date value (text "MM/DD/YYYY" or Date) */
function parseSheetDate_(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val)) return val;

  const s = S_(val);
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2]));
  const dt = new Date(s);
  return isNaN(dt) ? null : dt;
}

/** Strictly get sheet by GID */
function getStrictSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();
  for (const sh of sheets) if (sh.getSheetId() === gid) return sh;
  throw new Error(`Sheet with GID ${gid} not found in ${CFG.SHEET_ID}`);
}

/**
 * Get transport appointments for a specific date string (YYYY-MM-DD)
 * Returns array of records
 */
function getTransportAppointments(targetDateStr) {
  const tz = CFG.TZ;
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const sh = getStrictSheetByGid_(ss, CFG.SHEET_GID);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(h => S_(h));
  const rows = values.slice(1);
  const idx = {};
  header.forEach((h, i) => idx[h] = i);

  const [y, m, d] = targetDateStr.split('-').map(Number);
  const selectedDate = new Date(y, m - 1, d);
  const selectedStr = Utilities.formatDate(selectedDate, tz, 'yyyyMMdd');

  const out = [];

  rows.forEach((r, i) => {
    const apptStatus = S_(r[idx[CFG.COLS.APPT_STATUS]]);
    const transNeeded = S_(r[idx[CFG.COLS.TRANSPORT_NEEDED]]);
    const dateObj = parseSheetDate_(r[idx[CFG.COLS.DATE]]);
    if (apptStatus !== 'Scheduled' || transNeeded.toLowerCase() !== 'yes' || !dateObj) return;

    const rowDateStr = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');
    if (rowDateStr !== selectedStr) return;

    const first = S_(r[idx[CFG.COLS.FIRST]]);
    const last = S_(r[idx[CFG.COLS.LAST]]);
    const name = joinNonEmpty_([first, last], ' ');
    const address1 = S_(r[idx[CFG.COLS.ADDRESS]]);
    const address2 = joinNonEmpty_([
      S_(r[idx[CFG.COLS.CITY]]),
      S_(r[idx[CFG.COLS.STATE]]),
      S_(r[idx[CFG.COLS.ZIP]])
    ], ', ');

    const phone = S_(r[idx[CFG.COLS.PHONE]]);
    const email = S_(r[idx[CFG.COLS.EMAIL]]);
    const petName = S_(r[idx[CFG.COLS.PET_NAME]]);
    const apptType = S_(r[idx[CFG.COLS.APPT_TYPE]]);
    const species = S_(r[idx[CFG.COLS.SPECIES]]);
    const b1 = S_(r[idx[CFG.COLS.BREED1]]);
    const b2 = S_(r[idx[CFG.COLS.BREED2]]);
    const speciesBreed = joinNonEmpty_([species, joinNonEmpty_([b1, b2], ' / ')], ' • ');
    const age = S_(r[idx[CFG.COLS.AGE]]);
    const sex = S_(r[idx[CFG.COLS.SEX]]);
    const color = S_(r[idx[CFG.COLS.COLOR]]);
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

  Logger.log(`getTransportAppointments(${targetDateStr}) → ${out.length} rows`);
  return out;
}

/**
 * Create Transportation Contracts for the given date (passed from frontend)
 */
function createTransportContracts(targetDateStr) {
  const tz = CFG.TZ;
  const appts = getTransportAppointments(targetDateStr);
  if (!appts.length) {
    return { ok: false, message: `No transport appointments for ${targetDateStr}.`, individuals: [], merged: null };
  }

  const tempFolder = DriveApp.getFolderById(CFG.TEMP_FOLDER_ID);
  const indivFolder = DriveApp.getFolderById(CFG.INDIVIDUAL_PDFS_FOLDER_ID);
  const mergedFolder = DriveApp.getFolderById(CFG.MERGED_PDFS_FOLDER_ID);
  const templateFile = DriveApp.getFileById(CFG.SLIDES_TEMPLATE_ID);

  const individualPdfs = [];
  const tempClones = [];

  try {
    appts.forEach(a => {
      // 1) Clone the template
      const cloneName = `TransportContract_${sanitizeName_(a.name)}_${Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss')}`;
      const clone = templateFile.makeCopy(cloneName, tempFolder);
      tempClones.push(clone);

      // 2) Replace placeholders
      const pres = SlidesApp.openById(clone.getId());
      const map = {
        [CFG.PLACEHOLDERS.DATE]: a.date,
        [CFG.PLACEHOLDERS.NAME]: a.name,
        [CFG.PLACEHOLDERS.ADDRESS]: a.address1,
        [CFG.PLACEHOLDERS.ADDRESS2]: a.address2,
        [CFG.PLACEHOLDERS.PHONE]: a.phone,
        [CFG.PLACEHOLDERS.EMAIL]: a.email,
        [CFG.PLACEHOLDERS.PET_NAME]: a.petName,
        [CFG.PLACEHOLDERS.SPECIES_BREED]: a.speciesBreed,
        [CFG.PLACEHOLDERS.AGE_SEX_COLOR]: a.ageSexColor,
        [CFG.PLACEHOLDERS.APPT_TYPE]: a.apptType
      };
      replaceInPresentation_(pres, map);

      // 3) Export as PDF → Individual folder
      Utilities.sleep(400);
      const pdfBlob = DriveApp.getFileById(clone.getId()).getAs(MimeType.PDF);
      const pdfFile = indivFolder.createFile(pdfBlob).setName(`${cloneName}.pdf`);
      if (pdfFile.getSize() <= 0) throw new Error(`Empty PDF generated for ${a.name}`);

      individualPdfs.push({ id: pdfFile.getId(), name: pdfFile.getName(), url: pdfFile.getUrl() });
    });

    Logger.log(`Created ${individualPdfs.length} individual PDFs`);

    // 4) Clean up temp clones
    tempClones.forEach(f => { try { f.setTrashed(true); } catch (e) { Logger.log(`Failed to trash temp clone: ${e}`); } });

    // 5) Merge PDFs
    Utilities.sleep(1500);
    const outputName = `Transportation_Contracts_${targetDateStr.replace(/-/g, '')}.pdf`;
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
    Logger.log(`createTransportContracts(${targetDateStr}) error: ${err.stack || err}`);
    return { ok: false, message: `Error creating contracts for ${targetDateStr}.`, error: String(err), individuals: individualPdfs, merged: null };
  }
}

/** Replace placeholders across entire presentation */
function replaceInPresentation_(presentation, map) {
  Object.keys(map).forEach(k => {
    try {
      presentation.replaceAllText(k, map[k] || '');
    } catch (e) {
      Logger.log(`replaceAllText error on ${k} → ${e}`);
    }
  });
  Utilities.sleep(300);
}

/** Merge PDFs via Render service */
function mergePDFs_(pdfs, outputName) {
  const url = getMergeApiUrl_();
  Logger.log(`Merging ${pdfs.length} PDFs via ${url}`);
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
  Logger.log(`Merge response ${code}: ${text.substring(0, 300)}`);
  if (code >= 200 && code < 300) {
    try { return JSON.parse(text); }
    catch (_) { throw new Error('Invalid JSON from merge API'); }
  }
  throw new Error(`Merge API error: ${code} — ${text}`);
}

/** Testing helper */
function testTransport() {
  const todayStr = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd');
  const data = getTransportAppointments(todayStr);
  Logger.log(JSON.stringify(data, null, 2));
}