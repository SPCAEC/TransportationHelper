/**
 * Transportation Helper — Backend (v2)
 * Generates transport contracts for a selected date only.
 * Sheet is read-only. Exports clean per-pet PDFs, then merges them.
 */

const CFG = {
  TZ: 'America/New_York',

  SHEET_ID: '110OZsGAWmndDo07REdKQIrdR92XDBLwKgMvtfZ1oboU',
  SHEET_GID: 0,

  SLIDES_TEMPLATE_ID: '1eb_JRWgowvKckVGZ5JYc76Z0P-D1jE0MI4zqmz8JvbE',

  TEMP_FOLDER_ID: '1-JoMz-afUsCUYJu7hEx95NOpNGVa1Uht',
  INDIVIDUAL_PDFS_FOLDER_ID: '1z3XOvYJAcwpWlXddGHySMHy_BubYarlm',
  MERGED_PDFS_FOLDER_ID: '1e_NlS-TLwM4IuKmk3l6OVXXPRC7a43lc',

  MERGE_API_URL: 'https://pdf-merge-service.onrender.com/merge',

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

/** Serve UI */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Transportation Helper')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getMainHubLink() {
  return PropertiesService.getScriptProperties().getProperty('MAIN_HUB_LINK') || '';
}

/** Basic utils */
const S_ = v => (v == null) ? '' : String(v).trim();
const join_ = (arr, sep) => arr.filter(Boolean).join(sep);
const sanitize_ = s => S_(s).replace(/[^\w\-. ]+/g, '_').slice(0, 80);

/** Parse sheet date text like 10/15/2025 */
function parseSheetDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  const m = String(v).trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(+m[3], +m[1]-1, +m[2]);
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

/** Get filtered appointments for specific date (YYYY-MM-DD) */
function getTransportAppointments(targetDateStr) {
  const tz = CFG.TZ;
  if (!targetDateStr) throw new Error('Missing target date');

  const [y, m, d] = targetDateStr.split('-').map(Number);
  const targetDate = new Date(y, m - 1, d);
  const targetStr = Utilities.formatDate(targetDate, tz, 'yyyyMMdd');

  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const sh = ss.getSheets().find(s => s.getSheetId() === CFG.SHEET_GID);
  const values = sh.getDataRange().getValues();
  const header = values[0].map(S_);
  const rows = values.slice(1);

  const idx = {};
  header.forEach((h, i) => idx[h] = i);

  const out = [];

  rows.forEach((r, i) => {
    const apptStatus = S_(r[idx[CFG.COLS.APPT_STATUS]]);
    const transNeeded = S_(r[idx[CFG.COLS.TRANSPORT_NEEDED]]);
    const dateObj = parseSheetDate_(r[idx[CFG.COLS.DATE]]);
    if (apptStatus !== 'Scheduled' || transNeeded.toLowerCase() !== 'yes' || !dateObj) return;

    const rowStr = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');
    if (rowStr !== targetStr) return;

    const name = join_([S_(r[idx[CFG.COLS.FIRST]]), S_(r[idx[CFG.COLS.LAST]])], ' ');
    const address1 = S_(r[idx[CFG.COLS.ADDRESS]]);
    const address2 = join_([
      S_(r[idx[CFG.COLS.CITY]]),
      S_(r[idx[CFG.COLS.STATE]]),
      S_(r[idx[CFG.COLS.ZIP]])
    ], ', ');
    const speciesBreed = join_([
      S_(r[idx[CFG.COLS.SPECIES]]),
      join_([S_(r[idx[CFG.COLS.BREED1]]), S_(r[idx[CFG.COLS.BREED2]])], ' / ')
    ], ' • ');
    const ageSexColor = join_([
      S_(r[idx[CFG.COLS.AGE]]),
      S_(r[idx[CFG.COLS.SEX]]),
      S_(r[idx[CFG.COLS.COLOR]])
    ], ' • ');

    out.push({
      date: Utilities.formatDate(dateObj, tz, 'MMMM d, yyyy'),
      name,
      address1,
      address2,
      phone: S_(r[idx[CFG.COLS.PHONE]]),
      email: S_(r[idx[CFG.COLS.EMAIL]]),
      petName: S_(r[idx[CFG.COLS.PET_NAME]]),
      apptType: S_(r[idx[CFG.COLS.APPT_TYPE]]),
      speciesBreed,
      ageSexColor
    });
  });

  Logger.log(`getTransportAppointments(${targetDateStr}) → ${out.length} rows`);
  return out;
}

/** Create contracts for the selected date only */
function createTransportContracts(targetDateStr) {
  const tz = CFG.TZ;
  const appts = getTransportAppointments(targetDateStr);
  if (!appts.length)
    return { ok: false, message: `No transport appointments for ${targetDateStr}.` };

  const tempFolder = DriveApp.getFolderById(CFG.TEMP_FOLDER_ID);
  const indivFolder = DriveApp.getFolderById(CFG.INDIVIDUAL_PDFS_FOLDER_ID);
  const mergedFolder = DriveApp.getFolderById(CFG.MERGED_PDFS_FOLDER_ID);
  const template = DriveApp.getFileById(CFG.SLIDES_TEMPLATE_ID);

  const individualPdfs = [];

  appts.forEach(a => {
    const clone = template.makeCopy(
      `Transport_${sanitize_(a.name)}_${Utilities.formatDate(new Date(), tz, 'HHmmss')}`,
      tempFolder
    );
    const pres = SlidesApp.openById(clone.getId());

    // Reliable replacement at presentation level
    Object.entries(CFG.PLACEHOLDERS).forEach(([k, ph]) => {
      const val = a[k.toLowerCase()] || a[k] || '';
      try { pres.replaceAllText(ph, val); } catch (err) {}
    });
    Utilities.sleep(200);

    const blob = clone.getAs(MimeType.PDF);
    if (blob.getBytes().length === 0) throw new Error('Empty PDF export');

    const file = indivFolder.createFile(blob).setName(`${a.petName || 'Pet'}_${a.name}.pdf`);
    individualPdfs.push({ id: file.getId(), name: file.getName(), url: file.getUrl() });

    try { clone.setTrashed(true); } catch (e) {}
  });

  Utilities.sleep(1500);
  const merged = mergePDFs_(individualPdfs, `Transport_${targetDateStr}.pdf`);
  if (!merged || !merged.contentBase64)
    return { ok: false, message: 'Merge failed', individuals: individualPdfs };

  const mergedBlob = Utilities.newBlob(
    Utilities.base64Decode(merged.contentBase64),
    MimeType.PDF,
    merged.fileName || `Transport_${targetDateStr}.pdf`
  );
  const mergedFile = DriveApp.getFolderById(CFG.MERGED_PDFS_FOLDER_ID).createFile(mergedBlob);
  return { ok: true, count: appts.length, merged: { url: mergedFile.getUrl() } };
}

/** Merge PDFs via Render service */
function mergePDFs_(pdfs, outputName) {
  const files = pdfs.map(p => {
    const blob = DriveApp.getFileById(p.id).getBlob();
    return { name: p.name, contentBase64: Utilities.base64Encode(blob.getBytes()) };
  });
  const payload = JSON.stringify({ outputName, files });
  const res = UrlFetchApp.fetch(CFG.MERGE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    payload,
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  Logger.log(`Merge response ${code}: ${text.substring(0, 300)}`);
  if (code >= 200 && code < 300) return JSON.parse(text);
  throw new Error(`Merge API error: ${code}`);
}