/**
 * Reference Data Setup — One-time functions to create and populate
 * the platform reference spreadsheet (ARDA religious data + ZIP-to-county)
 *
 * Run in order:
 *   1. createReferenceSpreadsheet() — creates the spreadsheet, logs the ID
 *   2. Set REFERENCE_SPREADSHEET_ID below, clasp push
 *   3. importZipData() — fetches ZIP-to-county (public URL, Drive fallback)
 *   4. importArdaData() — fetches ARDA Excel from public source, processes it
 *   5. verifyReferenceData() — confirms everything matches
 *
 * No manual uploads needed — everything fetches from public sources.
 * Drive fallback CSVs (arda_religious_data.csv, zip_to_county_raw.csv)
 * can be uploaded as backup if public sources go down.
 */

// REFERENCE_SPREADSHEET_ID is defined in Code.js

// Denominations we track (ARDA group codes → column names)
var TRACKED_DENOMINATIONS = {
  '019': 'American Baptist',
  '053': 'Assemblies of God',
  '081': 'Catholic Church',
  '097': 'Christian Churches',
  '123': 'Church of God (Anderson IN)',
  '127': 'Church of God (Cleveland TN)',
  '141': 'Church of God in Christ',
  '151': 'Latter-day Saints',
  '165': 'Church of the Nazarene',
  '167': 'Churches of Christ',
  '193': 'Episcopal Church',
  '207': 'Evangelical Lutheran (ELCA)',
  '283': 'Lutheran Church Missouri Synod',
  '355': 'Presbyterian Church USA',
  '413': 'Seventh-day Adventist',
  '419': 'Southern Baptist Convention',
  '449': 'United Methodist Church',
  '500': 'Non-denominational'
};


// ── STEP 1: Create the spreadsheet ──

function createReferenceSpreadsheet() {
  var ss = SpreadsheetApp.create('Church Health Platform — Reference Data');
  var id = ss.getId();

  var sheet1 = ss.getSheets()[0];
  sheet1.setName('ReligiousData');

  ss.insertSheet('ZipToCounty');

  var orgSheet = ss.insertSheet('Organizations');
  orgSheet.getRange(1, 1, 1, 10).setValues([[
    'OrgID', 'OrgName', 'State', 'DenominationCode', 'DenominationName',
    'AdminName', 'AdminEmail', 'Status', 'DateCreated', 'Counties'
  ]]);
  orgSheet.getRange(1, 1, 1, 10).setFontWeight('bold');

  var denomSheet = ss.insertSheet('DenominationLookup');
  var denomData = [
    ['Code', 'Name', 'ShortName'],
    ['019', 'American Baptist Churches in the USA', 'American Baptist'],
    ['053', 'Assemblies of God', 'Assemblies of God'],
    ['081', 'Catholic Church', 'Catholic'],
    ['097', 'Christian Churches and Churches of Christ', 'Christian Churches'],
    ['123', 'Church of God (Anderson, Indiana)', 'Church of God (Anderson)'],
    ['127', 'Church of God (Cleveland, Tennessee)', 'Church of God (Cleveland)'],
    ['141', 'Church of God in Christ', 'Church of God in Christ'],
    ['151', 'Church of Jesus Christ of Latter-day Saints', 'Latter-day Saints'],
    ['165', 'Church of the Nazarene', 'Church of the Nazarene'],
    ['167', 'Churches of Christ', 'Churches of Christ'],
    ['193', 'Episcopal Church', 'Episcopal'],
    ['207', 'Evangelical Lutheran Church in America', 'ELCA Lutheran'],
    ['283', 'Lutheran Church--Missouri Synod', 'LCMS Lutheran'],
    ['355', 'Presbyterian Church (U.S.A.)', 'Presbyterian (USA)'],
    ['413', 'Seventh-day Adventist Church', 'Seventh-day Adventist'],
    ['419', 'Southern Baptist Convention', 'Southern Baptist'],
    ['449', 'United Methodist Church', 'United Methodist'],
    ['500', 'Non-denominational Christian Churches', 'Non-denominational']
  ];
  denomSheet.getRange(1, 1, denomData.length, 3).setValues(denomData);
  denomSheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  Logger.log('Reference spreadsheet created: ' + id);
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('Set REFERENCE_SPREADSHEET_ID = "' + id + '" then clasp push');
  return id;
}


// ── STEP 2: Import ZIP-to-county data ──

function importZipData() {
  if (!REFERENCE_SPREADSHEET_ID) {
    throw new Error('Set REFERENCE_SPREADSHEET_ID first — run createReferenceSpreadsheet()');
  }

  var csvText = '';

  // Try public URL first
  try {
    Logger.log('Fetching ZIP data from public source...');
    var response = UrlFetchApp.fetch(
      'https://raw.githubusercontent.com/ninken/US-ZipCode-Data/master/US_ZIP_FIPS_DataExtract.csv',
      { muteHttpExceptions: true }
    );
    if (response.getResponseCode() === 200) {
      csvText = response.getContentText();
      Logger.log('Fetched from public URL: ' + csvText.length + ' bytes');
    } else {
      throw new Error('HTTP ' + response.getResponseCode());
    }
  } catch (e) {
    Logger.log('Public URL failed (' + e.message + ') — trying Drive fallback...');
    var files = DriveApp.searchFiles(
      'title contains "zip_to_county" and trashed = false'
    );
    if (!files.hasNext()) {
      throw new Error(
        'Public URL failed and no zip_to_county file on Drive. ' +
        'Upload zip_to_county_raw.csv to Drive and retry.'
      );
    }
    csvText = files.next().getBlob().getDataAsString();
    Logger.log('Loaded from Drive fallback.');
  }

  var parsed = Utilities.parseCsv(csvText);
  Logger.log('Parsed ' + parsed.length + ' rows. Processing...');

  // Build: ZipCode, CountyFIPS, State, County, City
  var rows = [['ZipCode', 'CountyFIPS', 'State', 'County', 'City']];
  for (var i = 1; i < parsed.length; i++) {
    var raw = parsed[i];
    if (!raw[0]) continue;

    var zip = String(raw[0]).replace(/\s/g, '');
    while (zip.length < 5) zip = '0' + zip;
    var stateFips = String(raw[1]).replace(/\s/g, '');
    while (stateFips.length < 2) stateFips = '0' + stateFips;
    var countyFips3 = String(raw[4]).replace(/\s/g, '');
    while (countyFips3.length < 3) countyFips3 = '0' + countyFips3;

    rows.push([zip, stateFips + countyFips3, raw[3] || '', raw[5] || '', raw[6] || '']);
  }

  writeToSheet_(REFERENCE_SPREADSHEET_ID, 'ZipToCounty', rows, 5);
  Logger.log('ZIP import complete: ' + (rows.length - 1) + ' ZIP codes.');
}


// ── STEP 3: Import ARDA religious data ──
//
// Fetches the ARDA 2020 group-detail Excel from usreligioncensus.org,
// converts to Google Sheets via Drive API, reads denomination data per county,
// pivots into our flat format, and writes to the ReligiousData sheet.
//
// Falls back to a pre-processed CSV on Drive if the public source fails.

function importArdaData() {
  if (!REFERENCE_SPREADSHEET_ID) {
    throw new Error('Set REFERENCE_SPREADSHEET_ID first — run createReferenceSpreadsheet()');
  }

  // ── Try fetching and processing from public ARDA source ──
  try {
    Logger.log('Fetching ARDA data from public source...');
    var rows = fetchAndProcessArda_();
    writeToSheet_(REFERENCE_SPREADSHEET_ID, 'ReligiousData', rows, rows[0].length);
    Logger.log('ARDA import complete: ' + (rows.length - 1) + ' counties from public source.');
    return;
  } catch (e) {
    Logger.log('Public ARDA fetch failed: ' + e.message);
    Logger.log('Trying Drive fallback...');
  }

  // ── Fallback: pre-processed CSV on Drive ──
  var files = DriveApp.searchFiles(
    'title contains "arda_religious_data" and trashed = false'
  );
  if (!files.hasNext()) {
    throw new Error(
      'Public ARDA source failed and no arda_religious_data CSV on Drive. ' +
      'Upload arda_religious_data.csv to Drive and retry.'
    );
  }

  var csvText = files.next().getBlob().getDataAsString();
  var parsed = Utilities.parseCsv(csvText);
  Logger.log('Loaded from Drive: ' + parsed.length + ' rows');

  // Convert numeric columns
  var top3Col = parsed[0].length - 1;
  for (var i = 1; i < parsed.length; i++) {
    for (var j = 3; j < top3Col; j++) {
      parsed[i][j] = Number(parsed[i][j]) || 0;
    }
  }

  writeToSheet_(REFERENCE_SPREADSHEET_ID, 'ReligiousData', parsed, parsed[0].length);
  Logger.log('ARDA import complete: ' + (parsed.length - 1) + ' counties from Drive fallback.');
}


/**
 * Fetches ARDA Excel files from usreligioncensus.org, converts to Sheets,
 * reads county summary + group detail, pivots into our flat format.
 */
function fetchAndProcessArda_() {
  // Fetch county summary (population, total adherents)
  Logger.log('  Fetching county summary...');
  var summaryUrl = 'https://www.usreligioncensus.org/sites/default/files/2023-06/2020_USRC_Summaries.xlsx';
  var summaryData = fetchExcelAsSheetData_(summaryUrl, '2020 County Summary');

  // Build county lookup: FIPS → {state, county, pop, adherents, congs}
  var counties = {};
  for (var i = 1; i < summaryData.length; i++) {
    var row = summaryData[i];
    var fips = String(row[0]).trim();
    if (!fips) continue;
    while (fips.length < 5) fips = '0' + fips;
    counties[fips] = {
      state: row[1] || '',
      county: row[2] || '',
      pop: Number(row[3]) || 0,
      adherents: Number(row[5]) || 0,
      congs: Number(row[4]) || 0
    };
  }
  Logger.log('  County summaries loaded: ' + Object.keys(counties).length);

  // Fetch group detail (denomination data per county)
  Logger.log('  Fetching group detail (this may take a moment)...');
  var detailUrl = 'https://www.usreligioncensus.org/sites/default/files/2023-06/2020_USRC_Group_Detail.xlsx';
  var detailData = fetchExcelAsSheetData_(detailUrl, '2020 Group by County');

  // Build denomination data per county
  // detailData columns: FIPS, State, County, GroupCode, GroupName, Congs, Adherents, ...
  var countyDenom = {};  // fips → { code → {congs, adherents} }
  var allDenoms = {};    // fips → [{name, adherents}] for top 3

  for (var i = 1; i < detailData.length; i++) {
    var row = detailData[i];
    var fips = String(row[0]).trim();
    if (!fips) continue;
    while (fips.length < 5) fips = '0' + fips;

    var code = String(row[3]).trim();
    while (code.length < 3) code = '0' + code;
    var congs = Number(row[5]) || 0;
    var adherents = Number(row[6]) || 0;

    // Track for top 3
    if (adherents > 0) {
      if (!allDenoms[fips]) allDenoms[fips] = [];
      allDenoms[fips].push({ name: row[4], adherents: adherents });
    }

    // Track our target denominations
    if (TRACKED_DENOMINATIONS[code]) {
      if (!countyDenom[fips]) countyDenom[fips] = {};
      countyDenom[fips][code] = { congs: congs, adherents: adherents };
    }
  }
  Logger.log('  Denomination detail processed.');

  // Build sorted denomination codes for consistent column order
  var sortedCodes = Object.keys(TRACKED_DENOMINATIONS).sort();

  // Build header
  var header = ['FIPS', 'State', 'County', 'Population', 'TotalAdherents', 'TotalCongregations'];
  for (var c = 0; c < sortedCodes.length; c++) {
    var name = TRACKED_DENOMINATIONS[sortedCodes[c]];
    header.push(name + ' Congs');
    header.push(name + ' Adherents');
  }
  header.push('Top3');

  // Build data rows
  var rows = [header];
  var fipsList = Object.keys(counties).sort();
  for (var f = 0; f < fipsList.length; f++) {
    var fips = fipsList[f];
    var cs = counties[fips];
    var row = [fips, cs.state, cs.county, cs.pop, cs.adherents, cs.congs];

    var denomData = countyDenom[fips] || {};
    for (var c = 0; c < sortedCodes.length; c++) {
      var d = denomData[sortedCodes[c]] || { congs: 0, adherents: 0 };
      row.push(d.congs);
      row.push(d.adherents);
    }

    // Top 3
    var top3 = '';
    if (allDenoms[fips]) {
      allDenoms[fips].sort(function(a, b) { return b.adherents - a.adherents; });
      var top = [];
      for (var t = 0; t < Math.min(3, allDenoms[fips].length); t++) {
        var d = allDenoms[fips][t];
        top.push(d.name + ' (' + d.adherents.toLocaleString() + ')');
      }
      top3 = top.join('; ');
    }
    row.push(top3);
    rows.push(row);
  }

  Logger.log('  Built ' + (rows.length - 1) + ' county rows with ' + header.length + ' columns.');
  return rows;
}


/**
 * Fetches an Excel file from a URL, saves to Drive as Google Sheets,
 * reads a specific sheet tab, then cleans up the temp file.
 * Returns 2D array of values.
 */
function fetchExcelAsSheetData_(url, sheetName) {
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to fetch ' + url + ' — HTTP ' + response.getResponseCode());
  }

  var blob = response.getBlob();
  // Create temp file on Drive, converting xlsx → Google Sheets
  var resource = {
    title: '_temp_arda_import_' + Date.now(),
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };
  var tempFile = Drive.Files.insert(resource, blob, {
    convert: true
  });

  var tempSS = SpreadsheetApp.openById(tempFile.id);
  var sheet = tempSS.getSheetByName(sheetName);
  if (!sheet) {
    // Clean up and fail
    DriveApp.getFileById(tempFile.id).setTrashed(true);
    throw new Error('Sheet "' + sheetName + '" not found in downloaded file.');
  }

  var data = sheet.getDataRange().getValues();
  Logger.log('  Read ' + data.length + ' rows from "' + sheetName + '"');

  // Clean up temp file
  DriveApp.getFileById(tempFile.id).setTrashed(true);
  return data;
}


// ── Helper: write 2D array to a sheet in batches ──

function writeToSheet_(ssId, sheetName, rows, numCols) {
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(sheetName);
  sheet.clearContents();

  var batchSize = 5000;
  for (var start = 0; start < rows.length; start += batchSize) {
    var end = Math.min(start + batchSize, rows.length);
    var batch = rows.slice(start, end);
    sheet.getRange(start + 1, 1, batch.length, numCols).setValues(batch);
    SpreadsheetApp.flush();
    Logger.log('  Written rows ' + (start + 1) + ' to ' + end);
  }
  sheet.getRange(1, 1, 1, numCols).setFontWeight('bold');
}


// ── VERIFICATION ──

function verifyReferenceData() {
  if (!REFERENCE_SPREADSHEET_ID) {
    throw new Error('Set REFERENCE_SPREADSHEET_ID first');
  }

  var ss = SpreadsheetApp.openById(REFERENCE_SPREADSHEET_ID);

  var rdSheet = ss.getSheetByName('ReligiousData');
  var rdRows = rdSheet.getLastRow();
  Logger.log('ReligiousData: ' + (rdRows - 1) + ' counties');

  var zcSheet = ss.getSheetByName('ZipToCounty');
  var zcRows = zcSheet.getLastRow();
  Logger.log('ZipToCounty: ' + (zcRows - 1) + ' ZIP codes');

  // Verify Dallas County (FIPS 48113)
  var rdData = rdSheet.getDataRange().getValues();
  var headers = rdData[0];
  for (var i = 1; i < rdData.length; i++) {
    if (String(rdData[i][0]) === '48113') {
      Logger.log('\nDallas County (48113):');
      Logger.log('  Population: ' + rdData[i][3]);
      Logger.log('  Total Adherents: ' + rdData[i][4]);
      for (var j = 0; j < headers.length; j++) {
        if (String(headers[j]).indexOf('Assemblies of God') >= 0) {
          Logger.log('  ' + headers[j] + ': ' + rdData[i][j]);
        }
      }
      break;
    }
  }

  // Verify ZIP 75001 → 48113
  var zcData = zcSheet.getDataRange().getValues();
  for (var i = 1; i < zcData.length; i++) {
    if (String(zcData[i][0]) === '75001') {
      var match = String(zcData[i][1]) === '48113' ? 'MATCH' : 'MISMATCH';
      Logger.log('\nZIP 75001 → County FIPS: ' + zcData[i][1] + ' (expect 48113) ' + match);
      break;
    }
  }

  Logger.log('\nDenominations tracked: ' + Object.keys(TRACKED_DENOMINATIONS).length);
  var denomSheet = ss.getSheetByName('DenominationLookup');
  var denomData = denomSheet.getDataRange().getValues();
  for (var i = 1; i < denomData.length; i++) {
    Logger.log('  ' + denomData[i][0] + ': ' + denomData[i][2]);
  }

  Logger.log('\nVerification complete.');
}
