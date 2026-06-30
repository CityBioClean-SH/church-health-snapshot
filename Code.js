/**
 * Church Health Snapshot — Backend
 * Google Apps Script server-side code
 * Scoring formulas are server-side only — never exposed to the client
 */

const SPREADSHEET_ID = '1f_zyugFWUkuepoGJA0MEGx2lM3i3xDHb8PcOJiNqXCQ';
const REFERENCE_SPREADSHEET_ID = '1LprfzdsT3IbRxfYFAWjdphRl9rqweovcVj4oF8K1WSM';
const ADMIN_EMAIL = ''; // Set via setupAdmin()

// ── SETUP ──

function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Users sheet
  let users = ss.getSheetByName('Users');
  if (!users) {
    users = ss.insertSheet('Users');
    users.appendRow(['PIN', 'Name', 'Email', 'Role', 'Church', 'DateRegistered', 'LastLogin', 'Active', 'FirstLogin']);
    users.getRange('A1:I1').setFontWeight('bold').setBackground('#1b2541').setFontColor('white');
    // Add demo users
    users.appendRow(['1234', 'Steve Harper', '', 'Admin', '', new Date(), '', 'Y', 'N']);
    users.appendRow(['5678', 'Pastor Jones', '', 'Pastor', 'First Assembly', new Date(), '', 'Y', 'N']);
    users.appendRow(['0000', 'District Lead', '', 'District', '', new Date(), '', 'Y', 'N']);
    users.setColumnWidth(1, 80);
    users.setColumnWidth(2, 160);
    users.setColumnWidth(3, 200);
    users.setColumnWidth(4, 80);
    users.setColumnWidth(5, 160);
  }

  // Assessments sheet
  let assess = ss.getSheetByName('Assessments');
  if (!assess) {
    assess = ss.insertSheet('Assessments');
    const headers = [
      'Timestamp', 'AssessedBy', 'Role', 'ChurchName', 'PastorName',
      'Q1_ChildrenMinistry', 'Q2_YouthMinistry', 'Q3_PastorAge', 'Q4_CongAge',
      'Q5_Attendance', 'Q6_AttendanceTrend', 'Q7_ChurchStatus', 'Q8_FinanceTrend',
      'Q9_Income', 'Q10_Conversions', 'Q11_Baptisms', 'Q12_SpiritBaptisms',
      'Q13_Guests', 'Q14_CityType', 'Q15_Population', 'Q16_PopChange',
      'Q17_MedianAge', 'Q18_PerCapitaIncome', 'Q19_DominantFamily',
      'Q20_Facilities', 'Q21_StrategicThinking', 'Q22_Volunteers',
      'Q23_Denial', 'Q24_PrayerHours', 'Q25_Vision', 'Q26_Leadership',
      'Q27_SamePage', 'Q28_StatusQuo', 'Q29_Decor',
      'Notes1', 'Notes2', 'Notes3',
      'Score', 'Rating', 'S1Score', 'S2Score', 'S3Score'
    ];
    assess.appendRow(headers);
    assess.getRange('A1:AO1').setFontWeight('bold').setBackground('#1b2541').setFontColor('white');
    assess.setFrozenRows(1);
  }

  // Audit Log sheet
  let audit = ss.getSheetByName('AuditLog');
  if (!audit) {
    audit = ss.insertSheet('AuditLog');
    audit.appendRow(['Timestamp', 'User', 'Action', 'Details']);
    audit.getRange('A1:D1').setFontWeight('bold').setBackground('#1b2541').setFontColor('white');
    audit.setFrozenRows(1);
  }

  // Remove default Sheet1 if it exists
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }

  Logger.log('Setup complete. Spreadsheet: ' + SPREADSHEET_ID);
}

function resetUsersSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var existing = ss.getSheetByName('Users');
  if (existing) ss.deleteSheet(existing);

  var users = ss.insertSheet('Users');
  users.appendRow(['PIN', 'Name', 'Email', 'Role', 'Church', 'DateRegistered', 'LastLogin', 'Active', 'FirstLogin']);
  users.getRange('A1:I1').setFontWeight('bold').setBackground('#1b2541').setFontColor('white');
  users.appendRow(['1234', 'Steve Harper', 'steve@citybioclean.com', 'Admin', '', new Date(), '', 'Y', 'N']);
  users.appendRow(['5678', 'Mike Harper', 'mharper@northtexas.ag', 'Admin', '', new Date(), '', 'Y', 'N']);
  users.setColumnWidth(1, 80);
  users.setColumnWidth(2, 160);
  users.setColumnWidth(3, 200);
  users.setColumnWidth(4, 80);
  users.setColumnWidth(5, 160);
  users.setColumnWidth(9, 80);
  Logger.log('Users sheet reset with FirstLogin column. Steve + Mike added as Admin.');
}

// ── MULTI-TENANT MIGRATION (run once) ──

function setupMultiTenant() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Add OrgID column (J) to Users if not present
  var users = ss.getSheetByName('Users');
  var headers = users.getRange(1, 1, 1, users.getLastColumn()).getValues()[0];
  if (headers.indexOf('OrgID') === -1) {
    var nextCol = headers.length + 1;
    users.getRange(1, nextCol).setValue('OrgID').setFontWeight('bold');
    // Set existing users to 'NT' (North Texas)
    var lastRow = users.getLastRow();
    if (lastRow > 1) {
      for (var i = 2; i <= lastRow; i++) {
        users.getRange(i, nextCol).setValue('NT');
      }
    }
    Logger.log('Added OrgID column to Users. Existing users set to NT.');
  }

  // Create Organizations sheet if not present
  if (!ss.getSheetByName('Organizations')) {
    var orgSheet = ss.insertSheet('Organizations');
    orgSheet.appendRow([
      'OrgID', 'OrgName', 'State', 'DenominationCode', 'DenominationName',
      'AdminName', 'AdminEmail', 'AdminNotifyEmail', 'Status', 'DateCreated', 'Counties'
    ]);
    orgSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1b2541').setFontColor('white');
    // Add North Texas as first org
    orgSheet.appendRow([
      'NT', 'North Texas District', 'Texas', '053', 'Assemblies of God',
      'Mike Harper', 'mharper@northtexas.ag', 'mharper@northtexas.ag',
      'Active', new Date(), ''
    ]);
    Logger.log('Created Organizations sheet with North Texas.');
  }

  // Create Churches sheet if not present
  if (!ss.getSheetByName('Churches')) {
    var churchSheet = ss.insertSheet('Churches');
    churchSheet.appendRow([
      'ChurchID', 'OrgID', 'ChurchName', 'City', 'State', 'ZIP',
      'PastorName', 'PastorEmail', 'SectionLeaderRow', 'Status', 'DateAdded'
    ]);
    churchSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1b2541').setFontColor('white');
    Logger.log('Created Churches sheet.');
  }

  // Add OrgID and ChurchID columns to Assessments if not present
  var assess = ss.getSheetByName('Assessments');
  var aHeaders = assess.getRange(1, 1, 1, assess.getLastColumn()).getValues()[0];
  if (aHeaders.indexOf('OrgID') === -1) {
    var nextCol = aHeaders.length + 1;
    assess.getRange(1, nextCol).setValue('OrgID').setFontWeight('bold');
    assess.getRange(1, nextCol + 1).setValue('ChurchID').setFontWeight('bold');
    assess.getRange(1, nextCol + 2).setValue('AssessorRole').setFontWeight('bold');
    // Tag existing assessments as NT
    var lastRow = assess.getLastRow();
    if (lastRow > 1) {
      for (var i = 2; i <= lastRow; i++) {
        assess.getRange(i, nextCol).setValue('NT');
      }
    }
    Logger.log('Added OrgID, ChurchID, AssessorRole columns to Assessments.');
  }

  Logger.log('Multi-tenant migration complete.');
}

function setupCensusKey() {
  PropertiesService.getScriptProperties().setProperty('CENSUS_API_KEY', 'b084691f5c7c1ab4b82caec5dbf0988f2d31dca8');
  Logger.log('Census API key stored in Script Properties.');
}

function lookupCensusData(zip) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('CENSUS_API_KEY');
    if (!apiKey) return { success: false, error: 'Census API key not configured. Run setupCensusKey() in Code.js.' };

    zip = String(zip).trim().replace(/\D/g, '');
    if (zip.length !== 5) return { success: false, error: 'Invalid ZIP code' };

    // ACS 5-year estimates — most recent available
    // B01003_001E = Total population
    // B01002_001E = Median age
    // B19013_001E = Median household income
    // B19301_001E = Per capita income
    var vars = 'NAME,B01003_001E,B01002_001E,B19013_001E,B19301_001E';
    var url2023 = 'https://api.census.gov/data/2023/acs/acs5?get=' + vars + '&for=zip%20code%20tabulation%20area:' + zip + '&key=' + apiKey;
    var url2020 = 'https://api.census.gov/data/2020/acs/acs5?get=' + vars + '&for=zip%20code%20tabulation%20area:' + zip + '&key=' + apiKey;

    // Fetch current data
    var response = UrlFetchApp.fetch(url2023, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      // Fall back to 2022
      var url2022 = url2023.replace('/2023/', '/2022/');
      response = UrlFetchApp.fetch(url2022, { muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        return { success: false, error: 'No Census data found for ZIP ' + zip };
      }
    }

    var data = JSON.parse(response.getContentText());
    if (data.length < 2) return { success: false, error: 'No data returned for ZIP ' + zip };

    var row = data[1];
    var placeName = row[0];
    var population = parseInt(row[1]) || 0;
    var medianAge = parseFloat(row[2]) || 0;
    var medianHouseholdIncome = parseInt(row[3]) || 0;
    var perCapitaIncome = parseInt(row[4]) || 0;

    // Determine city type from population
    var cityType = '';
    if (population < 8000) cityType = 'Rural (Under 8000)';
    else if (population <= 50000) cityType = 'Small city (8k-50k)';
    else if (population <= 100000) cityType = 'Large City (over 50k)';
    else cityType = 'Urban/Metro (100k+)';

    // Fetch older data for population change comparison
    var popChange = 'Static';
    try {
      var oldResponse = UrlFetchApp.fetch(url2020, { muteHttpExceptions: true });
      if (oldResponse.getResponseCode() === 200) {
        var oldData = JSON.parse(oldResponse.getContentText());
        if (oldData.length >= 2) {
          var oldPop = parseInt(oldData[1][1]) || 0;
          if (oldPop > 0 && population > 0) {
            var pctChange = ((population - oldPop) / oldPop) * 100;
            if (pctChange > 2) popChange = 'Increasing';
            else if (pctChange < -2) popChange = 'Declining';
            else popChange = 'Static';
          }
        }
      }
    } catch(e) {
      // If older data fails, default to Static
    }

    // Get city name, state, and city population
    var cityName = '';
    var stateName = '';
    var stateAbbr = '';
    var cityPop = 0;
    try {
      var zipInfoResp = UrlFetchApp.fetch('https://api.zippopotam.us/us/' + zip, { muteHttpExceptions: true });
      if (zipInfoResp.getResponseCode() === 200) {
        var zipInfo = JSON.parse(zipInfoResp.getContentText());
        if (zipInfo.places && zipInfo.places.length > 0) {
          cityName = zipInfo.places[0]['place name'];
          stateName = zipInfo.places[0]['state'];
          stateAbbr = zipInfo.places[0]['state abbreviation'] || '';

          // Get city population from Census using the correct state FIPS
          var stateFipsMap = {'AL':'01','AK':'02','AZ':'04','AR':'05','CA':'06','CO':'08','CT':'09','DE':'10','DC':'11','FL':'12','GA':'13','HI':'15','ID':'16','IL':'17','IN':'18','IA':'19','KS':'20','KY':'21','LA':'22','ME':'23','MD':'24','MA':'25','MI':'26','MN':'27','MS':'28','MO':'29','MT':'30','NE':'31','NV':'32','NH':'33','NJ':'34','NM':'35','NY':'36','NC':'37','ND':'38','OH':'39','OK':'40','OR':'41','PA':'42','RI':'44','SC':'45','SD':'46','TN':'47','TX':'48','UT':'49','VT':'50','VA':'51','WA':'53','WV':'54','WI':'55','WY':'56'};
          var sFips = stateFipsMap[stateAbbr] || '48';
          var placesUrl = 'https://api.census.gov/data/2022/acs/acs5?get=NAME,B01003_001E&for=place:*&in=state:' + sFips + '&key=' + apiKey;
          var placesResp = UrlFetchApp.fetch(placesUrl, { muteHttpExceptions: true });
          if (placesResp.getResponseCode() === 200) {
            var placesData = JSON.parse(placesResp.getContentText());
            var cityLower = cityName.toLowerCase();
            for (var p = 1; p < placesData.length; p++) {
              var pName = placesData[p][0].toLowerCase();
              if (pName.indexOf(cityLower + ' city') === 0 || pName.indexOf(cityLower + ' cdp') === 0 || pName.indexOf(cityLower + ' town') === 0) {
                cityPop = parseInt(placesData[p][1]) || 0;
                break;
              }
            }
          }
        }
      }
    } catch(e) {
      // City lookup is optional — don't fail the whole request
    }

    logAudit('SYSTEM', 'CENSUS_LOOKUP', 'ZIP: ' + zip + ' — ' + (cityName || placeName) + (stateAbbr ? ', ' + stateAbbr : '') + ', Pop: ' + population + (cityPop ? ', City Pop: ' + cityPop : ''));

    return {
      success: true,
      placeName: placeName,
      cityName: cityName,
      stateName: stateName,
      stateAbbr: stateAbbr,
      cityPop: cityPop,
      population: population,
      medianAge: Math.round(medianAge * 10) / 10,
      medianHouseholdIncome: medianHouseholdIncome,
      perCapitaIncome: perCapitaIncome,
      cityType: cityType,
      popChange: popChange
    };
  } catch(e) {
    return { success: false, error: 'Census lookup failed: ' + e.message };
  }
}

function lookupZipFromCity(cityName, stateCode) {
  try {
    cityName = String(cityName).trim();
    if (!cityName) return { success: false, error: 'City name is required' };
    stateCode = String(stateCode || '').toLowerCase().trim();

    if (!stateCode) {
      return { success: false, error: 'Enter a state abbreviation with the city, or use a ZIP code instead.' };
    }

    // Accept full state names — convert to abbreviation
    var stateMap = {'alabama':'al','alaska':'ak','arizona':'az','arkansas':'ar','california':'ca','colorado':'co','connecticut':'ct','delaware':'de','district of columbia':'dc','florida':'fl','georgia':'ga','hawaii':'hi','idaho':'id','illinois':'il','indiana':'in','iowa':'ia','kansas':'ks','kentucky':'ky','louisiana':'la','maine':'me','maryland':'md','massachusetts':'ma','michigan':'mi','minnesota':'mn','mississippi':'ms','missouri':'mo','montana':'mt','nebraska':'ne','nevada':'nv','new hampshire':'nh','new jersey':'nj','new mexico':'nm','new york':'ny','north carolina':'nc','north dakota':'nd','ohio':'oh','oklahoma':'ok','oregon':'or','pennsylvania':'pa','rhode island':'ri','south carolina':'sc','south dakota':'sd','tennessee':'tn','texas':'tx','utah':'ut','vermont':'vt','virginia':'va','washington':'wa','west virginia':'wv','wisconsin':'wi','wyoming':'wy'};
    if (stateCode.length > 2 && stateMap[stateCode]) {
      stateCode = stateMap[stateCode];
    }

    var citySlug = cityName.toLowerCase().replace(/\s+/g, '%20');
    var resp = UrlFetchApp.fetch('https://api.zippopotam.us/us/' + stateCode + '/' + citySlug, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      return { success: false, error: 'City "' + cityName + '" not found in ' + stateCode.toUpperCase() + '. Try entering a ZIP code.' };
    }

    var data = JSON.parse(resp.getContentText());
    if (!data.places || data.places.length === 0) {
      return { success: false, error: 'No ZIP codes found for ' + cityName };
    }

    // Return the first ZIP code for the city
    return {
      success: true,
      zip: data.places[0]['post code'],
      cityName: data['place name'] || cityName,
      allZips: data.places.map(function(p) { return p['post code']; })
    };
  } catch(e) {
    return { success: false, error: 'City lookup failed: ' + e.message };
  }
}

function setupAdmin() {
  const adminEmail = 'mharper@northtexas.ag';
  PropertiesService.getScriptProperties().setProperty('ADMIN_EMAIL', adminEmail);
  Logger.log('Admin email set to: ' + adminEmail);
}

function setupPlatformOwner() {
  // Set Steve as PlatformOwner and configure notification email
  PropertiesService.getScriptProperties().setProperty('PLATFORM_OWNER_EMAIL', 'steve@citybioclean.com');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var users = ss.getSheetByName('Users');
  var data = users.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === '1234') { // Steve's PIN
      users.getRange(i + 1, 4).setValue('PlatformOwner');
      Logger.log('Steve Harper (row ' + (i + 1) + ') set to PlatformOwner');
      break;
    }
  }
  Logger.log('PLATFORM_OWNER_EMAIL set to steve@citybioclean.com');
}

// ── WEB APP ──

function doGet(e) {
  var page = (e && e.parameter && e.parameter.p) ? e.parameter.p : 'app';
  var file = 'App';
  var title = 'Church Health Snapshot';

  if (page === 'onboard') {
    file = 'Onboard';
    title = 'Join Church Health Snapshot';
  }

  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── AUTH ──

function verifyPIN(pin) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const users = ss.getSheetByName('Users');
    const data = users.getDataRange().getValues();
    var headers = data[0];
    var orgIdCol = headers.indexOf('OrgID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(pin) && data[i][7] === 'Y') {
        users.getRange(i + 1, 7).setValue(new Date());

        var isFirstLogin = (data[i][8] === 'Y');
        if (isFirstLogin) {
          users.getRange(i + 1, 9).setValue('N');
        }

        var orgId = orgIdCol >= 0 ? String(data[i][orgIdCol]) : '';
        var org = orgId ? getOrgById_(orgId) : null;

        logAudit(data[i][1], 'LOGIN', 'PIN login successful' + (isFirstLogin ? ' (first time)' : '') + (orgId ? ' [' + orgId + ']' : ''));

        return {
          success: true,
          name: data[i][1],
          email: data[i][2],
          role: data[i][3],
          church: data[i][4],
          firstLogin: isFirstLogin,
          orgId: orgId,
          orgName: org ? org.orgName : '',
          denomName: org ? org.denominationName : 'Assemblies of God',
          userRow: i + 1
        };
      }
    }

    logAudit('Unknown', 'LOGIN_FAILED', 'PIN: ' + pin);
    return { success: false, error: 'Invalid PIN' };
  } catch(e) {
    return { success: false, error: 'Authentication error' };
  }
}

function verifyEmail(email, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const users = ss.getSheetByName('Users');
    const data = users.getDataRange().getValues();
    var headers = data[0];
    var orgIdCol = headers.indexOf('OrgID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === String(email).toLowerCase() && data[i][7] === 'Y') {
        users.getRange(i + 1, 7).setValue(new Date());

        var isFirstLogin = (data[i][8] === 'Y');
        if (isFirstLogin) {
          users.getRange(i + 1, 9).setValue('N');
        }

        var orgId = orgIdCol >= 0 ? String(data[i][orgIdCol]) : '';
        var org = orgId ? getOrgById_(orgId) : null;

        logAudit(data[i][1], 'LOGIN', 'Email login: ' + email + (isFirstLogin ? ' (first time)' : '') + (orgId ? ' [' + orgId + ']' : ''));

        return {
          success: true,
          name: data[i][1],
          email: data[i][2],
          role: data[i][3],
          church: data[i][4],
          firstLogin: isFirstLogin,
          orgId: orgId,
          orgName: org ? org.orgName : '',
          denomName: org ? org.denominationName : 'Assemblies of God',
          userRow: i + 1
        };
      }
    }

    logAudit('Unknown', 'LOGIN_FAILED', 'Email: ' + email);
    return { success: false, error: 'Invalid credentials' };
  } catch(e) {
    return { success: false, error: 'Authentication error' };
  }
}

// ── SCORING ENGINE (server-side only — never exposed to client) ──

function calculateScore(answers) {
  const v = key => answers[key] || '';
  const n = key => { const x = parseFloat(answers[key]); return isNaN(x) ? 0 : x; };

  let s1 = 0, s2 = 0, s3 = 0;
  const pastorAge = n('q3'), congAge = n('q4'), attendance = n('q5');
  const population = n('q14'), cityType = v('q13');
  const conversions = n('q9'), income = n('q17'), totalIncome = n('q8b');
  const medianResAge = n('q16');

  // ── Section 1: Church Statistical ──
  if (v('q1') === 'Yes') s1 += 4;
  if (v('q2') === 'Yes') s1 += 4;
  if (pastorAge > 0) { [60,65,70,75].forEach(t => { if (pastorAge < t) s1++; }); }
  if (congAge > 0) { [60,65,70,75].forEach(t => { if (congAge < t) s1++; }); }

  let cityMult = 0;
  if (cityType === 'Rural (Under 8000)') cityMult = 0.1;
  else if (cityType === 'Small city (8k-50k)') cityMult = 1;
  else if (cityType === 'Large City (over 50k)') cityMult = 0.1;
  else if (cityType === 'Urban/Metro (100k+)') cityMult = 0.1;

  if (cityMult > 0 && population > 0 && attendance > 0) {
    const r = attendance / (population * cityMult);
    [0.01, 0.05, 0.10, 0.15, 0.25].forEach(t => { if (r >= t) s1++; });
  }

  if (v('q6') === 'Static') s1 += 3; else if (v('q6') === 'Increasing') s1 += 5;
  if (v('q7') === 'General Council' || v('q7') === 'Parent Affiliated') s1 += 1;
  if (v('q8') === 'Static') s1 += 3; else if (v('q8') === 'Increasing') s1 += 5;

  if (attendance > 0 && conversions > 0) {
    [0.01, 0.03, 0.05, 0.08, 0.10].forEach(t => { if (conversions >= t * attendance) s1++; });
  }
  const baptisms = n('q10');
  if (conversions > 0 && baptisms > 0) {
    [0.05, 0.10, 0.25, 0.50, 0.75].forEach(t => { if (baptisms >= t * conversions) s1++; });
  }
  const spiritBap = n('q11');
  if (conversions > 0 && spiritBap > 0) {
    [0.10, 0.15, 0.20, 0.25, 0.30].forEach(t => { if (spiritBap >= t * conversions) s1++; });
  }
  const guests = n('q12');
  if (attendance > 0 && guests > 0) {
    [0.50, 0.75, 1.00, 1.50, 2.00].forEach(t => { if (guests >= t * attendance) s1++; });
  }

  // ── Section 2: Community Demographics ──
  if (v('q15') === 'Static') s2 += 1; else if (v('q15') === 'Increasing') s2 += 3;

  let congAgeDiff = 0;
  if (medianResAge > 0 && congAge > 0) congAgeDiff = congAge - medianResAge;

  // G42 + G43 both use congregation age gap (matches spreadsheet)
  if (medianResAge > 0 && congAge > 0) {
    // G42
    if (congAgeDiff - 10 <= 1) s2++;
    if (congAgeDiff - 15 <= 1) s2++;
    if (congAgeDiff - 20 <= 1) s2 += 2;
    // G43
    if (congAgeDiff - 10 <= 1) s2++;
    if (congAgeDiff - 15 <= 1) s2++;
    if (congAgeDiff - 20 <= 1) s2 += 2;
  }

  const potentialPC = income > 0 ? income * 0.1 : 0;
  // Row 46: points when potential giving is BELOW threshold
  if (income > 0) {
    [0, 250, 500, 750, 1000].forEach(t => { if (!(potentialPC > t)) s2++; });
  }

  // ── Section 3: Church Subjective ──
  if (v('q18') === 'No') s3 += 2;
  if (v('q19') === 'Static') s3 += 2; else if (v('q19') === 'Increasing') s3 += 3;
  if (v('q20') === 'Moderate') s3 += 1; else if (v('q20') === 'Good') s3 += 2; else if (v('q20') === 'Excellent') s3 += 3;
  if (v('q21') === 'Moderate') s3 += 1; else if (v('q21') === 'Good') s3 += 2; else if (v('q21') === 'Excellent') s3 += 3;
  if (v('q22') === 'No') s3 += 2;

  const prayerHrs = n('q23');
  let prayerScore = 0;
  if (attendance > 0 && prayerHrs > 0) {
    [0.025, 0.05, 0.065, 0.085].forEach(t => { if (prayerHrs > t * attendance) prayerScore++; });
    if (prayerHrs >= 0.095 * attendance) prayerScore++;
    if (prayerHrs > 0.15 * attendance) prayerScore--;
    if (prayerHrs > 0.175 * attendance) prayerScore -= 3;
    if (prayerHrs > 0.2 * attendance) prayerScore--;
  }
  s3 += Math.max(0, prayerScore);

  if (v('q24') === 'Yes') s3 += 1;
  if (v('q25') === 'Yes') s3 += 2;
  if (v('q26') === 'Yes') s3 += 2;
  if (v('q27') === 'No') s3 += 2;
  if (v('q28') === 'Current Decade') s3 += 3; else if (v('q28') === 'Last Decade') s3 += 1;

  const totalScore = s1 + s2 + s3;
  let rating;
  if (totalScore >= 76) rating = 'GREEN';
  else if (totalScore > 60) rating = 'YELLOW';
  else rating = 'RED';

  // Findings
  const perCapitaGiving = attendance > 0 ? totalIncome / attendance : 0;
  let pastorAgeDiff = 0;
  if (medianResAge > 0 && pastorAge > 0) pastorAgeDiff = pastorAge - medianResAge;

  // Category breakdown for radar chart
  let catLeadership = 0;
  if (v('q8') === 'Static') catLeadership += 3; else if (v('q8') === 'Increasing') catLeadership += 5;
  if (v('q19') === 'Static') catLeadership += 2; else if (v('q19') === 'Increasing') catLeadership += 3;
  if (v('q20') === 'Moderate') catLeadership += 1; else if (v('q20') === 'Good') catLeadership += 2; else if (v('q20') === 'Excellent') catLeadership += 3;
  if (v('q22') === 'No') catLeadership += 2;
  if (v('q24') === 'Yes') catLeadership += 1;
  if (v('q25') === 'Yes') catLeadership += 2;
  if (v('q26') === 'Yes') catLeadership += 2;
  if (v('q27') === 'No') catLeadership += 2;

  let catDiscipleship = 0;
  if (v('q1') === 'Yes') catDiscipleship += 4;
  if (v('q2') === 'Yes') catDiscipleship += 4;
  if (conversions > 0 && baptisms > 0) { [0.05,0.10,0.25,0.50,0.75].forEach(t => { if (baptisms >= t * conversions) catDiscipleship++; }); }
  if (v('q21') === 'Moderate') catDiscipleship += 1; else if (v('q21') === 'Good') catDiscipleship += 2; else if (v('q21') === 'Excellent') catDiscipleship += 3;

  let catOutreach = 0;
  if (attendance > 0 && conversions > 0) { [0.01,0.03,0.05,0.08,0.10].forEach(t => { if (conversions >= t * attendance) catOutreach++; }); }
  if (attendance > 0 && guests > 0) { [0.50,0.75,1.00,1.50,2.00].forEach(t => { if (guests >= t * attendance) catOutreach++; }); }
  if (pastorAge > 0) { [60,65,70,75].forEach(t => { if (pastorAge < t) catOutreach++; }); }

  let catSpiritual = 0;
  if (conversions > 0 && spiritBap > 0) { [0.10,0.15,0.20,0.25,0.30].forEach(t => { if (spiritBap >= t * conversions) catSpiritual++; }); }
  if (attendance > 0 && prayerHrs > 0) {
    [0.025,0.05,0.065,0.085].forEach(t => { if (prayerHrs > t * attendance) catSpiritual++; });
    if (prayerHrs >= 0.095 * attendance) catSpiritual++;
  }

  let catCommunity = 0;
  if (v('q15') === 'Static') catCommunity += 1; else if (v('q15') === 'Increasing') catCommunity += 3;
  if (v('q6') === 'Static') catCommunity += 3; else if (v('q6') === 'Increasing') catCommunity += 5;
  if (congAge > 0) { [60,65,70,75].forEach(t => { if (congAge < t) catCommunity++; }); }
  if (v('q18') === 'No') catCommunity += 2;
  if (v('q28') === 'Current Decade') catCommunity += 2; else if (v('q28') === 'Last Decade') catCommunity += 1;

  const categories = {
    'Leadership': { raw: catLeadership, max: 20, pct: Math.round((catLeadership / 20) * 100), color: '#2563eb' },
    'Discipleship': { raw: catDiscipleship, max: 16, pct: Math.round((catDiscipleship / 16) * 100), color: '#7c3aed' },
    'Outreach': { raw: catOutreach, max: 14, pct: Math.round((catOutreach / 14) * 100), color: '#059669' },
    'Spiritual': { raw: catSpiritual, max: 10, pct: Math.round((catSpiritual / 10) * 100), color: '#dc2626' },
    'Community': { raw: catCommunity, max: 16, pct: Math.round((catCommunity / 16) * 100), color: '#d97706' }
  };

  return {
    score: totalScore,
    rating: rating,
    s1: s1,
    s2: s2,
    s3: s3,
    categories: categories,
    findings: {
      perCapitaGiving: Math.round(perCapitaGiving * 100) / 100,
      potentialPerCapita: Math.round(potentialPC * 100) / 100,
      givingGap: Math.round((perCapitaGiving - potentialPC) * 100) / 100,
      pastorAgeDiff: Math.round(pastorAgeDiff * 100) / 100,
      congAgeDiff: Math.round(congAgeDiff * 100) / 100
    }
  };
}

// ── SUBMIT ASSESSMENT ──

function submitAssessment(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const assess = ss.getSheetByName('Assessments');

    const result = calculateScore(data.answers);

    const a = data.answers;
    const row = [
      new Date(),
      data.userName,
      data.userRole,
      data.churchName,
      data.pastorName,
      a.q1, a.q2, a.q3, a.q4, a.q5, a.q6, a.q7, a.q8,
      a.q8b, a.q9, a.q10, a.q11, a.q12,
      a.q13, a.q14, a.q15, a.q16, a.q17,
      a.q18, a.q19, a.q20, a.q21, a.q22, a.q23,
      a.q24, a.q25, a.q26, a.q27, a.q28,
      data.notes1 || '', data.notes2 || '', data.notes3 || '',
      result.score, result.rating, result.s1, result.s2, result.s3,
      data.orgId || '', data.churchId || '', data.userRole || ''
    ];

    assess.appendRow(row);

    logAudit(data.userName, 'SUBMIT',
      data.churchName + ' — Score: ' + result.score + ' (' + result.rating + ')' +
      (data.orgId ? ' [' + data.orgId + ']' : ''));

    // Send admin email to org's admin, not hardcoded
    sendAdminCopy(data, result);

    return {
      success: true,
      score: result.score,
      rating: result.rating,
      s1: result.s1,
      s2: result.s2,
      s3: result.s3,
      categories: result.categories,
      findings: result.findings
    };
  } catch(e) {
    logAudit(data.userName || 'Unknown', 'SUBMIT_ERROR', e.message);
    return { success: false, error: 'Failed to submit assessment' };
  }
}

// ── SILENT ADMIN EMAIL ──

function sendAdminCopy(data, result) {
  const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
  if (!adminEmail) return;

  try {
    const ratingColor = result.rating === 'GREEN' ? '#1a6b3c' : result.rating === 'YELLOW' ? '#a16207' : '#991b1b';
    const ratingWord = result.rating === 'GREEN' ? 'Healthy' : result.rating === 'YELLOW' ? 'Caution' : 'Critical';

    const a = data.answers;
    const body = `
      <div style="font-family:sans-serif;max-width:600px;margin:0 auto;">
        <div style="background:#1b2541;color:white;padding:20px;border-radius:8px 8px 0 0;">
          <h2 style="margin:0;">New Assessment Submitted</h2>
        </div>
        <div style="padding:20px;border:1px solid #e5e1d8;border-top:none;border-radius:0 0 8px 8px;">
          <p><strong>Church:</strong> ${data.churchName || 'Not specified'}</p>
          <p><strong>Pastor:</strong> ${data.pastorName || 'Not specified'}</p>
          <p><strong>Assessed By:</strong> ${data.userName} (${data.userRole})</p>
          <p><strong>Date:</strong> ${new Date().toLocaleDateString('en-US', {month:'long',day:'numeric',year:'numeric'})}</p>
          <hr style="border:none;border-top:1px solid #e5e1d8;margin:16px 0;">
          <div style="text-align:center;padding:20px;">
            <div style="font-size:48px;font-weight:bold;color:#1b2541;">${result.score}</div>
            <div style="color:#6b7280;">of 95</div>
            <div style="display:inline-block;padding:8px 24px;border-radius:20px;background:${ratingColor}20;color:${ratingColor};font-weight:bold;margin-top:8px;">${ratingWord}</div>
          </div>
          <p style="font-size:12px;color:#9ca3af;text-align:center;">
            Statistical: ${result.s1} | Demographics: ${result.s2} | Subjective: ${result.s3}
          </p>
          <hr style="border:none;border-top:1px solid #e5e1d8;margin:16px 0;">
          <p style="font-size:12px;color:#9ca3af;">This is an automated notification. The assessor was not notified of this copy. View all submissions in the <a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}">tracking spreadsheet</a>.</p>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: adminEmail,
      subject: 'Assessment: ' + (data.churchName || 'Unknown Church') + ' — ' + ratingWord + ' (' + result.score + '/95)',
      htmlBody: body,
      noReply: true
    });
  } catch(e) {
    // Don't fail the submission if email fails
    logAudit('SYSTEM', 'EMAIL_ERROR', 'Admin copy failed: ' + e.message);
  }
}

// ── ADMIN: GET ALL ASSESSMENTS ──

/**
 * Get assessments scoped by role:
 *  - PlatformOwner: all assessments across all orgs
 *  - Admin: all assessments for their org
 *  - SectionLeader: assessments for their assigned churches (theirs + pastor's)
 *  - Pastor: only their own assessments
 */
function getAssessments(params) {
  try {
    var userRole = params.userRole || params;
    var orgId = params.orgId || '';
    var userName = params.userName || '';
    var userRow = params.userRow || 0;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const assess = ss.getSheetByName('Assessments');
    const data = assess.getDataRange().getValues();

    if (data.length <= 1) return { success: true, assessments: [] };

    var headers = data[0];
    var orgIdCol = headers.indexOf('OrgID');
    var assessorRoleCol = headers.indexOf('AssessorRole');

    // Get section leader's churches if applicable
    var myChurches = [];
    if (userRole === 'SectionLeader' && userRow) {
      myChurches = getChurchesForSectionLeader_(userRow);
    }

    const assessments = [];
    for (let i = 1; i < data.length; i++) {
      var rowOrgId = orgIdCol >= 0 ? String(data[i][orgIdCol]) : '';
      var churchName = data[i][3];
      var assessedBy = data[i][1];

      // Filter by role
      if (userRole === 'PlatformOwner') {
        // See everything
      } else if (userRole === 'Admin') {
        if (orgId && rowOrgId !== orgId) continue;
      } else if (userRole === 'SectionLeader') {
        if (orgId && rowOrgId !== orgId) continue;
        // Only see assessments for churches they oversee
        var isMyChurch = myChurches.some(function(c) {
          return c.toLowerCase() === String(churchName).toLowerCase();
        });
        if (!isMyChurch) continue;
      } else if (userRole === 'Pastor') {
        // Only see their own assessments
        if (assessedBy !== userName) continue;
      } else {
        // Unknown role — show nothing
        continue;
      }

      assessments.push({
        timestamp: data[i][0],
        assessedBy: data[i][1],
        role: data[i][2],
        assessorRole: assessorRoleCol >= 0 ? data[i][assessorRoleCol] : data[i][2],
        churchName: churchName,
        pastorName: data[i][4],
        score: data[i][37],
        rating: data[i][38],
        s1: data[i][39],
        s2: data[i][40],
        s3: data[i][41],
        orgId: rowOrgId
      });
    }

    return { success: true, assessments: assessments };
  } catch(e) {
    return { success: false, error: 'Failed to load assessments: ' + e.message };
  }
}

/** Get church names assigned to a section leader */
function getChurchesForSectionLeader_(userRow) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var churches = ss.getSheetByName('Churches');
    if (!churches) return [];
    var data = churches.getDataRange().getValues();
    var names = [];
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][8]) === Number(userRow) && data[i][9] === 'Active') {
        names.push(String(data[i][2]));
      }
    }
    return names;
  } catch(e) {
    return [];
  }
}

// ── ADMIN: MANAGE USERS ──

function addUser(data) {
  try {
    var pin = data.pin, name = data.name, email = data.email, role = data.role, church = data.church;
    var orgId = data.orgId || '';
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var users = ss.getSheetByName('Users');

    var existing = users.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (String(existing[i][0]) === String(pin)) {
        return { success: false, error: 'PIN already in use' };
      }
    }

    users.appendRow([String(pin), name, email || '', role, church || '', new Date(), '', 'Y', 'Y', orgId]);
    logAudit('ADMIN', 'USER_ADDED', name + ' (' + role + ')' + (orgId ? ' [' + orgId + ']' : ''));

    if (email) {
      sendWelcomeEmail(name, email, pin, role, orgId);
    }

    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function sendWelcomeEmail(name, email, pin, role, orgId) {
  try {
    var appUrl = ScriptApp.getService().getUrl();
    var org = orgId ? getOrgById_(orgId) : null;
    var orgName = org ? org.orgName : 'Church Health Snapshot';
    var adminEmail = org ? (org.adminNotifyEmail || org.adminEmail) : (PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || '');

    var body = '<div style="font-family:\'Segoe UI\',sans-serif;max-width:560px;margin:0 auto;">' +
      '<div style="background:#1b2541;padding:28px 24px;border-radius:10px 10px 0 0;text-align:center;">' +
        '<h1 style="margin:0;color:white;font-size:22px;">Welcome to Church Health Snapshot</h1>' +
        '<p style="margin:6px 0 0;color:#c9a227;font-size:13px;">' + orgName + '</p>' +
      '</div>' +
      '<div style="padding:28px 24px;border:1px solid #e5e1d8;border-top:none;border-radius:0 0 10px 10px;background:#fdfcf8;">' +
        '<p style="font-size:15px;color:#1f2937;">Hi ' + name + ',</p>' +
        '<p style="font-size:14px;color:#4b5563;line-height:1.6;">You\'ve been registered for the Church Health Snapshot assessment tool. Here\'s everything you need to get started:</p>' +
        '<div style="background:white;border:1px solid #e5e1d8;border-radius:8px;padding:20px;margin:20px 0;text-align:center;">' +
          '<div style="font-size:12px;text-transform:uppercase;letter-spacing:1px;color:#9ca3af;margin-bottom:6px;">Your PIN</div>' +
          '<div style="font-size:36px;font-weight:700;color:#1b2541;letter-spacing:4px;">' + pin + '</div>' +
        '</div>' +
        '<div style="background:white;border:1px solid #e5e1d8;border-radius:8px;padding:16px;margin:20px 0;">' +
          '<div style="font-size:12px;text-transform:uppercase;letter-spacing:1px;color:#9ca3af;margin-bottom:8px;">How to Log In</div>' +
          '<ol style="margin:0;padding-left:20px;font-size:14px;color:#4b5563;line-height:1.8;">' +
            '<li>Open the link below</li>' +
            '<li>Enter your PIN</li>' +
            '<li>Complete the assessment for your church</li>' +
          '</ol>' +
        '</div>' +
        '<div style="text-align:center;margin:24px 0;">' +
          '<a href="' + appUrl + '" style="display:inline-block;background:#1b2541;color:white;padding:12px 32px;border-radius:6px;text-decoration:none;font-weight:600;font-size:14px;">Open Assessment Tool</a>' +
        '</div>' +
        '<p style="font-size:12px;color:#9ca3af;text-align:center;margin-top:24px;">Questions? Contact your district administrator' + (adminEmail ? ' at ' + adminEmail : '') + '.</p>' +
      '</div>' +
    '</div>';

    MailApp.sendEmail({
      to: email,
      subject: 'Welcome to Church Health Snapshot — Your Login Details',
      htmlBody: body,
      name: 'Church Health Snapshot',
      replyTo: adminEmail || ''
    });

    logAudit('SYSTEM', 'WELCOME_EMAIL', 'Sent to ' + email + ' for ' + name);
  } catch(e) {
    logAudit('SYSTEM', 'WELCOME_EMAIL_ERROR', 'Failed for ' + email + ': ' + e.message);
  }
}

function getUsers(params) {
  try {
    var orgId = (params && params.orgId) ? params.orgId : '';
    var callerRole = (params && params.userRole) ? params.userRole : '';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const users = ss.getSheetByName('Users');
    const data = users.getDataRange().getValues();
    var headers = data[0];
    var orgIdCol = headers.indexOf('OrgID');

    const userList = [];
    for (let i = 1; i < data.length; i++) {
      var rowOrgId = orgIdCol >= 0 ? String(data[i][orgIdCol]) : '';

      // Org scoping: Admin sees only their org, PlatformOwner sees all
      if (callerRole !== 'PlatformOwner' && orgId && rowOrgId !== orgId) continue;

      userList.push({
        row: i + 1,
        pin: String(data[i][0]),
        name: data[i][1],
        email: data[i][2],
        role: data[i][3],
        church: data[i][4],
        registered: data[i][5],
        lastLogin: data[i][6],
        active: data[i][7],
        orgId: rowOrgId
      });
    }

    return { success: true, users: userList };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function updateUser(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var users = ss.getSheetByName('Users');
    var all = users.getDataRange().getValues();
    var row = data.row;

    if (row < 2 || row > all.length) return { success: false, error: 'Invalid user row' };

    // Check for duplicate PIN (excluding this user's row)
    if (data.pin) {
      for (var i = 1; i < all.length; i++) {
        if ((i + 1) !== row && String(all[i][0]) === String(data.pin)) {
          return { success: false, error: 'PIN already in use by ' + all[i][1] };
        }
      }
      users.getRange(row, 1).setValue(String(data.pin));
    }
    if (data.name) users.getRange(row, 2).setValue(data.name);
    if (data.email !== undefined) users.getRange(row, 3).setValue(data.email);
    if (data.role) users.getRange(row, 4).setValue(data.role);
    if (data.church !== undefined) users.getRange(row, 5).setValue(data.church);

    logAudit('ADMIN', 'USER_UPDATED', data.name + ' (row ' + row + ')');
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function toggleUserActive(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var users = ss.getSheetByName('Users');
    var all = users.getDataRange().getValues();
    var row = data.row;

    if (row < 2 || row > all.length) return { success: false, error: 'Invalid user row' };

    var currentStatus = all[row - 1][7];
    var newStatus = (currentStatus === 'Y') ? 'N' : 'Y';
    users.getRange(row, 8).setValue(newStatus);

    var userName = all[row - 1][1];
    logAudit('ADMIN', newStatus === 'Y' ? 'USER_ENABLED' : 'USER_DISABLED', userName);
    return { success: true, active: newStatus, name: userName };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── CHURCH MANAGEMENT ──

function addChurch(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var churches = ss.getSheetByName('Churches');
    if (!churches) return { success: false, error: 'Churches sheet not found. Run setupMultiTenant().' };

    var existing = churches.getDataRange().getValues();
    var nextId = 'CH' + String(existing.length).padStart(4, '0');

    churches.appendRow([
      nextId,
      data.orgId || '',
      data.churchName,
      data.city || '',
      data.state || '',
      data.zip || '',
      data.pastorName || '',
      data.pastorEmail || '',
      data.sectionLeaderRow || '',
      'Active',
      new Date()
    ]);

    logAudit(data.addedBy || 'SYSTEM', 'CHURCH_ADDED',
      data.churchName + (data.orgId ? ' [' + data.orgId + ']' : ''));
    return { success: true, churchId: nextId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function addChurchesBulk(data) {
  try {
    var results = [];
    var churches = data.churches || [];
    for (var i = 0; i < churches.length; i++) {
      churches[i].orgId = data.orgId;
      churches[i].sectionLeaderRow = data.sectionLeaderRow;
      churches[i].addedBy = data.addedBy;
      results.push(addChurch(churches[i]));
    }
    var added = results.filter(function(r) { return r.success; }).length;
    return { success: true, added: added, total: churches.length };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getChurches(params) {
  try {
    var orgId = params.orgId || '';
    var sectionLeaderRow = params.sectionLeaderRow || '';
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var churches = ss.getSheetByName('Churches');
    if (!churches) return { success: true, churches: [] };

    var data = churches.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      if (orgId && String(data[i][1]) !== orgId) continue;
      if (sectionLeaderRow && String(data[i][8]) !== String(sectionLeaderRow)) continue;

      list.push({
        churchId: data[i][0],
        orgId: data[i][1],
        churchName: data[i][2],
        city: data[i][3],
        state: data[i][4],
        zip: data[i][5],
        pastorName: data[i][6],
        pastorEmail: data[i][7],
        sectionLeaderRow: data[i][8],
        status: data[i][9],
        dateAdded: data[i][10]
      });
    }
    return { success: true, churches: list };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function updateChurchStatus(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var churches = ss.getSheetByName('Churches');
    var all = churches.getDataRange().getValues();
    for (var i = 1; i < all.length; i++) {
      if (all[i][0] === data.churchId) {
        churches.getRange(i + 1, 10).setValue(data.status); // Active or Inactive
        logAudit(data.updatedBy || 'SYSTEM', 'CHURCH_STATUS',
          all[i][2] + ' → ' + data.status);
        return { success: true };
      }
    }
    return { success: false, error: 'Church not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/** Fuzzy search churches by name for the assessment flow */
function searchChurches(params) {
  try {
    var query = String(params.query || '').toLowerCase().trim();
    var orgId = params.orgId || '';
    if (query.length < 2) return { success: true, matches: [] };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var churches = ss.getSheetByName('Churches');
    if (!churches) return { success: true, matches: [] };

    var data = churches.getDataRange().getValues();
    var matches = [];
    for (var i = 1; i < data.length; i++) {
      if (orgId && String(data[i][1]) !== orgId) continue;
      if (data[i][9] !== 'Active') continue;

      var name = String(data[i][2]).toLowerCase();
      // Simple fuzzy: check if query words appear in the church name
      var queryWords = query.split(/\s+/);
      var matchCount = 0;
      for (var w = 0; w < queryWords.length; w++) {
        if (name.indexOf(queryWords[w]) >= 0) matchCount++;
      }
      if (matchCount > 0) {
        matches.push({
          churchId: data[i][0],
          churchName: data[i][2],
          city: data[i][3],
          state: data[i][4],
          pastorName: data[i][6],
          score: matchCount / queryWords.length // relevance
        });
      }
    }

    // Sort by relevance
    matches.sort(function(a, b) { return b.score - a.score; });
    return { success: true, matches: matches.slice(0, 10) };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── DASHBOARD DATA ──

/** Section leader dashboard: their churches with latest scores and trends */
function getSectionDashboard(params) {
  try {
    var orgId = params.orgId;
    var userRow = params.userRow;

    // Get churches assigned to this section leader
    var churchResult = getChurches({ orgId: orgId, sectionLeaderRow: userRow });
    if (!churchResult.success) return churchResult;

    // Get all assessments for this org
    var assessResult = getAssessments({
      userRole: 'SectionLeader',
      orgId: orgId,
      userRow: userRow
    });
    if (!assessResult.success) return assessResult;

    // Build church health map
    var churchHealth = [];
    for (var c = 0; c < churchResult.churches.length; c++) {
      var church = churchResult.churches[c];
      // Find all assessments for this church, sorted by date
      var churchAssessments = assessResult.assessments.filter(function(a) {
        return a.churchName && a.churchName.toLowerCase() === church.churchName.toLowerCase();
      }).sort(function(a, b) {
        return new Date(b.timestamp) - new Date(a.timestamp);
      });

      var latest = churchAssessments.length > 0 ? churchAssessments[0] : null;
      var previous = churchAssessments.length > 1 ? churchAssessments[1] : null;
      var trend = 'none';
      if (latest && previous) {
        if (latest.score > previous.score) trend = 'improving';
        else if (latest.score < previous.score) trend = 'declining';
        else trend = 'stable';
      }

      churchHealth.push({
        church: church,
        assessmentCount: churchAssessments.length,
        latest: latest ? {
          score: latest.score,
          rating: latest.rating,
          date: latest.timestamp,
          assessedBy: latest.assessedBy,
          assessorRole: latest.assessorRole
        } : null,
        trend: trend,
        history: churchAssessments.map(function(a) {
          return { score: a.score, rating: a.rating, date: a.timestamp, assessedBy: a.assessedBy, assessorRole: a.assessorRole };
        })
      });
    }

    // Summary counts
    var summary = { green: 0, yellow: 0, red: 0, unassessed: 0, total: churchHealth.length };
    for (var h = 0; h < churchHealth.length; h++) {
      if (!churchHealth[h].latest) summary.unassessed++;
      else if (churchHealth[h].latest.rating === 'GREEN') summary.green++;
      else if (churchHealth[h].latest.rating === 'YELLOW') summary.yellow++;
      else summary.red++;
    }

    return { success: true, churches: churchHealth, summary: summary };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/** District admin dashboard: all sections, all churches, district-wide view */
function getDistrictDashboard(params) {
  try {
    var orgId = params.orgId;

    // Get all churches for this org
    var churchResult = getChurches({ orgId: orgId });
    if (!churchResult.success) return churchResult;

    // Get all assessments for this org
    var assessResult = getAssessments({ userRole: 'Admin', orgId: orgId });
    if (!assessResult.success) return assessResult;

    // Get section leader names from Users sheet
    var sectionLeaders = {};
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var users = ss.getSheetByName('Users');
    var userData = users.getDataRange().getValues();
    var headers = userData[0];
    var orgIdCol = headers.indexOf('OrgID');
    for (var u = 1; u < userData.length; u++) {
      var uOrgId = orgIdCol >= 0 ? String(userData[u][orgIdCol]) : '';
      if (orgId && uOrgId !== orgId) continue;
      if (userData[u][3] === 'SectionLeader' && userData[u][7] === 'Y') {
        sectionLeaders[u + 1] = { name: userData[u][1], row: u + 1 };
      }
    }

    // Build health summary per church
    var churchMap = {};
    for (var c = 0; c < churchResult.churches.length; c++) {
      var ch = churchResult.churches[c];
      churchMap[ch.churchName.toLowerCase()] = {
        church: ch,
        assessments: []
      };
    }

    for (var a = 0; a < assessResult.assessments.length; a++) {
      var assessment = assessResult.assessments[a];
      var key = (assessment.churchName || '').toLowerCase();
      if (churchMap[key]) {
        churchMap[key].assessments.push(assessment);
      }
    }

    var summary = { green: 0, yellow: 0, red: 0, unassessed: 0, total: 0, totalAssessments: assessResult.assessments.length };
    var churches = [];
    var keys = Object.keys(churchMap);
    for (var k = 0; k < keys.length; k++) {
      var entry = churchMap[keys[k]];
      entry.assessments.sort(function(a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
      var latest = entry.assessments.length > 0 ? entry.assessments[0] : null;

      summary.total++;
      if (!latest) summary.unassessed++;
      else if (latest.rating === 'GREEN') summary.green++;
      else if (latest.rating === 'YELLOW') summary.yellow++;
      else summary.red++;

      // Include section leader name
      var slRow = entry.church.sectionLeaderRow;
      var slName = (slRow && sectionLeaders[slRow]) ? sectionLeaders[slRow].name : '';

      churches.push({
        church: entry.church,
        assessmentCount: entry.assessments.length,
        latest: latest ? { score: latest.score, rating: latest.rating, date: latest.timestamp } : null,
        sectionLeaderName: slName,
        history: entry.assessments.map(function(a) {
          return { score: a.score, rating: a.rating, date: a.timestamp, assessedBy: a.assessedBy, assessorRole: a.assessorRole };
        })
      });
    }

    // Build section leader groups
    var sections = [];
    var slKeys = Object.keys(sectionLeaders);
    for (var s = 0; s < slKeys.length; s++) {
      var sl = sectionLeaders[slKeys[s]];
      var slChurches = churches.filter(function(c) { return String(c.church.sectionLeaderRow) === String(sl.row); });
      var slSummary = { green: 0, yellow: 0, red: 0, unassessed: 0, total: slChurches.length };
      slChurches.forEach(function(c) {
        if (!c.latest) slSummary.unassessed++;
        else if (c.latest.rating === 'GREEN') slSummary.green++;
        else if (c.latest.rating === 'YELLOW') slSummary.yellow++;
        else slSummary.red++;
      });
      sections.push({ name: sl.name, row: sl.row, churches: slChurches, summary: slSummary });
    }

    // Churches not assigned to any section leader
    var unassigned = churches.filter(function(c) { return !c.sectionLeaderName; });

    return { success: true, churches: churches, summary: summary, sections: sections, unassigned: unassigned };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── ORG MANAGEMENT (Platform Owner only) ──

function getOrganizations() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var orgSheet = ss.getSheetByName('Organizations');
    if (!orgSheet) return { success: true, organizations: [] };

    var data = orgSheet.getDataRange().getValues();
    var orgs = [];
    for (var i = 1; i < data.length; i++) {
      orgs.push({
        row: i + 1,
        orgId: data[i][0],
        orgName: data[i][1],
        state: data[i][2],
        denominationCode: data[i][3],
        denominationName: data[i][4],
        adminName: data[i][5],
        adminEmail: data[i][6],
        status: data[i][8],
        dateCreated: data[i][9]
      });
    }
    return { success: true, organizations: orgs };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function approveOrganization(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var orgSheet = ss.getSheetByName('Organizations');
    var all = orgSheet.getDataRange().getValues();

    for (var i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(data.orgId)) {
        orgSheet.getRange(i + 1, 9).setValue('Active'); // Status column
        logAudit('PLATFORM', 'ORG_APPROVED', all[i][1] + ' (' + data.orgId + ')');

        // Create admin user for the org if PIN provided
        if (data.adminPin) {
          addUser({
            pin: data.adminPin,
            name: all[i][5],
            email: all[i][6],
            role: 'Admin',
            church: '',
            orgId: data.orgId
          });
        }
        return { success: true };
      }
    }
    return { success: false, error: 'Organization not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── ONBOARDING ──

/** Get list of counties for a given state (for intake form county picker) */
function getCountiesByState(stateName) {
  try {
    var refSS = SpreadsheetApp.openById(REFERENCE_SPREADSHEET_ID);
    var rdSheet = refSS.getSheetByName('ReligiousData');
    var data = rdSheet.getDataRange().getValues();
    var counties = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === String(stateName).toLowerCase()) {
        counties.push({
          fips: String(data[i][0]),
          county: String(data[i][2]).replace(' County', '').replace(' Parish', '')
        });
      }
    }

    counties.sort(function(a, b) { return a.county.localeCompare(b.county); });
    return { success: true, counties: counties, state: stateName };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/** Get denomination list for the intake dropdown */
function getDenominations() {
  try {
    var refSS = SpreadsheetApp.openById(REFERENCE_SPREADSHEET_ID);
    var denomSheet = refSS.getSheetByName('DenominationLookup');
    var data = denomSheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      list.push({ code: String(data[i][0]), name: data[i][1], shortName: data[i][2] });
    }
    return { success: true, denominations: list };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/** Submit onboarding request — creates pending org, notifies Steve */
function submitOnboardRequest(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var orgSheet = ss.getSheetByName('Organizations');
    if (!orgSheet) return { success: false, error: 'Organizations sheet not found.' };

    // Generate org ID from name (first letters, uppercase, max 6 chars)
    var orgId = String(data.orgName || '').replace(/[^a-zA-Z\s]/g, '').split(/\s+/)
      .map(function(w) { return w.charAt(0).toUpperCase(); }).join('').substring(0, 6);
    // Check for duplicate and append number if needed
    var existing = orgSheet.getDataRange().getValues();
    var baseId = orgId;
    var counter = 1;
    while (existing.some(function(row) { return String(row[0]) === orgId; })) {
      orgId = baseId + counter;
      counter++;
    }

    var counties = (data.counties || []).join(',');

    orgSheet.appendRow([
      orgId,
      data.orgName,
      data.state,
      data.denomCode,
      data.denomName,
      data.adminName,
      data.adminEmail,
      data.adminEmail,
      'Pending',
      new Date(),
      counties
    ]);

    logAudit('ONBOARD', 'ORG_REQUEST', data.orgName + ' (' + orgId + ') — ' + data.adminName + ' <' + data.adminEmail + '>');

    // Notify Steve
    try {
      var ownerEmail = PropertiesService.getScriptProperties().getProperty('PLATFORM_OWNER_EMAIL') || 'steve@citybioclean.com';
      MailApp.sendEmail({
        to: ownerEmail,
        subject: 'New District Onboard Request: ' + data.orgName,
        htmlBody: '<div style="font-family:sans-serif;max-width:500px;">' +
          '<h2 style="color:#1b2541;">New District Request</h2>' +
          '<p><strong>District:</strong> ' + data.orgName + '</p>' +
          '<p><strong>State:</strong> ' + data.state + '</p>' +
          '<p><strong>Denomination:</strong> ' + data.denomName + '</p>' +
          '<p><strong>Admin:</strong> ' + data.adminName + ' &lt;' + data.adminEmail + '&gt;</p>' +
          '<p><strong>Counties:</strong> ' + (data.counties || []).length + ' selected</p>' +
          '<p><strong>Org ID:</strong> ' + orgId + '</p>' +
          '<p style="margin-top:20px;">Log in to the platform to approve this request.</p>' +
          '</div>',
        noReply: true
      });
    } catch(emailErr) {
      logAudit('SYSTEM', 'ONBOARD_EMAIL_ERROR', emailErr.message);
    }

    return { success: true, orgId: orgId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── RELIGIOUS DATA (dynamic from Reference Spreadsheet) ──

// Legacy hardcoded data kept as fallback — will be removed after multi-tenant is verified
var NT_ZIP_TO_COUNTY_LEGACY = {"75001":"48113","75002":"48085","75006":"48113","75007":"48085","75009":"48085","75010":"48121","75013":"48085","75019":"48113","75020":"48181","75021":"48181","75022":"48121","75023":"48085","75024":"48085","75025":"48085","75028":"48121","75032":"48397","75033":"48085","75034":"48085","75035":"48085","75036":"48121","75038":"48113","75039":"48113","75040":"48113","75041":"48113","75042":"48113","75043":"48113","75044":"48085","75048":"48085","75050":"48113","75051":"48113","75052":"48113","75054":"48113","75056":"48121","75057":"48121","75058":"48181","75060":"48113","75061":"48113","75062":"48113","75063":"48113","75065":"48121","75067":"48113","75068":"48121","75069":"48085","75070":"48085","75071":"48085","75072":"48085","75074":"48085","75075":"48085","75076":"48181","75077":"48121","75078":"48085","75080":"48085","75081":"48113","75082":"48085","75087":"48085","75088":"48113","75089":"48113","75090":"48181","75092":"48181","75093":"48085","75094":"48085","75098":"48085","75101":"48139","75102":"48349","75104":"48113","75105":"48349","75109":"48349","75110":"48349","75114":"48257","75115":"48113","75116":"48113","75119":"48139","75124":"48213","75125":"48113","75126":"48257","75132":"48397","75134":"48113","75135":"48231","75137":"48113","75141":"48113","75142":"48257","75143":"48213","75144":"48349","75146":"48113","75147":"48213","75148":"48213","75149":"48113","75150":"48113","75152":"48139","75153":"48349","75154":"48113","75155":"48349","75156":"48213","75157":"48257","75158":"48257","75159":"48113","75160":"48231","75161":"48257","75163":"48213","75164":"48085","75165":"48139","75166":"48085","75167":"48139","75169":"48231","75172":"48113","75173":"48085","75180":"48113","75181":"48113","75182":"48113","75189":"48085","75201":"48113","75202":"48113","75203":"48113","75204":"48113","75205":"48113","75206":"48113","75207":"48113","75208":"48113","75209":"48113","75210":"48113","75211":"48113","75212":"48113","75214":"48113","75215":"48113","75216":"48113","75217":"48113","75218":"48113","75219":"48113","75220":"48113","75223":"48113","75224":"48113","75225":"48113","75226":"48113","75227":"48113","75228":"48113","75229":"48113","75230":"48113","75231":"48113","75232":"48113","75233":"48113","75234":"48113","75235":"48113","75236":"48113","75237":"48113","75238":"48113","75240":"48113","75241":"48113","75243":"48113","75244":"48113","75246":"48113","75247":"48113","75248":"48085","75249":"48113","75251":"48113","75252":"48085","75253":"48113","75254":"48113","75261":"48439","75270":"48113","75287":"48085","75390":"48113","75401":"48231","75402":"48231","75407":"48085","75409":"48085","75413":"48147","75414":"48181","75418":"48147","75422":"48231","75423":"48147","75424":"48085","75428":"48231","75429":"48231","75433":"48231","75438":"48147","75439":"48147","75442":"48085","75446":"48147","75447":"48147","75449":"48147","75452":"48085","75453":"48231","75454":"48085","75459":"48181","75469":"48147","75474":"48231","75475":"48147","75476":"48147","75479":"48147","75488":"48147","75489":"48181","75490":"48147","75491":"48085","75492":"48147","75495":"48085","75496":"48147","75751":"48213","75752":"48213","75756":"48213","75758":"48213","75763":"48213","75770":"48213","75778":"48213","75782":"48213","75803":"48213","75853":"48213","75859":"48349","76001":"48439","76002":"48439","76005":"48439","76006":"48439","76008":"48367","76009":"48251","76010":"48439","76011":"48439","76012":"48439","76013":"48439","76014":"48439","76015":"48439","76016":"48439","76017":"48439","76018":"48439","76020":"48367","76021":"48439","76022":"48439","76023":"48367","76028":"48251","76031":"48251","76033":"48221","76034":"48439","76035":"48221","76036":"48251","76039":"48439","76040":"48439","76041":"48139","76043":"48425","76044":"48251","76048":"48221","76049":"48221","76050":"48139","76051":"48113","76052":"48121","76053":"48439","76054":"48439","76058":"48251","76059":"48251","76060":"48439","76061":"48251","76063":"48139","76064":"48139","76065":"48113","76066":"48363","76067":"48363","76070":"48251","76071":"48439","76073":"48497","76077":"48425","76078":"48121","76082":"48367","76084":"48139","76085":"48367","76086":"48367","76087":"48221","76088":"48367","76092":"48121","76093":"48251","76102":"48439","76103":"48439","76104":"48439","76105":"48439","76106":"48439","76107":"48439","76108":"48367","76109":"48439","76110":"48439","76111":"48439","76112":"48439","76114":"48439","76115":"48439","76116":"48439","76117":"48439","76118":"48439","76119":"48439","76120":"48439","76123":"48439","76126":"48367","76127":"48439","76129":"48439","76131":"48439","76132":"48439","76133":"48439","76134":"48439","76135":"48439","76137":"48439","76140":"48439","76148":"48439","76155":"48439","76164":"48439","76177":"48121","76179":"48439","76180":"48439","76182":"48439","76201":"48121","76203":"48121","76205":"48121","76207":"48121","76208":"48121","76209":"48121","76210":"48121","76225":"48337","76226":"48121","76227":"48121","76228":"48077","76230":"48077","76233":"48097","76234":"48097","76238":"48097","76239":"48097","76240":"48097","76241":"48097","76244":"48439","76245":"48181","76247":"48121","76248":"48439","76249":"48121","76250":"48097","76251":"48337","76252":"48097","76253":"48097","76255":"48337","76258":"48097","76259":"48121","76261":"48077","76262":"48121","76263":"48097","76264":"48181","76265":"48097","76266":"48097","76267":"48497","76268":"48181","76270":"48337","76271":"48097","76272":"48097","76273":"48097","76301":"48485","76302":"48485","76305":"48077","76306":"48485","76308":"48485","76309":"48485","76310":"48077","76311":"48485","76354":"48485","76357":"48077","76360":"48485","76365":"48077","76367":"48485","76377":"48077","76389":"48077","76401":"48143","76402":"48143","76426":"48237","76427":"48237","76429":"48363","76431":"48237","76433":"48143","76436":"48143","76439":"48367","76444":"48143","76445":"48143","76446":"48143","76449":"48237","76450":"48363","76453":"48143","76457":"48143","76458":"48237","76459":"48237","76462":"48143","76463":"48143","76472":"48363","76475":"48363","76476":"48221","76484":"48363","76486":"48237","76487":"48237","76490":"48367","76623":"48139","76626":"48139","76639":"48349","76641":"48139","76648":"48349","76649":"48143","76651":"48139","76666":"48349","76670":"48139","76679":"48349","76681":"48349","76690":"48143","76693":"48349"};

var NT_RELIGIOUS_DATA_LEGACY = {
  "48113": {county:"Dallas",pop:2604053,totalAdherents:1541280,agCongs:114,agAdherents:56159,evangelical:573341,pentecostal:89478,catholic:431645,mainline:214466,blackProt:152732,top3:"Catholic Church (431,645); Non-denominational (228,162); Southern Baptist (210,795)"},
  "48439": {county:"Tarrant",pop:2113854,totalAdherents:1280384,agCongs:58,agAdherents:17420,evangelical:644592,pentecostal:38812,catholic:359705,mainline:115271,blackProt:66974,top3:"Catholic Church (359,705); Southern Baptist (278,899); Non-denominational (262,336)"},
  "48085": {county:"Collin",pop:1079153,totalAdherents:502689,agCongs:27,agAdherents:4432,evangelical:198484,pentecostal:5408,catholic:140562,mainline:60497,blackProt:6880,top3:"Catholic Church (140,562); Southern Baptist (87,523); Non-denominational (87,248)"},
  "48121": {county:"Denton",pop:914870,totalAdherents:368975,agCongs:39,agAdherents:12775,evangelical:209835,pentecostal:14002,catholic:92887,mainline:24911,blackProt:3734,top3:"Southern Baptist (122,611); Catholic Church (92,887); Non-denominational (57,330)"},
  "48139": {county:"Ellis",pop:195509,totalAdherents:97589,agCongs:17,agAdherents:3031,evangelical:66918,pentecostal:4205,catholic:15145,mainline:6643,blackProt:6204,top3:"Southern Baptist (38,101); Non-denominational (20,104); Catholic Church (15,145)"},
  "48251": {county:"Johnson",pop:182690,totalAdherents:65690,agCongs:17,agAdherents:4243,evangelical:52409,pentecostal:4541,catholic:3286,mainline:5666,blackProt:1060,top3:"Southern Baptist (27,276); Non-denominational (8,078); Seventh-day Adventist (4,963)"},
  "48257": {county:"Kaufman",pop:149773,totalAdherents:59414,agCongs:10,agAdherents:719,evangelical:37678,pentecostal:1319,catholic:9884,mainline:3057,blackProt:5198,top3:"Southern Baptist (27,450); Catholic Church (9,884); Non-denominational (6,396)"},
  "48367": {county:"Parker",pop:151188,totalAdherents:65923,agCongs:10,agAdherents:1334,evangelical:44106,pentecostal:2392,catholic:10528,mainline:7173,blackProt:510,top3:"Southern Baptist (28,354); Catholic Church (10,528); Non-denominational (10,420)"},
  "48397": {county:"Rockwall",pop:110631,totalAdherents:79946,agCongs:6,agAdherents:3404,evangelical:57438,pentecostal:3404,catholic:13495,mainline:5633,blackProt:540,top3:"Southern Baptist (48,858); Catholic Church (13,495); United Methodist (4,773)"},
  "48231": {county:"Hunt",pop:101596,totalAdherents:42590,agCongs:7,agAdherents:766,evangelical:31015,pentecostal:1356,catholic:2650,mainline:2849,blackProt:1712,top3:"Southern Baptist (21,129); Non-denominational (6,144); Muslim Estimate (3,018)"},
  "48497": {county:"Wise",pop:70062,totalAdherents:31623,agCongs:5,agAdherents:418,evangelical:23815,pentecostal:418,catholic:4971,mainline:1692,blackProt:0,top3:"Southern Baptist (13,242); Non-denominational (8,720); Catholic Church (4,971)"},
  "48221": {county:"Hood",pop:62459,totalAdherents:31041,agCongs:5,agAdherents:527,evangelical:20899,pentecostal:1156,catholic:5358,mainline:3481,blackProt:0,top3:"Southern Baptist (12,402); Non-denominational (5,510); Catholic Church (5,358)"},
  "48181": {county:"Grayson",pop:137008,totalAdherents:61722,agCongs:10,agAdherents:2139,evangelical:43574,pentecostal:3619,catholic:6543,mainline:5746,blackProt:2850,top3:"Southern Baptist (28,981); Catholic Church (6,543); Non-denominational (6,454)"},
  "48097": {county:"Cooke",pop:41860,totalAdherents:28654,agCongs:3,agAdherents:230,evangelical:13629,pentecostal:372,catholic:12297,mainline:1506,blackProt:482,top3:"Catholic Church (12,297); Southern Baptist (6,342); Non-denominational (5,360)"},
  "48147": {county:"Fannin",pop:36052,totalAdherents:17755,agCongs:5,agAdherents:235,evangelical:14007,pentecostal:1008,catholic:1090,mainline:1443,blackProt:668,top3:"Southern Baptist (10,962); United Methodist (1,091); Catholic Church (1,090)"},
  "48485": {county:"Wichita",pop:129584,totalAdherents:88399,agCongs:11,agAdherents:3606,evangelical:54916,pentecostal:4096,catholic:22376,mainline:5469,blackProt:2152,top3:"Southern Baptist (37,103); Catholic Church (22,376); Non-denominational (8,984)"},
  "48077": {county:"Clay",pop:10290,totalAdherents:5798,agCongs:1,agAdherents:253,evangelical:5224,pentecostal:468,catholic:178,mainline:394,blackProt:0,top3:"Southern Baptist (3,439); Non-denominational (950); United Methodist (394)"},
  "48337": {county:"Montague",pop:20197,totalAdherents:8718,agCongs:5,agAdherents:664,evangelical:6434,pentecostal:664,catholic:1280,mainline:683,blackProt:126,top3:"Southern Baptist (3,863); Catholic Church (1,280); Non-denominational (830)"},
  "48237": {county:"Jack",pop:8588,totalAdherents:3918,agCongs:2,agAdherents:56,evangelical:2463,pentecostal:56,catholic:749,mainline:462,blackProt:236,top3:"Southern Baptist (2,028); Catholic Church (749); United Methodist (331)"},
  "48363": {county:"Palo Pinto",pop:28569,totalAdherents:15793,agCongs:1,agAdherents:311,evangelical:11523,pentecostal:876,catholic:2345,mainline:940,blackProt:362,top3:"Southern Baptist (8,762); Catholic Church (2,345); Non-denominational (964)"},
  "48143": {county:"Erath",pop:42788,totalAdherents:20253,agCongs:4,agAdherents:598,evangelical:13456,pentecostal:661,catholic:3110,mainline:2514,blackProt:490,top3:"Southern Baptist (9,713); Catholic Church (3,110); United Methodist (2,269)"},
  "48425": {county:"Somervell",pop:9337,totalAdherents:5199,agCongs:1,agAdherents:69,evangelical:3731,pentecostal:69,catholic:765,mainline:702,blackProt:0,top3:"Southern Baptist (2,654); Non-denominational (850); Catholic Church (765)"},
  "48349": {county:"Navarro",pop:52834,totalAdherents:23703,agCongs:3,agAdherents:282,evangelical:14466,pentecostal:1421,catholic:2388,mainline:3717,blackProt:2182,top3:"Southern Baptist (10,280); United Methodist (3,109); Catholic Church (2,388)"},
  "48213": {county:"Henderson",pop:82627,totalAdherents:39838,agCongs:13,agAdherents:1116,evangelical:26596,pentecostal:1704,catholic:6659,mainline:3834,blackProt:1802,top3:"Southern Baptist (16,039); Catholic Church (6,659); Non-denominational (6,530)"}
};

/**
 * Dynamic religious data lookup — reads from Reference Spreadsheet.
 * Works for any ZIP in the US, any denomination.
 * @param {string} zip - 5-digit ZIP code
 * @param {string} orgId - (optional) organization ID for denomination-specific highlight
 */
function lookupReligiousData(zip, orgId) {
  try {
    zip = String(zip).trim().replace(/\D/g, '');
    if (zip.length !== 5) return { success: false, error: 'Invalid ZIP code' };

    var refSS = SpreadsheetApp.openById(REFERENCE_SPREADSHEET_ID);

    // Step 1: ZIP → County FIPS
    var zcSheet = refSS.getSheetByName('ZipToCounty');
    var zcData = zcSheet.getDataRange().getValues();
    var countyFips = '';
    for (var i = 1; i < zcData.length; i++) {
      if (String(zcData[i][0]) === zip) {
        countyFips = String(zcData[i][1]);
        break;
      }
    }
    if (!countyFips) return { success: false, error: 'ZIP code not found in coverage area' };

    // Step 2: Get county religious data
    var rdSheet = refSS.getSheetByName('ReligiousData');
    var rdData = rdSheet.getDataRange().getValues();
    var headers = rdData[0];
    var countyRow = null;
    for (var i = 1; i < rdData.length; i++) {
      if (String(rdData[i][0]) === countyFips) {
        countyRow = rdData[i];
        break;
      }
    }
    if (!countyRow) return { success: false, error: 'No religious data for county ' + countyFips };

    var pop = Number(countyRow[3]) || 0;
    var totalAdherents = Number(countyRow[4]) || 0;
    var unchurchedPct = pop > 0 ? Math.round((1 - totalAdherents / pop) * 100) : 0;
    if (unchurchedPct < 0) unchurchedPct = 0;

    // Step 3: Get org's denomination for highlight card
    var denomCode = '053'; // Default to AG
    var denomName = 'Assemblies of God';
    if (orgId) {
      var org = getOrgById_(orgId);
      if (org) {
        denomCode = org.denominationCode || '053';
        denomName = org.denominationName || 'Assemblies of God';
      }
    }

    // Find the org's denomination columns in the data
    var denomCongs = 0, denomAdherents = 0;
    var denomShort = getDenomShortName_(denomName);
    for (var j = 6; j < headers.length - 1; j += 2) {
      if (String(headers[j]).indexOf(denomShort) >= 0 && String(headers[j]).indexOf('Congs') >= 0) {
        denomCongs = Number(countyRow[j]) || 0;
        denomAdherents = Number(countyRow[j + 1]) || 0;
        break;
      }
    }

    // Get top3 (last column)
    var top3 = countyRow[headers.length - 1] || '';

    // Get evangelical total (scan for it)
    var evangelical = 0;
    for (var j = 0; j < headers.length; j++) {
      if (String(headers[j]) === 'Evangelical Lutheran (ELCA) Adherents') {
        evangelical = Number(countyRow[j]) || 0;
      }
    }

    return {
      success: true,
      county: String(countyRow[2]).replace(' County', ''),
      countyFips: countyFips,
      countyPop: pop,
      totalAdherents: totalAdherents,
      unchurchedPct: unchurchedPct,
      unchurchedCount: Math.max(0, pop - totalAdherents),
      denomName: denomName,
      denomCongregations: denomCongs,
      denomAdherents: denomAdherents,
      // Legacy field names for backward compatibility with existing frontend
      agCongregations: denomCongs,
      agAdherents: denomAdherents,
      evangelical: evangelical,
      pentecostal: 0,
      catholic: 0,
      mainline: 0,
      blackProtestant: 0,
      top3: top3,
      source: 'ARDA 2020 US Religion Census'
    };
  } catch(e) {
    return { success: false, error: 'Religious data lookup failed: ' + e.message };
  }
}

/** Get org details by OrgID */
function getOrgById_(orgId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var orgSheet = ss.getSheetByName('Organizations');
    if (!orgSheet) return null;
    var data = orgSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(orgId)) {
        return {
          orgId: data[i][0],
          orgName: data[i][1],
          state: data[i][2],
          denominationCode: String(data[i][3]),
          denominationName: data[i][4],
          adminName: data[i][5],
          adminEmail: data[i][6],
          adminNotifyEmail: data[i][7],
          status: data[i][8]
        };
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

/** Map full denomination name to the short name used in ReligiousData column headers */
function getDenomShortName_(denomName) {
  var map = {
    'Assemblies of God': 'Assemblies of God',
    'Southern Baptist Convention': 'Southern Baptist Convention',
    'Catholic Church': 'Catholic Church',
    'United Methodist Church': 'United Methodist Church',
    'Non-denominational Christian Churches': 'Non-denominational',
    'Church of God (Cleveland, Tennessee)': 'Church of God (Cleveland TN)',
    'Church of God (Anderson, Indiana)': 'Church of God (Anderson IN)',
    'Church of God in Christ': 'Church of God in Christ',
    'Church of the Nazarene': 'Church of the Nazarene',
    'Churches of Christ': 'Churches of Christ',
    'Christian Churches and Churches of Christ': 'Christian Churches',
    'Evangelical Lutheran Church in America': 'Evangelical Lutheran (ELCA)',
    'Lutheran Church--Missouri Synod': 'Lutheran Church Missouri Synod',
    'Episcopal Church': 'Episcopal Church',
    'Presbyterian Church (U.S.A.)': 'Presbyterian Church USA',
    'American Baptist Churches in the USA': 'American Baptist',
    'Seventh-day Adventist Church': 'Seventh-day Adventist',
    'Church of Jesus Christ of Latter-day Saints': 'Latter-day Saints'
  };
  return map[denomName] || denomName;
}

// ── AUDIT LOG ──

function logAudit(user, action, details) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const audit = ss.getSheetByName('AuditLog');
    audit.appendRow([new Date(), user, action, details]);
  } catch(e) {
    // Silent fail — don't break the app over logging
  }
}
