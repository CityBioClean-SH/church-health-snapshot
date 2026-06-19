/**
 * Church Health Snapshot — Backend
 * Google Apps Script server-side code
 * Scoring formulas are server-side only — never exposed to the client
 */

const SPREADSHEET_ID = '1f_zyugFWUkuepoGJA0MEGx2lM3i3xDHb8PcOJiNqXCQ';
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

    logAudit('SYSTEM', 'CENSUS_LOOKUP', 'ZIP: ' + zip + ' — ' + placeName + ', Pop: ' + population);

    return {
      success: true,
      placeName: placeName,
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

function setupAdmin() {
  // Set your admin email here, then run this function once
  const adminEmail = 'mharper@northtexas.ag';
  PropertiesService.getScriptProperties().setProperty('ADMIN_EMAIL', adminEmail);
  Logger.log('Admin email set to: ' + adminEmail);
}

// ── WEB APP ──

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('App')
    .setTitle('Church Health Snapshot')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── AUTH ──

function verifyPIN(pin) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const users = ss.getSheetByName('Users');
    const data = users.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(pin) && data[i][7] === 'Y') {
        // Update last login
        users.getRange(i + 1, 7).setValue(new Date());

        // Check firstLogin flag
        var isFirstLogin = (data[i][8] === 'Y');
        if (isFirstLogin) {
          users.getRange(i + 1, 9).setValue('N'); // Flip to N after first login
        }

        // Audit log
        logAudit(data[i][1], 'LOGIN', 'PIN login successful' + (isFirstLogin ? ' (first time)' : ''));

        return {
          success: true,
          name: data[i][1],
          email: data[i][2],
          role: data[i][3],
          church: data[i][4],
          firstLogin: isFirstLogin
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

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === String(email).toLowerCase() && data[i][7] === 'Y') {
        // Update last login
        users.getRange(i + 1, 7).setValue(new Date());

        // Check firstLogin flag
        var isFirstLogin = (data[i][8] === 'Y');
        if (isFirstLogin) {
          users.getRange(i + 1, 9).setValue('N');
        }

        logAudit(data[i][1], 'LOGIN', 'Email login: ' + email + (isFirstLogin ? ' (first time)' : ''));

        return {
          success: true,
          name: data[i][1],
          email: data[i][2],
          role: data[i][3],
          church: data[i][4],
          firstLogin: isFirstLogin
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
      perCapitaGiving: Math.round(perCapitaGiving),
      potentialPerCapita: Math.round(potentialPC),
      givingGap: Math.round(perCapitaGiving - potentialPC),
      pastorAgeDiff: pastorAgeDiff,
      congAgeDiff: congAgeDiff
    }
  };
}

// ── SUBMIT ASSESSMENT ──

function submitAssessment(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const assess = ss.getSheetByName('Assessments');

    // Calculate score server-side
    const result = calculateScore(data.answers);

    // Build row
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
      result.score, result.rating, result.s1, result.s2, result.s3
    ];

    assess.appendRow(row);

    // Audit log
    logAudit(data.userName, 'SUBMIT', data.churchName + ' — Score: ' + result.score + ' (' + result.rating + ')');

    // Silent admin email
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

function getAssessments(userRole) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const assess = ss.getSheetByName('Assessments');
    const data = assess.getDataRange().getValues();

    if (data.length <= 1) return { success: true, assessments: [] };

    const headers = data[0];
    const assessments = [];

    for (let i = 1; i < data.length; i++) {
      assessments.push({
        timestamp: data[i][0],
        assessedBy: data[i][1],
        role: data[i][2],
        churchName: data[i][3],
        pastorName: data[i][4],
        score: data[i][37],
        rating: data[i][38],
        s1: data[i][39],
        s2: data[i][40],
        s3: data[i][41]
      });
    }

    return { success: true, assessments: assessments };
  } catch(e) {
    return { success: false, error: 'Failed to load assessments' };
  }
}

// ── ADMIN: MANAGE USERS ──

function addUser(data) {
  try {
    var pin = data.pin, name = data.name, email = data.email, role = data.role, church = data.church;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var users = ss.getSheetByName('Users');

    // Check for duplicate PIN
    var existing = users.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (String(existing[i][0]) === String(pin)) {
        return { success: false, error: 'PIN already in use' };
      }
    }

    users.appendRow([String(pin), name, email || '', role, church || '', new Date(), '', 'Y', 'Y']);
    logAudit('ADMIN', 'USER_ADDED', name + ' (' + role + ')');

    // Send welcome email
    if (email) {
      sendWelcomeEmail(name, email, pin, role);
    }

    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function sendWelcomeEmail(name, email, pin, role) {
  try {
    var appUrl = ScriptApp.getService().getUrl();
    var adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || '';

    var body = '<div style="font-family:\'Segoe UI\',sans-serif;max-width:560px;margin:0 auto;">' +
      '<div style="background:#1b2541;padding:28px 24px;border-radius:10px 10px 0 0;text-align:center;">' +
        '<h1 style="margin:0;color:white;font-size:22px;">Welcome to Church Health Snapshot</h1>' +
        '<p style="margin:6px 0 0;color:#c9a227;font-size:13px;">North Texas Assemblies of God</p>' +
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

function getUsers() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const users = ss.getSheetByName('Users');
    const data = users.getDataRange().getValues();

    const userList = [];
    for (let i = 1; i < data.length; i++) {
      userList.push({
        row: i + 1,
        pin: String(data[i][0]),
        name: data[i][1],
        email: data[i][2],
        role: data[i][3],
        church: data[i][4],
        registered: data[i][5],
        lastLogin: data[i][6],
        active: data[i][7]
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
