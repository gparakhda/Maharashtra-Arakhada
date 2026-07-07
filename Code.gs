/**
 * ग्रामपंचायत लाभार्थी कार्य-व्यवस्थापन प्रणाली — बॅकएंड (Google Apps Script)
 *
 * ==== डिप्लॉय कसं करायचं ====
 * 1) नवीन Google Sheet बनवा
 * 2) Extensions > Apps Script मध्ये हा संपूर्ण कोड पेस्ट करा
 * 3) setupSheets() फंक्शन एकदा Run करा
 * 4) Deploy > New deployment > Web app (Execute as: Me, Access: Anyone)
 * 5) मिळालेली URL index.html मधल्या API_URL मध्ये टाका
 *
 * प्रत्येक ग्रामपंचायत + वर्षासाठी कृती आराखडा व युक्तधारा प्रणाली आराखडा हे स्वतंत्र, पूर्ण
 * फॉरमॅट केलेले शीट-टॅब आपोआप तयार होतात (एकदा तयार झाल्यावर तीच शीट पुढेही वापरली/भरली जाते —
 * त्याच वर्षासाठी त्याच ग्रामपंचायतीची नवी वेगळी शीट पुन्हा तयार होत नाही).
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEETS = {
  CONFIG: 'Config',
  DISTRICTS: 'Districts',
  TALUKAS: 'Talukas',
  GRAMPANCHAYATS: 'GramPanchayats',
  LOGIN_HISTORY: 'LoginHistory',
  OTP_TEMP: 'OTP_Temp',
  WORK_CATALOG: 'WorkCatalog',
  BENEFICIARIES: 'Beneficiaries',
  KRUTI_FORMAT: 'KrutiAarakhada_Format',
  YUKTDHARA_FORMAT: 'YuktdharaPranali_Format'
};

const HEADER_GREEN_BG = '#93c47d';

// ================= पहिल्यांदा एकदाच चालवा =================
function setupSheets() {
  const defs = {
    [SHEETS.CONFIG]: ['Key', 'Value'],
    [SHEETS.DISTRICTS]: ['District'],
    [SHEETS.TALUKAS]: ['District', 'Taluka'],
    [SHEETS.GRAMPANCHAYATS]: ['District', 'Taluka', 'GramPanchayat', 'Address', 'GP_Officer_Name', 'GP_Officer_Mobile', 'RojgarSahayak_Name', 'RojgarSahayak_Mobile', 'LastPriority'],
    [SHEETS.LOGIN_HISTORY]: ['Timestamp', 'Mobile', 'Name', 'GramPanchayat', 'OTP', 'Status'],
    [SHEETS.OTP_TEMP]: ['Mobile', 'OTP', 'CreatedAt'],
    [SHEETS.WORK_CATALOG]: ['BeneficiaryType', 'WorkName', 'WorkMainType', 'WorkSubType', 'Category', 'ParimaanUnit', 'ParimaanBaseQty', 'SkilledRatePerUnit', 'UnskilledRatePerUnit'],
    [SHEETS.BENEFICIARIES]: ['Timestamp', 'District', 'Taluka', 'GramPanchayat', 'GPAddress', 'GPOfficerName', 'GPOfficerMobile', 'RojgarSahayakName', 'RojgarSahayakMobile', 'Yantra', 'BeneficiaryName', 'BeneficiaryType', 'SocialCategory', 'FamilyComposition', 'WorkName', 'WorkMainType', 'WorkSubType', 'Category', 'ParimaanUnit', 'Quantity', 'SkilledAmount', 'UnskilledAmount', 'TotalAmount', 'ManDays', 'Priority', 'LocationType', 'PlotNumber', 'Latitude', 'Longitude', 'SavedBy', 'RecordId'],
    [SHEETS.KRUTI_FORMAT]: ['ColumnHeader', 'MapsFromField', 'Order'],
    [SHEETS.YUKTDHARA_FORMAT]: ['ColumnHeader', 'MapsFromField', 'Order']
  };

  Object.keys(defs).forEach(name => {
    let sh = SS.getSheetByName(name);
    if (!sh) sh = SS.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(defs[name]);
      sh.setFrozenRows(1);
    } else {
      migrateSheetHeaders(sh, defs[name]);
    }
  });

  const cfg = SS.getSheetByName(SHEETS.CONFIG);
  if (cfg.getLastRow() < 2) {
    const defaults = [
      ['AppName', 'ग्रामपंचायत कार्य व्यवस्थापन'],
      ['LogoUrl', ''],
      ['HeadingText', 'लाभार्थी कार्य नोंद प्रणाली'],
      ['SchemeNameEnglish', 'Viksit Bharat-Guarantee for Rozgar and Ajeevika Mission (Gramin) (VB-G RAM G) Act, 2025'],
      ['SchemeNameMarathi', 'विकसित भारत–रोजगार आणि आजीविका हमी मिशन (ग्रामीण)'],
      ['SchemeYear', '2025-26'],
      ['DeveloperName', ''],
      ['DeveloperMobile', ''],
      ['SyncIntervalMinutes', '10'],
      ['MapProvider', 'osm'],
      ['OtpMode', 'demo']
    ];
    defaults.forEach(r => cfg.appendRow(r));
  }

  const yuktFormat = SS.getSheetByName(SHEETS.YUKTDHARA_FORMAT);
  refreshFormatSheet(yuktFormat, [
    ['अ.क्र.', 'srNo', 1],
    ['कामाचे नाव', 'workName', 2],
    ['लाभार्थीचे नाव', 'beneficiaryName', 3],
    ['अक्षांश', 'latitude', 4],
    ['रेखांश', 'longitude', 5],
    ['प्राधान्यक्रम', 'priority', 6],
    ['कामाचा मुख्य प्रकार', 'workMainType', 7],
    ['कामाचा उप प्रकार', 'workSubType', 8],
    ['कामाची कॅटेगरी', 'category', 9]
  ]);

  const krutiFormat = SS.getSheetByName(SHEETS.KRUTI_FORMAT);
  refreshFormatSheet(krutiFormat, [
    ['अ.क्र.', 'srNo', 1],
    ['ग्रामपंचायत', 'gp', 2],
    ['यंत्रणा', 'yantra', 3],
    ['कामाचा प्रकार', 'workMainType', 4],
    ['सार्वजनिक कामे - कामाचे नाव', 'publicWorkName', 5],
    ['सार्वजनिक कामे - परिमाण', 'publicResult', 6],
    ['सार्वजनिक कामे - संभाव्य किंमत (लाखात) - अकुशल', 'publicUnskilledLakh', 7],
    ['सार्वजनिक कामे - संभाव्य किंमत (लाखात) - कुशल', 'publicSkilledLakh', 8],
    ['सार्वजनिक कामे - संभाव्य किंमत (लाखात) - इतर योजना अभिसरणकामांसाठी', 'publicOtherLakh', 9],
    ['सार्वजनिक कामे - संभाव्य किंमत (लाखात) - एकुण', 'publicTotalLakh', 10],
    ['सार्वजनिक कामे - संभाव्य मनुष्यदिन निर्मीती', 'publicManDays', 11],
    ['वैयक्तीक कामे - कुटूंब रचणा', 'familyComposition', 12],
    ['वैयक्तीक कामे - कामाचे नाव', 'individualWorkName', 13],
    ['वैयक्तीक कामे - लाभार्थीचे नाव', 'individualBeneficiaryName', 14],
    ['वैयक्तीक कामे - प्रवर्ग', 'socialCategory', 15],
    ['वैयक्तीक कामे - परिमाण', 'individualResult', 16],
    ['वैयक्तीक कामे - संभाव्य किंमत (लाखात) - अकुशल', 'individualUnskilledLakh', 17],
    ['वैयक्तीक कामे - संभाव्य किंमत (लाखात) - कुशल', 'individualSkilledLakh', 18],
    ['वैयक्तीक कामे - संभाव्य किंमत (लाखात) - एकुण', 'individualTotalLakh', 19],
    ['वैयक्तीक कामे - संभाव्य मनुष्यदिन निर्मीती', 'individualManDays', 20],
    ['एकुण एकत्र - परिमाण', 'totalResult', 21],
    ['एकुण एकत्र - संभाव्य किंमत (लाखात) - अकुशल', 'totalUnskilledLakh', 22],
    ['एकुण एकत्र - संभाव्य किंमत (लाखात) - कुशल', 'totalSkilledLakh', 23],
    ['एकुण एकत्र - संभाव्य किंमत (लाखात) - इतर योजना अभिसरणकामांसाठी', 'totalOtherLakh', 24],
    ['एकुण एकत्र - संभाव्य किंमत (लाखात) - एकुण', 'totalCostLakh', 25],
    ['एकुण एकत्र - संभाव्य मनुष्यदिन निर्मीती', 'totalManDays', 26]
  ]);

  Logger.log('Setup complete. आता Deploy > New deployment करा.');
}

// Format शीट प्रत्येक वेळी setupSheets चालवताना ताज्या/बरोबर व्याख्येने अपडेट करतो —
// जुन्या फेरीतला विसंगत (mismatched) डेटा शिल्लक राहून गोंधळ होऊ नये म्हणून
function refreshFormatSheet(sheet, rows) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  rows.forEach(r => sheet.appendRow(r));
}

// ================= मदतनीस फंक्शन्स =================
// आधीपासून अस्तित्वात असलेल्या शीटमध्ये कोड-अपडेटमुळे नवे कॉलम जोडले गेले असतील, तर ते हरवलेले
// कॉलम शेवटी आपोआप जोडतो — जुना डेटा अजिबात न बदलता. (उदा. WorkCatalog मध्ये ParimaanUnit हरवलेलं असेल तर हे ते जोडतं.)
function migrateSheetHeaders(sheet, desiredHeaders) {
  const lastCol = sheet.getLastColumn();
  const existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim()) : [];
  desiredHeaders.forEach(h => {
    if (!existing.includes(h)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      existing.push(h);
    }
  });
}

function sheetToObjects(sheetName) {
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  return data.filter(r => r.join('') !== '').map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (typeof v === 'string') v = v.trim();
      obj[h] = v;
    });
    return obj;
  });
}

function getConfigMap() {
  const rows = sheetToObjects(SHEETS.CONFIG);
  const map = {};
  rows.forEach(r => map[r.Key] = r.Value);
  return map;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function genOtp() { return String(Math.floor(1000 + Math.random() * 9000)); }
function genId() { return Utilities.getUuid(); }

function getFormatSorted(formatSheetName) {
  return sheetToObjects(formatSheetName).sort((a, b) => Number(a.Order) - Number(b.Order));
}

// शीट-नाव म्हणून चालणार नाहीत असे वर्ण काढून टाकतो, Unicode NFC मध्ये normalize करतो (जेणेकरून
// एकाच नावाचे वेगळे टायपिंग/एन्कोडिंग असले तरी तेच शीट सापडेल, नवीन डुप्लिकेट तयार होणार नाही),
// व 100 अक्षरांच्या मर्यादेत ठेवतो
function sanitizeSheetName(name) {
  return String(name).normalize('NFC').replace(/[\[\]\*\?\/\\:]/g, '').substring(0, 90);
}

// शीटचं नाव तंतोतंत जुळत नसलं (वेगळं normalization) तरी शोधून काढतो
// आधीची "एकूण" रांग असेल तर काढून टाकतो (नवा डेटा जोडण्याआधी हे करणं गरजेचं — नाहीतर नवा डेटा
// चुकून त्या बेरीज-रांगेखाली जाईल)
function removeTotalRowIfExists(sheet) {
  const startRow = 8;
  const lastRow = sheet.getLastRow();
  for (let r = startRow; r <= lastRow; r++) {
    if (String(sheet.getRange(r, 1).getValue()).trim() === 'एकूण') {
      sheet.deleteRow(r);
      return;
    }
  }
}

// सर्व डेटा-रांगांच्या शेवटी अकुशल/कुशल/एकुण/मनुष्यदिन (तिन्ही गटांतील) यांची बेरीज दाखवणारी हिरवी रांग टाकतो
function addOrUpdateTotalRow(sheet, totalCols) {
  const startRow = 8;
  removeTotalRowIfExists(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return; // अजून डेटाच नाही

  const sumCols = [7, 8, 9, 10, 11, 17, 18, 19, 20, 22, 23, 24, 25, 26].filter(c => c <= totalCols);
  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, totalCols).getValues();
  const sums = {};
  sumCols.forEach(c => {
    sums[c] = dataRange.reduce((acc, row) => acc + (Number(row[c - 1]) || 0), 0);
  });

  const totalRow = lastRow + 1;
  sheet.getRange(totalRow, 1).setValue('एकूण');
  sumCols.forEach(c => sheet.getRange(totalRow, c).setValue(sums[c]));

  const range = sheet.getRange(totalRow, 1, 1, totalCols);
  range.setBackground(HEADER_GREEN_BG);
  range.setFontWeight('bold');
  range.setBorder(true, true, true, true, true, true);
}

function findSheetByNormalizedName(name) {
  const target = String(name).normalize('NFC');
  const sheets = SS.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().normalize('NFC') === target) return sheets[i];
  }
  return null;
}

// एका मर्ज केलेल्या रेंजमध्ये मजकूर + स्टाईल टाकण्यासाठी मदतनीस फंक्शन
function mergeSet(sheet, r1, c1, r2, c2, value, style) {
  const range = sheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  if (r2 > r1 || c2 > c1) range.merge();
  range.setValue(value);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setWrap(true);
  if (style) {
    if (style.bold !== undefined) range.setFontWeight(style.bold ? 'bold' : 'normal');
    if (style.size) range.setFontSize(style.size);
    if (style.color) range.setFontColor(style.color);
    if (style.bg) range.setBackground(style.bg);
  }
  return range;
}

// ================= कृती आराखडा — पूर्ण फॉरमॅट केलेलं शीट हेडर (तुम्ही दिलेल्या नमुन्याप्रमाणे तंतोतंत) =================
function buildKrutiHeader(sheet, cfg, gp, taluka, district) {
  // पहिल्या 2 रांगांत उजवीकडचे 2 स्तंभ डेव्हलपर-तपशिलासाठी मोकळे ठेवतो, बाकी योजनेच्या नावासाठी
  mergeSet(sheet, 1, 1, 1, 24, cfg.SchemeNameEnglish || '', { bold: true, size: 16, color: '#FF0000' });
  mergeSet(sheet, 2, 1, 2, 24, `${cfg.SchemeNameMarathi || ''} वर्ष - ${cfg.SchemeYear || ''} आराखडा`, { bold: true, size: 16, color: '#0000FF' });
  mergeSet(sheet, 3, 1, 3, 26, `सन ${cfg.SchemeYear || ''} साठी ग्रामपंचायत निहाय कामाचे नियोजन  (WORK PROJECTION) ५०  टक्के ग्रामपंचायत यंत्रणा + ५०  टक्के शासकीय यंत्रणा)`, { bold: true, size: 12, color: '#000000' });
  mergeSet(sheet, 4, 1, 4, 26, `ग्रामपंचायत ${gp} , तालुका ${taluka} , जि. ${district}   (रक्कम लाखात)`, { bold: true, size: 11, color: '#000000' });

  insertDeveloperCredit(sheet, cfg, 25, 26);

  mergeSet(sheet, 5, 1, 7, 1, 'अ.क्र.');
  mergeSet(sheet, 5, 2, 7, 2, 'ग्रामपंचायत');
  mergeSet(sheet, 5, 3, 7, 3, 'यंत्रणा');
  mergeSet(sheet, 5, 4, 7, 4, 'कामाचा प्रकार');

  mergeSet(sheet, 5, 5, 5, 11, 'सार्वजनिक कामे');
  mergeSet(sheet, 6, 5, 7, 5, 'कामाचे नाव');
  mergeSet(sheet, 6, 6, 7, 6, 'परिमाण');
  mergeSet(sheet, 6, 7, 6, 10, 'संभाव्य किंमत');
  sheet.getRange(7, 7).setValue('अकुशल');
  sheet.getRange(7, 8).setValue('कुशल');
  sheet.getRange(7, 9).setValue('इतर योजना अभिसरणकामांसाठी');
  sheet.getRange(7, 10).setValue('एकुण');
  mergeSet(sheet, 6, 11, 7, 11, 'संभाव्य मनुष्यदिन निर्मीती');

  mergeSet(sheet, 5, 12, 5, 20, 'वैयक्तीक कामे');
  mergeSet(sheet, 6, 12, 7, 12, 'कुटूंब रचणा');
  mergeSet(sheet, 6, 13, 7, 13, 'कामाचे नाव');
  mergeSet(sheet, 6, 14, 7, 14, 'लाभार्थीचे नाव');
  mergeSet(sheet, 6, 15, 7, 15, 'प्रवर्ग');
  mergeSet(sheet, 6, 16, 7, 16, 'परिमाण');
  mergeSet(sheet, 6, 17, 6, 19, 'संभाव्य किंमत');
  sheet.getRange(7, 17).setValue('अकुशल');
  sheet.getRange(7, 18).setValue('कुशल');
  sheet.getRange(7, 19).setValue('एकुण');
  mergeSet(sheet, 6, 20, 7, 20, 'संभाव्य मनुष्यदिन निर्मीती');

  mergeSet(sheet, 5, 21, 5, 26, 'एकुण एकत्र');
  mergeSet(sheet, 6, 21, 7, 21, 'परिमाण');
  mergeSet(sheet, 6, 22, 6, 25, 'संभाव्य किंमत');
  sheet.getRange(7, 22).setValue('अकुशल');
  sheet.getRange(7, 23).setValue('कुशल');
  sheet.getRange(7, 24).setValue('इतर योजना अभिसरणकामांसाठी');
  sheet.getRange(7, 25).setValue('एकुण');
  mergeSet(sheet, 6, 26, 7, 26, 'संभाव्य मनुष्यदिन निर्मीती');

  const headRange = sheet.getRange(5, 1, 3, 26);
  headRange.setBackground(HEADER_GREEN_BG);
  headRange.setFontWeight('bold');
  headRange.setFontColor('#000000');
  headRange.setFontSize(10);
  headRange.setHorizontalAlignment('center');
  headRange.setVerticalAlignment('middle');
  headRange.setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(7);
  for (let c = 1; c <= 26; c++) sheet.setColumnWidth(c, 95);
  insertLogoIfPresent(sheet, cfg);
}

// ================= युक्तधारा प्रणाली आराखडा — पूर्ण फॉरमॅट केलेलं शीट हेडर =================
function buildYuktdharaHeader(sheet, cfg, gp, taluka, district, officer, sahayak) {
  const englishLine = cfg.SchemeNameEnglish || '';
  const marathiLine = `${cfg.SchemeNameMarathi || ''} वर्ष - ${cfg.SchemeYear || ''} आराखडा`;
  const gpLine = `ग्रामपंचायत - ${gp}, ता. - ${taluka}, जि. - ${district}`;
  const officerLine = `ग्रामपंचायत अधिकारी - ${officer}`;
  const sahayakLine = `ग्राम रोजगार सहाय्यक - ${sahayak}`;

  mergeSet(sheet, 1, 1, 1, 7, englishLine, { bold: true, size: 16 });
  mergeSet(sheet, 2, 1, 2, 7, marathiLine, { bold: true, size: 16, color: '#0000CC' });
  mergeSet(sheet, 3, 1, 3, 9, gpLine, { bold: true, size: 12, color: '#000000' });
  mergeSet(sheet, 4, 1, 4, 9, officerLine, { bold: true, size: 12, color: '#000000' });
  mergeSet(sheet, 5, 1, 5, 9, sahayakLine, { bold: true, size: 12, color: '#000000' });
  // रो 6 रिकामी

  insertDeveloperCredit(sheet, cfg, 8, 9);

  const headers = ['अ.क्र.', 'कामाचे नाव', 'लाभार्थीचे नाव', 'अक्षांश', 'रेखांश', 'प्राधान्यक्रम', 'कामाचा मुख्य प्रकार', 'कामाचा उप प्रकार', 'कामाची कॅटेगरी'];
  sheet.getRange(7, 1, 1, 9).setValues([headers]);
  const headRange = sheet.getRange(7, 1, 1, 9);
  headRange.setBackground(HEADER_GREEN_BG);
  headRange.setFontWeight('bold');
  headRange.setFontColor('#000000');
  headRange.setFontSize(10);
  headRange.setHorizontalAlignment('center');
  headRange.setVerticalAlignment('middle');
  headRange.setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(7);
  for (let c = 1; c <= 9; c++) sheet.setColumnWidth(c, 120);
  insertLogoIfPresent(sheet, cfg);
}

// डेव्हलपरचं नाव + मोबाईल शीटच्या उजव्या वरच्या कोपऱ्यात लहान अक्षरात दाखवतो
function insertDeveloperCredit(sheet, cfg, colStart, colEnd) {
  if (!cfg.DeveloperName && !cfg.DeveloperMobile) return;
  const text = `विकसित: ${cfg.DeveloperName || ''}${cfg.DeveloperMobile ? ' (' + cfg.DeveloperMobile + ')' : ''}`;
  mergeSet(sheet, 1, colStart, 1, colEnd, text, { bold: false, size: 8, color: '#555555' });
}

// डेव्हलपरने Config मध्ये LogoUrl दिला असेल तर तो शीटच्या डाव्या बाजूला, मुख्य हेडिंगशेजारी बसवतो
function insertLogoIfPresent(sheet, cfg) {
  if (!cfg.LogoUrl) return;
  try {
    const img = sheet.insertImage(cfg.LogoUrl, 1, 1);
    img.setWidth(50).setHeight(50);
  } catch (err) {
    // लोगो URL उपलब्ध नसल्यास किंवा चूक असल्यास शांतपणे पुढे जातो — बाकी शीट व्यवस्थित तयार होते
  }
}

// GP + वर्षासाठी आधीच शीट असेल तर तीच वापरतो (नवी डुप्लिकेट तयार होत नाही — पुढे भर घालणं हेच "एडिट")
function getOrCreateReportSheet(type, gp, taluka, district, gpRowData) {
  const cfg = getConfigMap();
  const year = cfg.SchemeYear || '';
  const prefix = type === 'kruti' ? 'कृती' : 'युक्तधारा';
  const sheetName = sanitizeSheetName(`${prefix}_${gp}_${year}`);

  let sh = findSheetByNormalizedName(sheetName);
  let isNew = false;
  if (!sh) {
    sh = SS.insertSheet(sheetName);
    isNew = true;
    if (type === 'kruti') {
      buildKrutiHeader(sh, cfg, gp, taluka, district);
    } else {
      const officer = `${gpRowData.GP_Officer_Name || ''} मो. ${gpRowData.GP_Officer_Mobile || ''}`;
      const sahayak = `${gpRowData.RojgarSahayak_Name || ''} मो. ${gpRowData.RojgarSahayak_Mobile || ''}`;
      buildYuktdharaHeader(sh, cfg, gp, taluka, district, officer, sahayak);
    }
  }
  return { sheet: sh, isNew: isNew, name: sheetName };
}

// प्रत्येक रेकॉर्डचा recordId शेवटच्या (न दिसणाऱ्या) कॉलममध्ये ठेवतो, जेणेकरून तोच लाभार्थी पुन्हा
// (एडिट करून) पाठवला तर नवी रांग न बनता तीच रांग अपडेट होते — डुप्लिकेट कधीही तयार होत नाही
function upsertRowsBelowHeader(sheet, format, records) {
  const startRow = 8;
  const idCol = format.length + 1;
  const lastRow = sheet.getLastRow();

  // सध्याच्या recordId कॉलममधली मूल्ये वाचून rowIndex चा नकाशा तयार करतो
  const idMap = {};
  if (lastRow >= startRow) {
    const idValues = sheet.getRange(startRow, idCol, lastRow - startRow + 1, 1).getValues();
    idValues.forEach((v, i) => { if (v[0]) idMap[v[0]] = startRow + i; });
  }

  let nextRow = Math.max(lastRow + 1, startRow);
  records.forEach(r => {
    const row = format.map(f => r[f.MapsFromField] !== undefined ? r[f.MapsFromField] : '');
    const targetRow = (r.recordId && idMap[r.recordId]) ? idMap[r.recordId] : nextRow;
    const range = sheet.getRange(targetRow, 1, 1, row.length);
    range.setValues([row]);
    range.setBorder(true, true, true, true, true, true);
    sheet.getRange(targetRow, idCol).setValue(r.recordId || '');
    if (targetRow === nextRow) nextRow++;
  });
}

// ================= doGet - डेटा वाचणे =================
function doGet(e) {
  const action = e.parameter.action;
  try {
    switch (action) {
      case 'config':
        return jsonOut({ ok: true, data: getConfigMap() });

      case 'districts':
        return jsonOut({ ok: true, data: sheetToObjects(SHEETS.DISTRICTS) });

      case 'talukas': {
        const district = e.parameter.district;
        let rows = sheetToObjects(SHEETS.TALUKAS);
        if (district) rows = rows.filter(r => r.District === district);
        return jsonOut({ ok: true, data: rows });
      }

      case 'gramPanchayats': {
        const district = e.parameter.district;
        const taluka = e.parameter.taluka;
        let rows = sheetToObjects(SHEETS.GRAMPANCHAYATS);
        if (district) rows = rows.filter(r => r.District === district);
        if (taluka) rows = rows.filter(r => r.Taluka === taluka);
        return jsonOut({ ok: true, data: rows });
      }

      case 'allLocations':
        return jsonOut({
          ok: true,
          data: {
            districts: sheetToObjects(SHEETS.DISTRICTS),
            talukas: sheetToObjects(SHEETS.TALUKAS),
            gramPanchayats: sheetToObjects(SHEETS.GRAMPANCHAYATS)
          }
        });

      case 'gpDetails': {
        const gp = e.parameter.gp;
        const rows = sheetToObjects(SHEETS.GRAMPANCHAYATS).filter(r => r.GramPanchayat === gp);
        return jsonOut({ ok: true, data: rows[0] || null });
      }

      case 'workCatalog': {
        const type = e.parameter.beneficiaryType;
        let rows = sheetToObjects(SHEETS.WORK_CATALOG);
        if (type) rows = rows.filter(r => r.BeneficiaryType === type);
        return jsonOut({ ok: true, data: rows });
      }

      case 'loginHistory': {
        const gp = e.parameter.gp;
        let rows = sheetToObjects(SHEETS.LOGIN_HISTORY);
        if (gp) rows = rows.filter(r => r.GramPanchayat === gp);
        return jsonOut({ ok: true, data: rows.reverse() });
      }

      case 'beneficiaries': {
        const gp = e.parameter.gp;
        let rows = sheetToObjects(SHEETS.BENEFICIARIES);
        if (gp) rows = rows.filter(r => r.GramPanchayat === gp);
        return jsonOut({ ok: true, data: rows });
      }

      case 'reportFormat': {
        const which = e.parameter.type;
        const sheetName = which === 'kruti' ? SHEETS.KRUTI_FORMAT : SHEETS.YUKTDHARA_FORMAT;
        return jsonOut({ ok: true, data: getFormatSorted(sheetName) });
      }

      case 'checkExistingPlan': {
        const gp = e.parameter.gp;
        const cfg = getConfigMap();
        const year = cfg.SchemeYear || '';
        const krutiName = sanitizeSheetName(`कृती_${gp}_${year}`);
        const yuktName = sanitizeSheetName(`युक्तधारा_${gp}_${year}`);
        const krutiExists = !!findSheetByNormalizedName(krutiName);
        const yuktExists = !!findSheetByNormalizedName(yuktName);
        return jsonOut({ ok: true, exists: krutiExists || yuktExists, year });
      }

      default:
        return jsonOut({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ================= doPost - डेटा लिहिणे / OTP =================
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  try {
    switch (action) {
      case 'sendOtp': return sendOtp(body);
      case 'verifyOtp': return verifyOtp(body);
      case 'saveBeneficiary': return saveBeneficiary(body);
      case 'finalizeSave': return finalizeSave(body);
      case 'upsertGpOfficer': return upsertGpOfficer(body);
      case 'getNextPriority': return getNextPriority(body);
      case 'deleteBeneficiaryRecord': return deleteBeneficiaryRecord(body);
      case 'updateBeneficiaryRecord': return updateBeneficiaryRecord(body);
      default:
        return jsonOut({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function sendOtp(body) {
  const mobile = String(body.mobile || '').trim();
  if (mobile.length !== 10) return jsonOut({ ok: false, error: 'योग्य 10 अंकी मोबाईल क्रमांक टाका' });
  const otp = genOtp();
  SS.getSheetByName(SHEETS.OTP_TEMP).appendRow([mobile, otp, new Date()]);
  const cfg = getConfigMap();
  if (cfg.OtpMode === 'live') {
    // ==== इथे खरी WhatsApp Business API कॉल टाका ====
  }
  return jsonOut({ ok: true, name: body.name || '', demoOtp: cfg.OtpMode === 'demo' ? otp : undefined });
}

function verifyOtp(body) {
  const mobile = String(body.mobile || '').trim();
  const otp = String(body.otp || '').trim();
  const sh = SS.getSheetByName(SHEETS.OTP_TEMP);
  const data = sh.getDataRange().getValues();
  let matched = false;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === mobile && String(data[i][1]) === otp) {
      const created = new Date(data[i][2]);
      if ((new Date() - created) < 10 * 60 * 1000) matched = true;
      break;
    }
  }
  if (!matched) return jsonOut({ ok: false, error: 'चुकीचा किंवा कालबाह्य OTP' });
  const user = { Mobile: mobile, Name: body.name || 'यूजर', Role: 'फिल्ड यूजर', GramPanchayat: '' };
  SS.getSheetByName(SHEETS.LOGIN_HISTORY).appendRow([new Date(), mobile, user.Name, user.GramPanchayat, otp, 'Success']);
  return jsonOut({ ok: true, user });
}

function upsertGpOfficer(body) {
  const { district, taluka, gp, address, officerName, officerMobile, sahayakName, sahayakMobile } = body;
  const sh = SS.getSheetByName(SHEETS.GRAMPANCHAYATS);
  const data = sh.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === district && data[i][1] === taluka && data[i][2] === gp) { rowIndex = i; break; }
  }
  if (rowIndex > -1) {
    sh.getRange(rowIndex + 1, 4, 1, 4).setValues([[address || '', officerName || '', officerMobile || '', sahayakName || '']]);
    sh.getRange(rowIndex + 1, 8, 1, 1).setValues([[sahayakMobile || '']]);
  } else {
    sh.appendRow([district, taluka, gp, address || '', officerName || '', officerMobile || '', sahayakName || '', sahayakMobile || '', 0]);
  }
  return jsonOut({ ok: true });
}

// प्रत्येक ग्रामपंचायतीचा प्राधान्यक्रम-आकडा वेगळा व अनुक्रमे वाढता ठेवतो — कधीही डबल होत नाही,
// आणि ग्रामपंचायत बदलली तरच पुन्हा तोच आकडा (नव्या गावासाठी) वापरता येतो
function getNextPriority(body) {
  const { district, taluka, gp } = body;
  const sh = SS.getSheetByName(SHEETS.GRAMPANCHAYATS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === district && data[i][1] === taluka && data[i][2] === gp) {
      const current = Number(data[i][8]) || 0;
      const next = current + 1;
      sh.getRange(i + 1, 9).setValue(next);
      return jsonOut({ ok: true, priority: next });
    }
  }
  // ग्रामपंचायत सापडली नाही तर नवी नोंद तयार करून 1 पासून सुरू करतो
  sh.appendRow([district, taluka, gp, '', '', '', '', '', 1]);
  return jsonOut({ ok: true, priority: 1 });
}

// चुकून चुकीची माहिती सेव्ह झाली असेल तर (अंतिम सेव्हपूर्वी) ती रेकॉर्ड-आयडी वरून कच्च्या लॉगमधून काढून टाकतो
function deleteBeneficiaryRecord(body) {
  const { recordId } = body;
  const sh = SS.getSheetByName(SHEETS.BENEFICIARIES);
  const data = sh.getDataRange().getValues();
  const idCol = data[0].indexOf('RecordId');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idCol] === recordId) {
      sh.deleteRow(i + 1);
      return jsonOut({ ok: true });
    }
  }
  return jsonOut({ ok: false, error: 'रेकॉर्ड सापडलं नाही' });
}

// आधी सेव्ह झालेली नोंद एडिट करून पुन्हा पाठवली की, नवी रांग न बनवता तीच रांग अपडेट करतो
function updateBeneficiaryRecord(body) {
  const d = body.data;
  const sh = SS.getSheetByName(SHEETS.BENEFICIARIES);
  const data = sh.getDataRange().getValues();
  const idCol = data[0].indexOf('RecordId');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === d.recordId) {
      sh.getRange(i + 1, 1, 1, 30).setValues([[
        data[i][0], d.district, d.taluka, d.gp, d.gpAddress, d.gpOfficerName, d.gpOfficerMobile,
        d.rojgarSahayakName, d.rojgarSahayakMobile, d.yantra, d.beneficiaryName, d.beneficiaryType,
        d.socialCategory, d.familyComposition, d.workName, d.workMainType, d.workSubType, d.category,
        d.parimaanUnit, d.quantity, d.skilledAmount, d.unskilledAmount, d.totalAmount, d.manDays, d.priority,
        d.locationType, d.plotNumber, d.latitude, d.longitude, d.savedBy
      ]]);
      return jsonOut({ ok: true, recordId: d.recordId });
    }
  }
  return jsonOut({ ok: false, error: 'रेकॉर्ड सापडलं नाही, नवीन म्हणून सेव्ह करा' });
}

function saveBeneficiary(body) {
  const d = body.data;
  const sh = SS.getSheetByName(SHEETS.BENEFICIARIES);
  const id = genId();
  sh.appendRow([
    new Date(), d.district, d.taluka, d.gp, d.gpAddress, d.gpOfficerName, d.gpOfficerMobile,
    d.rojgarSahayakName, d.rojgarSahayakMobile, d.yantra, d.beneficiaryName, d.beneficiaryType,
    d.socialCategory, d.familyComposition, d.workName, d.workMainType, d.workSubType, d.category,
    d.parimaanUnit, d.quantity, d.skilledAmount, d.unskilledAmount, d.totalAmount, d.manDays, d.priority,
    d.locationType, d.plotNumber, d.latitude, d.longitude, d.savedBy, id
  ]);
  return jsonOut({ ok: true, recordId: id });
}

// "संपूर्ण माहिती सेव्ह करा" — त्या ग्रामपंचायत + वर्षाच्या स्वतःच्या फॉरमॅट केलेल्या शीटमध्ये (नवी नसेल तर आधीच्या शीटमध्ये पुढे) डेटा जातो
function finalizeSave(body) {
  const records = body.records || [];
  if (records.length === 0) return jsonOut({ ok: false, error: 'सेव्ह करण्यासाठी लाभार्थी नाहीत' });

  const first = records[0];
  const gpRows = sheetToObjects(SHEETS.GRAMPANCHAYATS).filter(r => r.GramPanchayat === first.gp);
  const gpRowData = gpRows[0] || {};

  const krutiFormat = getFormatSorted(SHEETS.KRUTI_FORMAT);
  const yuktFormat = getFormatSorted(SHEETS.YUKTDHARA_FORMAT);

  const krutiRes = getOrCreateReportSheet('kruti', first.gp, first.taluka, first.district, gpRowData);
  const yuktRes = getOrCreateReportSheet('yuktdhara', first.gp, first.taluka, first.district, gpRowData);

  removeTotalRowIfExists(krutiRes.sheet);
  upsertRowsBelowHeader(krutiRes.sheet, krutiFormat, records);
  upsertRowsBelowHeader(yuktRes.sheet, yuktFormat, records);
  addOrUpdateTotalRow(krutiRes.sheet, 26);

  return jsonOut({
    ok: true,
    message: `सर्व माहिती सेव्ह झाली — एकूण ${records.length} लाभार्थी "${krutiRes.name}" व "${yuktRes.name}" या शीटमध्ये गेले` +
      (krutiRes.isNew ? '' : ' (आधीच्या आराखड्यात पुढे भर घातली)')
  });
}
