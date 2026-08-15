const ACTIVITY_SHEET = 'Activities';
const CATEGORY_SHEET = 'Categories';
const SETTING_SHEET = 'Settings';
const LOG_SHEET = 'Logs';

const DEFAULT_YEAR = 2026;
const TIMEZONE = 'Asia/Bangkok';

// ======================================================
// SETUP SYSTEM
// ======================================================

function setupCalendarSystem() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  PropertiesService
    .getScriptProperties()
    .setProperty('SPREADSHEET_ID', ss.getId());

  // Activities
  let sheet = ss.getSheetByName(ACTIVITY_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(ACTIVITY_SHEET);
  }

  sheet.clear();

  const headers = [
    'ID',
    'Date',
    'Start Time',
    'End Time',
    'Title',
    'Category',
    'Description',
    'Location',
    'Priority',
    'Status',
    'Owner',
    'Color',
    'Created At',
    'Updated At'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#172554')
    .setFontColor('#ffffff');

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 240);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 300);
  sheet.setColumnWidth(8, 180);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 150);

  // Categories
  let cat = ss.getSheetByName(CATEGORY_SHEET);

  if (!cat) {
    cat = ss.insertSheet(CATEGORY_SHEET);
  }

  cat.clear();

  cat.getRange(1, 1, 1, 3)
    .setValues([[
      'Category',
      'Color',
      'Description'
    ]])
    .setFontWeight('bold')
    .setBackground('#172554')
    .setFontColor('#ffffff');

  cat.getRange(2, 1, 6, 3).setValues([
    ['งาน', '#2563eb', 'งานทั่วไป'],
    ['ประชุม', '#7c3aed', 'Meeting / Appointment'],
    ['Deadline', '#dc2626', 'กำหนดส่งงาน'],
    ['ส่วนตัว', '#16a34a', 'Personal'],
    ['Project', '#ea580c', 'Project Activity'],
    ['อื่นๆ', '#64748b', 'Other']
  ]);

  cat.setFrozenRows(1);

  // Settings
  let setting = ss.getSheetByName(SETTING_SHEET);

  if (!setting) {
    setting = ss.insertSheet(SETTING_SHEET);
  }

  setting.clear();

  setting.getRange(1, 1, 1, 2)
    .setValues([['Setting', 'Value']])
    .setFontWeight('bold')
    .setBackground('#172554')
    .setFontColor('#ffffff');

  setting.getRange(2, 1, 3, 2).setValues([
    ['YEAR', 2026],
    ['TIMEZONE', 'Asia/Bangkok'],
    ['SYSTEM_NAME', 'Annual Planner 2026']
  ]);

  // Logs
  let log = ss.getSheetByName(LOG_SHEET);

  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
  }

  log.clear();

  log.getRange(1, 1, 1, 5)
    .setValues([[
      'Timestamp',
      'Action',
      'Activity ID',
      'Title',
      'User'
    ]])
    .setFontWeight('bold')
    .setBackground('#172554')
    .setFontColor('#ffffff');

  SpreadsheetApp.flush();

  return 'Calendar System 2026 setup completed.';
}


// ======================================================
// WEB APP
// ======================================================

function doGet() {

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Annual Planner 2026')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}


// ======================================================
// DATABASE
// ======================================================

function getDB() {

  const id = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}


// ======================================================
// INITIAL DATA
// ======================================================

function getInitialData() {

  return {
    year: DEFAULT_YEAR,
    activities: getActivities(DEFAULT_YEAR),
    categories: getCategories()
  };
}


// ======================================================
// GET ACTIVITIES
// ======================================================

function getActivities(year) {

  const sheet = getDB().getSheetByName(ACTIVITY_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const data = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      14
    )
    .getDisplayValues();

  return data
    .filter(r => {

      if (!r[1]) return false;

      const y = extractYear(r[1]);

      return Number(y) === Number(year);

    })
    .map(r => ({
      id: r[0],
      date: normalizeDate(r[1]),
      startTime: r[2],
      endTime: r[3],
      title: r[4],
      category: r[5],
      description: r[6],
      location: r[7],
      priority: r[8],
      status: r[9],
      owner: r[10],
      color: r[11],
      createdAt: r[12],
      updatedAt: r[13]
    }));
}


function extractYear(dateStr) {

  const value = normalizeDate(dateStr);

  if (!value) return '';

  return value.substring(0, 4);
}


function normalizeDate(value) {

  if (!value) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }

  const parts = value.split('/');

  if (parts.length === 3) {

    let year = Number(parts[2]);

    if (year > 2400) {
      year -= 543;
    }

    return (
      year +
      '-' +
      String(parts[1]).padStart(2, '0') +
      '-' +
      String(parts[0]).padStart(2, '0')
    );
  }

  return value;
}


// ======================================================
// GET CATEGORIES
// ======================================================

function getCategories() {

  const sheet = getDB().getSheetByName(CATEGORY_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const rows = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      3
    )
    .getValues();

  return rows
    .filter(r => r[0])
    .map(r => ({
      name: r[0],
      color: r[1],
      description: r[2]
    }));
}


// ======================================================
// SAVE ACTIVITY
// ======================================================

function saveActivity(data) {

  const lock = LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    const sheet = getDB().getSheetByName(ACTIVITY_SHEET);

    const now = Utilities.formatDate(
      new Date(),
      TIMEZONE,
      'yyyy-MM-dd HH:mm:ss'
    );

    const id = data.id || generateID();

    const color =
      data.color ||
      getCategoryColor(data.category);

    const row = [
      id,
      data.date || '',
      data.startTime || '',
      data.endTime || '',
      data.title || '',
      data.category || '',
      data.description || '',
      data.location || '',
      data.priority || 'กลาง',
      data.status || 'วางแผน',
      data.owner || '',
      color || '#2563eb',
      data.createdAt || now,
      now
    ];

    let action = 'CREATE';

    if (data.id) {

      const ids = sheet
        .getRange(
          2,
          1,
          Math.max(sheet.getLastRow() - 1, 1),
          1
        )
        .getValues()
        .flat();

      const index = ids.indexOf(data.id);

      if (index !== -1) {

        sheet
          .getRange(index + 2, 1, 1, row.length)
          .setValues([row]);

        action = 'UPDATE';

      } else {

        sheet.appendRow(row);

      }

    } else {

      sheet.appendRow(row);

    }

    writeLog(
      action,
      id,
      data.title || ''
    );

    return {
      success: true,
      id: id
    };

  } finally {

    lock.releaseLock();

  }
}


// ======================================================
// DELETE
// ======================================================

function deleteActivity(id) {

  const sheet = getDB().getSheetByName(ACTIVITY_SHEET);

  if (sheet.getLastRow() < 2) {
    return false;
  }

  const ids = sheet
    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      1
    )
    .getValues()
    .flat();

  const index = ids.indexOf(id);

  if (index === -1) {
    return false;
  }

  const title =
    sheet.getRange(index + 2, 5).getValue();

  sheet.deleteRow(index + 2);

  writeLog(
    'DELETE',
    id,
    title
  );

  return true;
}


// ======================================================
// HELPERS
// ======================================================

function generateID() {

  return (
    'ACT-' +
    Utilities.formatDate(
      new Date(),
      TIMEZONE,
      'yyyyMMddHHmmss'
    ) +
    '-' +
    Math.floor(Math.random() * 900 + 100)
  );
}


function getCategoryColor(category) {

  const categories = getCategories();

  const found = categories.find(
    x => x.name === category
  );

  return found
    ? found.color
    : '#2563eb';
}


function writeLog(action, id, title) {

  const sheet = getDB().getSheetByName(LOG_SHEET);

  let user = '';

  try {
    user = Session.getActiveUser().getEmail();
  } catch (e) {}

  sheet.appendRow([
    Utilities.formatDate(
      new Date(),
      TIMEZONE,
      'yyyy-MM-dd HH:mm:ss'
    ),
    action,
    id,
    title,
    user
  ]);
}
