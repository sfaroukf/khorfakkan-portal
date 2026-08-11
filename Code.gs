/*************************************************************
 * نظام إدارة قاعات الاجتماعات وطلبات التنسيق والمتابعة
 * مستشفى خورفكان — Khorfakkan Hospital
 * Backend — Code.gs
 * CODE_VERSION = "v2026-08-10-settings-cache-perf"  ← دوّر عليها بـ Ctrl+F للتأكد إنك ناسخ آخر نسخة
 *************************************************************/

/* ============ إعدادات عامة ============ */
const SHEET_NAMES = {
  USERS: 'Users',
  ROOMS: 'Rooms',
  BOOKINGS: 'Bookings',
  REQUESTS: 'Requests',
  CATEGORIES: 'Categories',
  SETTINGS: 'Settings',
  DIRECTORY: 'Directory',
  AUDITLOG: 'AuditLog'
};

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* ============ التثبيت الأولي — شغّلها مرة واحدة فقط ============ */
function setupDatabase() {
  const ss = getSS();

  const sheetsConfig = {
    Users: ['Email', 'Name', 'Role', 'Department', 'AssignedCategories', 'Status', 'PasswordHash', 'Salt'],
    Rooms: ['RoomID', 'RoomName', 'RoomNameEn', 'Location', 'Capacity', 'ManagerEmail', 'Status'],
    Bookings: ['BookingID', 'RoomID', 'RoomName', 'RequesterEmail', 'RequesterName', 'RequesterPhone', 'Date', 'StartTime', 'EndTime', 'Purpose', 'Status', 'CreatedAt', 'ApprovedBy', 'Notes'],
    Requests: ['RequestID', 'RequesterEmail', 'RequesterName', 'RequesterPhone', 'Category', 'Description', 'AssignedTo', 'Status', 'CreatedAt', 'CompletedAt', 'AINotes', 'Priority', 'AttachmentUrl', 'AttachmentName', 'CompletionNotes', 'WasTransferred', 'AdminNote'],
    Categories: ['CategoryName', 'DefaultAssigneeEmail'],
    Settings: ['Key', 'Value'],
    Directory: ['Phone', 'Name', 'Email', 'Department', 'UpdatedAt', 'PasswordHash', 'Salt'],
    AuditLog: ['Timestamp', 'ActorEmail', 'ActorName', 'Action', 'Details']
  };

  Object.keys(sheetsConfig).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear();
    sheet.getRange(1, 1, 1, sheetsConfig[name].length).setValues([sheetsConfig[name]]);
    sheet.getRange(1, 1, 1, sheetsConfig[name].length).setFontWeight('bold').setBackground('#0F766E').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  });

  // مستخدم أدمن افتراضي — عدّل البريد إلى بريدك
  const usersSheet = ss.getSheetByName('Users');
  usersSheet.appendRow([Session.getActiveUser().getEmail(), 'مدير النظام', 'Admin', 'الإدارة', '', 'Active']);

  // فئات طلبات افتراضية
  const catSheet = ss.getSheetByName('Categories');
  catSheet.appendRow(['رسالة رسمية', '']);
  catSheet.appendRow(['تصميم', '']);
  catSheet.appendRow(['ميديا', '']);
  catSheet.appendRow(['أخرى', '']);

  // قاعات افتراضية
  const roomSheet = ss.getSheetByName('Rooms');
  roomSheet.appendRow(['R1', 'قاعة الاجتماعات الكبرى', 'المبنى الرئيسي - الطابق الأول', 20, '', 'Active']);
  roomSheet.appendRow(['R2', 'قاعة التدريب', 'المبنى الإداري - الطابق الثاني', 15, '', 'Active']);

  // إعدادات — اسم البرنامج والشعار قابلان للتعديل لاحقاً من لوحة تحكم الأدمن
  const settingsSheet = ss.getSheetByName('Settings');
  settingsSheet.appendRow(['GEMINI_API_KEY', 'ضع_مفتاح_API_هنا']);
  settingsSheet.appendRow(['APP_NAME', 'بوابة خورفكان الإدارية']);
  settingsSheet.appendRow(['LOGO_URL', '']);
  // بريد صاحب الموقع — الشخص الوحيد اللي هيتعرف عليه النظام تلقائيًا بجلسة Google؛
  // أي حد تاني (حتى لو أدمن) لازم يسجّل دخول ببريده الشخصي وكلمة سر
  settingsSheet.appendRow(['OWNER_EMAIL', Session.getActiveUser().getEmail()]);

  SpreadsheetApp.getUi().alert('تم إعداد قاعدة البيانات بنجاح. لا تنسَ إضافة مفتاح Gemini API في ورقة Settings.');
}

/* ============ ترقية ذاتية للمخطط — تضيف ورقة Directory وأعمدة الهاتف تلقائياً
 * للمشتركين الحاليين دون الحاجة لإعادة تشغيل setupDatabase (لا تمسح أي بيانات) ============ */
function ensureSchemaUpgrades() {
  // الفحص الكامل مكلف (~10 استدعاءات Sheets API) — نخزّن نتيجة "تم الفحص" مؤقتاً
  // عشان ميتكررش في كل تحميل صفحة، وده كان السبب الرئيسي لبطء التنقل بين الصفحات
  const cache = CacheService.getScriptCache();
  if (cache.get('schema_checked_v1')) return;

  const ss = getSS(); // خارج كل try/catch عشان يفضل متاح لكل الخطوات التالية

  try {
    if (!ss.getSheetByName(SHEET_NAMES.DIRECTORY)) {
      const sheet = ss.insertSheet(SHEET_NAMES.DIRECTORY);
      const headers = ['Phone', 'Name', 'Email', 'Department', 'UpdatedAt'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0F766E').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  } catch (err) { Logger.log('ensureSchemaUpgrades (Directory) error: ' + err); }

  try {
    if (!ss.getSheetByName(SHEET_NAMES.AUDITLOG)) {
      const sheet = ss.insertSheet(SHEET_NAMES.AUDITLOG);
      const headers = ['Timestamp', 'ActorEmail', 'ActorName', 'Action', 'Details'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0F766E').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  } catch (err) { Logger.log('ensureSchemaUpgrades (AuditLog) error: ' + err); }

  try { addColumnIfMissing(SHEET_NAMES.BOOKINGS, 'RequesterPhone'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.REQUESTS, 'RequesterPhone'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.REQUESTS, 'AttachmentUrl'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.REQUESTS, 'AttachmentName'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.REQUESTS, 'CompletionNotes'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.REQUESTS, 'WasTransferred'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.REQUESTS, 'AdminNote'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.DIRECTORY, 'PasswordHash'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.DIRECTORY, 'Salt'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.USERS, 'PasswordHash'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.USERS, 'Salt'); } catch (err) { Logger.log(err); }
  try { addColumnIfMissing(SHEET_NAMES.ROOMS, 'RoomNameEn'); } catch (err) { Logger.log(err); }

  // للمشتركين الحاليين الذين ثبّتوا النظام قبل هذا التحديث — عيّن صاحب الموقع تلقائيًا
  // (أول شخص بجلسة Google يفتح الرابط بعد هذا التحديث)
  try {
    const settings = getSheetData(SHEET_NAMES.SETTINGS);
    if (!settings.find(s => s.Key === 'OWNER_EMAIL')) {
      const currentEmail = Session.getActiveUser().getEmail();
      if (currentEmail) setSetting('OWNER_EMAIL', currentEmail);
    }
  } catch (err) { Logger.log(err); }

  cache.put('schema_checked_v1', 'true', 7200); // ساعتين — كافية جداً لأن بنية الأعمدة نادراً ما تتغير
}

function addColumnIfMissing(sheetName, columnName) {
  const sheet = getSS().getSheetByName(sheetName);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (headers.indexOf(columnName) === -1) {
    sheet.getRange(1, headers.length + 1).setValue(columnName);
  }
}

/* أضف قيماً لصف جديد حسب أسماء الأعمدة الفعلية بالورقة — يمنع كسر البيانات عند تعديل المخطط لاحقاً */
function appendRowByHeaders(sheetName, valuesObj) {
  const sheet = getSS().getSheetByName(sheetName);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const row = headers.map(h => (valuesObj[h] !== undefined ? valuesObj[h] : ''));
  sheet.appendRow(row);
}

/* ============ توجيه الطلبات — Web App Router ============
 * الدالة كلها محمية بـ try/catch: أي خطأ غير متوقع يظهر برسالة واضحة للمستخدم
 * بدل صفحة فاضية بيضاء لا تفسّر شيئاً — هذا يسهّل تشخيص أي عطل مستقبلي فوراً. */
function doGet(e) {
  try {
    ensureSchemaUpgrades();
    const page = e.parameter.page || 'Dashboard';
    const staffToken = e.parameter.staffToken || '';
    const user = getCurrentUser(staffToken);

    if (!user) {
      return HtmlService.createTemplateFromFile('AccessDenied').evaluate()
        .setTitle('غير مصرح — نظام إدارة القاعات')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    let template;
    let resolvedPage = 'Dashboard';
    // ضيف لسه ما سجّلش دخول خالص (لا Google ولا فريق عمل) — يُسمح له بمشاهدة شاشة الدخول
    // في صفحات الأدمن/الإحصائيات، بدل ما يتحول تلقائيًا لصفحة الحجز
    const isUnauthenticatedGuest = (!user.Email && !user.Registered);
    if (page === 'Admin' && (user.Role === 'Admin' || user.Role === 'RoomManager' || isUnauthenticatedGuest)) {
      template = HtmlService.createTemplateFromFile('Admin');
      resolvedPage = 'Admin';
    } else if (page === 'Statistics' && (user.Role === 'Admin' || user.Role === 'RoomManager' || user.Role === 'Coordinator' || isUnauthenticatedGuest)) {
      template = HtmlService.createTemplateFromFile('Statistics');
      resolvedPage = 'Statistics';
    } else if (page === 'Requests') {
      template = HtmlService.createTemplateFromFile('Requests');
      resolvedPage = 'Requests';
    } else if (page === 'Assistant') {
      template = HtmlService.createTemplateFromFile('Assistant');
      resolvedPage = 'Assistant';
    } else {
      template = HtmlService.createTemplateFromFile('Dashboard');
      resolvedPage = 'Dashboard';
    }

    const appSettings = getAppSettings();
    template.currentUser = user;
    template.appSettings = appSettings;
    template.currentPage = resolvedPage;
    template.staffToken = staffToken;
    template.hasGoogleSession = !!Session.getActiveUser().getEmail();
    return template.evaluate()
      .setTitle(appSettings.appName)
      .setFaviconUrl(appSettings.logoUrl || 'https://www.gstatic.com/images/branding/product/2x/apps_script_48dp.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    const stackInfo = (err && err.stack) ? err.stack : 'لا تتوفر تفاصيل إضافية.';
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;direction:rtl;text-align:right;padding:40px;max-width:700px;margin:auto;">' +
      '<h2 style="color:#b91c1c;">حدث خطأ أثناء تحميل الصفحة</h2>' +
      '<p style="color:#444;">تفاصيل الخطأ التقنية (أرسلها لمسؤول النظام إن استمرت المشكلة):</p>' +
      '<pre style="background:#f3f4f6;padding:12px;border-radius:8px;white-space:pre-wrap;">' + (err && err.message ? err.message : err) + '</pre>' +
      '<p style="color:#888;font-size:12px;margin-top:12px;">مكان الخطأ بالتفصيل:</p>' +
      '<pre style="background:#fef2f2;padding:12px;border-radius:8px;white-space:pre-wrap;font-size:11px;color:#991b1b;">' + stackInfo + '</pre>' +
      '</div>'
    ).setTitle('خطأ').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * API bridge for the Cloudflare Pages frontend.
 * The browser never calls Google Apps Script directly.
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').trim();
    const args = Array.isArray(body.args) ? body.args : [];

    const allowed = ["addAdminNoteToRequest", "addCategoryManually", "addOrUpdateRoom", "addOrUpdateUser", "adminResetPassword", "adminSetStaffPassword", "confirmPasswordReset", "createBooking", "createRequest", "deleteBooking", "deleteDirectoryUser", "deleteRequest", "deleteRoom", "deleteUser", "getAISummary", "getAllRequests", "getAllRoomsForAdmin", "getAllUsers", "getAssignedRequests", "getAuditLog", "getBackupsList", "getBookingsForCalendar", "getBookingsHistory", "getBookingsReport", "getCategories", "getCoordinatorsList", "getDashboardStats", "getDirectoryUsers", "getGeminiKeyStatus", "getMyBookings", "getMyNotificationCount", "getMyRequests", "getMyStatsBundle", "getOwnerEmail", "getOwnerEmailForAdmin", "getPendingBookings", "getRequestsReport", "getRooms", "handleAIAssistantMessage", "loginAccount", "manualBackupNow", "reassignRequest", "registerAccount", "requestPasswordReset", "saveAppName", "saveAssistantSettings", "saveEmailNotificationSetting", "saveGeminiApiKey", "saveLogo", "saveOwnerEmail", "saveVoiceRecorderUrl", "selfRegisterEmail", "staffLogin", "staffLogout", "testGeminiConnection", "updateBookingDetails", "updateBookingStatus", "updateRequestDetails", "updateRequestStatus", "uploadRequestAttachment"];
    if (!allowed.includes(action)) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: { message: 'العملية غير مسموح بها.' } }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const fn = globalThis[action];
    if (typeof fn !== 'function') {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: { message: 'الدالة غير موجودة: ' + action } }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = fn.apply(null, args);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: {
          message: err && err.message ? err.message : String(err),
          name: err && err.name ? err.name : 'Error'
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}



function include(filename, data) {
  const tmpl = HtmlService.createTemplateFromFile(filename);
  if (data) {
    Object.keys(data).forEach(key => { tmpl[key] = data[key]; });
  }
  return tmpl.evaluate().getContent();
}

/* ============ دليل الهاتف — تسجيل سريع: اكتب الهاتف مرة، وباقي المرات يتعرف عليك تلقائياً ============
 * التطبيع يعتمد على آخر 9 أرقام فقط (يتجاهل مفتاح الدولة 971 والصفر الأول تلقائياً)
 * حتى لو اختلفت طريقة كتابة نفس الرقم بين مرة وأخرى (05xxxxxxxx / 5xxxxxxxx / 9715xxxxxxxx). */

// تحويل الأرقام العربية الهندية (٠١٢٣٤٥٦٧٨٩) والفارسية/الأردو (۰۱۲۳۴۵۶۷۸۹) للأرقام الإنجليزية —
// ضروري لأن أي جهاز أو كيبورد شغال بنظام أرقام عربي/هندي هيبعت الأرقام دي بدل 0-9 العادية،
// وكل مقارنات النظام (رقم الهاتف، الأوقات، رموز التحقق) بتتوقع أرقام إنجليزية فقط.
function normalizeDigits(str) {
  if (str === null || str === undefined) return str;
  const arabicIndic = '٠١٢٣٤٥٦٧٨٩';
  const persianIndic = '۰۱۲۳۴۵۶۷۸۹';
  return String(str).replace(/[٠-٩۰-۹]/g, d => {
    let idx = arabicIndic.indexOf(d);
    if (idx > -1) return String(idx);
    idx = persianIndic.indexOf(d);
    return idx > -1 ? String(idx) : d;
  });
}

function normalizePhone(phone) {
  const digits = normalizeDigits(phone).replace(/[^0-9]/g, '');
  return digits.slice(-9);
}

function lookupByPhone(phone) {
  try {
    const normalized = normalizePhone(phone);
    if (!normalized || normalized.length < 6) return null;
    const sheet = getSS().getSheetByName(SHEET_NAMES.DIRECTORY);
    if (!sheet) return null;
    const dir = getSheetData(SHEET_NAMES.DIRECTORY);
    const row = dir.find(d => normalizePhone(d.Phone) === normalized);
    return row || null;
  } catch (err) {
    Logger.log('lookupByPhone error: ' + err);
    return null;
  }
}

function saveToDirectory(phone, name, email, department) {
  try {
    const normalized = normalizePhone(phone);
    if (!normalized || !name) return { success: false };
    const sheet = getSS().getSheetByName(SHEET_NAMES.DIRECTORY);
    if (!sheet) return { success: false }; // الورقة لسه ما اتنشأتش — لن توقف عملية الحجز/الطلب
    const data = sheet.getDataRange().getValues();
    const phoneCol = data[0].indexOf('Phone');
    for (let i = 1; i < data.length; i++) {
      if (normalizePhone(data[i][phoneCol]) === normalized) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[phone, name, email || '', department || '', new Date()]]);
        return { success: true };
      }
    }
    sheet.appendRow([phone, name, email || '', department || '', new Date()]);
    return { success: true };
  } catch (err) {
    Logger.log('saveToDirectory error: ' + err);
    return { success: false }; // خطأ هنا لا يجب أبداً أن يمنع نجاح الحجز أو الطلب
  }
}
/* ============ المصادقة والمستخدمين ============
 * فلسفة الصلاحيات (مُحدَّثة):
 * - جلسة Google تُعتمَد تلقائياً فقط لصاحب الموقع (البريد المحدد في OWNER_EMAIL بورقة Settings).
 * - أي شخص آخر — حتى لو كان يملك جلسة Google — يجب أن يسجّل دخوله ببريده الشخصي وكلمة مرور
 *   (staffLogin) أو يُنشئ حساباً جديداً (selfRegisterEmail)، مع خيار "تذكرني" لتفادي تكرار الدخول.
 * - المستخدم الجديد الذي يسجّل بنفسه يُمنح تلقائياً دور "موظف" (Employee)، ويمكن للأدمن ترقيته لاحقاً.
 */
function getOwnerEmail() {
  return String(getAppSettings().ownerEmail || '').toLowerCase();
}

// تخزين رمز الجلسة — نستخدم PropertiesService بدل CacheService لأن الأخيرة محدودة بـ 6 ساعات
// كحد أقصى (قيد من جوجل نفسها)، بينما "تذكرني" يحتاج جلسة تدوم أيامًا أو أسابيع
function createSessionToken(payload, remember) {
  const token = Utilities.getUuid();
  const ttlMs = remember ? 30 * 24 * 60 * 60 * 1000 : 8 * 60 * 60 * 1000; // 30 يوم لو "تذكرني"، وإلا 8 ساعات
  payload.expiresAt = Date.now() + ttlMs;
  PropertiesService.getScriptProperties().setProperty('stok_' + token, JSON.stringify(payload));
  return token;
}

function readSessionToken(token) {
  if (!token) return null;
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('stok_' + token);
    if (!raw) return null;
    const data = JSON.parse(raw);
    if (data.expiresAt && Date.now() > data.expiresAt) {
      PropertiesService.getScriptProperties().deleteProperty('stok_' + token);
      return null;
    }
    return data;
  } catch (err) {
    return null;
  }
}

function revokeSessionToken(token) {
  if (!token) return;
  try { PropertiesService.getScriptProperties().deleteProperty('stok_' + token); } catch (err) { /* تجاهل */ }
}

function getCurrentUser(staffToken) {
  const email = Session.getActiveUser().getEmail();
  const ownerEmail = getOwnerEmail();

  // فقط صاحب الموقع (البريد المحدد في الإعدادات) يتعرف عليه النظام تلقائيًا بجلسة Google
  if (email && ownerEmail && email.toLowerCase() === ownerEmail) {
    const data = getSheetData(SHEET_NAMES.USERS);
    const row = data.find(r => r.Email && r.Email.toLowerCase() === email.toLowerCase() && r.Status === 'Active');
    if (row) { row.Registered = true; return row; }
    return { Email: email, Name: email.split('@')[0], Role: 'Admin', Department: '', Status: 'Active', Registered: false };
  }

  // أي شخص آخر — حتى لو عنده جلسة Google — يحتاج رمز جلسة صادر من تسجيل دخول صريح
  if (staffToken) {
    const cached = readSessionToken(staffToken);
    if (cached) {
      return { Email: cached.Email, Name: cached.Name, Role: cached.Role, Department: cached.Department || '', Status: 'Active', Registered: true };
    }
  }

  // بدون أي جلسة صريحة — ضيف مجهول (بيُطلب منه تسجيل الدخول عبر بوابة الدخول)
  return { Email: '', Name: 'ضيف', Role: 'Employee', Department: '', Status: 'Active', Registered: false };
}

function getAllUsers(staffToken) {
  requireAdmin(staffToken);
  return getSheetData(SHEET_NAMES.USERS);
}

function addOrUpdateUser(userObj, staffToken) {
  requireAdmin(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  const emailCol = data[0].indexOf('Email');
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol].toLowerCase() === userObj.Email.toLowerCase()) { rowIndex = i + 1; break; }
  }
  const rowValues = [userObj.Email, userObj.Name, userObj.Role, userObj.Department, userObj.AssignedCategories || '', userObj.Status || 'Active'];
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  logAudit('إضافة/تعديل مستخدم', `${userObj.Name} (${userObj.Email}) — الدور: ${userObj.Role}`);
  return { success: true };
}

// حذف مستخدم دور (أدمن/مسئول قاعة/تنسيق) — يعود موظفًا عاديًا افتراضيًا بمجرد حذفه من هنا
function deleteUser(email, staffToken) {
  const admin = requireAdmin(staffToken);
  if (admin.Email && admin.Email.toLowerCase() === String(email).toLowerCase()) {
    return { success: false, message: 'لا يمكنك حذف حسابك الخاص من هنا.' };
  }
  const sheet = getSS().getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  const emailCol = data[0].indexOf('Email');
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] && String(data[i][emailCol]).toLowerCase() === String(email).toLowerCase()) {
      logAudit('حذف مستخدم', `${data[i][data[0].indexOf('Name')]} (${email})`);
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'المستخدم غير موجود.' };
}

// ============ تسجيل دخول فريق العمل بالبريد وكلمة السر ============
// بديل لجلسة Google — يُستخدم لما جلسة Google غير متاحة (نشر Anyone، أو مشاكل شبكة)
function staffLogin(email, password, remember) {
  const users = getSheetData(SHEET_NAMES.USERS);
  const user = users.find(u => u.Email && u.Email.toLowerCase() === String(email || '').toLowerCase() && u.Status === 'Active');
  if (!user) return { success: false, message: 'لا يوجد حساب بهذا البريد. لو أول مرة تستخدم النظام، اضغط "إنشاء حساب جديد".' };
  if (!user.PasswordHash) return { success: false, message: 'هذا الحساب لم يُفعَّل بكلمة مرور دخول بعد. تواصل مع مدير النظام لتفعيله من صفحة "المستخدمون".' };

  const hash = hashPassword(String(password || ''), user.Salt);
  if (hash !== user.PasswordHash) return { success: false, message: 'كلمة المرور غير صحيحة.' };

  const token = createSessionToken({ Email: user.Email, Name: user.Name, Role: user.Role, Department: user.Department || '' }, !!remember);
  return { success: true, token: token, user: { Email: user.Email, Name: user.Name, Role: user.Role } };
}

// تسجيل حساب جديد ببريد شخصي — لأي موظف ليس له دور رسمي بعد؛ يُمنح تلقائياً دور "موظف"
// ويقدر الأدمن يرقّيه لاحقاً من صفحة المستخدمين. لو الحساب موجود بالفعل بدون كلمة مرور
// (أضافه الأدمن يدويًا بدور معيّن)، تُستخدم هذه الدالة لتفعيله بنفس الدور المحدد له.
function selfRegisterEmail(name, email, password, remember) {
  name = String(name || '').trim();
  email = String(email || '').trim().toLowerCase();
  if (!name) return { success: false, message: 'يرجى إدخال الاسم.' };
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { success: false, message: 'بريد إلكتروني غير صالح.' };
  const pwError = passwordStrengthError(password);
  if (pwError) return { success: false, message: pwError };

  const sheet = getSS().getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const nameCol = headers.indexOf('Name');
  const roleCol = headers.indexOf('Role');
  const passCol = headers.indexOf('PasswordHash');
  const saltCol = headers.indexOf('Salt');
  const statusCol = headers.indexOf('Status');

  const salt = Utilities.getUuid();
  const hash = hashPassword(password, salt);

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] && String(data[i][emailCol]).toLowerCase() === email) {
      if (data[i][passCol]) {
        return { success: false, message: 'هذا البريد مسجّل بالفعل. استخدم "تسجيل الدخول" بدلاً من ذلك.' };
      }
      // حساب أضافه الأدمن يدويًا بدور معيّن، بدون كلمة مرور بعد — فعّله الآن بنفس الدور
      sheet.getRange(i + 1, nameCol + 1).setValue(name);
      sheet.getRange(i + 1, passCol + 1).setValue(hash);
      sheet.getRange(i + 1, saltCol + 1).setValue(salt);
      const role = data[i][roleCol] || 'Employee';
      const token = createSessionToken({ Email: email, Name: name, Role: role, Department: '' }, !!remember);
      logAudit('تفعيل حساب موظف (ذاتي)', `${name} (${email})`);
      return { success: true, token: token, user: { Email: email, Name: name, Role: role } };
    }
  }

  // مفيش حساب بهذا البريد — أنشئ موظفًا جديدًا بدور "Employee" افتراضيًا
  appendRowByHeaders(SHEET_NAMES.USERS, {
    Email: email, Name: name, Role: 'Employee', Department: '', AssignedCategories: '', Status: 'Active',
    PasswordHash: hash, Salt: salt
  });
  const token = createSessionToken({ Email: email, Name: name, Role: 'Employee', Department: '' }, !!remember);
  logAudit('تسجيل حساب موظف جديد (ذاتي)', `${name} (${email})`);
  return { success: true, token: token, user: { Email: email, Name: name, Role: 'Employee' } };
}

// تسجيل الخروج — إبطال رمز الجلسة فعليًا من السيرفر
function staffLogout(staffToken) {
  revokeSessionToken(staffToken);
  return { success: true };
}

// الأدمن يفعّل/يعيد تعيين كلمة مرور دخول أي مستخدم دور (أدمن/مسئول قاعة/تنسيق) مباشرة
function adminSetStaffPassword(email, newPassword, staffToken) {
  requireAdmin(staffToken);
  const pwError = passwordStrengthError(newPassword);
  if (pwError) return { success: false, message: pwError };

  const sheet = getSS().getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const passCol = headers.indexOf('PasswordHash');
  const saltCol = headers.indexOf('Salt');

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] && String(data[i][emailCol]).toLowerCase() === String(email).toLowerCase()) {
      const salt = Utilities.getUuid();
      const hash = hashPassword(String(newPassword), salt);
      sheet.getRange(i + 1, passCol + 1).setValue(hash);
      sheet.getRange(i + 1, saltCol + 1).setValue(salt);
      logAudit('تفعيل/تعديل كلمة مرور دخول فريق العمل', `للبريد: ${email}`);
      return { success: true, message: 'تم حفظ كلمة المرور بنجاح — يقدر يدخل بيها من "دخول فريق العمل".' };
    }
  }
  return { success: false, message: 'المستخدم غير موجود في ورقة Users.' };
}


function requireAdmin(staffToken) {
  const user = getCurrentUser(staffToken);
  if (!user || user.Role !== 'Admin') {
    throw new Error('غير مصرح لك بتنفيذ هذا الإجراء — للأدمن فقط');
  }
  return user;
}

/* ============ سجل التدقيق — يسجّل كل إجراء إداري مهم: مين عمل إيه وإمتى ============
 * آمنة تماماً: أي خطأ في التسجيل نفسه لا يوقف العملية الأصلية أبداً. */
function logAudit(action, details) {
  try {
    const user = getCurrentUser();
    const sheet = getSS().getSheetByName(SHEET_NAMES.AUDITLOG);
    if (!sheet) return;
    sheet.appendRow([new Date(), user.Email || '(بدون جلسة)', user.Name || '', action, details || '']);
  } catch (err) {
    Logger.log('logAudit error: ' + err);
  }
}

function getAuditLog(limit, staffToken) {
  requireAdmin(staffToken);
  ensureSchemaUpgrades(); // يضمن وجود ورقة AuditLog قبل القراءة، لو كانت غائبة لأي سبب
  const rows = getSheetData(SHEET_NAMES.AUDITLOG).sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
  return rows.slice(0, limit || 200);
}

// اعتماد/رفض الحجوزات مسموح للأدمن (كل القاعات) ولمسئول القاعة (قاعاته فقط)
function requireApprover(staffToken) {
  const user = getCurrentUser(staffToken);
  if (!user || (user.Role !== 'Admin' && user.Role !== 'RoomManager')) {
    throw new Error('غير مصرح لك بتنفيذ هذا الإجراء');
  }
  return user;
}

/* ============ أدوات عامة لقراءة الأوراق كـ JSON ============
 * ملاحظة مهمة: Google Sheets يحوّل قيم التاريخ/الوقت المُدخلة كنص (مثل "14:00") تلقائياً
 * إلى كائنات Date عند القراءة عبر getValues() — لو تركناها كما هي، تفشل مقارنات النصوص
 * (مثل بناء "2026-01-01T14:00" للتقويم) بصمت، فتختفي بعض الحجوزات من التقويم والتقارير.
 * لذلك نطبّع أعمدة Date/StartTime/EndTime دائماً لنصوص موحدة هنا في نقطة قراءة واحدة. */
function getSheetData(sheetName) {
  const sheet = getSS().getSheetByName(sheetName);
  if (!sheet) return []; // الورقة غير موجودة — نرجع مصفوفة فاضية بدل الانهيار (زي حالة AuditLog قبل إنشائها)
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const tz = Session.getScriptTimeZone() || 'Asia/Dubai';
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        if (h === 'Date') val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
        else if (h === 'StartTime' || h === 'EndTime') val = Utilities.formatDate(val, tz, 'HH:mm');
        // أي عمود تاريخ/وقت آخر (CreatedAt, CompletedAt, UpdatedAt, ApprovedBy...) يتحول لنص موحّد
        // أيضاً — إرسال كائن Date خام عبر google.script.run كان يسبب فشل صامت في وصول البيانات للواجهة
        else val = Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      obj[h] = val;
    });
    return obj;
  });
}

/* ============ القاعات ============ */
function getRooms() {
  return getSheetData(SHEET_NAMES.ROOMS).filter(r => r.Status === 'Active');
}

/* ============ الحجوزات ============ */
function getBookingsForCalendar() {
  const bookings = getSheetData(SHEET_NAMES.BOOKINGS).filter(b => b.Status === 'Approved' || b.Status === 'Pending');
  const roomsMap = {};
  getSheetData(SHEET_NAMES.ROOMS).forEach(r => { roomsMap[r.RoomID] = r; });

  return bookings.map(b => {
    const room = roomsMap[b.RoomID];
    const roomNameEn = (room && room.RoomNameEn) ? room.RoomNameEn : b.RoomName;
    return {
      id: b.BookingID,
      start: b.Date + 'T' + b.StartTime,
      end: b.Date + 'T' + b.EndTime,
      color: b.Status === 'Pending' ? '#F59E0B' : '#0F766E',
      extendedProps: {
        status: b.Status, requester: b.RequesterName, purpose: b.Purpose,
        roomNameAr: b.RoomName, roomNameEn: roomNameEn
      }
    };
  });
}

function createBooking(booking) {
  const user = getCurrentUser();
  const requesterPhone = String(booking.phone || '').trim();
  const requesterName = (booking.guestName || user.Name || '').trim();
  const requesterEmail = (booking.guestEmail || user.Email || '').trim();

  if (!requesterName) {
    return { success: false, message: 'يرجى إدخال الاسم لإتمام الحجز.' };
  }
  // رقم الهاتف مطلوب بس للضيوف (بدون جلسة Google) — أصحاب حساب Google هويتهم معروفة أصلاً
  if (!user.Email && !requesterPhone) {
    return { success: false, message: 'يرجى إدخال رقم الهاتف لإتمام الحجز.' };
  }

  if (isDoubleBooked(booking.roomId, booking.date, booking.startTime, booking.endTime)) {
    return { success: false, message: 'هذا الموعد محجوز مسبقاً أو معلّق الموافقة لنفس القاعة.' };
  }

  const bookingId = 'BK-' + new Date().getTime();
  const room = getRooms().find(r => r.RoomID === booking.roomId);

  appendRowByHeaders(SHEET_NAMES.BOOKINGS, {
    BookingID: bookingId, RoomID: booking.roomId, RoomName: room ? room.RoomName : booking.roomId,
    RequesterEmail: requesterEmail, RequesterName: requesterName, RequesterPhone: requesterPhone,
    Date: booking.date, StartTime: booking.startTime, EndTime: booking.endTime,
    Purpose: booking.purpose, Status: 'Pending', CreatedAt: new Date(), ApprovedBy: '', Notes: ''
  });

  // تسجيل/تحديث تلقائي في دليل الهاتف — المرة القادمة يكفي رقم الهاتف لتعبئة الباقي
  saveToDirectory(requesterPhone, requesterName, requesterEmail, user.Department || '');

  notifyAdmins('طلب حجز قاعة جديد', `${requesterName} طلب حجز "${room ? room.RoomName : ''}" بتاريخ ${booking.date} من ${booking.startTime} إلى ${booking.endTime}.\nالغرض: ${booking.purpose}`);

  return { success: true, message: 'تم إرسال طلب الحجز وهو الآن بانتظار موافقة الأدمن.' };
}

function isDoubleBooked(roomId, date, startTime, endTime) {
  const bookings = getSheetData(SHEET_NAMES.BOOKINGS).filter(b =>
    b.RoomID === roomId && b.Date === date && (b.Status === 'Approved' || b.Status === 'Pending')
  );
  return bookings.some(b => timeOverlap(startTime, endTime, b.StartTime, b.EndTime));
}

function timeOverlap(s1, e1, s2, e2) {
  return (s1 < e2) && (s2 < e1);
}

function getPendingBookings(staffToken) {
  const user = requireApprover(staffToken);
  const pending = getSheetData(SHEET_NAMES.BOOKINGS).filter(b => b.Status === 'Pending');
  if (user.Role === 'Admin') return pending;
  // مسئول القاعة يرى فقط طلبات قاعاته (يدعم أكثر من مسئول لنفس القاعة)
  const myRoomIds = getMyManagedRoomIds(user.Email);
  return pending.filter(b => myRoomIds.includes(b.RoomID));
}

function updateBookingStatus(bookingId, newStatus, staffToken) {
  const approver = requireApprover(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.BOOKINGS);
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('BookingID');
  const statusCol = data[0].indexOf('Status');
  const approvedByCol = data[0].indexOf('ApprovedBy');
  const emailCol = data[0].indexOf('RequesterEmail');
  const roomCol = data[0].indexOf('RoomName');
  const dateCol = data[0].indexOf('Date');
  const roomIdCol = data[0].indexOf('RoomID');
  const startCol = data[0].indexOf('StartTime');
  const endCol = data[0].indexOf('EndTime');
  const purposeCol = data[0].indexOf('Purpose');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === bookingId) {
      // مسئول القاعة لا يستطيع اعتماد حجوزات قاعة ليست تابعة له
      if (approver.Role === 'RoomManager') {
        const room = getSheetData(SHEET_NAMES.ROOMS).find(r => r.RoomID === data[i][roomIdCol]);
        if (!room || !isRoomManagerOf(room, approver.Email)) {
          throw new Error('غير مصرح لك باعتماد حجوزات هذه القاعة');
        }
      }
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      sheet.getRange(i + 1, approvedByCol + 1).setValue(approver.Name);

      if (isEmailNotificationsEnabled() && data[i][emailCol]) {
        try {
          const statusArabic = newStatus === 'Approved' ? 'تمت الموافقة على' : 'تم رفض';
          let subject = 'تحديث حالة حجز القاعة';
          let body = `${statusArabic} طلب حجزك لقاعة "${data[i][roomCol]}" بتاريخ ${data[i][dateCol]}.`;

          const ai = generateEmailContent(newStatus === 'Approved' ? 'booking_approved' : 'booking_rejected', {
            roomName: data[i][roomCol], date: data[i][dateCol],
            startTime: data[i][startCol], endTime: data[i][endCol], purpose: data[i][purposeCol]
          });
          if (ai) { subject = ai.subject; body = ai.body; }

          MailApp.sendEmail(data[i][emailCol], subject, body);
        } catch (err) { /* تجاهل خطأ البريد — لا يجب أن يوقف عملية الاعتماد */ }
      }

      logAudit(newStatus === 'Approved' ? 'اعتماد حجز' : 'رفض حجز', `${data[i][roomCol]} — ${data[i][dateCol]} — مقدّم الطلب: ${data[i][emailCol]}`);
      return { success: true };
    }
  }
  return { success: false };
}

// عدّاد التنبيهات الحية للقائمة الجانبية — يختلف المحتوى حسب دور المستخدم
function getMyNotificationCount(staffToken) {
  const user = getCurrentUser(staffToken);
  let bookingsCount = 0;
  let requestsCount = 0;

  if (user.Role === 'Admin') {
    bookingsCount = getSheetData(SHEET_NAMES.BOOKINGS).filter(b => b.Status === 'Pending').length;
    requestsCount = getSheetData(SHEET_NAMES.REQUESTS).filter(r => r.Status === 'Pending').length;
  } else if (user.Role === 'RoomManager') {
    const roomIds = getMyManagedRoomIds(user.Email);
    bookingsCount = getSheetData(SHEET_NAMES.BOOKINGS).filter(b => b.Status === 'Pending' && roomIds.includes(b.RoomID)).length;
  } else if (user.Role === 'Coordinator') {
    requestsCount = getSheetData(SHEET_NAMES.REQUESTS).filter(r =>
      r.Status !== 'Completed' && r.AssignedTo && user.Email && r.AssignedTo.toLowerCase() === user.Email.toLowerCase()
    ).length;
  }

  return { bookings: bookingsCount, requests: requestsCount, total: bookingsCount + requestsCount };
}

// يدعم أكثر من مسئول قاعة واحد — البريد الإلكتروني مخزَّن كقائمة مفصولة بفاصلة في نفس الخلية
function roomManagerEmailsList(room) {
  return String(room.ManagerEmail || '').split(',').map(e => e.trim().toLowerCase()).filter(Boolean);
}

function isRoomManagerOf(room, email) {
  if (!email) return false;
  return roomManagerEmailsList(room).includes(String(email).toLowerCase());
}

function getMyManagedRoomIds(email) {
  return getSheetData(SHEET_NAMES.ROOMS)
    .filter(r => isRoomManagerOf(r, email))
    .map(r => r.RoomID);
}

function canManageBookingRoom(approver, roomId) {
  if (approver.Role === 'Admin') return true;
  if (approver.Role === 'RoomManager') return getMyManagedRoomIds(approver.Email).includes(roomId);
  return false;
}

/* ============ سجل الحجوزات الكامل (كل الحالات) — للأدمن كل القاعات، ولمسئول القاعة قاعاته فقط ============ */
function getBookingsHistory(staffToken) {
  const approver = requireApprover(staffToken);
  const all = getSheetData(SHEET_NAMES.BOOKINGS);
  const filtered = approver.Role === 'Admin' ? all : all.filter(b => canManageBookingRoom(approver, b.RoomID));
  return filtered.sort((a, b) => new Date(b.Date) - new Date(a.Date));
}

/* ============ تعديل حجز — مع إعادة فحص تعارض الأوقات ============ */
function updateBookingDetails(bookingId, updates, staffToken) {
  const approver = requireApprover(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.BOOKINGS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('BookingID');
  const roomIdCol = headers.indexOf('RoomID');
  const dateCol = headers.indexOf('Date');
  const startCol = headers.indexOf('StartTime');
  const endCol = headers.indexOf('EndTime');
  const purposeCol = headers.indexOf('Purpose');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === bookingId) {
      const roomId = data[i][roomIdCol];
      if (!canManageBookingRoom(approver, roomId)) {
        throw new Error('غير مصرح لك بتعديل حجوزات هذه القاعة');
      }
      const newDate = updates.date || data[i][dateCol];
      const newStart = updates.startTime || data[i][startCol];
      const newEnd = updates.endTime || data[i][endCol];

      const conflict = getSheetData(SHEET_NAMES.BOOKINGS).some(b =>
        b.BookingID !== bookingId && b.RoomID === roomId && b.Date === newDate &&
        (b.Status === 'Approved' || b.Status === 'Pending') && timeOverlap(newStart, newEnd, b.StartTime, b.EndTime)
      );
      if (conflict) return { success: false, message: 'تعارض مع حجز آخر لنفس القاعة في هذا الوقت.' };

      sheet.getRange(i + 1, dateCol + 1).setValue(newDate);
      sheet.getRange(i + 1, startCol + 1).setValue(newStart);
      sheet.getRange(i + 1, endCol + 1).setValue(newEnd);
      if (updates.purpose) sheet.getRange(i + 1, purposeCol + 1).setValue(updates.purpose);
      logAudit('تعديل حجز', `${bookingId} — ${newDate} ${newStart}-${newEnd}`);
      return { success: true };
    }
  }
  return { success: false, message: 'الحجز غير موجود.' };
}

/* ============ حذف حجز ============ */
function deleteBooking(bookingId, staffToken) {
  const approver = requireApprover(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.BOOKINGS);
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('BookingID');
  const roomIdCol = data[0].indexOf('RoomID');
  const roomNameCol = data[0].indexOf('RoomName');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === bookingId) {
      if (!canManageBookingRoom(approver, data[i][roomIdCol])) {
        throw new Error('غير مصرح لك بحذف حجوزات هذه القاعة');
      }
      logAudit('حذف حجز', `${bookingId} — ${data[i][roomNameCol]}`);
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'الحجز غير موجود.' };
}

/* ============ إدارة القاعات وتعيين مسئوليها — الأدمن فقط ============
 * يقدر الأدمن هنا يربط أي موظف (بريده) كمسئول لقاعة واحدة أو أكثر،
 * فيصبح هذا الموظف يشوف فقط حجوزات القاعات المخصصة له (سجل كامل + اعتماد/رفض + تعديل/حذف). */
function getAllRoomsForAdmin(staffToken) {
  requireAdmin(staffToken);
  return getSheetData(SHEET_NAMES.ROOMS);
}

function addOrUpdateRoom(roomObj, staffToken) {
  requireAdmin(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.ROOMS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('RoomID');
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === roomObj.RoomID) { rowIndex = i + 1; break; }
  }
  const valuesObj = {
    RoomID: roomObj.RoomID, RoomName: roomObj.RoomName, RoomNameEn: roomObj.RoomNameEn || '',
    Location: roomObj.Location, Capacity: roomObj.Capacity, ManagerEmail: roomObj.ManagerEmail || '', Status: roomObj.Status || 'Active'
  };
  const rowValues = headers.map(h => (valuesObj[h] !== undefined ? valuesObj[h] : ''));
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  logAudit('إضافة/تعديل قاعة', `${roomObj.RoomName} — مسئول القاعة: ${roomObj.ManagerEmail || 'بدون'}`);
  return { success: true };
}

// حذف قاعة نهائيًا — يمنع الحذف لو فيها حجوزات معلّقة أو معتمدة مستقبلية (لتفادي فقدان التزامات قائمة)
function deleteRoom(roomId, staffToken) {
  requireAdmin(staffToken);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Dubai', 'yyyy-MM-dd');
  const hasActiveBookings = getSheetData(SHEET_NAMES.BOOKINGS).some(b =>
    b.RoomID === roomId && (b.Status === 'Pending' || b.Status === 'Approved') && b.Date >= today
  );
  if (hasActiveBookings) {
    return { success: false, message: 'لا يمكن حذف هذه القاعة — يوجد بها حجوزات معلّقة أو معتمدة قادمة. ألغِها أولاً.' };
  }
  const sheet = getSS().getSheetByName(SHEET_NAMES.ROOMS);
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('RoomID');
  const nameCol = data[0].indexOf('RoomName');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === roomId) {
      logAudit('حذف قاعة', data[i][nameCol]);
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'القاعة غير موجودة.' };
}

function getCategories() {
  return getSheetData(SHEET_NAMES.CATEGORIES);
}

// يحفظ فئة جديدة في القائمة الرئيسية لو مش موجودة بالفعل — بيسمح لأي موظف إنه يضيف
// نوع طلب جديد وقت ما بيقدّم طلب، ويظهر تلقائيًا لكل الموظفين من بعد كده
function addCategoryIfNew(categoryName) {
  categoryName = String(categoryName || '').trim();
  if (!categoryName) return { success: false };
  const cats = getSheetData(SHEET_NAMES.CATEGORIES);
  const exists = cats.some(c => String(c.CategoryName || '').trim().toLowerCase() === categoryName.toLowerCase());
  if (exists) return { success: true, alreadyExists: true };

  appendRowByHeaders(SHEET_NAMES.CATEGORIES, { CategoryName: categoryName, DefaultAssigneeEmail: '' });
  logAudit('إضافة نوع طلب جديد', categoryName);
  return { success: true, alreadyExists: false };
}

// إضافة نوع طلب جديد يدويًا — للأدمن وموظف التنسيق (المسئول عن القسم) فقط، مش لأي موظف عادي
function addCategoryManually(categoryName, staffToken) {
  const user = getCurrentUser(staffToken);
  if (user.Role !== 'Admin' && user.Role !== 'Coordinator') {
    throw new Error('غير مصرح لك بإضافة نوع طلب جديد — للأدمن وموظف التنسيق فقط.');
  }
  return addCategoryIfNew(categoryName);
}

/* ============ رفع مرفق (ملف/صورة) لطلب تنسيق ============ */
function uploadRequestAttachment(base64Data, mimeType, filename) {
  const cleanBase64 = base64Data.split(',').pop();
  const bytes = Utilities.base64Decode(cleanBase64);
  const blob = Utilities.newBlob(bytes, mimeType, filename || 'attachment');

  const folderName = 'مرفقات طلبات التنسيق والمتابعة';
  const existingFolders = DriveApp.getFoldersByName(folderName);
  const folder = existingFolders.hasNext() ? existingFolders.next() : DriveApp.createFolder(folderName);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { success: true, url: file.getUrl(), name: filename || file.getName() };
}

function createRequest(reqObj) {
  const user = getCurrentUser();
  const requesterPhone = String(reqObj.phone || '').trim();
  const requesterName = (reqObj.guestName || user.Name || '').trim();
  const requesterEmail = (reqObj.guestEmail || user.Email || '').trim();

  if (!requesterName) {
    return { success: false, message: 'يرجى إدخال الاسم لإتمام تقديم الطلب.' };
  }
  // رقم الهاتف مطلوب بس للضيوف (بدون جلسة Google) — أصحاب حساب Google هويتهم معروفة أصلاً
  if (!user.Email && !requesterPhone) {
    return { success: false, message: 'يرجى إدخال رقم الهاتف لإتمام تقديم الطلب.' };
  }

  const requestId = 'RQ-' + new Date().getTime();

  // تصنيف ذكي عبر AI إذا لم يحدد المستخدم الفئة أو لتأكيدها
  const aiResult = classifyRequestWithAI(reqObj.description, reqObj.category);
  const finalCategory = reqObj.category || aiResult.category;
  addCategoryIfNew(finalCategory); // يحفظ الفئة تلقائيًا لو جديدة عشان تظهر لكل الموظفين لاحقًا
  const assignee = getAssigneeForCategory(finalCategory);

  appendRowByHeaders(SHEET_NAMES.REQUESTS, {
    RequestID: requestId, RequesterEmail: requesterEmail, RequesterName: requesterName, RequesterPhone: requesterPhone,
    Category: finalCategory, Description: reqObj.description, AssignedTo: assignee, Status: 'Pending',
    CreatedAt: new Date(), CompletedAt: '', AINotes: aiResult.notes, Priority: aiResult.priority,
    AttachmentUrl: reqObj.attachmentUrl || '', AttachmentName: reqObj.attachmentName || ''
  });

  // تسجيل/تحديث تلقائي في دليل الهاتف
  saveToDirectory(requesterPhone, requesterName, requesterEmail, user.Department || '');

  if (assignee) {
    notifyUser(assignee, 'طلب تنسيق جديد مسند إليك', `طلب جديد من ${requesterName}\nالفئة: ${finalCategory}\nالوصف: ${reqObj.description}`);
  }

  return { success: true, message: 'تم إرسال الطلب بنجاح.', category: finalCategory, aiNotes: aiResult.notes };
}

function getAssigneeForCategory(category) {
  const cats = getCategories();
  const match = cats.find(c => c.CategoryName === category);
  return match ? match.DefaultAssigneeEmail : '';
}

// يرفق اسم الموظف المسند إليه كل طلب (بدل البريد فقط) — لعرضه بوضوح في الواجهة
function attachAssignedToName(requestsList) {
  const users = getSheetData(SHEET_NAMES.USERS);
  return requestsList.map(r => {
    const match = r.AssignedTo ? users.find(u => u.Email.toLowerCase() === String(r.AssignedTo).toLowerCase()) : null;
    r.AssignedToName = match ? match.Name : (r.AssignedTo || '');
    return r;
  });
}

function getMyRequests(guestEmail, guestPhone) {
  const user = getCurrentUser();
  const email = user.Email || guestEmail || '';
  const phoneNorm = normalizePhone(guestPhone);
  const list = getSheetData(SHEET_NAMES.REQUESTS).filter(r =>
    (email && r.RequesterEmail === email) || (phoneNorm && normalizePhone(r.RequesterPhone) === phoneNorm)
  ).sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
  return attachAssignedToName(list);
}

// حجوزات الموظف الشخصية — لعرضها في صفحة القاعات ("حجوزاتي")
function getMyBookings(guestEmail, guestPhone) {
  const user = getCurrentUser();
  const email = user.Email || guestEmail || '';
  const phoneNorm = normalizePhone(guestPhone);
  const roomsMap = {};
  getSheetData(SHEET_NAMES.ROOMS).forEach(r => { roomsMap[r.RoomID] = r; });

  const list = getSheetData(SHEET_NAMES.BOOKINGS).filter(b =>
    (email && b.RequesterEmail === email) || (phoneNorm && normalizePhone(b.RequesterPhone) === phoneNorm)
  ).sort((a, b) => new Date(b.Date) - new Date(a.Date));

  return list.map(b => {
    const room = roomsMap[b.RoomID];
    b.RoomNameEn = (room && room.RoomNameEn) ? room.RoomNameEn : b.RoomName;
    return b;
  });
}

function getAssignedRequests(staffToken) {
  const user = getCurrentUser(staffToken);
  const list = getSheetData(SHEET_NAMES.REQUESTS).filter(r => r.AssignedTo === user.Email)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
  return attachAssignedToName(list);
}

function getAllRequests(staffToken) {
  requireAdmin(staffToken);
  const list = getSheetData(SHEET_NAMES.REQUESTS).sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
  return attachAssignedToName(list);
}

// تعديل بيانات طلب (الفئة/الوصف) — للأدمن فقط
function updateRequestDetails(requestId, updates, staffToken) {
  requireAdmin(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.REQUESTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('RequestID');
  const categoryCol = headers.indexOf('Category');
  const descriptionCol = headers.indexOf('Description');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === requestId) {
      if (updates.category) sheet.getRange(i + 1, categoryCol + 1).setValue(updates.category);
      if (updates.description) sheet.getRange(i + 1, descriptionCol + 1).setValue(updates.description);
      logAudit('تعديل بيانات طلب', requestId);
      return { success: true };
    }
  }
  return { success: false, message: 'الطلب غير موجود.' };
}

// إضافة/تعديل ملاحظة الأدمن على طلب (منفصلة عن ملاحظة الإنجاز) — تُستخدم لأي غرض إداري
function addAdminNoteToRequest(requestId, note, staffToken) {
  requireAdmin(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.REQUESTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('RequestID');
  let notesCol = headers.indexOf('AdminNote');
  if (notesCol === -1) { addColumnIfMissing(SHEET_NAMES.REQUESTS, 'AdminNote'); notesCol = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('AdminNote'); }
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === requestId) {
      sheet.getRange(i + 1, notesCol + 1).setValue(note);
      logAudit('إضافة ملاحظة أدمن على طلب', requestId);
      return { success: true };
    }
  }
  return { success: false, message: 'الطلب غير موجود.' };
}

// حذف طلب نهائيًا — للأدمن فقط
function deleteRequest(requestId, staffToken) {
  requireAdmin(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.REQUESTS);
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('RequestID');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === requestId) {
      logAudit('حذف طلب', requestId);
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'الطلب غير موجود.' };
}

function updateRequestStatus(requestId, newStatus, completionNote, staffToken) {
  const user = getCurrentUser(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.REQUESTS);
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('RequestID');
  const statusCol = data[0].indexOf('Status');
  const completedCol = data[0].indexOf('CompletedAt');
  const requesterEmailCol = data[0].indexOf('RequesterEmail');
  const categoryCol = data[0].indexOf('Category');
  const descriptionCol = data[0].indexOf('Description');
  const assignedToCol = data[0].indexOf('AssignedTo');
  const notesCol = data[0].indexOf('CompletionNotes');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === requestId) {
      // فقط الأدمن أو موظف التنسيق المُسنَد إليه هذا الطلب تحديداً يمكنه تغيير حالته
      const isAssignee = user.Email && data[i][assignedToCol] && user.Email.toLowerCase() === String(data[i][assignedToCol]).toLowerCase();
      if (user.Role !== 'Admin' && !isAssignee) {
        throw new Error('غير مصرح لك — هذا الطلب مسند لموظف آخر');
      }
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      if (newStatus === 'Completed') {
        sheet.getRange(i + 1, completedCol + 1).setValue(new Date());
        if (completionNote && notesCol > -1) {
          sheet.getRange(i + 1, notesCol + 1).setValue(completionNote);
        }

        if (isEmailNotificationsEnabled() && data[i][requesterEmailCol]) {
          let subject = 'تم إنجاز طلبك';
          let body = `تم الانتهاء من طلبك في فئة "${data[i][categoryCol]}". شكراً لتواصلك مع قسم التنسيق والمتابعة.`;

          const ai = generateEmailContent('request_completed', {
            category: data[i][categoryCol], description: data[i][descriptionCol], completionNote: completionNote || ''
          });
          if (ai) { subject = ai.subject; body = ai.body; }

          notifyUser(data[i][requesterEmailCol], subject, body);
        }
      }
      logAudit('تحديث حالة طلب', `${requestId} → ${newStatus}`);
      return { success: true };
    }
  }
  return { success: false };
}

function reassignRequest(requestId, newAssigneeEmail, staffToken) {
  const user = getCurrentUser(staffToken);
  const sheet = getSS().getSheetByName(SHEET_NAMES.REQUESTS);
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('RequestID');
  const assignedCol = data[0].indexOf('AssignedTo');
  const transferredCol = data[0].indexOf('WasTransferred');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === requestId) {
      // الأدمن أو الموظف المُسنَد إليه الطلب حالياً فقط يقدر يحوّله لزميل آخر
      const isCurrentAssignee = user.Email && data[i][assignedCol] && user.Email.toLowerCase() === String(data[i][assignedCol]).toLowerCase();
      if (user.Role !== 'Admin' && !isCurrentAssignee) {
        throw new Error('غير مصرح لك — هذا الطلب مسند لموظف آخر');
      }
      sheet.getRange(i + 1, assignedCol + 1).setValue(newAssigneeEmail);
      if (transferredCol > -1) sheet.getRange(i + 1, transferredCol + 1).setValue(true);
      notifyUser(newAssigneeEmail, 'تم إسناد طلب جديد إليك', `تم تحويل الطلب رقم ${requestId} إليك من قبل ${user.Name}.`);
      logAudit('تحويل طلب', `${requestId} → ${newAssigneeEmail}`);
      return { success: true };
    }
  }
  return { success: false };
}

// قائمة موظفي التنسيق والمتابعة (للتحويل بينهم) — تُستثنى المستخدم الحالي
function getCoordinatorsList(staffToken) {
  const user = getCurrentUser(staffToken);
  return getSheetData(SHEET_NAMES.USERS)
    .filter(u => (u.Role === 'Coordinator' || u.Role === 'Admin') && u.Status === 'Active' && u.Email.toLowerCase() !== (user.Email || '').toLowerCase())
    .map(u => ({ Email: u.Email, Name: u.Name }));
}

/* ============ الإحصائيات ============ */
function getDashboardStats() {
  const requests = getSheetData(SHEET_NAMES.REQUESTS);
  const bookings = getSheetData(SHEET_NAMES.BOOKINGS);
  return {
    pendingRequests: requests.filter(r => r.Status === 'Pending').length,
    inProgressRequests: requests.filter(r => r.Status === 'InProgress').length,
    completedRequests: requests.filter(r => r.Status === 'Completed').length,
    pendingBookings: bookings.filter(b => b.Status === 'Pending').length,
    approvedBookings: bookings.filter(b => b.Status === 'Approved').length,
    totalRequests: requests.length
  };
}

/* ============ الإشعارات ============ */
function notifyAdmins(subject, body) {
  const admins = getSheetData(SHEET_NAMES.USERS).filter(u => u.Role === 'Admin');
  admins.forEach(a => notifyUser(a.Email, subject, body));
}

function notifyUser(email, subject, body) {
  if (!email) return;
  try {
    MailApp.sendEmail(email, subject, body);
  } catch (err) {
    Logger.log('فشل إرسال البريد: ' + err);
  }
}

/* ============ الذكاء الاصطناعي — Gemini API ============ */
function getGeminiApiKey() {
  const settings = getSheetData(SHEET_NAMES.SETTINGS);
  const row = settings.find(s => s.Key === 'GEMINI_API_KEY');
  return row ? row.Value : '';
}

// حفظ مفتاح Gemini API من لوحة تحكم الأدمن مباشرة — بدل تعديل ورقة Settings يدويًا
function saveGeminiApiKey(apiKey, staffToken) {
  requireAdmin(staffToken);
  if (!apiKey || !apiKey.trim()) return { success: false, message: 'المفتاح فارغ.' };
  setSetting('GEMINI_API_KEY', apiKey.trim());
  logAudit('تحديث مفتاح Gemini API', 'تم تحديث المفتاح');
  return { success: true };
}

// حالة المفتاح الحالي (مُخفى جزئيًا لأمان) — لعرضها في الواجهة بدون كشف المفتاح كاملاً
function getGeminiKeyStatus(staffToken) {
  requireAdmin(staffToken);
  const key = getGeminiApiKey();
  if (!key || key === 'ضع_مفتاح_API_هنا') return { configured: false, masked: '' };
  const masked = key.length > 8 ? key.slice(0, 4) + '••••••••' + key.slice(-4) : '••••••••';
  return { configured: true, masked: masked };
}

// اختبار الاتصال بمفتاح Gemini الحالي — رسالة بسيطة للتأكد إنه شغال فعليًا
function testGeminiConnection(staffToken) {
  requireAdmin(staffToken);
  const result = callGeminiAI('أجب بكلمة واحدة فقط: تم');
  if (result.error) return { success: false, message: result.text };
  return { success: true, message: 'الاتصال شغال ✔ — رد الذكاء الاصطناعي: ' + result.text };
}

// ترتيب الموديلات من الأحدث للأقدم — لو الأول اتقفل من جوجل (بيحصل بشكل متكرر)، يجرّب اللي بعده تلقائيًا
// عيلة 3.x فقط — موديلات 2.5 بقت متقفلة تمامًا للمشاريع/المفاتيح الجديدة من جوجل (حتى قبل تاريخ إغلاقها الرسمي)
const GEMINI_MODELS_FALLBACK = ['gemini-3.5-flash', 'gemini-3.6-flash', 'gemini-3.5-flash-lite'];

// استخراج قوي لـ JSON من رد الذكاء الاصطناعي — يتحمل أي نص زيادة (مقدمة/خاتمة) قد يضيفه النموذج
// حول الـ JSON، ويتحمل أسوار ```json``` بأي شكل، بدل ما يفشل التحليل بالكامل
function extractJsonFromAIResponse(text) {
  if (!text) throw new Error('رد فارغ من الذكاء الاصطناعي');
  let cleaned = normalizeDigits(text).replace(/```json/gi, '').replace(/```/g, '').trim();
  const firstBrace = cleaned.indexOf('{');
  const lastBrace = cleaned.lastIndexOf('}');
  if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
    cleaned = cleaned.substring(firstBrace, lastBrace + 1);
  }
  return JSON.parse(cleaned);
}

function callGeminiAI(prompt, expectJson) {
  const apiKey = getGeminiApiKey();
  if (!apiKey || apiKey === 'ضع_مفتاح_API_هنا') {
    return { error: true, text: 'لم يتم إعداد مفتاح Gemini API بعد.' };
  }

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: expectJson
      ? { temperature: 0.2, maxOutputTokens: 3000, responseMimeType: 'application/json' } // يجبر الموديل يرجّع JSON نضيف فعليًا
      : { temperature: 0.3, maxOutputTokens: 800 }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const allErrors = [];
  for (let m = 0; m < GEMINI_MODELS_FALLBACK.length; m++) {
    const model = GEMINI_MODELS_FALLBACK[m];
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());
      const text = json.candidates && json.candidates[0] && json.candidates[0].content.parts[0].text;
      if (text) return { error: false, text: text };
      const msg = (json.error && json.error.message) ? json.error.message : 'لا يوجد رد صالح';
      allErrors.push(model + ': ' + msg);
    } catch (err) {
      allErrors.push(model + ': ' + String(err));
    }
    // كمّل للموديل التالي في القائمة بدل ما توقف فورًا
  }

  return { error: true, text: 'تعذر الحصول على رد من أي موديل متاح. تفاصيل كل موديل: ' + allErrors.join(' | ') };
}

// توليد محتوى إيميل احترافي بالذكاء الاصطناعي — عنوان ونص ديناميكيين حسب السياق
// يرجع null عند أي فشل (مفتاح غير مُعد، خطأ شبكة، رد غير متوقع) للسماح بالرجوع لرسالة ثابتة بديلة
function generateEmailContent(type, context) {
  let prompt;
  if (type === 'booking_approved') {
    prompt = `اكتب بريداً إلكترونياً قصيراً واحترافياً باللغة العربية لموظف تم اعتماد حجزه لقاعة اجتماعات.
تفاصيل الحجز: القاعة "${context.roomName}"، بتاريخ ${context.date}، من الساعة ${context.startTime} إلى ${context.endTime}، والغرض: "${context.purpose}".
النبرة ودودة ومهنية، والنص لا يتجاوز 4 أسطر. أجب بصيغة JSON فقط بدون أي نص إضافي بهذا الشكل بالضبط:
{"subject":"عنوان مناسب وديناميكي","body":"نص الرسالة"}`;
  } else if (type === 'booking_rejected') {
    prompt = `اكتب بريداً إلكترونياً قصيراً ولبقاً باللغة العربية لموظف تم رفض حجزه لقاعة اجتماعات.
تفاصيل الحجز: القاعة "${context.roomName}"، بتاريخ ${context.date}، من ${context.startTime} إلى ${context.endTime}.
النبرة مهذبة ومهنية بدون اعتذار مبالغ فيه، ولا يتجاوز النص 3 أسطر. أجب بصيغة JSON فقط:
{"subject":"عنوان مناسب","body":"نص الرسالة"}`;
  } else if (type === 'request_completed') {
    prompt = `اكتب بريداً إلكترونياً قصيراً واحترافياً باللغة العربية لموظف تم إنجاز طلبه في قسم التنسيق والمتابعة.
نوع الطلب: "${context.category}"، وتفاصيله: "${context.description}".
${context.completionNote ? `ملاحظة الموظف الذي أنجز الطلب: "${context.completionNote}" — اذكر هذه الملاحظة أو مضمونها بإيجاز في الرسالة لو كانت مفيدة للمستلم.` : ''}
اجعل الرسالة تشير تحديداً لنوع الإنجاز المناسب لهذا الطلب (مثلاً: تم إعداد التصميم، تم صياغة الرسالة، تم تنفيذ المهمة... حسب طبيعة الطلب).
النبرة ودودة ومهنية، ولا يتجاوز النص 4 أسطر. أجب بصيغة JSON فقط بدون أي نص إضافي بهذا الشكل بالضبط:
{"subject":"عنوان مناسب وديناميكي","body":"نص الرسالة"}`;
  } else {
    return null;
  }

  const result = callGeminiAI(prompt, true);
  if (result.error) return null;
  try {
    const parsed = extractJsonFromAIResponse(result.text);
    if (!parsed.subject || !parsed.body) return null;
    return parsed;
  } catch (err) {
    return null;
  }
}


// تصنيف تلقائي لطلبات التنسيق حسب المحتوى
function classifyRequestWithAI(description, userSelectedCategory) {
  const categories = getCategories().map(c => c.CategoryName).join('، ');
  const prompt = `أنت مساعد إداري في مستشفى. صنّف الطلب التالي إلى واحدة من هذه الفئات فقط: ${categories}.
حدد أيضاً مستوى الأولوية (منخفضة/متوسطة/عالية) وأعطِ ملاحظة قصيرة جداً (سطر واحد) للموظف المسؤول.
الطلب: "${description}"
أجب بصيغة JSON فقط بدون أي نص إضافي بهذا الشكل بالضبط:
{"category":"...","priority":"...","notes":"..."}`;

  const result = callGeminiAI(prompt, true);
  if (result.error) {
    return { category: userSelectedCategory || 'أخرى', priority: 'متوسطة', notes: '' };
  }
  try {
    const parsed = extractJsonFromAIResponse(result.text);
    return {
      category: userSelectedCategory || parsed.category || 'أخرى',
      priority: parsed.priority || 'متوسطة',
      notes: parsed.notes || ''
    };
  } catch (err) {
    return { category: userSelectedCategory || 'أخرى', priority: 'متوسطة', notes: '' };
  }
}

// المساعد الذكي — يجيب عن أسئلة توافر القاعات والطلبات
// المساعد الذكي "الفعّال" — يفهم لو المستخدم طالب إنشاء طلب تنسيق ويعمله تلقائيًا،
// أو لو مجرد سؤال عادي يرد عليه بالطريقة المعتادة
function handleAIAssistantMessage(message, phone, guestName, guestEmail, uiLang, staffToken) {
  const isEn = uiLang === 'en';
  // لو فيه جلسة حقيقية (Google لصاحب الموقع، أو دخول فريق العمل بالبريد الشخصي)، هويتك معروفة أصلاً — لا داعي لرقم الهاتف خالص
  const sessionUser = getCurrentUser(staffToken);
  const hasRealIdentity = !!sessionUser.Email;
  const effectivePhone = String(phone || '').trim();
  const effectiveName = hasRealIdentity ? sessionUser.Name : (guestName || '');
  const effectiveEmail = hasRealIdentity ? sessionUser.Email : (guestEmail || '');

  const categories = getCategories().map(c => c.CategoryName).join('، ');
  const rooms = getRooms();
  const roomNames = rooms.map(r => r.RoomName).join('، ');
  const tz = Session.getScriptTimeZone() || 'Asia/Dubai';
  const today = new Date();
  const todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, tz, 'yyyy-MM-dd');
  const langInstruction = isEn ? '\nمهم: اكتب كل الأوصاف (purpose, description) بالإنجليزية لأن المستخدم مستخدم الواجهة بالإنجليزية.' : '';
  const customRules = getAppSettings().assistantCustomRules || '';
  const rulesSection = customRules
    ? `\n\nقواعد وتعليمات إضافية من الإدارة — يجب مراعاتها دائمًا وبأولوية قصوى:\n${customRules}\nلو رسالة المستخدم بتطلب حجز أو طلب يتعارض مع القاعدة دي، صنّف الرسالة كـ "blocked" واشرح السبب بوضوح ولباقة في حقل reason.`
    : '';

  const classifyPrompt = `أنت مساعد ذكي في نظام حجز قاعات وطلبات تنسيق ومتابعة بمستشفى. حلل رسالة المستخدم التالية وحدد نوعها بدقة:
- "book_room": لو المستخدم طالب حجز قاعة اجتماعات (مثل: احجز لي قاعة كذا، عايز أحجز غدًا الساعة كذا).
- "create_request": لو المستخدم طالب تنفيذ عمل يحتاج قسم التنسيق والمتابعة (مثل: اكتب لي رسالة رسمية، صمم لي بوستر، محتاج ميديا).
- "question": لو مجرد سؤال أو استفسار عادي (عن توافر قاعة، حالة طلب سابق، معلومة عامة) بدون طلب حجز فعلي.
- "blocked": لو الطلب يتعارض مع قاعدة إدارية محددة أدناه (إن وُجدت) ويجب رفضه.

معلومات مرجعية مهمة:
- اليوم بالضبط: ${todayStr} — وبكرة/غدًا بالضبط: ${tomorrowStr}
- أسماء القاعات المتاحة بالضبط (استخدم الاسم زي ما هو من القائمة دي فقط): ${roomNames}
- لو الوقت مذكور بعبارة عامة استخدم: "صباحًا"=09:00 إلى 10:00، "الظهر"=13:00 إلى 14:00، "بعد الظهر"=15:00 إلى 16:00، "مساءً"=18:00 إلى 19:00. لو مفيش أي إشارة للوقت، استخدم 10:00 إلى 11:00.
الفئات المتاحة لطلبات التنسيق: ${categories}${langInstruction}${rulesSection}

رسالة المستخدم: "${message}"

أجب بصيغة JSON فقط بدون أي نص إضافي، بأحد الأشكال بالضبط:
{"intent":"book_room","roomName":"الاسم بالضبط من القائمة أعلاه","date":"yyyy-MM-dd","startTime":"HH:MM","endTime":"HH:MM","purpose":"وصف قصير للغرض"}
{"intent":"create_request","category":"الفئة الأنسب من القائمة أعلاه","description":"وصف واضح ومنظم للمطلوب مبني على رسالة المستخدم"}
{"intent":"blocked","reason":"سبب الرفض بوضوح ولباقة"}
{"intent":"question"}`;

  const classifyResult = callGeminiAI(classifyPrompt, true);
  if (!classifyResult.error) {
    try {
      const parsed = extractJsonFromAIResponse(classifyResult.text);

      if (parsed.intent === 'blocked' && parsed.reason) {
        return { type: 'blocked', text: parsed.reason };
      }

      if (parsed.intent === 'book_room' && parsed.roomName) {
        if (!hasRealIdentity && !effectivePhone) {
          return { type: 'need_identity', text: isEn
            ? 'Sure, I\'d love to book this room for you right away 👍 but I need your phone number and first name (from the "New Booking" form, or after you log in) so I can register the booking under your name.'
            : 'تمام، حابب أحجزلك القاعة دي فورًا 👍 بس محتاج تكتب رقم هاتفك واسمك الأول (من نموذج "حجز قاعة جديد" أو بعد ما تسجّل دخولك) عشان أقدر أسجّل الحجز باسمك.' };
        }
        const room = rooms.find(r => r.RoomName === parsed.roomName)
          || rooms.find(r => parsed.roomName && r.RoomName.includes(parsed.roomName))
          || rooms.find(r => parsed.roomName && parsed.roomName.includes(r.RoomName));
        if (!room) {
          return { type: 'error', text: isEn
            ? `Couldn't find a room named exactly "${parsed.roomName}". Available rooms: ${roomNames}`
            : `مش لاقي قاعة اسمها "${parsed.roomName}" بالظبط. القاعات المتاحة عندنا: ${roomNames}` };
        }
        const bookRes = createBooking({
          roomId: room.RoomID, date: parsed.date, startTime: parsed.startTime, endTime: parsed.endTime,
          purpose: parsed.purpose || (isEn ? 'Booked via Smart Assistant' : 'حجز عبر المساعد الذكي'), phone: effectivePhone, guestName: effectiveName, guestEmail: effectiveEmail
        });
        if (bookRes.success) {
          return {
            type: 'booking_created',
            text: isEn
              ? `Done ✔ Requested booking of "${room.RoomName}" on ${parsed.date} from ${parsed.startTime} to ${parsed.endTime}.\n${bookRes.message}`
              : `تمام ✔ طلبت حجز "${room.RoomName}" بتاريخ ${parsed.date} من ${parsed.startTime} إلى ${parsed.endTime}.\n${bookRes.message}`
          };
        }
        return { type: 'error', text: (isEn ? 'I tried to book it but ran into a problem: ' : 'حاولت أحجزلك بس حصلت مشكلة: ') + bookRes.message };
      }

      if (parsed.intent === 'create_request' && parsed.description) {
        if (!hasRealIdentity && !effectivePhone) {
          return { type: 'need_identity', text: isEn
            ? 'Sure, I\'d love to create this request for you right away 👍 but I need your phone number and first name (from the "New Request" form, or after you log in) so I can register it under your name.'
            : 'تمام، حابب أنشئلك الطلب ده فورًا 👍 بس محتاج تكتب رقم هاتفك واسمك الأول (من نموذج "طلب جديد" أو بعد ما تسجّل دخولك) عشان أقدر أسجّله باسمك.' };
        }
        const reqRes = createRequest({
          category: parsed.category, description: parsed.description,
          phone: effectivePhone, guestName: effectiveName, guestEmail: effectiveEmail
        });
        if (reqRes.success) {
          return {
            type: 'request_created',
            text: isEn
              ? `Done ✔ Created a new coordination request in category "${reqRes.category}":\n"${parsed.description}"\n\nYou can track it from the "Coordination Requests" page.`
              : `تمام ✔ أنشأت لك طلب تنسيق جديد في فئة "${reqRes.category}":\n"${parsed.description}"\n\nتقدر تتابع حالته من صفحة "طلبات التنسيق والمتابعة".`
          };
        }
        return { type: 'error', text: (isEn ? 'I tried to create the request but got an error: ' : 'حاولت أنشئ الطلب بس حصل خطأ: ') + (reqRes.message || (isEn ? 'unknown' : 'غير معروف')) };
      }
    } catch (err) {
      // فشل تحليل الرد كـ JSON — كمّل كسؤال عادي بدل ما توقف
    }
  }

  // سؤال عادي — استخدم منطق المساعد المعتاد
  return { type: 'answer', text: askAIAssistant(message, uiLang) };
}

function askAIAssistant(question, uiLang) {
  const rooms = getRooms();
  const bookings = getSheetData(SHEET_NAMES.BOOKINGS).filter(b => b.Status === 'Approved');
  const bookingsSummary = bookings.slice(-40).map(b => `${b.RoomName} | ${b.Date} | ${b.StartTime}-${b.EndTime}`).join('\n');
  const roomsSummary = rooms.map(r => `${r.RoomName} (السعة: ${r.Capacity}, ${r.Location})`).join('\n');
  const langLine = uiLang === 'en' ? 'أجب بإيجاز ووضوح باللغة الإنجليزية فقط (Respond only in English).' : 'أجب بإيجاز ووضوح باللغة العربية.';
  const appSettings = getAppSettings();
  const assistantNameLine = appSettings.assistantName ? `اسمك هو "${appSettings.assistantName}" — عرّف نفسك بيه لو سُئلت.` : '';
  const rulesLine = appSettings.assistantCustomRules ? `\n\nقواعد وتعليمات إضافية من الإدارة — اتبعها دائمًا:\n${appSettings.assistantCustomRules}` : '';

  const prompt = `أنت مساعد ذكي داخل نظام إدارة قاعات الاجتماعات في مستشفى خورفكان. ${langLine} ${assistantNameLine}
القاعات المتاحة:
${roomsSummary}

الحجوزات المعتمدة الحالية (عيّنة):
${bookingsSummary || 'لا توجد حجوزات مسجلة'}${rulesLine}

سؤال المستخدم: "${question}"
أجب بدقة، وإذا كان السؤال عن توافر قاعة في وقت معين، تحقق من قائمة الحجوزات أعلاه قبل الإجابة.`;

  const result = callGeminiAI(prompt);
  return result.text;
}

// ملخص ذكي يومي للأدمن
function getAISummary() {
  const stats = getDashboardStats();
  const prompt = `أنت محلل بيانات إداري. بناءً على الإحصائيات التالية لنظام إدارة قاعات ومتابعة الطلبات في مستشفى، اكتب ملخصاً تنفيذياً موجزاً (3 أسطر كحد أقصى) باللغة العربية مع أي توصية عملية إن وجدت:
طلبات معلقة: ${stats.pendingRequests}
طلبات قيد التنفيذ: ${stats.inProgressRequests}
طلبات مكتملة: ${stats.completedRequests}
حجوزات معلقة الموافقة: ${stats.pendingBookings}
حجوزات معتمدة: ${stats.approvedBookings}`;

  const result = callGeminiAI(prompt);
  return result.text;
}

/* ============ تسجيل الدخول برقم الهاتف وكلمة سر ============
 * كلمة السر لا تُخزَّن أبداً كنص صريح — تُشفَّر بـ SHA-256 مع Salt عشوائي خاص بكل مستخدم. */
function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + '::' + salt);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

// سياسة كلمة مرور قوية: 6 أحرف على الأقل + حرف واحد ورقم واحد على الأقل
function passwordStrengthError(password) {
  password = String(password || '');
  if (password.length < 6) return 'كلمة المرور يجب أن تكون 6 أحرف على الأقل.';
  if (!/[a-zA-Z\u0600-\u06FF]/.test(password)) return 'كلمة المرور يجب أن تحتوي على حرف واحد على الأقل.';
  if (!/[0-9]/.test(password)) return 'كلمة المرور يجب أن تحتوي على رقم واحد على الأقل.';
  return null;
}

function registerAccount(name, phone, email, password) {
  name = String(name || '').trim();
  password = String(password || '');
  const normalized = normalizePhone(phone);

  if (!normalized || normalized.length < 6) return { success: false, message: 'رقم الهاتف غير صالح.' };
  if (!name) return { success: false, message: 'يرجى إدخال الاسم.' };
  const pwError = passwordStrengthError(password);
  if (pwError) return { success: false, message: pwError };

  const sheet = getSS().getSheetByName(SHEET_NAMES.DIRECTORY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneCol = headers.indexOf('Phone');
  const passCol = headers.indexOf('PasswordHash');
  const saltCol = headers.indexOf('Salt');
  const nameCol = headers.indexOf('Name');
  const emailCol = headers.indexOf('Email');
  const updatedCol = headers.indexOf('UpdatedAt');

  const salt = Utilities.getUuid();
  const hash = hashPassword(password, salt);

  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(data[i][phoneCol]) === normalized) {
      // الرقم موجود بالفعل بدون كلمة مرور (من الاستخدام السريع القديم) — نسمح بتفعيل حساب عليه
      if (data[i][passCol]) {
        return { success: false, message: 'هذا الرقم مسجّل بالفعل بحساب. استخدم تسجيل الدخول بدلاً من ذلك.' };
      }
      sheet.getRange(i + 1, nameCol + 1).setValue(name);
      sheet.getRange(i + 1, emailCol + 1).setValue(email || '');
      sheet.getRange(i + 1, passCol + 1).setValue(hash);
      sheet.getRange(i + 1, saltCol + 1).setValue(salt);
      sheet.getRange(i + 1, updatedCol + 1).setValue(new Date());
      return { success: true, user: { Name: name, Phone: phone, Email: email || '' } };
    }
  }

  appendRowByHeaders(SHEET_NAMES.DIRECTORY, {
    Phone: phone, Name: name, Email: email || '', Department: '', UpdatedAt: new Date(),
    PasswordHash: hash, Salt: salt
  });
  return { success: true, user: { Name: name, Phone: phone, Email: email || '' } };
}

function loginAccount(phone, password) {
  const normalized = normalizePhone(phone);
  if (!normalized) return { success: false, message: 'رقم الهاتف غير صالح.' };

  // حماية بسيطة من محاولات الدخول المتكررة: قفل مؤقت 5 دقائق بعد 5 محاولات فاشلة
  const cache = CacheService.getScriptCache();
  const attemptsKey = 'loginAttempts_' + normalized;
  const attempts = Number(cache.get(attemptsKey) || 0);
  if (attempts >= 5) {
    return { success: false, message: 'محاولات دخول كثيرة فاشلة. حاول مرة أخرى بعد 5 دقائق.' };
  }

  const dir = getSheetData(SHEET_NAMES.DIRECTORY);
  const row = dir.find(d => normalizePhone(d.Phone) === normalized);
  if (!row) return { success: false, message: 'لا يوجد حساب بهذا الرقم. أنشئ حساباً جديداً.' };
  if (!row.PasswordHash) return { success: false, message: 'هذا الرقم غير مفعّل بكلمة مرور بعد. اضغط "حساب جديد" لتفعيله.' };

  const hash = hashPassword(String(password || ''), row.Salt);
  if (hash !== row.PasswordHash) {
    cache.put(attemptsKey, String(attempts + 1), 300); // 5 دقائق
    return { success: false, message: 'كلمة المرور غير صحيحة.' };
  }

  cache.remove(attemptsKey);
  return { success: true, user: { Name: row.Name, Phone: row.Phone, Email: row.Email || '' } };
}

/* ============ نسيت كلمة المرور — رمز تحقق عبر البريد الإلكتروني (ذاتي الخدمة) ============ */
function requestPasswordReset(phone) {
  const normalized = normalizePhone(phone);
  if (!normalized) return { success: false, message: 'رقم الهاتف غير صالح.' };

  const dir = getSheetData(SHEET_NAMES.DIRECTORY);
  const row = dir.find(d => normalizePhone(d.Phone) === normalized);
  if (!row) return { success: false, message: 'لا يوجد حساب بهذا الرقم.' };
  if (!row.Email) return { success: false, message: 'لا يوجد بريد إلكتروني مسجّل لهذا الحساب. تواصل مع مدير النظام لإعادة التعيين يدوياً.' };

  const otp = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('pwreset_' + normalized, otp, 600); // صالح 10 دقائق

  try {
    MailApp.sendEmail(row.Email, 'رمز إعادة تعيين كلمة المرور', `مرحباً ${row.Name}،\n\nرمز التحقق الخاص بك لإعادة تعيين كلمة المرور: ${otp}\nصالح لمدة 10 دقائق فقط.\n\nإذا لم تطلب هذا، تجاهل الرسالة.`);
  } catch (err) {
    return { success: false, message: 'تعذر إرسال البريد الإلكتروني. تواصل مع مدير النظام.' };
  }
  // إخفاء جزء من البريد لأسباب خصوصية بسيطة عند عرضه للمستخدم
  const maskedEmail = row.Email.replace(/(.{2}).+(@.+)/, '$1***$2');
  return { success: true, message: 'تم إرسال رمز التحقق إلى ' + maskedEmail };
}

function confirmPasswordReset(phone, otp, newPassword) {
  const normalized = normalizePhone(phone);
  const cached = CacheService.getScriptCache().get('pwreset_' + normalized);
  const normalizedOtp = normalizeDigits(otp).trim();
  if (!cached || cached !== normalizedOtp) {
    return { success: false, message: 'رمز التحقق غير صحيح أو منتهي الصلاحية.' };
  }
  const pwError1 = passwordStrengthError(newPassword);
  if (pwError1) return { success: false, message: pwError1 };

  const sheet = getSS().getSheetByName(SHEET_NAMES.DIRECTORY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneCol = headers.indexOf('Phone');
  const passCol = headers.indexOf('PasswordHash');
  const saltCol = headers.indexOf('Salt');

  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(data[i][phoneCol]) === normalized) {
      const salt = Utilities.getUuid();
      const hash = hashPassword(String(newPassword), salt);
      sheet.getRange(i + 1, passCol + 1).setValue(hash);
      sheet.getRange(i + 1, saltCol + 1).setValue(salt);
      CacheService.getScriptCache().remove('pwreset_' + normalized);
      return { success: true, message: 'تم تغيير كلمة المرور بنجاح. سجّل الدخول بكلمة المرور الجديدة.' };
    }
  }
  return { success: false, message: 'حدث خطأ غير متوقع.' };
}

/* ============ إعادة تعيين كلمة المرور من قبل الأدمن مباشرة (بدون رمز تحقق) ============ */
function adminResetPassword(phone, newPassword, staffToken) {
  requireAdmin(staffToken);
  const pwError2 = passwordStrengthError(newPassword);
  if (pwError2) return { success: false, message: pwError2 };
  const normalized = normalizePhone(phone);
  const sheet = getSS().getSheetByName(SHEET_NAMES.DIRECTORY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneCol = headers.indexOf('Phone');
  const passCol = headers.indexOf('PasswordHash');
  const saltCol = headers.indexOf('Salt');

  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(data[i][phoneCol]) === normalized) {
      const salt = Utilities.getUuid();
      const hash = hashPassword(String(newPassword), salt);
      sheet.getRange(i + 1, passCol + 1).setValue(hash);
      sheet.getRange(i + 1, saltCol + 1).setValue(salt);
      logAudit('إعادة تعيين كلمة مرور (أدمن)', `للهاتف: ${phone}`);
      return { success: true };
    }
  }
  return { success: false, message: 'المستخدم غير موجود.' };
}

// إعادة تعيين كلمة مرور لمستخدم "دور" (أدمن/مسئول قاعة/تنسيق) — بالبحث عن حساب الهاتف المرتبط بنفس بريده.
// لو مفيش حساب هاتف مسجّل لهذا البريد بعد، تُنشئ واحداً جديداً بالرقم المُدخل.
function adminResetPasswordByEmail(email, phone, newPassword, staffToken) {
  requireAdmin(staffToken);
  const pwError = passwordStrengthError(newPassword);
  if (pwError) return { success: false, message: pwError };
  if (!email) return { success: false, message: 'بريد إلكتروني غير صالح.' };

  const sheet = getSS().getSheetByName(SHEET_NAMES.DIRECTORY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const phoneCol = headers.indexOf('Phone');
  const passCol = headers.indexOf('PasswordHash');
  const saltCol = headers.indexOf('Salt');
  const nameCol = headers.indexOf('Name');
  const updatedCol = headers.indexOf('UpdatedAt');

  const salt = Utilities.getUuid();
  const hash = hashPassword(String(newPassword), salt);

  // ابحث عن حساب هاتف موجود بنفس البريد بالفعل
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] && String(data[i][emailCol]).toLowerCase() === String(email).toLowerCase()) {
      sheet.getRange(i + 1, passCol + 1).setValue(hash);
      sheet.getRange(i + 1, saltCol + 1).setValue(salt);
      if (phone) sheet.getRange(i + 1, phoneCol + 1).setValue(phone);
      logAudit('إعادة تعيين كلمة مرور (أدمن)', `للبريد: ${email}`);
      return { success: true, message: 'تم تحديث كلمة المرور بنجاح.' };
    }
  }

  // مفيش حساب هاتف بعد — أنشئ واحداً جديداً (يحتاج رقم هاتف)
  if (!phone) {
    return { success: false, message: 'لا يوجد حساب هاتف مرتبط بهذا البريد بعد. أدخل رقم هاتف لإنشاء حساب جديد له.' };
  }
  const normalized = normalizePhone(phone);
  if (!normalized || normalized.length < 6) return { success: false, message: 'رقم الهاتف غير صالح.' };

  const userRow = getSheetData(SHEET_NAMES.USERS).find(u => u.Email.toLowerCase() === email.toLowerCase());
  appendRowByHeaders(SHEET_NAMES.DIRECTORY, {
    Phone: phone, Name: userRow ? userRow.Name : email, Email: email, Department: userRow ? userRow.Department : '',
    UpdatedAt: new Date(), PasswordHash: hash, Salt: salt
  });
  logAudit('إنشاء حساب هاتف جديد (أدمن)', `للبريد: ${email} — الهاتف: ${phone}`);
  return { success: true, message: 'تم إنشاء حساب هاتف جديد لهذا المستخدم بكلمة المرور المحددة.' };
}


/* ============ عرض كل الموظفين المسجّلين بالهاتف — للأدمن فقط، بدون كشف كلمات المرور أبداً ============ */
function getDirectoryUsers(staffToken) {
  requireAdmin(staffToken);
  return getSheetData(SHEET_NAMES.DIRECTORY)
    .map(d => ({
      Name: d.Name, Phone: d.Phone, Email: d.Email || '',
      HasPassword: !!d.PasswordHash, UpdatedAt: d.UpdatedAt
    }))
    .sort((a, b) => String(a.Name || '').localeCompare(String(b.Name || ''), 'ar'));
}

// حذف حساب هاتف — يحذف بيانات الدخول فقط، حجوزاته/طلباته القديمة تفضل محفوظة بالسجل
function deleteDirectoryUser(phone, staffToken) {
  requireAdmin(staffToken);
  const normalized = normalizePhone(phone);
  const sheet = getSS().getSheetByName(SHEET_NAMES.DIRECTORY);
  const data = sheet.getDataRange().getValues();
  const phoneCol = data[0].indexOf('Phone');
  const nameCol = data[0].indexOf('Name');
  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(data[i][phoneCol]) === normalized) {
      logAudit('حذف حساب هاتف', `${data[i][nameCol]} (${phone})`);
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'الحساب غير موجود.' };
}


function getAppSettings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('app_settings_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (err) { /* تجاهل وأعد القراءة من الشيت */ }
  }
  const settings = getSheetData(SHEET_NAMES.SETTINGS);
  const get = k => { const r = settings.find(s => s.Key === k); return r ? r.Value : ''; };
  const result = {
    appName: get('APP_NAME') || 'بوابة خورفكان الإدارية',
    logoUrl: get('LOGO_URL') || '',
    emailNotificationsEnabled: get('EMAIL_NOTIFICATIONS_ENABLED') !== 'false', // مفعّلة افتراضياً
    ownerEmail: get('OWNER_EMAIL') || '',
    voiceRecorderUrl: get('VOICE_RECORDER_URL') || '',
    assistantName: get('ASSISTANT_NAME') || '',
    assistantCustomRules: get('ASSISTANT_CUSTOM_RULES') || ''
  };
  cache.put('app_settings_v1', JSON.stringify(result), 300); // 5 دقائق
  return result;
}

// تفريغ ذاكرة الإعدادات المؤقتة — تُستدعى تلقائيًا بعد أي تعديل على الإعدادات
// عشان التغيير يظهر فورًا بدل انتظار انتهاء الخمس دقائق
function invalidateSettingsCache() {
  try { CacheService.getScriptCache().remove('app_settings_v1'); } catch (err) { /* تجاهل */ }
}

function isEmailNotificationsEnabled() {
  return getAppSettings().emailNotificationsEnabled;
}

// عرض بريد صاحب الموقع الحالي — للأدمن فقط
function getOwnerEmailForAdmin(staffToken) {
  requireAdmin(staffToken);
  return { ownerEmail: getOwnerEmail() };
}

// تغيير بريد صاحب الموقع — للأدمن فقط (يُستخدم لو الموقع انتقل لشخص آخر)
function saveOwnerEmail(email, staffToken) {
  requireAdmin(staffToken);
  email = String(email || '').trim().toLowerCase();
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { success: false, message: 'بريد إلكتروني غير صالح.' };
  setSetting('OWNER_EMAIL', email);
  logAudit('تغيير بريد صاحب الموقع', email);
  return { success: true };
}

// حفظ رابط صفحة التسجيل الصوتي المستضافة على GitHub Pages (خارج قيد المايك في إطار الويب أب)
function saveVoiceRecorderUrl(url, staffToken) {
  requireAdmin(staffToken);
  url = String(url || '').trim();
  if (url && !/^https:\/\//.test(url)) return { success: false, message: 'الرابط لازم يبدأ بـ https://' };
  setSetting('VOICE_RECORDER_URL', url);
  logAudit('تحديث رابط التسجيل الصوتي', url);
  return { success: true };
}

// حفظ اسم المساعد الذكي المخصص وقواعده الإضافية — للأدمن فقط
// القواعد دي بتتضاف تلقائيًا لكل محادثات المساعد الذكي (تصنيف النية والردود العادية)
function saveAssistantSettings(name, customRules, staffToken) {
  requireAdmin(staffToken);
  setSetting('ASSISTANT_NAME', String(name || '').trim());
  setSetting('ASSISTANT_CUSTOM_RULES', String(customRules || '').trim());
  logAudit('تحديث إعدادات المساعد الذكي', 'الاسم: ' + name);
  return { success: true };
}

function saveEmailNotificationSetting(enabled, staffToken) {
  requireAdmin(staffToken);
  setSetting('EMAIL_NOTIFICATIONS_ENABLED', enabled ? 'true' : 'false');
  logAudit('تعديل إعداد الإشعارات', enabled ? 'تم تفعيل إرسال الإيميلات' : 'تم إيقاف إرسال الإيميلات');
  return { success: true };
}

function setSetting(key, value) {
  const sheet = getSS().getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  const keyCol = data[0].indexOf('Key');
  const valCol = data[0].indexOf('Value');
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) { sheet.getRange(i + 1, valCol + 1).setValue(value); invalidateSettingsCache(); return; }
  }
  sheet.appendRow([key, value]);
  invalidateSettingsCache();
}

function saveAppName(name, staffToken) {
  requireAdmin(staffToken);
  if (!name || !name.trim()) throw new Error('الاسم مطلوب');
  setSetting('APP_NAME', name.trim());
  logAudit('تعديل اسم البرنامج', name.trim());
  return { success: true };
}

// رفع الشعار: يُحفظ في Google Drive (مجلد مخصص) لتفادي حدود حجم خلايا الشيت،
// ويُخزَّن رابط العرض فقط في ورقة Settings
function saveLogo(base64Data, mimeType, staffToken) {
  requireAdmin(staffToken);
  const cleanBase64 = base64Data.split(',').pop();
  const bytes = Utilities.base64Decode(cleanBase64);
  const blob = Utilities.newBlob(bytes, mimeType, 'app-logo');

  const folderName = 'شعارات بوابة خورفكان الإدارية';
  const existingFolders = DriveApp.getFoldersByName(folderName);
  const folder = existingFolders.hasNext() ? existingFolders.next() : DriveApp.createFolder(folderName);

  // حذف الشعارات السابقة للحفاظ على شعار واحد فقط
  const oldFiles = folder.getFiles();
  while (oldFiles.hasNext()) { oldFiles.next().setTrashed(true); }

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  setSetting('LOGO_URL', url);
  logAudit('تحديث الشعار', url);
  return { success: true, url: url };
}

/* ============ أداة مساعدة: هل التاريخ ضمن المدى المحدد؟ ============ */
function inDateRange(rawDate, startDate, endDate) {
  if (!rawDate) return false;
  const d = new Date(rawDate);
  if (startDate) {
    const s = new Date(startDate);
    if (d < s) return false;
  }
  if (endDate) {
    const en = new Date(endDate);
    en.setHours(23, 59, 59, 999);
    if (d > en) return false;
  }
  return true;
}

/* ============ إحصائيات الحجوزات (بمدى تاريخي، مع فلترة قاعات اختيارية) ============ */
function getBookingStats(startDate, endDate, roomIdsFilter) {
  const bookings = getSheetData(SHEET_NAMES.BOOKINGS).filter(b =>
    inDateRange(b.Date, startDate, endDate) && (!roomIdsFilter || roomIdsFilter.includes(b.RoomID))
  );

  const byRoomMap = {};
  const byStatus = { Pending: 0, Approved: 0, Rejected: 0 };
  let totalHours = 0;

  bookings.forEach(b => {
    byRoomMap[b.RoomName] = (byRoomMap[b.RoomName] || 0) + 1;
    if (byStatus[b.Status] !== undefined) byStatus[b.Status]++;
    totalHours += hoursBetween(b.StartTime, b.EndTime);
  });

  return {
    total: bookings.length,
    byStatus: byStatus,
    byRoom: Object.keys(byRoomMap).map(k => ({ room: k, count: byRoomMap[k] })).sort((a, b) => b.count - a.count),
    totalHours: Math.round(totalHours * 10) / 10
  };
}

function hoursBetween(start, end) {
  if (!start || !end) return 0;
  const [sh, sm] = String(start).split(':').map(Number);
  const [eh, em] = String(end).split(':').map(Number);
  return ((eh * 60 + em) - (sh * 60 + sm)) / 60;
}

/* ============ إحصائيات قسم التنسيق والمتابعة — لكل موظف ============ */
function getCoordinationStats(startDate, endDate, staffToken) {
  requireAdmin(staffToken);
  const requests = getSheetData(SHEET_NAMES.REQUESTS).filter(r => inDateRange(r.CreatedAt, startDate, endDate));

  const byEmployeeMap = {};
  const byCategoryMap = {};

  requests.forEach(r => {
    const emp = r.AssignedTo || 'غير مسند';
    if (!byEmployeeMap[emp]) byEmployeeMap[emp] = { employee: emp, total: 0, pending: 0, inProgress: 0, completed: 0, avgDays: [] };
    byEmployeeMap[emp].total++;
    if (r.Status === 'Pending') byEmployeeMap[emp].pending++;
    if (r.Status === 'InProgress') byEmployeeMap[emp].inProgress++;
    if (r.Status === 'Completed') {
      byEmployeeMap[emp].completed++;
      if (r.CompletedAt) {
        const days = (new Date(r.CompletedAt) - new Date(r.CreatedAt)) / (1000 * 60 * 60 * 24);
        byEmployeeMap[emp].avgDays.push(days);
      }
    }
    byCategoryMap[r.Category] = (byCategoryMap[r.Category] || 0) + 1;
  });

  const byEmployee = Object.values(byEmployeeMap).map(e => ({
    employee: e.employee, total: e.total, pending: e.pending, inProgress: e.inProgress, completed: e.completed,
    avgCompletionDays: e.avgDays.length ? Math.round((e.avgDays.reduce((a, b) => a + b, 0) / e.avgDays.length) * 10) / 10 : null
  })).sort((a, b) => b.total - a.total);

  return {
    total: requests.length,
    completed: requests.filter(r => r.Status === 'Completed').length,
    pending: requests.filter(r => r.Status === 'Pending').length,
    inProgress: requests.filter(r => r.Status === 'InProgress').length,
    byEmployee: byEmployee,
    byCategory: Object.keys(byCategoryMap).map(k => ({ category: k, count: byCategoryMap[k] }))
  };
}

// إحصائيات موظف التنسيق لنفسه فقط
function getMyCoordinationStats(startDate, endDate, staffToken) {
  const user = getCurrentUser(staffToken);
  if (!user.Email) throw new Error('غير مصرح');
  const requests = getSheetData(SHEET_NAMES.REQUESTS).filter(r =>
    r.AssignedTo === user.Email && inDateRange(r.CreatedAt, startDate, endDate)
  );
  const byCategoryMap = {};
  requests.forEach(r => { byCategoryMap[r.Category] = (byCategoryMap[r.Category] || 0) + 1; });
  return {
    total: requests.length,
    pending: requests.filter(r => r.Status === 'Pending').length,
    inProgress: requests.filter(r => r.Status === 'InProgress').length,
    completed: requests.filter(r => r.Status === 'Completed').length,
    byCategory: Object.keys(byCategoryMap).map(k => ({ category: k, count: byCategoryMap[k] }))
  };
}

/* ============ التحليل الذكي الشامل بالذكاء الاصطناعي — للأدمن فقط ============
 * يحلل بيانات الحجوزات وطلبات التنسيق معاً: أوقات الذروة، استخدام القاعات،
 * توزيع عبء العمل بين الموظفين، وتوصيات عملية لتحسين الأداء. */
function getBookingTimePatterns(startDate, endDate) {
  const bookings = getSheetData(SHEET_NAMES.BOOKINGS).filter(b => inDateRange(b.Date, startDate, endDate));
  const dayNames = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

  const byHour = {};
  const byDay = {};
  bookings.forEach(b => {
    const hour = String(b.StartTime || '').split(':')[0];
    if (hour) byHour[hour] = (byHour[hour] || 0) + 1;
    if (b.Date) {
      const dayName = dayNames[new Date(b.Date).getDay()];
      byDay[dayName] = (byDay[dayName] || 0) + 1;
    }
  });

  return {
    byHour: Object.keys(byHour).sort().map(h => ({ hour: h + ':00', count: byHour[h] })),
    byDay: dayNames.map(d => ({ day: d, count: byDay[d] || 0 }))
  };
}

function generateAIInsights(startDate, endDate, staffToken) {
  requireAdmin(staffToken);

  const bookings = getSheetData(SHEET_NAMES.BOOKINGS).filter(b => inDateRange(b.Date, startDate, endDate));
  const requests = getSheetData(SHEET_NAMES.REQUESTS).filter(r => inDateRange(r.CreatedAt, startDate, endDate));
  const timePatterns = getBookingTimePatterns(startDate, endDate);

  const roomCounts = {};
  bookings.forEach(b => { roomCounts[b.RoomName] = (roomCounts[b.RoomName] || 0) + 1; });

  const employeeLoad = {};
  requests.forEach(r => {
    const emp = r.AssignedTo || 'غير مسند';
    if (!employeeLoad[emp]) employeeLoad[emp] = { total: 0, completed: 0, pendingOrInProgress: 0 };
    employeeLoad[emp].total++;
    if (r.Status === 'Completed') employeeLoad[emp].completed++;
    else employeeLoad[emp].pendingOrInProgress++;
  });

  if (bookings.length === 0 && requests.length === 0) {
    return { error: true, message: 'لا توجد بيانات كافية في هذه الفترة الزمنية لإجراء تحليل مفيد.' };
  }

  const dataSummary = {
    الفترة: `${startDate || 'البداية'} إلى ${endDate || 'اليوم'}`,
    إجمالي_الحجوزات: bookings.length,
    إجمالي_الطلبات: requests.length,
    الحجوزات_حسب_الساعة: timePatterns.byHour,
    الحجوزات_حسب_يوم_الأسبوع: timePatterns.byDay,
    استخدام_القاعات: roomCounts,
    عبء_عمل_موظفي_التنسيق: employeeLoad
  };

  const prompt = `أنت محلل بيانات إداري خبير في تحسين كفاءة العمل بمستشفى. حلّل بيانات نظام حجز القاعات وقسم التنسيق والمتابعة التالية، وأعطِ تحليلاً عمليًا ومباشرًا بدون مقدمات إنشائية.

البيانات:
${JSON.stringify(dataSummary)}

أجب بصيغة JSON فقط بدون أي نص إضافي بهذا الشكل بالضبط:
{
  "peakTimes": "فقرة من سطرين إلى ثلاثة عن أوقات وأيام الذروة في حجز القاعات، ومتى الحمل يكون أعلى",
  "roomUtilization": "فقرة من سطرين إلى ثلاثة عن القاعات الأكثر والأقل استخدامًا وأي ملاحظة عملية بخصوصها",
  "staffWorkload": "فقرة من سطرين إلى ثلاثة تحلل توزيع عبء العمل بين موظفي التنسيق، وتوضح بوضوح إن كان هناك تفاوت كبير أو موظف محمّل أكثر من غيره",
  "recommendations": ["توصية عملية ومحددة 1", "توصية عملية ومحددة 2", "توصية عملية ومحددة 3", "توصية عملية ومحددة 4"]
}
كل فقرة بالعربية، مباشرة، عملية، ومبنية على الأرقام الفعلية المذكورة أعلاه فقط — لا تخترع بيانات غير موجودة.`;

  const result = callGeminiAI(prompt, true);
  if (result.error) return { error: true, message: result.text };
  try {
    const parsed = extractJsonFromAIResponse(result.text);
    return { error: false, insights: parsed, timePatterns: timePatterns };
  } catch (err) {
    return { error: true, message: 'تعذر تحليل رد الذكاء الاصطناعي (' + err.message + '). حاول مرة أخرى.' };
  }
}

/* ============ نقطة دخول موحّدة لصفحة الإحصائيات — حسب دور المستخدم ============ */
function getMyStatsBundle(startDate, endDate, staffToken) {
  const user = getCurrentUser(staffToken);
  const bundle = { role: user.Role };

  if (user.Role === 'Admin') {
    bundle.bookingStats = getBookingStats(startDate, endDate, null);
    bundle.coordinationStats = getCoordinationStats(startDate, endDate, staffToken);
  } else if (user.Role === 'RoomManager') {
    const roomIds = getMyManagedRoomIds(user.Email);
    bundle.bookingStats = getBookingStats(startDate, endDate, roomIds);
  } else if (user.Role === 'Coordinator') {
    bundle.myCoordinationStats = getMyCoordinationStats(startDate, endDate, staffToken);
  } else {
    throw new Error('غير مصرح لك بعرض الإحصائيات');
  }
  return bundle;
}

/* ============ التقارير القابلة للطباعة/التصدير ============ */
function getBookingsReport(startDate, endDate, staffToken) {
  const user = getCurrentUser(staffToken);
  let roomIds = null;
  if (user.Role === 'RoomManager') {
    roomIds = getMyManagedRoomIds(user.Email);
  } else if (user.Role !== 'Admin') {
    throw new Error('غير مصرح لك بعرض هذا التقرير');
  }
  return getSheetData(SHEET_NAMES.BOOKINGS)
    .filter(b => inDateRange(b.Date, startDate, endDate) && (!roomIds || roomIds.includes(b.RoomID)))
    .sort((a, b) => new Date(a.Date) - new Date(b.Date));
}

function getRequestsReport(startDate, endDate, staffToken) {
  requireAdmin(staffToken);
  return getSheetData(SHEET_NAMES.REQUESTS)
    .filter(r => inDateRange(r.CreatedAt, startDate, endDate))
    .sort((a, b) => new Date(a.CreatedAt) - new Date(b.CreatedAt));
}

/* ============================================================
   النسخ الاحتياطي الأسبوعي التلقائي + التقرير الشهري بالإيميل
   ============================================================
   خطوة تثبيت لمرة واحدة: من محرر Apps Script، شغّل الدالتين التاليتين يدوياً مرة واحدة:
   - setupWeeklyBackupTrigger()
   - setupMonthlyReportTrigger()
   بعدها هيشتغلوا تلقائياً للأبد بدون أي تدخل. */

function setupWeeklyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'weeklyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
  SpreadsheetApp.getUi().alert('تم تفعيل النسخ الاحتياطي الأسبوعي التلقائي (كل يوم أحد الساعة 2 صباحاً).');
}

function setupMonthlyReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendMonthlyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendMonthlyReport').timeBased().onMonthDay(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('تم تفعيل التقرير الشهري التلقائي (أول كل شهر الساعة 7 صباحاً).');
}

// نسخة احتياطية كاملة لملف البيانات في مجلد Drive مخصص — يحذف النسخ الأقدم من 8 أسابيع تلقائياً
function weeklyBackup() {
  try {
    const folderName = 'نسخ احتياطية — بوابة خورفكان الإدارية';
    const existingFolders = DriveApp.getFoldersByName(folderName);
    const folder = existingFolders.hasNext() ? existingFolders.next() : DriveApp.createFolder(folderName);

    const original = DriveApp.getFileById(getSS().getId());
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Dubai', 'yyyy-MM-dd');
    original.makeCopy(`نسخة احتياطية ${stamp}`, folder);

    // حذف النسخ الأقدم من 8 أسابيع لتفادي تراكم الملفات
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 56);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getDateCreated() < cutoff) f.setTrashed(true);
    }
    logAudit('نسخة احتياطية أسبوعية', 'تم إنشاء نسخة تلقائية بنجاح');
  } catch (err) {
    Logger.log('weeklyBackup error: ' + err);
  }
}

// نسخة احتياطية يدوية فورية — يشغّلها الأدمن بضغطة زر من لوحة التحكم، وترجع رابط الملف مباشرة
function manualBackupNow(staffToken) {
  requireAdmin(staffToken);
  const folderName = 'نسخ احتياطية — بوابة خورفكان الإدارية';
  const existingFolders = DriveApp.getFoldersByName(folderName);
  const folder = existingFolders.hasNext() ? existingFolders.next() : DriveApp.createFolder(folderName);

  const original = DriveApp.getFileById(getSS().getId());
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Dubai', "yyyy-MM-dd HH:mm");
  const copy = original.makeCopy(`نسخة احتياطية يدوية ${stamp}`, folder);

  logAudit('نسخة احتياطية يدوية', 'تم إنشاء نسخة فورية بنجاح');
  return { success: true, url: copy.getUrl(), name: copy.getName() };
}

// قائمة آخر النسخ الاحتياطية (تلقائية ويدوية) — للعرض في لوحة التحكم
function getBackupsList(staffToken) {
  requireAdmin(staffToken);
  const folderName = 'نسخ احتياطية — بوابة خورفكان الإدارية';
  const existingFolders = DriveApp.getFoldersByName(folderName);
  if (!existingFolders.hasNext()) return [];

  const folder = existingFolders.next();
  const files = folder.getFiles();
  const list = [];
  while (files.hasNext()) {
    const f = files.next();
    list.push({ name: f.getName(), url: f.getUrl(), date: f.getDateCreated().toISOString() });
  }
  return list.sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 15);
}

// تقرير شهري تلقائي بالإيميل لكل الأدمن — إحصائيات الشهر الماضي كاملة
function sendMonthlyReport() {
  try {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const end = new Date(now.getFullYear(), now.getMonth(), 0);
    const startStr = Utilities.formatDate(start, Session.getScriptTimeZone() || 'Asia/Dubai', 'yyyy-MM-dd');
    const endStr = Utilities.formatDate(end, Session.getScriptTimeZone() || 'Asia/Dubai', 'yyyy-MM-dd');
    const monthLabel = Utilities.formatDate(start, Session.getScriptTimeZone() || 'Asia/Dubai', 'MMMM yyyy');

    const bookingStats = getBookingStats(startStr, endStr, null);
    const coordStats = getCoordinationStats(startStr, endStr);
    const appSettings = getAppSettings();

    const html = `
      <div style="font-family:Arial,sans-serif;direction:rtl;text-align:right;max-width:600px;margin:auto;">
        <h2 style="color:#0f766e;">📊 التقرير الشهري — ${appSettings.appName}</h2>
        <p style="color:#666;">الفترة: ${startStr} إلى ${endStr}</p>
        <h3 style="color:#0f766e;">القاعات والحجوزات</h3>
        <ul>
          <li>إجمالي الحجوزات: <b>${bookingStats.total}</b></li>
          <li>معتمدة: <b>${bookingStats.byStatus.Approved || 0}</b> — مرفوضة: <b>${bookingStats.byStatus.Rejected || 0}</b></li>
          <li>إجمالي ساعات الاستخدام: <b>${bookingStats.totalHours}</b> ساعة</li>
          <li>الأكثر حجزاً: <b>${bookingStats.byRoom[0] ? bookingStats.byRoom[0].room : '—'}</b></li>
        </ul>
        <h3 style="color:#0f766e;">قسم التنسيق والمتابعة</h3>
        <ul>
          <li>إجمالي الطلبات: <b>${coordStats.total}</b></li>
          <li>مكتملة: <b>${coordStats.completed}</b> — قيد التنفيذ: <b>${coordStats.inProgress}</b> — معلّقة: <b>${coordStats.pending}</b></li>
        </ul>
        <p style="color:#999;font-size:12px;">تقرير تلقائي — لمزيد من التفاصيل افتح صفحة الإحصائيات والتقارير في النظام.</p>
      </div>`;

    const admins = getSheetData(SHEET_NAMES.USERS).filter(u => u.Role === 'Admin' && u.Status === 'Active');
    admins.forEach(a => {
      try { MailApp.sendEmail({ to: a.Email, subject: `التقرير الشهري — ${monthLabel}`, htmlBody: html }); } catch (e) { /* تجاهل */ }
    });
    logAudit('تقرير شهري تلقائي', `تم الإرسال لـ ${admins.length} أدمن`);
  } catch (err) {
    Logger.log('sendMonthlyReport error: ' + err);
  }
}
