/**
 * ==============================================================================
 * حاضنة المبادرات النوعية للمعلمات - الثانوية الرابعة بالجبيل 🏫
 * كود Google Apps Script المطور والآمن (الخادم البرمجي وقاعدة البيانات المركزية)
 * مديرة المدرسة: أ. فريدة بنت سعود العتيبي
 * ==============================================================================
 * 
 * الميزات والتحصينات الأمنية المحدثة:
 * 1. مصادقة إلزامية برمز API Key لمنع أي وصول أو تلاعب خارجي غير مصرح به.
 * 2. عمليات القراءة والكتابة (CRUD) الكاملة للمبادرات وروابط الشواهد مع مزامنة حية.
 * 3. حماية متقدمة ضد هجمات حقن المعادلات (Formula Injection).
 * 4. حجب وتطهير أي بيانات حساسة أو كلمات مرور من الردود العامة.
 * 5. فحص أمني دقيق لحجم ونوع الملفات المرفوعة إلى Google Drive.
 */

// مفتاح الوصول السري المعتمد لواجهة برمجة التطبيقات (API Key)
const DEFAULT_API_KEY = "SEC4_JUBAIL_API_KEY_2026";

function getApiKey() {
  const prop = PropertiesService.getScriptProperties().getProperty("API_KEY");
  return prop || DEFAULT_API_KEY;
}

// اسم مجلد الشواهد في Google Drive
const EVIDENCE_FOLDER_NAME = "شواهد_مبادرات_الثانوية_الرابعة_بالجبيل";

/**
 * دالة التحقق الأمني من الـ API Key للطلبات الواردة
 */
function isAuthorized(providedKey) {
  if (!providedKey) return false;
  return String(providedKey).trim() === getApiKey();
}

/**
 * معالجة طلبات GET (جلب المبادرات أو التحقق من الربط مع التحقق الأمني)
 */
function doGet(e) {
  try {
    const key = (e.parameter && (e.parameter.key || e.parameter.apiKey)) || "";
    
    // التحقق الأمني من مفتاح الـ API
    if (!isAuthorized(key)) {
      return createJsonResponse({
        status: "error",
        code: "UNAUTHORIZED",
        message: "عذراً، رمز الوصول السري للواجهة البرمجية غير صحيح أو مفقود (Unauthorized API Key)"
      });
    }

    const action = (e.parameter && e.parameter.action) || "GET_ALL";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheetsIfNeeded(ss);

    if (action === "GET_ALL") {
      const sheet = ss.getSheetByName("Initiatives");
      const data = sheet.getDataRange().getValues();
      let initiatives = [];
      
      if (data.length > 1) {
        const headers = data[0];
        const rows = data.slice(1);
        initiatives = rows.map(row => {
          const item = {};
          headers.forEach((h, idx) => {
            item[h] = row[idx];
          });
          return item;
        });
      }

      // جلب المعلمات مع استبعاد وتطهير أي حقول حساسة
      let teachers = [];
      const usersSheet = ss.getSheetByName("Users");
      if (usersSheet) {
        const uData = usersSheet.getDataRange().getValues();
        if (uData.length > 1) {
          const uHeaders = uData[0];
          teachers = uData.slice(1).map(row => {
            const t = {};
            uHeaders.forEach((h, idx) => {
              // استبعاد كلمات المرور من الاستجابة الخارجية تماماً
              if (h !== "password" && h !== "passwordHash" && h !== "salt") {
                t[h] = row[idx];
              }
            });
            return t;
          });
        }
      }

      return createJsonResponse({
        status: "success",
        school: "الثانوية الرابعة بالجبيل",
        principal: "أ. فريدة بنت سعود العتيبي",
        initiatives: initiatives,
        teachers: teachers,
        syncedAt: Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss")
      });
    }

    return createJsonResponse({ 
      status: "success", 
      message: "Google Apps Script API متصل ويعمل بنجاح ومحمي برمز المرور السري" 
    });

  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * معالجة طلبات POST (إضافة مبادرة، تحديث الحالة، مزامنة شاملة، رفع الملفات)
 */
function doPost(e) {
  try {
    let body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    const key = body.key || body.apiKey || (e.parameter && e.parameter.key) || "";

    // التحقق الأمني الإلزامي من مفتاح الـ API
    if (!isAuthorized(key)) {
      return createJsonResponse({
        status: "error",
        code: "UNAUTHORIZED",
        message: "رفض الطلب: يتطلب تنفيذ هذه العملية تزويد رمز المصادقة المعتمد (Unauthorized API Key)"
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheetsIfNeeded(ss);

    const action = body.action || "CREATE_INITIATIVE";
    const sheet = ss.getSheetByName("Initiatives");

    // 1. إضافة مبادرة جديدة مع تعقيم ضد حقن المعادلات
    if (action === "CREATE_INITIATIVE") {
      const init = body.data || {};
      const newRow = [
        sanitizeCell(init.id || "INIT-" + Utilities.formatDate(new Date(), "GMT+3", "yyyyMMdd-HHmm")),
        sanitizeCell(init.title || ""),
        sanitizeCell(init.teacherName || ""),
        sanitizeCell(init.teacherUsername || ""),
        sanitizeCell(init.category || ""),
        sanitizeCell(init.target || ""),
        Number(init.beneficiaries) || 0,
        sanitizeCell(init.startDate || ""),
        sanitizeCell(init.endDate || ""),
        sanitizeCell(init.problem || ""),
        sanitizeCell(init.goal || ""),
        sanitizeCell(init.plan || ""),
        sanitizeCell(init.metric || ""),
        sanitizeCell(init.budget || ""),
        sanitizeCell(init.status || "قيد المراجعة"),
        Number(init.progress) || 0,
        sanitizeCell(init.supervisorNotes || ""),
        sanitizeCell(init.evidenceLink || ""),
        sanitizeCell(init.createdAt || Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd"))
      ];

      sheet.appendRow(newRow);
      return createJsonResponse({ status: "success", message: "تم تسجيل المبادرة في Google Sheets بنجاح", id: newRow[0] });
    }

    // 2. تحديث حالة المبادرة وملاحظات المشرفة
    if (action === "UPDATE_STATUS") {
      const initId = body.id;
      const newStatus = body.status;
      const notes = body.supervisorNotes;
      const progress = body.progress;

      const data = sheet.getDataRange().getValues();
      let updated = false;

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == initId) {
          if (newStatus !== undefined) sheet.getRange(i + 1, 15).setValue(sanitizeCell(newStatus));
          if (progress !== undefined) sheet.getRange(i + 1, 16).setValue(Number(progress) || 0);
          if (notes !== undefined) sheet.getRange(i + 1, 17).setValue(sanitizeCell(notes));
          updated = true;
          break;
        }
      }

      if (updated) {
        return createJsonResponse({ status: "success", message: "تم تحديث المبادرة بنجاح" });
      } else {
        return createJsonResponse({ status: "not_found", message: "لم يتم العثور على كود المبادرة" });
      }
    }

    // 3. مزامنة شاملة لكافة المبادرات
    if (action === "UPDATE_ALL") {
      const list = body.data;
      if (Array.isArray(list)) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
        }

        const rowsToAdd = list.map(init => [
          sanitizeCell(init.id || ""),
          sanitizeCell(init.title || ""),
          sanitizeCell(init.teacherName || ""),
          sanitizeCell(init.teacherUsername || ""),
          sanitizeCell(init.category || ""),
          sanitizeCell(init.target || ""),
          Number(init.beneficiaries) || 0,
          sanitizeCell(init.startDate || ""),
          sanitizeCell(init.endDate || ""),
          sanitizeCell(init.problem || ""),
          sanitizeCell(init.goal || ""),
          sanitizeCell(init.plan || ""),
          sanitizeCell(init.metric || ""),
          sanitizeCell(init.budget || ""),
          sanitizeCell(init.status || "قيد المراجعة"),
          Number(init.progress) || 0,
          sanitizeCell(init.supervisorNotes || ""),
          sanitizeCell(init.evidenceLink || ""),
          sanitizeCell(init.createdAt || Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd"))
        ]);

        if (rowsToAdd.length > 0) {
          sheet.getRange(2, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
        }

        return createJsonResponse({ status: "success", message: "تمت المزامنة الشاملة مع Google Sheets بنجاح" });
      }
    }

    // 4. رفع ملف أو شاهد إلى Google Drive
    if (action === "UPLOAD_FILE") {
      const fileData = body.fileData; // Base64
      const fileName = body.fileName || "شاهد_مبادرة_" + Utilities.formatDate(new Date(), "GMT+3", "yyyyMMdd");
      const mimeType = body.mimeType || "application/pdf";

      const allowedMimes = [
        "application/pdf",
        "image/png",
        "image/jpeg",
        "image/webp",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/msword"
      ];

      if (!allowedMimes.includes(mimeType)) {
        return createJsonResponse({ status: "error", message: "نوع الملف غير مصرح به أمنياً. يرجى رفع ملفات PDF أو صور أو Word فقط." });
      }

      const folder = getOrCreateEvidenceFolder();
      const decodedBytes = Utilities.base64Decode(fileData.split(',')[1] || fileData);
      
      if (decodedBytes.length > 15 * 1024 * 1024) {
        return createJsonResponse({ status: "error", message: "حجم الملف يتجاوز الحد المسموح (15 ميغابايت)." });
      }

      const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      return createJsonResponse({
        status: "success",
        fileUrl: file.getUrl(),
        downloadUrl: file.getDownloadUrl(),
        message: "تم رفع الملف إلى Google Drive بنجاح"
      });
    }

    return createJsonResponse({ status: "error", message: "إجراء غير معروف" });

  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * تهيئة الأوراق والعناوين تلقائياً إذا لم تكن موجودة
 */
function initSheetsIfNeeded(ss) {
  let sheet = ss.getSheetByName("Initiatives");
  if (!sheet) {
    sheet = ss.insertSheet("Initiatives");
    const headers = [
      "id",
      "title",
      "teacherName",
      "teacherUsername",
      "category",
      "target",
      "beneficiaries",
      "startDate",
      "endDate",
      "problem",
      "goal",
      "plan",
      "metric",
      "budget",
      "status",
      "progress",
      "supervisorNotes",
      "evidenceLink",
      "createdAt"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#15445A").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }

  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    const userHeaders = ["id", "name", "username", "specialty", "role"];
    usersSheet.appendRow(userHeaders);
    usersSheet.getRange(1, 1, 1, userHeaders.length).setFontWeight("bold").setBackground("#0DA9A6").setFontColor("#FFFFFF");
    
    usersSheet.appendRow(["T-101", "أ. سارة راشد الحربي", "sara_h", "حاسب آلي", "teacher"]);
    usersSheet.appendRow(["T-102", "أ. نورة سعد الخالدي", "noura_k", "علوم وبيئة", "teacher"]);
    usersSheet.appendRow(["T-103", "أ. منى فهد الدوسري", "mona_d", "توجيه وإرشاد", "teacher"]);
    usersSheet.appendRow(["ADM-01", "أ. فريدة بنت سعود العتيبي", "admin_principal", "مديرة المدرسة", "admin"]);
    usersSheet.setFrozenRows(1);
  }
}

function getOrCreateEvidenceFolder() {
  const folders = DriveApp.getFoldersByName(EVIDENCE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(EVIDENCE_FOLDER_NAME);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sanitizeCell(val) {
  if (val === null || val === undefined) return "";
  let s = String(val);
  if (/^[=\+\-@]/.test(s)) {
    return "'" + s;
  }
  return s;
}

/**
 * جلب أو إنشاء مجلد الشواهد في Google Drive
 */
function getOrCreateEvidenceFolder() {
  const folders = DriveApp.getFoldersByName(EVIDENCE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(EVIDENCE_FOLDER_NAME);
}

/**
 * تنسيق رد JSON مع ترويسات الدعم
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * حماية خلايا Google Sheets من هجمات حقن الصيغ والمعادلات الخبيثة (Formula Injection Defense)
 */
function sanitizeCell(val) {
  if (val === null || val === undefined) return "";
  let s = String(val);
  // إذا كانت القيمة تبدأ برمز معادلة، يتم إبطالها بإضافة فاصلة عليا مفردة
  if (/^[=\+\-@]/.test(s)) {
    return "'" + s;
  }
  return s;
}
