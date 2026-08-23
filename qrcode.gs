/**
 * ============================================================
 * سكريبت "استخراج كود الطالب" — لصفحة الطلاب اللي هتترفع على GitHub
 * ============================================================
 * ده سكريبت مستقل تمامًا (Standalone) — مش هيتحط جوه شيت "قاعدة البيانات"
 * ومش هيلمس أو يعدل أي حاجة في السكريبت التاني اللي شغال على شيت الحضور،
 * فمفيش أي تعارض بينهم. الاتنين بيقروا بس (Read) من نفس شيت الداتا بيز،
 * وده حاجة عادية وآمنة جوجل شيت بتسمح بيها من غير أي مشكلة.
 *
 * ليه سكريبت منفصل ومش إضافة على القديم؟
 * - الشيت الواحد مينفعش يبقى ليه غير سكريبت "مربوط" (Bound) واحد بس، وهو
 *   أصلاً محجوز للسكريبت التاني بتاع الحضور.
 * - سكريبت مستقل زي ده معناه صفر تعارض: أي تعديل أو نشر جديد هنا مش
 *   هيأثر ولا هيلمس السكريبت التاني خالص.
 *
 * ============================================================
 * الخطوات:
 * 1) روح https://script.google.com > مشروع جديد (New project)
 * 2) امسح أي كود موجود، والصق هذا الكود بالكامل
 * 3) من فوق: Deploy > New deployment > اختر Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    ثم Deploy، سيبك أول مرة هيطلب يوافق على الصلاحيات (عادي، دوس Allow)
 * 4) انسخ رابط الـ Web app (Web app URL) وحطه في CONFIG.API_URL
 *    جوه ملف student-code-lookup.html
 * 5) في كل مرة تعدّل فيها الكود بعد كده، لازم:
 *    Deploy > Manage deployments > ✏️ (تعديل) > Version: New version > Deploy
 *    (غير كده اللينك القديم هيفضل شغال بالكود القديم)
 * ============================================================
 *
 * ملاحظة أمان مهمة جدًا:
 * الدالة دي بترجع بس اسم الطالب وكوده لما تلاقي تطابق مضبوط مع الكود أو
 * رقم الهاتف اللي المستخدم كتبه. هي بتاخد المدخل ده وتقارنه سطر سطر جوه
 * جوجل (Server-side)، ومفيش أي نقطة في الكود بترجع القائمة كاملة أو رقم
 * تليفون أي حد للمتصفح. حتى لو حد فتح الصفحة وقرا الكود من جواها أو من
 * الـ Network tab، أقصى حاجة هيشوفها هي رد بيانات الطالب اللي هو بنفسه
 * كتب كوده أو رقمه، مش أي طالب تاني.
 */

// ------------------------- الإعدادات -------------------------
const DB_SHEET_ID = "1TPFayp29KDfz9XzUqYxBZT04S5yD7y1YytAbFk1V6NI"; // آي دي شيت قاعدة البيانات
const DB_TAB_NAME = "Database";                                    // اسم التبويبة اللي فيها بيانات الطلاب

// أسماء الأعمدة زي ما هي في الشيت بالظبط (نفس أسماء السكريبت التاني عشان التوافق)
const COL_NAME = "اسم";
const COL_CODE = "الكود";
const COL_ID_FALLBACK = "الرقم التعريفى"; // لو عمود "الكود" فاضي، هناخد الرقم التعريفي بدل منه
const COL_PHONE = "الهاتف المحمول";

// ------------------------- نقطة الدخول -------------------------
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === "lookup") {
      const query = String(e.parameter.query || "").trim();
      return jsonOutput(lookupStudent(query));
    }
    return jsonOutput({ found: false, error: "طلب غير معروف" });
  } catch (err) {
    return jsonOutput({ found: false, error: "حصل خطأ، حاول تاني" });
  }
}

// ------------------------- البحث عن طالب واحد بس -------------------------
// بيدور على تطابق مضبوط (مش جزئي) إما مع "الكود" أو مع "رقم الهاتف" بعد
// توحيد شكله. أول ما يلاقي تطابق بيرجّعه ويوقف فورًا — من غير ما يمرّ على
// باقي الشيت ومن غير ما يبني أي قائمة كاملة في الذاكرة يرجعها للمتصفح.
function lookupStudent(query) {
  if (!query) return { found: false };

  const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(DB_TAB_NAME);
  if (!sheet) return { found: false, error: "التبويبة غير موجودة" };

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values);
  const headers = values[headerRowIdx].map(v => String(v).trim());

  const idxName = headers.indexOf(COL_NAME);
  const idxCode = headers.indexOf(COL_CODE);
  const idxIdFallback = headers.indexOf(COL_ID_FALLBACK);
  const idxPhone = headers.indexOf(COL_PHONE);

  const queryCodeLower = query.toLowerCase();
  const queryPhoneKey = normalizePhoneKey(query);

  for (let i = headerRowIdx + 1; i < values.length; i++) {
    const row = values[i];
    const name = idxName > -1 ? String(row[idxName] || "").trim() : "";
    if (!name) continue; // صف فاضي، تجاهله

    let code = idxCode > -1 ? String(row[idxCode] || "").trim() : "";
    if (!code && idxIdFallback > -1) {
      code = String(row[idxIdFallback] || "").trim();
    }
    const phone = idxPhone > -1 ? String(row[idxPhone] || "").trim() : "";

    const codeMatches = code && code.toLowerCase() === queryCodeLower;
    const phoneMatches = phone && queryPhoneKey && normalizePhoneKey(phone) === queryPhoneKey;

    if (codeMatches || phoneMatches) {
      if (!code) return { found: false }; // مالوش كود لسه (طالب جديد ماسجلش كود)
      // مهم: برجّع بس الاسم والكود. رقم الهاتف وأي عمود تاني في الشيت
      // (زي أعمدة السنتر أو الدرجات لو موجودة) بيفضل جوه جوجل ومبيتبعتش
      // للمتصفح خالص.
      return { found: true, name: name, code: code };
    }
  }
  return { found: false };
}

// ------------------------- إيجاد صف العناوين الحقيقي -------------------------
function findHeaderRowIndex(values) {
  const limit = Math.min(5, values.length);
  for (let r = 0; r < limit; r++) {
    const row = values[r].map(v => String(v).trim());
    if (row.indexOf(COL_NAME) > -1) return r;
  }
  return 0;
}

// ------------------------- توحيد شكل رقم الهاتف للمقارنة -------------------------
// نفس منطق السكريبت التاني بالظبط عشان التوافق: بنشيل أي حاجة مش رقم، وبعدين
// نشيل كود الدولة (٢٠) أو الصفر الأول لو موجودين.
function normalizePhoneKey(phone) {
  return String(phone || "").replace(/[^0-9]/g, "").replace(/^20/, "").replace(/^0/, "");
}

// ------------------------- إخراج JSON -------------------------
function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}