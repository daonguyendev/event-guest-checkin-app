/**
 * Nhóm khách → Màu (hex)
 */
const GROUP_COLORS = {
  "🔴 Ban tổ chức / Organizer":    "#ff0000",
  "🟡 Keynote / Keynote Speaker":  "#ffd700",
  "🟠 Báo cáo viên / Presenter":   "#ffa500",
  "🔵 Khách tham dự / Participant":"#1e90ff"
};

const SHEET_ID   = "1itkIj1DO5VFk8R6lQPeCDFl1VFf1rCGmb7AyJLmcEfk"; // Google Sheet của bạn
const SHEET_NAME = "Checkin";
const TZ         = "Asia/Ho_Chi_Minh";

// Bật email thông báo (nếu bạn đã dùng trước đó)
const ENABLE_EMAIL = true;
const EMAIL_SUBJECT_PREFIX = "[Check-in] ";
const EMAIL_SENDER_NAME = "Event Team";
const EMAIL_REPLY_TO = "";

// Cấu trúc cột (1-based)
const COL = {
  ID: 1,            // A
  FULLNAME: 2,      // B
  JOB: 3,           // C
  ORG: 4,           // D
  GROUP: 5,         // E
  GROUP_COLOR: 6,   // F (tô nền)
  EMAIL: 7,         // G
  PHONE: 8,         // H
  TIME: 9           // I
};

// Quy tắc ID tự tăng
const ID_PREFIX = "G";
const ID_PAD = 3;  // G001

/** Lấy sheet + header chuẩn */
function getDataSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 9).setValues([[
      "ID",
      "Họ và tên",
      "Công việc",
      "Đơn vị công tác",
      "Nhóm khách",
      "Màu nhóm khách",
      "Email",
      "Số điện thoại",
      "Thời gian check-in"
    ]]);
  }
  return sheet;
}

/** Giao diện Web App */
function doGet() {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.groupColors = GROUP_COLORS;
  return tpl.evaluate()
    .setTitle('Check-in')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Check-in (nhập tay)
 * payload: { fullName, job?, org, group, email, phone? }
 * - Validate cơ bản (job/phone là tùy chọn)
 * - Chống duplicate theo Email hoặc Phone (nếu có phone)
 * - Tạo ID tự tăng
 * - Ghi dòng + tô màu cột GROUP_COLOR
 * - Phone lưu dạng chuỗi để giữ số 0 đầu (nếu có)
 * - Gửi email xác nhận (nếu ENABLE_EMAIL = true)
 */
function checkIn(payload) {
  try {
    const sheet = getDataSheet_();
    const { fullName, job, org, group, email, phone } = normalizePayloadManual_(payload);

    // ======== Backend validation (job/phone: tùy chọn) ========
    if (!fullName) throw new Error('Vui lòng nhập "Họ và tên".');
    if (!org)      throw new Error('Vui lòng nhập "Đơn vị công tác".');
    if (!group)    throw new Error('Vui lòng chọn "Nhóm khách".');
    if (!email)    throw new Error('Vui lòng nhập "Email".');
    if (!isValidEmail_(email)) throw new Error('Email không hợp lệ.');

    // Phone chỉ chuẩn hoá khi có nhập
    let normalizedPhone = "";
    if (phone) normalizedPhone = normalizePhone_(phone); // throws nếu sai
    // =========================================================

    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);

    // Duplicate theo email hoặc (phone nếu có)
    const dup = findDuplicateByEmailOrPhone_(sheet, email, normalizedPhone);
    if (dup) {
      const prevId   = sheet.getRange(dup.row, COL.ID).getDisplayValue();
      const prevTime = sheet.getRange(dup.row, COL.TIME).getDisplayValue();
      lock.releaseLock();
      throw new Error(`Email/Phone đã check-in trước đó (ID ${prevId}${prevTime ? ", " + prevTime : ""}). Không thể check-in trùng.`);
    }

    const id    = getNextId_(sheet);
    const color = GROUP_COLORS[group] || "#dddddd";
    const now   = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");

    const phoneStr = normalizedPhone ? "'" + normalizedPhone : "";

    sheet.appendRow([
      id,
      fullName,
      job || "",
      org,
      group,
      "",
      email,
      phoneStr,
      now
    ]);

    // Tô màu cột "Màu nhóm khách"
    const row = sheet.getLastRow();
    sheet.getRange(row, COL.GROUP_COLOR).setBackground(color);

    // Email xác nhận (tuỳ chọn)
    if (ENABLE_EMAIL) {
      try {
        sendConfirmationEmail_(email, {
          id, fullName, job, org, group,
          phone: normalizedPhone, time: now
        });
      } catch (mailErr) {
        lock.releaseLock();
        return makeResult_(true,
          `Check-in thành công: ${fullName} (${shortGroupName_(group)}) — *Lưu ý:* gửi email thất bại: ${mailErr.message || mailErr}`,
          group, now, id);
      }
    }

    lock.releaseLock();

    // Thông điệp + màu theo Nhóm khách
    return makeResult_(true, `Check-in thành công: ${fullName} (${shortGroupName_(group)})`, group, now, id);

  } catch (err) {
    return makeResult_(false, err.message || String(err));
  }
}

/** Tên nhóm gọn: lấy phần Việt trước dấu '/' và bỏ emoji đầu dòng */
function shortGroupName_(group) {
  if (!group) return "";
  let s = String(group).split('/')[0]; // trước dấu '/'
  s = s.replace(/^[^\wÀ-ỹ]+/u, '').trim(); // bỏ emoji/ký tự đầu
  return s;
}

/** Gửi email xác nhận check-in (nếu bật) */
function sendConfirmationEmail_(to, info) {
  const subject = `${EMAIL_SUBJECT_PREFIX}${info.fullName} (ID ${info.id})`;
  const clr = GROUP_COLORS[info.group] || "#148a39";
  const html = `
    <div style="font-family:system-ui,Segoe UI,Arial;line-height:1.55;color:#111">
      <div style="background:${clr};color:#fff;border-radius:10px;padding:14px 16px;font-weight:700">
        Check-in thành công: ${_esc(info.fullName)} (${_esc(shortGroupName_(info.group))})
      </div>
      <div style="padding:12px 6px 0">
        <p>Thông tin của bạn đã được ghi nhận.</p>
        <ul>
          <li><b>ID:</b> ${_esc(info.id)}</li>
          <li><b>Công việc:</b> ${_esc(info.job || '')}</li>
          <li><b>Đơn vị:</b> ${_esc(info.org)}</li>
          <li><b>Nhóm:</b> ${_esc(info.group)}</li>
          ${info.phone ? `<li><b>Điện thoại:</b> ${_esc(info.phone)}</li>` : ``}
          <li><b>Thời gian:</b> ${_esc(info.time)}</li>
        </ul>
      </div>
    </div>`;
  const opt = { to, subject, htmlBody: html, name: EMAIL_SENDER_NAME || undefined };
  if (EMAIL_REPLY_TO) opt.replyTo = EMAIL_REPLY_TO;
  MailApp.sendEmail(opt);
}

/** Tìm dòng trùng theo Email hoặc Phone (phone chỉ check nếu có) */
function findDuplicateByEmailOrPhone_(sheet, email, phoneNormalized) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const emailVals = sheet.getRange(2, COL.EMAIL, lastRow - 1, 1).getValues();
  const phoneVals = sheet.getRange(2, COL.PHONE, lastRow - 1, 1).getValues();

  const emailLower = String(email || "").trim().toLowerCase();

  for (let i = 0; i < emailVals.length; i++) {
    const rowEmail = String(emailVals[i][0] || "").trim().toLowerCase();
    const rowPhone = safeNormalizePhone_(String(phoneVals[i][0] || ""));

    if (rowEmail && emailLower && rowEmail === emailLower) return { row: i + 2 };
    if (phoneNormalized && rowPhone && rowPhone === phoneNormalized) return { row: i + 2 };
  }
  return null;
}

/** Sinh ID tiếp theo theo quy tắc G001, G002, ... */
function getNextId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ID_PREFIX + String(1).padStart(ID_PAD, "0");

  const ids = sheet.getRange(2, COL.ID, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;
  for (const v of ids) {
    const s = String(v || "");
    const m = s.match(new RegExp("^" + ID_PREFIX + "(\\d+)$"));
    if (m) {
      const n = parseInt(m[1], 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  }
  const next = maxNum + 1;
  return ID_PREFIX + String(next).padStart(ID_PAD, "0");
}

/** Chuẩn hóa payload nhập tay */
function normalizePayloadManual_(payload) {
  payload = payload || {};
  const fullName = (payload.fullName || "").trim();
  const job      = (payload.job || "").trim();        // tùy chọn
  const org      = (payload.org || "").trim();
  const group    = (payload.group || "").trim();
  const email    = (payload.email || "").trim();
  const phone    = (payload.phone || "").trim();      // tùy chọn
  return { fullName, job, org, group, email, phone };
}

/** Email regex đơn giản */
function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/** Chuẩn hoá số điện thoại (dùng khi người dùng có nhập) */
function normalizePhone_(raw) {
  const digits = String(raw || "").replace(/\D/g, "");
  if (!digits) throw new Error('Số điện thoại không hợp lệ.');
  let normalized = digits;
  if (normalized.startsWith("84")) normalized = "0" + normalized.slice(2);
  if (!normalized.startsWith("0")) normalized = "0" + normalized;
  if (normalized.length < 9 || normalized.length > 11) {
    throw new Error('Số điện thoại không hợp lệ (9–11 chữ số).');
  }
  if (!/^\d+$/.test(normalized)) throw new Error('Số điện thoại chỉ được chứa chữ số.');
  return normalized;
}

/** Chuẩn hoá số điện thoại từ Sheet (không ném lỗi) */
function safeNormalizePhone_(raw) {
  const digits = String(raw || "").replace(/\D/g, "");
  if (!digits) return "";
  let n = digits;
  if (n.startsWith("84")) n = "0" + n.slice(2);
  if (!n.startsWith("0")) n = "0" + n;
  return n;
}

/** Kết quả trả về frontend (color = màu nhóm) */
function makeResult_(ok, message, group, time, id) {
  const color = GROUP_COLORS[group] || "#dddddd";
  return { ok, message, group: group || "", color, time: time || null, id: id || null };
}

function _esc(s) {
  return String(s == null ? "" : s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

/** Include (nếu cần) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
