/**
 * Cấu hình nhóm → màu sắc hiển thị.
 */
const GROUP_COLORS = {
  "VIP": "#ffd700",         // vàng
  "Diễn giả": "#90ee90",    // xanh lá nhạt
  "Khách mời": "#add8e6",   // xanh dương nhạt
  "Nội bộ": "#d3d3d3",      // xám
  "Khác": "#ffe4b5"         // cam nhạt
};

const SHEET_NAME = "Checkin"; 
const TZ = "Asia/Ho_Chi_Minh";

function doGet() {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.groupColors = GROUP_COLORS;
  return tpl.evaluate()
    .setTitle('Check-in')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function checkIn(payload) {
  try {
    if (payload && payload.raw) {
      const parsed = parseQrPayload_(payload.raw);
      payload = Object.assign({}, payload, parsed);
    }

    const { id, fullName, org, group } = normalizePayload_(payload);

    if (!fullName || !org) throw new Error('Thiếu "Họ và tên" hoặc "Đơn vị công tác".');

    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Không tìm thấy sheet ${SHEET_NAME}`);

    const found = findRowById_(sheet, id);
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

    if (found) {
      const existingTime = sheet.getRange(found.row, 5).getDisplayValue();
      return makeResult_(true, `ĐÃ check-in trước đó: ${fullName}`, group, existingTime, id);
    }

    sheet.appendRow([id, fullName, org, group || 'Khác', now]);
    return makeResult_(true, `Check-in thành công: ${fullName}`, group, now, id);
  } catch (err) {
    return makeResult_(false, err.message || String(err));
  }
}

function findRowById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return { row: i + 2 };
  }
  return null;
}

function normalizePayload_(payload) {
  const fullName = (payload.fullName || payload.name || '').trim();
  const org = (payload.org || payload.company || payload.department || '').trim();
  const group = (payload.group || payload.segment || 'Khác').trim();
  let id = (payload.id || '').trim();
  if (!id) id = Utilities.base64EncodeWebSafe(`${fullName}|${org}`).slice(0, 16);
  return { id, fullName, org, group };
}

function parseQrPayload_(rawText) {
  if (!rawText) return {};
  const t = String(rawText).trim();
  if (t.startsWith('{') && t.endsWith('}')) {
    try {
      const obj = JSON.parse(t);
      return {
        id: obj.id,
        fullName: obj.fullName || obj.name,
        org: obj.org || obj.company || obj.department,
        group: obj.group || obj.segment
      };
    } catch (e) {}
  }
  const parts = t.split('|');
  if (parts.length >= 4) return { id: parts[0], fullName: parts[1], org: parts[2], group: parts[3] };
  return { id: t };
}

function makeResult_(ok, message, group, time, id) {
  const color = GROUP_COLORS[group] || GROUP_COLORS['Khác'] || '#eee';
  return { ok, message, group: group || 'Khác', color, time: time || null, id: id || null };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}