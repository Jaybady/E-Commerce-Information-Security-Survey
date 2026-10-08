// Google Apps Script - Security Survey 2026
// ===========================================
// Đồng bộ với index.html: 15 câu hỏi (Q1-Q15)
//
// Cài đặt:
// 1. Tạo Google Sheet mới.
// 2. Extensions → Apps Script.
// 3. Xóa code mặc định và dán toàn bộ code này.
// 4. Deploy → New deployment → Web app.
// 5. Execute as: Me.
// 6. Who has access: Anyone.
// 7. Copy URL Web App và đặt vào SCRIPT_URL trong index.html.

const SHEET_NAME = 'Security Survey Data 2026';

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = JSON.parse(e.postData.contents);

    // Chuẩn bị dữ liệu theo đúng thứ tự Q1 → Q15 trong form.
    const row = [
      data.timestamp || new Date().toLocaleString('vi-VN'),
      data.q1 || '',
      joinValue(data.q2),
      data.q3 || '',
      data.q4 || '',
      joinValue(data.q5),
      joinValue(data.q6),
      data.q7 || '',
      data.q8 || '',
      data.q9 || '',
      data.q10 || '',
      data.q11 || '',
      joinValue(data.q12),
      data.q13 || '',
      joinValue(data.q14),
      joinValue(data.q15)
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Data saved successfully'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function joinValue(value) {
  if (Array.isArray(value)) {
    return value.join('; ');
  }
  return value || '';
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    sheet.appendRow([
      'Timestamp',
      'Q1: Tần suất mua sắm',
      'Q2: Nền tảng mua sắm',
      'Q3: Chi tiêu hàng tháng',
      'Q4: Mức độ quan tâm bảo mật',
      'Q5: Thông tin cần bảo vệ',
      'Q6: Rủi ro bảo mật đáng lo ngại',
      'Q7: Sử dụng mật khẩu khác nhau',
      'Q8: Sử dụng 2FA/MFA/OTP',
      'Q9: Xử lý quyền truy cập ứng dụng',
      'Q10: Kiểm tra đăng nhập/thiết bị',
      'Q11: Tin nhắn/email/cuộc gọi giả mạo',
      'Q12: Sự cố an toàn thông tin từng gặp',
      'Q13: Ứng phó tin nhắn yêu cầu nhấn link',
      'Q14: Biện pháp giảm rủi ro ưu tiên',
      'Q15: Trách nhiệm bảo vệ thông tin'
    ]);

    // Định dạng hàng tiêu đề.
    const headerRange = sheet.getRange(1, 1, 1, 16);
    headerRange.setBackground('#667eea');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setWrap(true);
    sheet.setFrozenRows(1);

    // Định dạng độ rộng cơ bản để dễ đọc dữ liệu.
    sheet.autoResizeColumns(1, 16);
  }

  return sheet;
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'active',
      message: 'Security Survey 2026 API is running'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
