// Google Apps Script Code
// ========================================
// Khảo sát: Rủi ro bảo mật và quản trị an toàn thông tin trên ứng dụng TMĐT di động
//
// Hướng dẫn cài đặt:
// 1. Mở Google Sheet dùng để lưu kết quả.
// 2. Extensions → Apps Script.
// 3. Thay code hiện tại bằng code này.
// 4. Deploy → Manage deployments → Edit deployment (hoặc New deployment nếu cần).
// 5. Loại triển khai: Web app.
// 6. Execute as: Me.
// 7. Who has access: Anyone.
// 8. Đặt URL Web App trong biến SCRIPT_URL của index.html.

const SHEET_NAME = 'Security Survey Data 2026';

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = JSON.parse(e.postData.contents);

    // Row order is exactly the same as Q1 → Q15 in index_updated_security_survey.html.
    const row = [
      data.timestamp || new Date().toLocaleString('vi-VN'),
      data.q1 || '',                                                                 // Q1: Tần suất mua sắm trực tuyến
      toCell(data.q2),                                                               // Q2: Nền tảng mua sắm
      data.q3 || '',                                                                 // Q3: Mức chi tiêu mỗi tháng
      data.q4 || '',                                                                 // Q4: Mức độ quan tâm bảo mật
      toCell(data.q5),                                                               // Q5: Thông tin cần được bảo vệ
      toCell(data.q6),                                                               // Q6: Rủi ro bảo mật đáng lo ngại
      data.q7 || '',                                                                 // Q7: Sử dụng mật khẩu khác nhau
      data.q8 || '',                                                                 // Q8: Sử dụng 2FA/OTP
      data.q9 || '',                                                                 // Q9: Quyền truy cập ứng dụng
      data.q10 || '',                                                                // Q10: Kiểm tra thiết bị/phiên đăng nhập
      data.q11 || '',                                                                // Q11: Nhận tin nhắn/email/link đáng ngờ
      toCell(data.q12),                                                              // Q12: Sự cố bảo mật đã gặp
      data.q13 || '',                                                                // Q13: Cách xử lý tin nhắn đáng ngờ
      toCell(data.q14),                                                              // Q14: Biện pháp ưu tiên
      toCell(data.q15)                                                               // Q15: Trách nhiệm bảo vệ thông tin
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data saved'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function toCell(value) {
  if (Array.isArray(value)) return value.join('; ');
  return value || '';
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headers = [
    'Timestamp',
    'Q1: Tần suất mua sắm trực tuyến',
    'Q2: Nền tảng mua sắm',
    'Q3: Mức chi tiêu mỗi tháng',
    'Q4: Mức độ quan tâm bảo mật',
    'Q5: Thông tin cần được bảo vệ',
    'Q6: Rủi ro bảo mật đáng lo ngại',
    'Q7: Sử dụng mật khẩu khác nhau',
    'Q8: Sử dụng 2FA/OTP',
    'Q9: Quyền truy cập ứng dụng',
    'Q10: Kiểm tra thiết bị/phiên đăng nhập',
    'Q11: Nhận tin nhắn/email/link đáng ngờ',
    'Q12: Sự cố bảo mật đã gặp',
    'Q13: Cách xử lý tin nhắn đáng ngờ',
    'Q14: Biện pháp ưu tiên',
    'Q15: Trách nhiệm bảo vệ thông tin'
  ];

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const sameHeaders = headers.every((h, i) => current[i] === h);
  if (!sameHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#667eea');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }

  return sheet;
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'active',
    message: 'Security Survey API is running'
  })).setMimeType(ContentService.MimeType.JSON);
}
