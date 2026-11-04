/**
 * =========================================================================
 * MÃ NGUỒN GOOGLE APPS SCRIPT CHO CỔNG ĐĂNG KÝ THI ĐẤU THỂ THAO ĐA MÔN
 * =========================================================================
 * TỰ ĐỘNG LƯU TRỮ VÀO GOOGLE DRIVE:
 * 1. Tự động ghi vào Sheet [TongHop_TatCa] để Ban Tổ Chức theo dõi tổng quát.
 * 2. TỰ ĐỘNG TÁCH THÀNH CÁC SHEET RIÊNG CHO TỪNG MÔN THI ĐẤU (VD: Sheet "Bóng Đá", 
 *    Sheet "Cầu Lông", Sheet "Kéo Co"...) để BTC mở Google Drive ra là có ngay danh 
 *    sách chi tiết bốc thăm chia bảng cho từng môn!
 * 3. Tự động đổi tên file Google Sheet theo đúng tên giải mà bạn cấu hình.
 * =========================================================================
 * HƯỚNG DẪN 3 BƯỚC ĐƠN GIẢN:
 * 1. Mở trang tính Google Sheet mới tại https://sheets.new (lưu trên Google Drive của bạn).
 * 2. Vào menu: Tiện ích mở rộng (Extensions) -> Apps Script.
 * 3. Dán toàn bộ đoạn mã này vào -> Nhấn "Triển khai" (Deploy) -> "Tùy chọn triển khai mới"
 *    -> Loại: "Ứng dụng web" (Web App)
 *    -> Người thực thi: "Tôi" (Execute as: Me)
 *    -> Ai có quyền truy cập: "Bất kỳ ai" (Anyone)
 * 4. Copy link Web App URL (/exec) và dán vào phần "Quản trị BTC -> Lưu Trực Tiếp Vào Google Drive".
 * =========================================================================
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var body = JSON.parse(e.postData.contents);
    var tournamentName = body.tournamentName || "GIẢI THỂ THAO";
    var record = body.record || body;

    // Đổi tên file Google Sheet theo tên giải đấu cụ thể
    if (tournamentName && ss.getName().indexOf(tournamentName) === -1) {
      ss.rename("Danh Sách Đăng Ký - " + tournamentName);
    }

    // --- 1. GHI VÀO SHEET TỔNG HỢP TẤT CẢ CÁC ĐOÀN ---
    var sheetTongHop = getOrCreateSheet(ss, "TongHop_TatCa", [
      "Mã Hồ Sơ", "Tên Đơn Vị / CLB", "Môn Thi Đấu", "Số VĐV", 
      "Trưởng Đoàn", "Số Điện Thoại", "Màu Áo / Ghi Chú", "Thời Gian Nộp", "Trạng Thái"
    ]);

    sheetTongHop.appendRow([
      record.id || "",
      record.donVi || "",
      record.sportName || "",
      (record.athletes ? record.athletes.length : 0),
      record.truongDoan || "",
      "'" + (record.sdt || ""),
      record.ghiChu || "",
      record.createdAt || new Date().toLocaleString("vi-VN"),
      record.status || "CHO_DUYET"
    ]);

    // --- 2. TỰ ĐỘNG TẠO & GHI VÀO SHEET RIÊNG THEO TỪNG MÔN THI ĐẤU ---
    var safeSportSheetName = sanitizeSheetName(record.sportName || "MonKhac");
    var isFootball = (record.sportId && record.sportId.indexOf("bong-da") !== -1);

    var sportHeaders = isFootball 
      ? ["STT", "Mã Hồ Sơ", "Đơn Vị", "Họ và Tên Cầu Thủ", "Năm Sinh", "Giới Tính", "Số CCCD", "Vị Trí", "Số Áo", "Vai Trò"]
      : ["STT", "Mã Hồ Sơ", "Đơn Vị", "Họ và Tên VĐV", "Năm Sinh", "Giới Tính", "Số CCCD", "Nội Dung / Hạng Cân", "Số Áo", "Đội Hình"];

    var sheetSport = getOrCreateSheet(ss, safeSportSheetName, sportHeaders);

    if (record.athletes && record.athletes.length > 0) {
      var rows = [];
      var startStt = sheetSport.getLastRow(); // STT tiếp nối trong sheet môn đó
      
      record.athletes.forEach(function(a, idx) {
        if (isFootball) {
          rows.push([
            startStt + idx,
            record.id || "",
            record.donVi || "",
            a.hoTen || "",
            "'" + (a.namSinh || ""),
            a.gioiTinh || "",
            "'" + (a.cccd || ""),
            a.vaiTro || "",
            "'" + (a.soAo || ""),
            a.loaiVdv || "Chính thức"
          ]);
        } else {
          rows.push([
            startStt + idx,
            record.id || "",
            record.donVi || "",
            a.hoTen || "",
            "'" + (a.namSinh || ""),
            a.gioiTinh || "",
            "'" + (a.cccd || ""),
            a.vaiTro || "",
            "'" + (a.soAo || ""),
            a.loaiVdv || ""
          ]);
        }
      });

      sheetSport.getRange(sheetSport.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "Đã lưu hồ sơ thành công vào Google Drive của bạn!",
      id: record.id,
      tournament: tournamentName
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("TongHop_TatCa");
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    }
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      list.push({
        id: data[i][0],
        donVi: data[i][1],
        sportName: data[i][2],
        soVdv: data[i][3],
        truongDoan: data[i][4],
        sdt: data[i][5],
        ghiChu: data[i][6],
        createdAt: data[i][7],
        status: data[i][8]
      });
    }
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: list
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    var range = sheet.getRange(1, 1, 1, headers.length);
    range.setBackground("#0b1324");
    range.setFontColor("#ffffff");
    range.setFontWeight("bold");
    sheet.setFrozenRows(1);
    // Tự động căn chỉnh độ rộng cột
    for (var col = 1; col <= headers.length; col++) {
      sheet.autoResizeColumn(col);
    }
  }
  return sheet;
}

function sanitizeSheetName(name) {
  var clean = name.replace(/[\/\\\?\*\]\[:]/g, "").trim();
  if (clean.length > 28) clean = clean.substring(0, 28);
  return clean || "MonThiDau";
}
