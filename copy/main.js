const TEN_SHEET = 'noiDung';

/**
 * Trang Web App
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('QR Clipboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Lấy Google Sheet chứa dữ liệu.
 *
 * Nếu chưa có sheet "noiDung", hàm sẽ tự tạo.
 */
function laySheetNoiDung() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(TEN_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(TEN_SHEET);

    sheet.getRange(1, 1, 1, 8).setValues([[
      'id',
      'ngay',
      'thang',
      'nam',
      'gio',
      'phut',
      'giay',
      'noiDung'
    ]]);
  }

  return sheet;
}


/**
 * Lấy toàn bộ nội dung.
 *
 * Sắp xếp mới nhất lên đầu.
 */
function layDanhSachNoiDung() {
  const sheet = laySheetNoiDung();

  const soDong = sheet.getLastRow();

  if (soDong <= 1) {
    return [];
  }

  const duLieu = sheet
    .getRange(2, 1, soDong - 1, 8)
    .getValues();

  return duLieu
    .filter(dong => dong[0] !== '')
    .map(dong => ({
      id: String(dong[0]),
      ngay: Number(dong[1]),
      thang: Number(dong[2]),
      nam: Number(dong[3]),
      gio: Number(dong[4]),
      phut: Number(dong[5]),
      giay: Number(dong[6]),
      noiDung: String(dong[7] ?? '')
    }))
    .reverse();
}


/**
 * Lưu một nội dung mới.
 */
function luuNoiDung(noiDung) {
  if (noiDung === null || noiDung === undefined) {
    throw new Error('Nội dung không hợp lệ.');
  }

  noiDung = String(noiDung);

  if (!noiDung.trim()) {
    throw new Error('Nội dung không được để trống.');
  }

  const sheet = laySheetNoiDung();

  const thoiGian = new Date();

  const ngay = thoiGian.getDate();
  const thang = thoiGian.getMonth() + 1;
  const nam = thoiGian.getFullYear();

  const gio = thoiGian.getHours();
  const phut = thoiGian.getMinutes();
  const giay = thoiGian.getSeconds();

  const id =
    Utilities.getUuid();

  sheet.appendRow([
    id,
    ngay,
    thang,
    nam,
    gio,
    phut,
    giay,
    noiDung
  ]);

  return {
    thanhCong: true,
    id: id,
    ngay: ngay,
    thang: thang,
    nam: nam,
    gio: gio,
    phut: phut,
    giay: giay,
    noiDung: noiDung
  };
}


/**
 * Xóa một nội dung theo ID.
 */
function xoaNoiDung(id) {
  if (!id) {
    throw new Error('Thiếu ID cần xóa.');
  }

  const sheet = laySheetNoiDung();

  const soDong = sheet.getLastRow();

  if (soDong <= 1) {
    return false;
  }

  const danhSachId = sheet
    .getRange(2, 1, soDong - 1, 1)
    .getValues();

  for (let i = 0; i < danhSachId.length; i++) {
    if (String(danhSachId[i][0]) === String(id)) {

      sheet.deleteRow(i + 2);

      return true;
    }
  }

  return false;
}


/**
 * Xóa toàn bộ dữ liệu nhưng giữ lại Header.
 */
function xoaTatCaNoiDung() {
  const sheet = laySheetNoiDung();

  const soDong = sheet.getLastRow();

  if (soDong > 1) {
    sheet
      .getRange(2, 1, soDong - 1, 8)
      .clearContent();
  }

  return true;
}
