function sendReportUI() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  // Lấy worksheet cuối cùng
  const newestWorksheet = sheet.getSheets()[sheet.getSheets().length - 1];
  const newestWorksheetName = newestWorksheet.getName();

  // Hiển thị hộp thoại để lựa chọn báo cáo hoặc nhập tên worksheet
  const response = ui.prompt(
    'Gửi báo cáo',
    `Báo cáo mới nhất là "${newestWorksheetName}". Nhập tên báo cáo nếu muốn chọn khác hoặc để trống để sử dụng báo cáo mới nhất:`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const inputName = response.getResponseText().trim();
    let targetWorksheet;
    const validPattern = /^Báo cáo \d{2}\/\d{4}$/;

    if (inputName) {
      // Kiểm tra định dạng của tên báo cáo nhập vào
      if (!validPattern.test(inputName)) {
        ui.alert('Tên báo cáo không đúng mẫu! Vui lòng nhập tên theo mẫu: "Báo cáo 01/2025".');
        return;
      }

      // Tìm worksheet theo tên nhập vào
      targetWorksheet = sheet.getSheetByName(inputName);
      if (!targetWorksheet) {
        ui.alert(`Worksheet "${inputName}" không tồn tại. Vui lòng kiểm tra lại!`);
        return;
      }
    } else {
      // Sử dụng worksheet cuối cùng
      targetWorksheet = newestWorksheet;
      // Kiểm tra định dạng của tên báo cáo từ worksheet mới nhất
      if (!validPattern.test(targetWorksheet.getName())) {
        ui.alert(`Tên báo cáo "${targetWorksheet.getName()}" không đúng mẫu. Vui lòng đảm bảo tên theo mẫu: "Báo cáo 01/2025".`);
        return;
      }
    }

    // Google Sheet đích
    const targetSheetUrl = "https://docs.google.com/spreadsheets/d/1dTyvaTxTFUAQRtzjeVpuNSmlVtrBhjXcyZThIlprphg/edit?usp=sharing";
    const targetSheetId = targetSheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];

    // Gửi báo cáo
    try {
      const log = copyWorksheetDirectly(targetWorksheet, targetSheetId);
      ui.alert(`Đã gửi báo cáo tới phòng KTCN.`);
    } catch (error) {
      ui.alert(`Đã gửi LẠI báo cáo !!!`);
    }
  } else {
    ui.alert('Gửi báo cáo đã bị huỷ.');
  }
}

function copyWorksheetDirectly(sourceWorksheet, targetSpreadsheetId) {
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  if (!targetSpreadsheet) {
    throw new Error('Không thể mở Google Sheet đích. Vui lòng kiểm tra ID.');
  }

  // Sử dụng copyTo để sao chép toàn bộ worksheet
  const tempSheet = sourceWorksheet.copyTo(targetSpreadsheet);

  // Đặt tên với tiền tố "PXĐT_"
  const newSheetName = `PXĐT_${sourceWorksheet.getName()}`;
  tempSheet.setName(newSheetName);

  return `Báo cáo đã được gửi tới phòng KTCN với tên: ${newSheetName}`;
}
