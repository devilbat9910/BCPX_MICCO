/**
 * Module tổng hợp báo cáo cuối tháng
 */

/**
 * Hiển thị hộp thoại để tổng hợp báo cáo
 */
function showConsolidateReportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConsolidateReportDialog')
    .setWidth(400)
    .setHeight(300)
    .setTitle('Tổng hợp báo cáo cuối tháng');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Tổng hợp báo cáo cuối tháng');
}

/**
 * Lấy danh sách tháng có sẵn trong spreadsheet hiện tại
 * @return {Array} Danh sách các tháng có báo cáo
 */
function getAvailableMonths() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const availableMonths = new Set();
  
  // Duyệt qua từng sheet
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    // Kiểm tra tên sheet có khớp với định dạng MM/YYYY không
    if (/^\d{2}\/\d{4}$/.test(sheetName)) {
      availableMonths.add(sheetName);
    }
  }
  
  // Chuyển Set thành Array và sắp xếp theo thứ tự thời gian
  return Array.from(availableMonths).sort((a, b) => {
    const [monthA, yearA] = a.split('/');
    const [monthB, yearB] = b.split('/');
    
    // So sánh năm trước
    if (yearA !== yearB) {
      return yearA - yearB;
    }
    
    // Nếu năm giống nhau, so sánh tháng
    return monthA - monthB;
  });
}

/**
 * Tổng hợp báo cáo cuối tháng
 * @param {Object} data - Dữ liệu từ form
 * @return {Object} Kết quả tổng hợp báo cáo
 */
function consolidateReport(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Parse dữ liệu đầu vào
    const monthYear = data.monthYear;
    
    if (!monthYear) {
      throw new Error('Vui lòng chọn tháng/năm');
    }
    
    // Tạo hoặc lấy sheet báo cáo tổng hợp từ template BC_TCT
    const reportSheet = createConsolidatedReportSheet(monthYear);
    
    // Tìm sheet INPUT có tên là monthYear
    const inputSheet = ss.getSheetByName(monthYear);
    if (!inputSheet) {
      throw new Error(`Không tìm thấy sheet "${monthYear}" để lấy dữ liệu input`);
    }
    
    // Log để debug
    Logger.log(`Đang tổng hợp báo cáo từ sheet "${monthYear}" vào sheet "${reportSheet.getName()}"`);
    
    // Sao chép dữ liệu từ cột C của sheet INPUT sang cột E và G của sheet báo cáo
    copyValuesFromInputSheet(inputSheet, reportSheet);
    
    // Hiển thị sheet báo cáo (đảm bảo không bị ẩn)
    reportSheet.activate();
    reportSheet.showSheet();
    
    return {
      success: true,
      message: `Đã tổng hợp báo cáo cho tháng ${monthYear}`,
      sheetName: reportSheet.getName()
    };
    
  } catch (error) {
    Logger.log(`Lỗi khi tổng hợp báo cáo: ${error.message}`);
    return {
      success: false,
      message: `Lỗi: ${error.message}`
    };
  }
}

/**
 * Tạo hoặc lấy sheet báo cáo tổng hợp từ template BC_TCT
 * @param {string} monthYear - Tháng/năm (MM/YYYY)
 * @return {Sheet} Sheet báo cáo tổng hợp
 */
function createConsolidatedReportSheet(monthYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Đổi tên báo cáo thành "Báo cáo MM/YYYY"
  const reportSheetName = `Báo cáo ${monthYear}`;
  
  // Kiểm tra xem đã có sheet báo cáo chưa
  let reportSheet = ss.getSheetByName(reportSheetName);
  
  if (!reportSheet) {
    // Lấy template BC_TCT
    const templateSheet = ss.getSheetByName('BC_TCT');
    if (!templateSheet) {
      throw new Error('Không tìm thấy template BC_TCT');
    }
    
    // Tạo sheet mới từ template
    reportSheet = templateSheet.copyTo(ss);
    reportSheet.setName(reportSheetName);
    
    // Cập nhật tiêu đề với tháng/năm
    const parts = monthYear.split('/');
    if (parts.length === 2) {
      const month = parts[0];
      const year = parts[1];
      
      // Cập nhật tiêu đề
      reportSheet.getRange('A6').setValue(`THÁNG ${month} NĂM ${year}`);
    }
    
    // Hiển thị sheet mới (đảm bảo không bị ẩn theo sheet template)
    reportSheet.showSheet();
  }
  
  return reportSheet;
}

/**
 * Sao chép giá trị từ cột C của sheet INPUT sang cột E và G của sheet báo cáo
 * @param {Sheet} inputSheet - Sheet INPUT
 * @param {Sheet} reportSheet - Sheet báo cáo
 */
function copyValuesFromInputSheet(inputSheet, reportSheet) {
  // Lấy dữ liệu từ sheet INPUT
  const inputData = inputSheet.getDataRange().getValues();
  
  // Lấy dữ liệu từ sheet báo cáo
  const reportData = reportSheet.getDataRange().getValues();
  
  // Tạo map từ chỉ mục trong sheet INPUT
  const inputMap = new Map();
  for (let i = 0; i < inputData.length; i++) {
    const index = inputData[i][0]; // Cột A
    const productName = inputData[i][1]; // Cột B (tên sản phẩm)
    
    // Kiểm tra chỉ mục và tên sản phẩm
    if (index && typeof index === 'string' && productName && typeof productName === 'string') {
      // Kiểm tra điều kiện lọc: tên sản phẩm phải chứa "Thuốc nổ" hoặc "Sản xuất trên"
      if (productName.includes("Thuốc nổ") || productName.includes("Sản xuất trên")) {
        const trimmedIndex = index.toString().trim();
        // Lưu giá trị của cột C (index 2)
        const value = inputData[i][2];
        
        // Chỉ lưu vào map nếu có giá trị ở cột C
        if (value !== null && value !== undefined && value !== "") {
          inputMap.set(trimmedIndex, {
            value: value,
            rowIndex: i + 1,  // 1-based index
            name: productName  // Lưu thêm tên sản phẩm để debug
          });
        }
      }
    }
  }
  
  // Log để debug (có thể xóa khi triển khai)
  Logger.log("Dữ liệu từ sheet INPUT (đã lọc theo điều kiện):");
  inputMap.forEach((data, index) => {
    Logger.log(`Index: ${index}, Name: ${data.name}, Value: ${data.value}`);
  });
  
  // Duyệt qua từng hàng của sheet báo cáo (bắt đầu từ hàng 11)
  for (let i = 10; i < reportData.length; i++) {
    const index = reportData[i][0]; // Cột A
    
    if (index && typeof index === 'string') {
      const trimmedIndex = index.toString().trim();
      const parts = trimmedIndex.split('.');
      
      // Trường hợp 1: Chỉ mục level 2 là "1" (ví dụ: A.1, B.1, C.1)
      // Kiểm tra xem phần tử thứ 2 (level 2) có phải là "1" không và phần tử đầu tiên là chữ cái (A-Z)
      if (parts.length === 2 && parts[1] === '1' && /^[A-Z]$/.test(parts[0])) {
        // Nếu có dữ liệu tương ứng trong sheet INPUT
        if (inputMap.has(trimmedIndex)) {
          const value = inputMap.get(trimmedIndex).value;
          
          if (value !== null && value !== undefined && value !== "") {
            // Cập nhật cột E (index 4)
            reportSheet.getRange(i + 1, 5).setValue(value);
            
            // Cập nhật cột G (index 6)
            reportSheet.getRange(i + 1, 7).setValue(value);
            
            Logger.log(`Đã cập nhật chỉ mục ${trimmedIndex} với giá trị ${value} vào hàng ${i+1}`);
          }
        }
      }
      // Trường hợp 2: Các chỉ mục khác
      // Chỉ cập nhật dữ liệu nếu các chỉ mục hoàn toàn giống nhau
      else if (inputMap.has(trimmedIndex)) {
        const value = inputMap.get(trimmedIndex).value;
        
        if (value !== null && value !== undefined && value !== "") {
          // Cập nhật cột E (index 4)
          reportSheet.getRange(i + 1, 5).setValue(value);
          
          // Cập nhật cột G (index 6)
          reportSheet.getRange(i + 1, 7).setValue(value);
          
          Logger.log(`Đã cập nhật chỉ mục ${trimmedIndex} với giá trị ${value} vào hàng ${i+1}`);
        }
      }
    }
  }
}