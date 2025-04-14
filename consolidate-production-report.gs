/**
 * Module tổng hợp báo cáo từ các phân xưởng về phòng KTCN
 */

/**
 * Hiển thị hộp thoại để tổng hợp báo cáo từ các phân xưởng
 */
function showConsolidateReportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConsolidateReportDialog')
    .setWidth(600)
    .setHeight(450)
    .setTitle('Tổng hợp báo cáo từ các phân xưởng');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Tổng hợp báo cáo từ các phân xưởng');
}

/**
 * Lấy danh sách tháng có sẵn từ các phân xưởng
 * @return {Array} Danh sách các tháng có báo cáo
 */
function getAvailableMonths() {
  const workshopUrls = getWorkshopUrls();
  const availableMonths = new Set();
  
  // Duyệt qua từng phân xưởng
  for (const code in workshopUrls) {
    try {
      // Mở bảng tính của phân xưởng
      const workshopSS = SpreadsheetApp.openByUrl(workshopUrls[code]);
      
      // Lấy danh sách sheet
      const sheets = workshopSS.getSheets();
      
      // Tìm các sheet có tên dạng MM/YYYY
      for (const sheet of sheets) {
        const sheetName = sheet.getName();
        // Kiểm tra tên sheet có khớp với định dạng MM/YYYY không
        if (/^\d{2}\/\d{4}$/.test(sheetName)) {
          availableMonths.add(sheetName);
        }
      }
    } catch (error) {
      Logger.log(`Lỗi khi truy cập phân xưởng ${code}: ${error.message}`);
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
 * Lấy URLs của các phân xưởng
 * @return {Object} Danh sách URL theo mã phân xưởng
 */
function getWorkshopUrls() {
  return {
    'CP': 'https://docs.google.com/spreadsheets/d/1fS7bRnPy2xJChqoLVr1AgEmMyoeOJlaNC0Plt_JS7N8/edit?usp=sharing',
    'ĐN': 'https://docs.google.com/spreadsheets/d/1OxLqZDL6sWXa3vg0inM_0d8CbGvAQCrEWSKpQTx-U84/edit?usp=sharing',
    'TB': 'https://docs.google.com/spreadsheets/d/1Nnn3_ElEiYGs2eanwH5O8fv7YJIMiwpzUreU_pVCNP8/edit?usp=sharing',
    'QN': 'https://docs.google.com/spreadsheets/d/1R9lMIQjzL_eDkMCCEdUenImUBOwxdE_LA3d78h-QriQ/edit?usp=sharing',
    'NB': 'https://docs.google.com/spreadsheets/d/1QT7fJvY7573VB-UJNCq3uJxVd-UMWTg57tDBn4U7FqU/edit?usp=sharing',
    'VT': 'https://docs.google.com/spreadsheets/d/1ojKesIV8nDd495U28GBEUoDTsUDfUqSqv-MW-Xza8vU/edit?usp=sharing',
    'ĐT': 'https://docs.google.com/spreadsheets/d/1RRO_RK2dZJcEsGtYxM4OUYPcv5BXP_vr0Od_idap8PA/edit?usp=sharing'
  };
}

/**
 * Lấy danh sách phân xưởng đã có báo cáo cho một tháng cụ thể
 * @param {string} monthYear - Tháng/năm (MM/YYYY)
 * @return {Array} Danh sách các phân xưởng có báo cáo
 */
function getWorkshopsWithReport(monthYear) {
  const workshopUrls = getWorkshopUrls();
  const workshopsWithReport = [];
  
  // Duyệt qua từng phân xưởng
  for (const code in workshopUrls) {
    try {
      // Mở bảng tính của phân xưởng
      const workshopSS = SpreadsheetApp.openByUrl(workshopUrls[code]);
      
      // Kiểm tra xem có sheet tương ứng với tháng đã chọn không
      const sheet = workshopSS.getSheetByName(monthYear);
      
      if (sheet) {
        // Lấy tên phân xưởng
        const name = getWorkshopName(code);
        
        workshopsWithReport.push({
          code: code,
          name: name,
          url: workshopUrls[code]
        });
      }
    } catch (error) {
      Logger.log(`Lỗi khi truy cập phân xưởng ${code}: ${error.message}`);
    }
  }
  
  return workshopsWithReport;
}

/**
 * Lấy tên phân xưởng từ mã
 * @param {string} code - Mã phân xưởng
 * @return {string} Tên phân xưởng
 */
function getWorkshopName(code) {
  const names = {
    'CP': 'Cẩm Phả',
    'ĐN': 'Đà Nẵng',
    'TB': 'Tây Bắc',
    'QN': 'Quảng Ninh',
    'NB': 'Ninh Bình',
    'VT': 'Vũng Tàu',
    'ĐT': 'Đông Triều'
  };
  
  return names[code] || code;
}

/**
 * Tổng hợp báo cáo từ các phân xưởng
 * @param {Object} data - Dữ liệu từ form
 * @return {Object} Kết quả tổng hợp báo cáo
 */
function consolidateReports(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Parse dữ liệu đầu vào
    const monthYear = data.monthYear;
    const selectedWorkshops = data.selectedWorkshops || [];
    
    if (!monthYear || selectedWorkshops.length === 0) {
      throw new Error('Vui lòng chọn tháng/năm và ít nhất một phân xưởng');
    }
    
    // Tạo hoặc lấy sheet báo cáo tổng hợp từ template BC_TCT
    const reportSheet = createConsolidatedReportSheet(monthYear);
    
    // Thu thập dữ liệu từ các phân xưởng
    const workshopData = collectDataFromWorkshops(selectedWorkshops, monthYear);
    
    // Xây dựng cây chỉ mục và xác định các nút có dữ liệu
    const indexTree = buildIndexTree(reportSheet);
    const nodesWithData = identifyNodesWithData(indexTree, workshopData);
    
    // Tổng hợp dữ liệu vào báo cáo
    updateConsolidatedReport(reportSheet, workshopData, nodesWithData);
    
    // Ẩn các hàng không có dữ liệu
    hideEmptyRows(reportSheet, nodesWithData);
    
    // Sao chép giá trị từ cột L sang cột E và G cho các hàng có INDEX level 2 là "1"
    copyValuesFromColumnL(reportSheet);
    
    return {
      success: true,
      message: `Đã tổng hợp báo cáo từ ${selectedWorkshops.length} phân xưởng cho tháng ${monthYear}`,
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
  }
  
  return reportSheet;
}

/**
 * Thu thập dữ liệu từ các phân xưởng
 * @param {Array} workshops - Danh sách các phân xưởng
 * @param {string} monthYear - Tháng/năm (MM/YYYY)
 * @return {Object} Dữ liệu từ các phân xưởng
 */
function collectDataFromWorkshops(workshops, monthYear) {
  const workshopUrls = getWorkshopUrls();
  const workshopData = {};
  
  // Duyệt qua từng phân xưởng được chọn
  for (const workshopCode of workshops) {
    try {
      // Mở bảng tính của phân xưởng
      const url = workshopUrls[workshopCode];
      if (!url) {
        Logger.log(`Không tìm thấy URL cho phân xưởng ${workshopCode}`);
        continue;
      }
      
      const workshopSS = SpreadsheetApp.openByUrl(url);
      
      // Lấy sheet báo cáo
      const sheet = workshopSS.getSheetByName(monthYear);
      if (!sheet) {
        Logger.log(`Không tìm thấy báo cáo tháng ${monthYear} cho phân xưởng ${workshopCode}`);
        continue;
      }
      
      // Lấy dữ liệu từ sheet
      const data = sheet.getDataRange().getValues();
      
      // Tìm dữ liệu sản lượng (từ cột C)
      const productionData = {};
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const index = row[0]; // Cột A (chỉ mục)
        const name = row[1];  // Cột B (tên)
        const production = row[2]; // Cột C (sản lượng)
        
        // Kiểm tra nếu là chỉ mục hợp lệ và có sản lượng
        if (index && typeof index === 'string' && production) {
          productionData[index.toString().trim()] = {
            name: name,
            production: production,
            rowIndex: i + 1  // 1-based index
          };
        }
      }
      
      workshopData[workshopCode] = productionData;
      
    } catch (error) {
      Logger.log(`Lỗi khi thu thập dữ liệu từ phân xưởng ${workshopCode}: ${error.message}`);
    }
  }
  
  return workshopData;
}

/**
 * Xây dựng cây chỉ mục từ sheet báo cáo
 * @param {Sheet} sheet - Sheet báo cáo
 * @return {Object} Cây chỉ mục
 */
function buildIndexTree(sheet) {
  const data = sheet.getDataRange().getValues();
  const indexTree = {
    root: {
      children: {},
      parent: null,
      row: 0
    }
  };
  
  // Bắt đầu từ hàng 11 (index 10) - sau tiêu đề
  for (let i = 10; i < data.length; i++) {
    const index = data[i][0]; // Cột A
    
    if (index && typeof index === 'string') {
      const trimmedIndex = index.toString().trim();
      const parts = trimmedIndex.split('.');
      
      // Thêm vào cây
      let currentNode = indexTree.root;
      let currentPath = '';
      
      for (let j = 0; j < parts.length; j++) {
        const part = parts[j];
        currentPath = currentPath ? `${currentPath}.${part}` : part;
        
        if (!currentNode.children[part]) {
          currentNode.children[part] = {
            index: currentPath,
            children: {},
            parent: currentNode,
            row: i + 1, // 1-based index
            data: {
              name: data[i][1], // Cột B (tên)
            }
          };
        }
        
        currentNode = currentNode.children[part];
      }
    }
  }
  
  return indexTree;
}

/**
 * Xác định các nút trong cây có dữ liệu
 * @param {Object} indexTree - Cây chỉ mục
 * @param {Object} workshopData - Dữ liệu từ các phân xưởng
 * @return {Set} Tập hợp các nút có dữ liệu
 */
function identifyNodesWithData(indexTree, workshopData) {
  const nodesWithData = new Set();
  
  // Hàm đệ quy để duyệt cây và đánh dấu nút có dữ liệu
  function traverseTree(node, path = '') {
    if (!node) return false;
    
    let hasData = false;
    const nodePath = path ? path : node.index;
    
    // Kiểm tra dữ liệu từ các phân xưởng
    for (const workshopCode in workshopData) {
      const workshopProductionData = workshopData[workshopCode];
      
      if (nodePath && workshopProductionData[nodePath]) {
        hasData = true;
        nodesWithData.add(nodePath);
        
        // Đánh dấu tất cả nút cha
        let parent = node.parent;
        let parentPath = '';
        
        // Tạo lại path cho nút cha
        if (nodePath) {
          const parts = nodePath.split('.');
          parts.pop();
          parentPath = parts.join('.');
        }
        
        while (parent && parent !== indexTree.root) {
          if (parentPath) {
            nodesWithData.add(parentPath);
          }
          
          // Đi lên nút cha tiếp theo
          parent = parent.parent;
          if (parentPath) {
            const parts = parentPath.split('.');
            parts.pop();
            parentPath = parts.join('.');
          }
        }
      }
    }
    
    // Duyệt các nút con
    for (const childKey in node.children) {
      const childNode = node.children[childKey];
      const childPath = nodePath ? `${nodePath}.${childKey}` : childKey;
      
      const childHasData = traverseTree(childNode, childPath);
      
      if (childHasData) {
        hasData = true;
        
        if (nodePath) {
          nodesWithData.add(nodePath);
        }
      }
    }
    
    return hasData;
  }
  
  traverseTree(indexTree.root);
  
  return nodesWithData;
}

/**
 * Cập nhật báo cáo tổng hợp với dữ liệu từ các phân xưởng
 * @param {Sheet} sheet - Sheet báo cáo tổng hợp
 * @param {Object} workshopData - Dữ liệu từ các phân xưởng
 * @param {Set} nodesWithData - Tập hợp các nút có dữ liệu
 */
function updateConsolidatedReport(sheet, workshopData, nodesWithData) {
  // Lấy dữ liệu của sheet báo cáo
  const data = sheet.getDataRange().getValues();
  
  // Tìm vị trí của các phân xưởng trong báo cáo
  const workshopColumns = findWorkshopColumns(sheet);
  
  // Cập nhật dữ liệu sản lượng
  for (let i = 10; i < data.length; i++) { // Bắt đầu từ hàng 11 (index 10)
    const index = data[i][0]; // Cột A
    
    if (index && typeof index === 'string') {
      const trimmedIndex = index.toString().trim();
      
      // Chỉ cập nhật nếu nút này có dữ liệu
      if (nodesWithData.has(trimmedIndex)) {
        // Cập nhật dữ liệu từng phân xưởng
        for (const workshopCode in workshopData) {
          const workshopProductionData = workshopData[workshopCode];
          const columnIndex = workshopColumns[workshopCode];
          
          if (columnIndex && workshopProductionData[trimmedIndex]) {
            const production = workshopProductionData[trimmedIndex].production;
            
            if (production) {
              sheet.getRange(i + 1, columnIndex).setValue(production);
            }
          }
        }
      }
    }
  }
}

/**
 * Tìm vị trí cột của các phân xưởng trong báo cáo
 * @param {Sheet} sheet - Sheet báo cáo tổng hợp
 * @return {Object} Vị trí cột của các phân xưởng
 */
function findWorkshopColumns(sheet) {
  const headerRow = sheet.getRange(10, 1, 1, sheet.getLastColumn()).getValues()[0];
  const workshopColumns = {};
  
  // Tìm vị trí cột của các phân xưởng
  for (let i = 0; i < headerRow.length; i++) {
    const header = headerRow[i];
    
    if (header && typeof header === 'string') {
      // Giả sử header là mã phân xưởng (ĐT, CP, QN, ...)
      const workshopCode = header.toString().trim();
      
      if (/^[A-ZĐ]{2}$/.test(workshopCode)) {
        workshopColumns[workshopCode] = i + 1; // 1-based index
      }
    }
  }
  
  return workshopColumns;
}

/**
 * Ẩn các hàng không có dữ liệu
 * @param {Sheet} sheet - Sheet báo cáo tổng hợp
 * @param {Set} nodesWithData - Tập hợp các nút có dữ liệu
 */
function hideEmptyRows(sheet, nodesWithData) {
  const data = sheet.getDataRange().getValues();
  const rowsToHide = [];
  
  // Tìm các hàng không có dữ liệu
  for (let i = 10; i < data.length; i++) { // Bắt đầu từ hàng 11 (index 10)
    const index = data[i][0]; // Cột A
    
    if (index && typeof index === 'string') {
      const trimmedIndex = index.toString().trim();
      
      // Nếu nút không có dữ liệu, ẩn hàng
      if (!nodesWithData.has(trimmedIndex)) {
        rowsToHide.push(i + 1); // 1-based index
      }
    }
  }
  
  // Ẩn các hàng
  if (rowsToHide.length > 0) {
    // Nhóm các hàng liên tiếp để ẩn hiệu quả hơn
    let startRow = rowsToHide[0];
    let count = 1;
    
    for (let i = 1; i < rowsToHide.length; i++) {
      if (rowsToHide[i] === rowsToHide[i - 1] + 1) {
        count++;
      } else {
        // Ẩn nhóm hàng hiện tại
        sheet.hideRows(startRow, count);
        
        // Bắt đầu nhóm mới
        startRow = rowsToHide[i];
        count = 1;
      }
    }
    
    // Ẩn nhóm cuối cùng
    if (count > 0) {
      sheet.hideRows(startRow, count);
    }
  }
}

/**
 * Sao chép giá trị từ cột L sang cột E và G cho các hàng có INDEX level 2 là "1"
 * @param {Sheet} sheet - Sheet báo cáo tổng hợp
 */
function copyValuesFromColumnL(sheet) {
  // Lấy toàn bộ dữ liệu của sheet
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 12).getValues(); // Lấy dữ liệu từ cột A đến L
  
  // Mảng chứa các cập nhật cần thực hiện
  const updatesE = [];
  const updatesG = [];
  
  // Duyệt qua từng hàng (bắt đầu từ hàng 11 - sau phần tiêu đề)
  for (let i = 10; i < data.length; i++) {
    const index = data[i][0]; // Cột A - INDEX
    const valueL = data[i][11]; // Cột L (index 11)
    
    // Kiểm tra nếu có INDEX và có giá trị ở cột L
    if (index && typeof index === 'string' && valueL) {
      // Phân tích INDEX
      const parts = index.toString().trim().split('.');
      
      // Kiểm tra xem phần tử thứ 2 (level 2) có phải là "1" không
      if (parts.length >= 2 && parts[1] === '1') {
        // Lưu lại hàng và giá trị cần cập nhật
        updatesE.push({
          row: i + 1, // 1-based index
          value: valueL
        });
        
        updatesG.push({
          row: i + 1, // 1-based index
          value: valueL
        });
      }
    }
  }
  
  // Thực hiện các cập nhật cho cột E
  updatesE.forEach(update => {
    sheet.getRange(update.row, 5).setValue(update.value); // Cột E = 5
  });
  
  // Thực hiện các cập nhật cho cột G
  updatesG.forEach(update => {
    sheet.getRange(update.row, 7).setValue(update.value); // Cột G = 7
  });
  
  Logger.log(`Đã sao chép ${updatesE.length} giá trị từ cột L sang cột E và G.`);
}