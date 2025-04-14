const Helpers = {
  getProductList: function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Danh mục sản phẩm');
    if (!sheet) {
      throw new Error('Không tìm thấy sheet "Danh mục sản phẩm".');
    }

    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    if (data.length === 0) {
      return [];
    }

    const uniqueProducts = [...new Set(data.flat().filter(item => item && item.trim() !== ''))];
    return uniqueProducts;
  },

  generateFilteredSheet: function (filterCriteria) {
    // Mở file đang chạy
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Sheet gốc để nhân bản báo cáo
    const sourceSheet = ss.getSheetByName('Báo cáo tổng hợp');
    if (!sourceSheet) {
      throw new Error('Không tìm thấy sheet "Báo cáo tổng hợp".');
    }

    // Tách "yyyy-mm" thành [year, month]
    const [year, month] = filterCriteria.monthYear.split('-');
    // Dùng cho việc thay thế nội dung "mm / yyyy" trong ô A5, A7...
    const formattedMonthYear = ` ${month} / ${year}`;
    // Tên sheet báo cáo đích
    const targetSheetName = `Báo cáo ${month}/${year}`;

    // ==================
    // 1) NHÂN BẢN "Báo cáo tổng hợp" => "Báo cáo mm/yyyy"
    // ==================
    let targetSheet = ss.getSheetByName(targetSheetName);
    if (targetSheet) {
      ss.deleteSheet(targetSheet);
    }
    targetSheet = sourceSheet.copyTo(ss);
    // --- THÊM DÒNG SAU để đảm bảo sheet mới không bị ẩn ---
    targetSheet.showSheet();
    //----------------------------------------------------------
    targetSheet.setName(targetSheetName);

    // Ghi tháng/năm vào ô K1 (ví dụ: "01/2025")
    targetSheet.getRange("K1").setValue(`${month}/${year}`);

    // Thay thế "mm/yyyy" trong các ô A5, A7... (nếu có)
    const rangesToReplace = ['A7:I7', 'A5:I5'];
    rangesToReplace.forEach(rangeAddress => {
      try {
        const range = targetSheet.getRange(rangeAddress);
        const cellContent = range.getValue();
        const updatedContent = cellContent.replace(/mm\/yyyy/g, formattedMonthYear);
        range.setValue(updatedContent);
      } catch (error) {
        console.error(`Lỗi khi thay thế nội dung tại dải ô ${rangeAddress}: ${error.message}`);
      }
    });

    // Lấy dữ liệu từ sheet "Danh mục sản phẩm"
    const dataSheet = ss.getSheetByName('Danh mục sản phẩm');
    const allIndexes = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 1).getValues()
      .flat()
      .filter(index => index && index.trim() !== '');
    const allProducts = dataSheet.getRange(2, 2, dataSheet.getLastRow() - 1, 1).getValues()
      .flat();

    // Xác định các mục cần giữ (dựa trên selectedProducts)
    const indexesToKeep = filterCriteria.selectedProducts.map(product => {
      const rowIndex = allProducts.findIndex(item => item.trim() === product.trim());
      return rowIndex !== -1 ? allIndexes[rowIndex] : null;
    }).filter(Boolean);

    // Tạo danh sách các mục cần xóa
    const indexesToDelete = allIndexes.filter(index => {
      return !indexesToKeep.some(keepIndex => index.startsWith(keepIndex));
    });

    console.log('Indexes to Keep:', indexesToKeep);
    console.log('Indexes to Delete:', indexesToDelete);

    // Hủy gộp ô (nếu có) trước khi xóa
    const rangeToBreak = targetSheet.getRange(13, 1, targetSheet.getLastRow() - 12, targetSheet.getLastColumn());
    try {
      const mergedRanges = rangeToBreak.getMergedRanges();
      if (mergedRanges.length > 0) {
        mergedRanges.forEach(mergedRange => mergedRange.breakApart());
      }
    } catch (error) {
      console.error(`Lỗi khi hủy gộp ô: ${error.message}`);
      throw new Error('Không thể hủy gộp ô trước khi thực hiện xóa.');
    }

    // Lấy dữ liệu cột A từ dòng 13 trở đi
    const targetData = targetSheet.getRange(13, 1, targetSheet.getLastRow() - 12, 1).getValues().flat();

    // Xác định các hàng cần xóa
    let rowsToDelete = [];
    targetData.forEach((rowValue, index) => {
      const rowIndex = index + 13;
      const isToDelete = indexesToDelete.some(indexToDelete => rowValue && rowValue.startsWith(indexToDelete));
      if (isToDelete) {
        rowsToDelete.push(rowIndex);
      }
    });

    console.log('Rows to Delete:', rowsToDelete);

    // Xóa các hàng cần xóa (theo nhóm, từ dưới lên)
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // Sắp xếp giảm dần
      while (rowsToDelete.length > 0) {
        const startRow = rowsToDelete[0];
        let count = 1;
        // Kiểm tra các hàng liên tiếp
        for (let i = 1; i < rowsToDelete.length; i++) {
          if (rowsToDelete[i] === startRow - count) {
            count++;
          } else {
            break;
          }
        }
        // Xóa nhóm liên tiếp
        targetSheet.deleteRows(startRow - count + 1, count);
        // Bỏ các hàng vừa xóa khỏi mảng
        rowsToDelete.splice(0, count);
      }
    }

    // ==================
    // 2) NHÂN BẢN "Theo dõi sản lượng trong tháng" => "mm/yyyy"
    // ==================
    const sourceSheet2 = ss.getSheetByName('Theo dõi sản lượng trong tháng');
    if (!sourceSheet2) {
      throw new Error('Không tìm thấy sheet "Theo dõi sản lượng trong tháng".');
    }

    // Tên sheet thứ 2, chỉ ghi "01/2025" chẳng hạn
    const secondSheetName = `${month}/${year}`;

    // Xóa sheet cũ nếu đã tồn tại
    let secondSheet = ss.getSheetByName(secondSheetName);
    if (secondSheet) {
      ss.deleteSheet(secondSheet);
    }

    secondSheet = sourceSheet2.copyTo(ss);
    // --- THÊM DÒNG SAU để đảm bảo sheet mới không bị ẩn ---
    secondSheet.showSheet();
    //----------------------------------------------------------
    secondSheet.setName(secondSheetName);

    // (Nếu cần ghi thông tin tháng/năm vào sheet này, bạn có thể thêm ở đây)

    // Kết quả
    return `Sheet "${sourceSheet.getName()}" đã được nhân bản và lọc dữ liệu. Đồng thời đã tạo sheet "${secondSheetName}" từ "Theo dõi sản lượng trong tháng".`;
  }
};
