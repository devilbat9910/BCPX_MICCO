function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Tự Động Báo Cáo')
    .addItem('Tạo báo cáo', 'showFilterDialog')
    .addItem('Thu gọn báo cáo', 'shrinkReportUI')
    .addItem('Gửi báo cáo', 'sendReportUI') // Thêm mục gửi báo cáo
    //.addItem('Tô màu', 'highlightCells')
    .addItem('Báo cáo cuối tháng', 'showConsolidateReportDialog')
    .addToUi();
}

function showFilterDialog() {
  const html = HtmlService.createHtmlOutputFromFile('FilterDialog')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tạo báo cáo mới');
}

function getProductList() {
  return Helpers.getProductList();
}

function generateAndCopySheet(filterCriteria) {
  return Helpers.generateFilteredSheet(filterCriteria);
}
