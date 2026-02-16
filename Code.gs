function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ระบบจัดการอุปกรณ์')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPageHTML(pageName) {
  try {
    return HtmlService.createTemplateFromFile(pageName).evaluate().getContent();
  } catch (e) {
    return "<div class='alert alert-danger'>ไม่พบไฟล์ HTML ที่ชื่อว่า: <b>" + pageName + "</b><br>โปรดตรวจสอบการตั้งชื่อไฟล์ที่แถบด้านซ้ายมือ</div>";
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUserData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('User');
  const values = sheet.getDataRange().getValues();
  const header = values.shift(); 
  return values; 
}

function getEquipmentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Equipment List');
  return sheet.getDataRange().getValues();
}