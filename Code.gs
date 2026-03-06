function doGet(e) {
  // รับค่าหน้าจาก URL ถ้าไม่มีค่าหรือเป็นค่าว่างให้ไปที่ Home
  var page = e.parameter.p || 'Home';
  
  return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle('ระบบจัดการอุปกรณ์')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}