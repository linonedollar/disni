/**
 * 新加坡 & 迪士尼郵輪行程 - Google Sheet 雲端資料庫後端 (Google Apps Script)
 * 試算表 ID: 1i3Wa7EoNEEmZRHUSDXKVfwYTVVqhzDyQnZUQVvcdjmA
 * 
 * 部署說明：
 * 1. 打開 Google 試算表 -> 擴充功能 -> Apps Script
 * 2. 清空內容並貼上此程式碼
 * 3. 點選右上角「部署」->「新增部署」
 * 4. 種類選擇「網頁應用程式 (Web app)」
 * 5. 誰可以存取：選擇「所有人 (Anyone)」
 * 6. 點擊「部署」，並複製 Web App 網址貼回旅遊網頁中即可！
 */

function setupHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = ["天數", "任務ID", "分類", "項目名稱", "狀態(已完成)", "最後更新時間", "備註"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
  sheet.setFrozenRows(1);
}

// GET: 讀取所有天數的任務完成狀態
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    // 如果是全新空白表，自動建立表頭
    if (lastRow === 0) {
      setupHeaders();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        data: {},
        message: "Sheet initialized with headers"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = sheet.getDataRange().getValues();
    var result = {};
    
    // 判斷第1列是否為表頭
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // 支援兩種格式：有「天數」欄位（A欄）或以「任務ID」為A欄
      var day = row[0] ? String(row[0]).trim() : "Day 1";
      var taskId = row[1] ? String(row[1]).trim() : (row[0] ? String(row[0]).trim() : "");
      var status = false;
      
      // 尋找布林值或 TRUE 狀態
      for (var col = 2; col < row.length; col++) {
        if (typeof row[col] === "boolean") {
          status = row[col];
          break;
        } else if (String(row[col]).toUpperCase() === "TRUE") {
          status = true;
          break;
        }
      }
      
      if (taskId) {
        result[taskId] = status;
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      tasks: result,
      totalCount: Object.keys(result).length
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// POST: 記錄打卡狀態（新增或更新）
function doPost(e) {
  try {
    var postData;
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else {
      postData = e.parameter;
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow === 0) {
      setupHeaders();
      lastRow = 1;
    }
    
    var day = postData.day || "Day 1";
    var taskId = postData.id || postData.taskId || "";
    var category = postData.category || "行程";
    var title = postData.title || "";
    var checked = (postData.checked === true || postData.checked === "true" || postData.checked === "TRUE");
    var note = postData.note || "";
    var timeStr = Utilities.formatDate(new Date(), "Asia/Singapore", "yyyy-MM-dd HH:mm:ss");
    
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // 依據任務ID尋找該列
    for (var i = 1; i < data.length; i++) {
      // 檢查欄B（任務ID）或欄A（如果舊表無天數欄）
      if (data[i][1] == taskId || data[i][0] == taskId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex > 0) {
      // 更新現有列
      // 假設結構: A:天數, B:任務ID, C:分類, D:項目名稱, E:狀態, F:最後更新時間, G:備註
      sheet.getRange(rowIndex, 1).setValue(day);
      sheet.getRange(rowIndex, 2).setValue(taskId);
      if (category) sheet.getRange(rowIndex, 3).setValue(category);
      if (title) sheet.getRange(rowIndex, 4).setValue(title);
      sheet.getRange(rowIndex, 5).setValue(checked);
      sheet.getRange(rowIndex, 6).setValue(timeStr);
      if (note) sheet.getRange(rowIndex, 7).setValue(note);
    } else {
      // 新增一列
      sheet.appendRow([day, taskId, category, title, checked, timeStr, note]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      updatedId: taskId,
      checked: checked,
      updatedAt: timeStr
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
