// ============================================
// 設定：你嘅通知email同Web App URL
// ============================================
var NOTIFY_EMAIL = 'ken316092@gmail.com';
var WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyquusxDm9OooyQFJwVKdf5Gd5awO0cFWeDCmyRVr4t6TTF_WGbz8qDFgeQmNHzdYNv/exec';

// ============================================
// 主要功能：接收賓客RSVP，寫入Sheet + 通知你
// ============================================
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.attend || '',
      data.adults || '',
      data.children || '',
      data.highchair || '',
      data.side || ''          // ← 新增：親友方（G欄）
    ]);

    var attendText = data.attend === 'yes' ? '參加'
                    : data.attend === 'no' ? '不參加'
                    : '待定';

    // 冇小朋友嘅話前端會傳空白，唔可以當成「否」
    var highchairText = data.highchair === 'yes' ? '是'
                      : data.highchair === 'no' ? '否'
                      : '－（沒有小孩）';

    var sideText = data.side || '（未填）';

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      // 主旨直接帶埋親友方，兩個同名賓客一眼分得出
      subject: '【婚禮RSVP】新回覆：' + (data.name || '（未填姓名）') + '（' + sideText + '）',
      body:
        '你收到一個新嘅婚禮出席回覆：\n\n' +
        '姓名：' + (data.name || '') + '\n' +
        '親友方：' + sideText + '\n' +
        '是否參加：' + attendText + '\n' +
        '成人人數：' + (data.adults || '') + '\n' +
        '小孩人數：' + (data.children || '') + '\n' +
        '需要兒童餐椅：' + highchairText + '\n\n' +
        '（完整記錄已經自動存入Google Sheet）'
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // 一旦寫入或送email中途出錯，即刻警報俾你，連原始資料都夾埋
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: '⚠️【緊急】RSVP寫入失敗，請盡快檢查！',
        body:
          '有賓客提交咗RSVP，但系統寫入Sheet時發生錯誤：\n\n' +
          '錯誤內容：' + err.message + '\n\n' +
          '賓客原始提交資料（請手動補回Sheet）：\n' +
          (e && e.postData ? e.postData.contents : '（無法取得原始資料）')
      });
    } catch (mailErr) {
      // 連send email都失敗嘅話，都冇辦法，但唔應該再拋錯畀前端
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// 每日自動健康檢查：模擬賓客提交，驗證成個流程正常
// ============================================
function dailyHealthCheck() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var testMarker = '[SYSTEM_TEST]';
  var rowsBefore = sheet.getLastRow();

  var testPayload = {
    name: testMarker,
    side: '兩邊都認識',      // ← 新增：等健康檢查行埋新欄位條路
    attend: 'yes',
    adults: '1',
    children: '0',
    highchair: ''            // ← 配合前端：冇小朋友就傳空白
  };

  var success = false;
  var errorDetail = '';

  try {
    var response = UrlFetchApp.fetch(WEB_APP_URL, {
      method: 'post',
      contentType: 'text/plain;charset=utf-8',
      payload: JSON.stringify(testPayload),
      muteHttpExceptions: true
    });

    var responseText = response.getContentText();
    var rowsAfter = sheet.getLastRow();

    // 判斷是否成功：HTTP回應要正常，而且真係加多咗一行
    if (response.getResponseCode() === 200 && rowsAfter > rowsBefore) {
      var lastRowData = sheet.getRange(rowsAfter, 1, 1, 7).getValues()[0];   // ← 6改7（多咗G欄）
      if (lastRowData[1] === testMarker) {
        success = true;
        // 測試通過，清走呢一行測試資料，keep個Sheet乾淨
        sheet.deleteRow(rowsAfter);
      }
    }

    if (!success) {
      errorDetail = 'HTTP狀態: ' + response.getResponseCode() + '\n回應內容: ' + responseText;
    }

  } catch (err) {
    errorDetail = err.message;
  }

  // 更新狀態記錄（寫喺H1/H2/I1/I2，方便你隨時打開Sheet睇一眼就知狀態）
  sheet.getRange('H1').setValue('系統狀態');
  sheet.getRange('I1').setValue('上次檢查時間');

  if (success) {
    sheet.getRange('H2').setValue('✅ 正常運作');
    sheet.getRange('I2').setValue(new Date());
  } else {
    sheet.getRange('H2').setValue('❌ 檢查失敗');
    sheet.getRange('I2').setValue(new Date());

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: '⚠️【緊急】婚禮RSVP系統健康檢查失敗！',
      body:
        '自動健康檢查發現RSVP系統可能已經失效，請盡快檢查：\n\n' +
        '錯誤詳情：\n' + errorDetail + '\n\n' +
        '建議檢查步驟：\n' +
        '1. 打開Apps Script編輯器，檢查「執行項目」有冇報錯\n' +
        '2. 檢查Deploy狀態係咪仍然Active\n' +
        '3. 手動打開網站測試一次RSVP表格\n'
    });
  }
}
