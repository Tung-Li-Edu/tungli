/**
 * 東立官網・預約試聽表單
 * 收到官網送出的預約 → 存進這份試算表的「預約試聽」工作表 → 用東立 LINE 官方帳號推播通知主任
 *
 * 需要在「專案設定 → 指令碼屬性」新增兩個屬性：
 *   LINE_TOKEN    ：LINE 官方帳號的 Channel access token（長期）
 *   LINE_USER_ID  ：要收通知的人的 LINE userId（U 開頭的一長串）
 */

const SHEET_NAME = '預約試聽';
const HEADERS = ['送出時間', '來源', '家長姓名', '聯絡電話', '學生年級', '有興趣的課程', '想詢問的事情', '處理狀態'];

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d._honey) return json_({ success: true }); // 機器人填了隱藏欄位，直接略過

    const clean = s => String(s || '').trim().slice(0, 500);
    const row = {
      parent: clean(d.parent),
      phone: clean(d.phone),
      grade: clean(d.grade),
      course: clean(d.course),
      msg: clean(d.msg),
      source: clean(d.source) || '電腦'
    };
    if (!row.parent || !row.phone) return json_({ success: false, message: '缺少姓名或電話' });

    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
    const sh = sheet_();
    sh.appendRow([now, row.source, row.parent, "'" + row.phone, row.grade, row.course, row.msg, '未聯絡']);

    try {
      pushLine_([
        '📩 官網有新的預約試聽',
        '時間：' + now + '（' + row.source + '）',
        '家長：' + row.parent,
        '電話：' + row.phone,
        '年級：' + (row.grade || '（未填）'),
        '課程：' + (row.course || '（未填）'),
        '詢問：' + (row.msg || '（未填）'),
        row.source === '手機' ? '※ 家長是用手機送出，可能也會從 LINE 傳同一筆給你' : ''
      ].filter(Boolean).join('\n'));
    } catch (err) {
      console.error('LINE 通知失敗：' + err); // 資料已存進試算表，通知失敗不影響家長送出
    }
    return json_({ success: true });
  } catch (err) {
    console.error(err);
    return json_({ success: false, message: String(err) });
  }
}

// 用瀏覽器打開網址時顯示，確認部署成功
function doGet() {
  return json_({ ok: true, service: '東立官網預約試聽' });
}

/** 第一次先執行這個：建立工作表和標題列 */
function setup() {
  sheet_();
}

/** 執行這個測試 LINE 通知有沒有設定好 */
function testLine() {
  pushLine_('✅ 東立官網預約通知測試成功！之後有家長預約試聽，會傳到這裡。');
}

/** token 有問題時執行這個：會在下方「執行記錄」顯示檢查結果（不會顯示完整 token） */
function checkLine() {
  const p = PropertiesService.getScriptProperties();
  const raw = p.getProperty('LINE_TOKEN') || '';
  const token = raw.replace(/^Bearer\s+/i, '').replace(/["'\s]/g, '');
  const to = (p.getProperty('LINE_USER_ID') || '').replace(/["'\s]/g, '');
  console.log('指令碼屬性名稱：' + Object.keys(p.getProperties()).join('、'));
  console.log('LINE_TOKEN 長度 ' + token.length + '，開頭 ' + token.slice(0, 4) + '…結尾 ' + token.slice(-4) + (raw !== token ? '（原本有多餘空白或符號，已自動去掉）' : ''));
  console.log('LINE_USER_ID：' + (to ? to.slice(0, 3) + '…（長度 ' + to.length + '，正常是 U 開頭、33 個字）' : '還沒設定'));
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
    headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (res.getResponseCode() === 200) {
    const info = JSON.parse(res.getContentText());
    console.log('✅ token 正確，對應的官方帳號：' + info.displayName + '（' + info.basicId + '）');
  } else {
    console.log('❌ token 不正確，LINE 回應 ' + res.getResponseCode() + '：' + res.getContentText());
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#F3E7D6');
  }
  return sh;
}

function pushLine_(text) {
  const p = PropertiesService.getScriptProperties();
  const token = (p.getProperty('LINE_TOKEN') || '').replace(/^Bearer\s+/i, '').replace(/["'\s]/g, '');
  const to = (p.getProperty('LINE_USER_ID') || '').replace(/["'\s]/g, '');
  if (!token || !to) throw new Error('還沒設定 LINE_TOKEN 或 LINE_USER_ID');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('LINE 回應 ' + res.getResponseCode() + '：' + res.getContentText());
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
