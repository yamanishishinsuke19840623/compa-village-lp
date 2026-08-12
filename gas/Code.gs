// =============================================
//  COMPA VILLAGE ゲストハウス — 予約管理 GAS
//  予約リクエスト受付 / リアルタイム空室管理 / メール通知
// =============================================

var HOST_EMAIL = 'jannenenenene@gmail.com';
var CC_EMAIL   = 'yamanishishinsuke19840623@gmail.com';
var SHEET_ID   = '1z8eWW_b7Xa40eRgY7fGZAfVOI58J9Vt1TKFIsHTNQfU';

// 各プランの同時受付可能数（2026-08-12 ゆうさん確認済み：3プラン構成、ファミリールームは廃止し団体プランに統合）
var ROOM_CAPACITY = {
  standard: 3, // スタンダード ツインルーム（個室3室、¥3,250/泊）
  dorm:     4, // ドミトリールーム（4ベッド、1ベッドから予約可、¥2,250/泊）
  group:    1  // 団体様用プラン（1〜15名・一棟貸し感覚、¥27,000/泊）
};
var ROOM_NAMES = {
  standard: 'スタンダード ツインルーム',
  dorm:     'ドミトリールーム',
  group:    '団体様用プラン'
};

// =============================================
//  エントリーポイント
// =============================================

// LP の予約フォームから送信される
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var roomType = d.roomType;
    var checkin  = d.checkin;   // 'YYYY-MM-DD'
    var checkout = d.checkout;  // 'YYYY-MM-DD'

    if (!roomType || !checkin || !checkout || !d.name || !d.email) {
      return res({ok: false, reason: 'invalid_input'});
    }
    if (!ROOM_NAMES[roomType]) {
      return res({ok: false, reason: 'invalid_room'});
    }
    if (!isAvailable(roomType, checkin, checkout)) {
      return res({ok: false, reason: 'unavailable'});
    }

    logBooking(d);
    sendHostNotification(d);
    sendGuestConfirmation(d);

    return res({ok: true});
  } catch (err) {
    return res({ok: false, error: err.toString()});
  }
}

// LP がリアルタイム空室状況を取得する（JSONP）
function doGet(e) {
  var params   = e.parameter || {};
  var callback = params.callback || 'cb';
  var action   = params.action || 'availability';

  if (action === 'availability') {
    var booked = getBookedRanges();
    var json = JSON.stringify({booked: booked});
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  if (action === 'ics') {
    // 将来的にBooking.com/Airbnb側へ取り込んでもらうためのiCal出力（雛形）
    return ContentService.createTextOutput(buildIcs(params.room || ''))
      .setMimeType(ContentService.MimeType.PLAIN_TEXT);
  }

  return ContentService.createTextOutput(callback + '(' + JSON.stringify({ok: true}) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function res(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetUrl() {
  return SHEET_ID ? 'https://docs.google.com/spreadsheets/d/' + SHEET_ID : '（SHEET_ID未設定）';
}

// =============================================
//  初期セットアップ（1回だけ実行）
// =============================================

function setupSheet() {
  var ss    = SpreadsheetApp.create('COMPA VILLAGE 予約台帳');
  var sheet = ss.getActiveSheet();
  sheet.setName('予約台帳');

  var headers = [
    '申込日時', 'お名前', 'メール', '部屋タイプ',
    'チェックイン', 'チェックアウト', '泊数', 'ステータス', '備考'
  ];
  var hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight('bold');
  hRange.setBackground('#1d3420');
  hRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['リクエスト受付', '確定', 'キャンセル'], true).build();
  sheet.getRange('H2:H1000').setDataValidation(statusRule);

  sheet.autoResizeColumns(1, headers.length);

  Logger.log('セットアップ完了');
  Logger.log('Sheet ID: ' + ss.getId());
  Logger.log('Sheet URL: ' + ss.getUrl());
  Logger.log('');
  Logger.log('↑ この Sheet ID を、コード冒頭の SHEET_ID に貼り付けてください');
}

// =============================================
//  空室判定・台帳操作
// =============================================

function isAvailable(roomType, checkin, checkout) {
  var capacity = ROOM_CAPACITY[roomType] || 0;
  if (capacity <= 0) return false;

  var overlapCount = 0;
  var rows = getBookingRows();
  var ci = new Date(checkin);
  var co = new Date(checkout);

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row[3] !== roomType) continue;
    if (row[7] === 'キャンセル') continue;
    var existCi = new Date(row[4]);
    var existCo = new Date(row[5]);
    // 期間が重なっているか（片方の開始が相手の終了より前 かつ 相手の開始が自分の終了より前）
    if (ci < existCo && existCi < co) overlapCount++;
  }
  return overlapCount < capacity;
}

// 部屋タイプごとの予約済み日程一覧（LPのカレンダー表示用）
function getBookedRanges() {
  var rows = getBookingRows();
  var byRoom = {};
  Object.keys(ROOM_NAMES).forEach(function(k){ byRoom[k] = []; });

  rows.forEach(function(row){
    var roomType = row[3];
    if (row[7] === 'キャンセル') return;
    if (!byRoom[roomType]) return;
    byRoom[roomType].push({
      checkin:  formatDate_(row[4]),
      checkout: formatDate_(row[5])
    });
  });
  return byRoom;
}

function getBookingRows() {
  if (!SHEET_ID) return [];
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('予約台帳');
  var values = sheet.getDataRange().getValues();
  return values.slice(1); // ヘッダーを除く
}

function logBooking(d) {
  if (!SHEET_ID) return;
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('予約台帳');
  sheet.appendRow([
    new Date(),
    d.name,
    d.email,
    d.roomType,
    d.checkin,
    d.checkout,
    d.nights || '',
    'リクエスト受付',
    ''
  ]);
}

function formatDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

// =============================================
//  メール通知
// =============================================

function sendHostNotification(d) {
  var body = '新しい予約リクエストが届きました。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += 'お名前　：' + d.name + '\n';
  body += 'メール　：' + d.email + '\n';
  body += '部屋タイプ：' + (ROOM_NAMES[d.roomType] || d.roomType) + '\n';
  body += 'チェックイン：' + d.checkin + '\n';
  body += 'チェックアウト：' + d.checkout + '\n';
  body += '泊数　　：' + (d.nights || '') + '泊\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += 'スプレッドシートの「ステータス」列を「確定」に変更すると、\n';
  body += '確定扱いになり、以後の空室判定にも反映されます。\n\n';
  body += '▶ 予約台帳: ' + sheetUrl();

  MailApp.sendEmail({
    to: HOST_EMAIL,
    cc: CC_EMAIL,
    subject: '【COMPA VILLAGE】予約リクエスト — ' + d.name + '様（' + (ROOM_NAMES[d.roomType] || '') + '）',
    body: body
  });
}

function sendGuestConfirmation(d) {
  var body = d.name + ' 様\n\n';
  body += 'COMPA VILLAGE ゲストハウスへのご予約リクエスト、ありがとうございます！\n';
  body += '以下の内容で承りました。ホストが内容を確認のうえ、改めてご連絡いたします。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '部屋タイプ：' + (ROOM_NAMES[d.roomType] || d.roomType) + '\n';
  body += 'チェックイン：' + d.checkin + '\n';
  body += 'チェックアウト：' + d.checkout + '\n';
  body += '泊数　　：' + (d.nights || '') + '泊\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += 'ご不明な点があれば、このメールに返信いただくか、\n';
  body += 'Instagram（@compa0601）のDMでもお気軽にご連絡ください。\n\n';
  body += '旅の途中に、最高の出会いを。\n';
  body += 'COMPA VILLAGE\n';

  MailApp.sendEmail({
    to: d.email,
    replyTo: HOST_EMAIL,
    subject: '【COMPA VILLAGE】ご予約リクエストを受け付けました',
    body: body
  });
}

// =============================================
//  スプレッドシート カスタムメニュー
// =============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌵 COMPA予約')
    .addItem('✅ 選択行を「確定」にして確定メールを送信', 'menuConfirmBooking')
    .addToUi();
}

function menuConfirmBooking() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert('データ行を選択してください'); return; }

  var data = sheet.getRange(row, 1, 1, 9).getValues()[0];
  var d = {
    name: data[1], email: data[2], roomType: data[3],
    checkin: formatDate_(data[4]), checkout: formatDate_(data[5]), nights: data[6]
  };
  if (!d.email) { SpreadsheetApp.getUi().alert('メールアドレスが空です'); return; }

  var confirm = SpreadsheetApp.getUi().alert(
    d.name + ' 様（' + (ROOM_NAMES[d.roomType] || d.roomType) + '）を確定にしますか？',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  sheet.getRange(row, 8).setValue('確定');

  var body = d.name + ' 様\n\n';
  body += 'ご予約が確定しましたのでご連絡いたします。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '部屋タイプ：' + (ROOM_NAMES[d.roomType] || d.roomType) + '\n';
  body += 'チェックイン：' + d.checkin + '\n';
  body += 'チェックアウト：' + d.checkout + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += '当日を楽しみにお待ちしております！\n\n';
  body += 'COMPA VILLAGE\n';

  MailApp.sendEmail({
    to: d.email, replyTo: HOST_EMAIL,
    subject: '【COMPA VILLAGE】ご予約が確定しました',
    body: body
  });

  SpreadsheetApp.getUi().alert('確定メールを送信しました');
}

// =============================================
//  iCal 出力（将来、他OTAへ取り込んでもらう用の雛形）
// =============================================

function buildIcs(roomType) {
  var rows = getBookingRows();
  var lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//COMPA VILLAGE//Booking//JA'];
  rows.forEach(function(row, i){
    if (roomType && row[3] !== roomType) return;
    if (row[7] === 'キャンセル') return;
    var ci = formatDate_(row[4]).replace(/-/g, '');
    var co = formatDate_(row[5]).replace(/-/g, '');
    lines.push('BEGIN:VEVENT');
    lines.push('UID:compa-village-' + i + '@compa-village-lp');
    lines.push('DTSTART;VALUE=DATE:' + ci);
    lines.push('DTEND;VALUE=DATE:' + co);
    lines.push('SUMMARY:予約あり（' + (ROOM_NAMES[row[3]] || row[3]) + '）');
    lines.push('END:VEVENT');
  });
  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}
