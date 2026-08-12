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
// 確定メールの決済案内に使うStripe本番決済リンク（2026-08-12 やまちゃん確認済み、index.htmlのPAYMENT_LINKSと同一）
var STRIPE_LINKS = {
  standard: 'https://buy.stripe.com/14A8wOewD4yC4mG0qM87K01',
  dorm:     'https://buy.stripe.com/00w5kCfAHc147ySa1m87K03',
  group:    'https://buy.stripe.com/8x24gygELc14f1k7Te87K02'
};
var ACCESS_MAP_URL = 'https://www.google.com/maps?q=COMPA+VILLAGE+%E4%B8%8B%E9%96%A2%E5%B8%82';
var IG_URL = 'https://www.instagram.com/compa0601/';

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

// 団体様用プラン（一棟貸し感覚）は建物全体（standard/dormと同じ部屋）を使うため、
// 「団体プランが入っている日はstandard/dormも予約不可」「standard/dormが1件でも入っていたら団体プランは予約不可」
// というルールを両方向でチェックする（2026-08-12 打合せで確認済みの業務ルール）
var GROUND_FLOOR_TYPES = ['standard', 'dorm', 'group'];

function isAvailable(roomType, checkin, checkout) {
  var capacity = ROOM_CAPACITY[roomType] || 0;
  if (capacity <= 0) return false;

  var ci = new Date(checkin);
  var co = new Date(checkout);
  var rows = getBookingRows();

  if (roomType === 'group') {
    // standard/dorm/group、いずれかで重なる予約が1件でもあれば不可
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (row[7] === 'キャンセル') continue;
      if (GROUND_FLOOR_TYPES.indexOf(row[3]) === -1) continue;
      var existCi = new Date(row[4]);
      var existCo = new Date(row[5]);
      if (ci < existCo && existCi < co) return false;
    }
    return true;
  }

  // standard/dorm: 同タイプの重複数がcapacity未満、かつ団体プランと重なっていないこと
  var overlapCount = 0;
  for (var j = 0; j < rows.length; j++) {
    var r = rows[j];
    if (r[7] === 'キャンセル') continue;
    var eCi = new Date(r[4]);
    var eCo = new Date(r[5]);
    if (!(ci < eCo && eCi < co)) continue; // 重なっていない
    if (r[3] === 'group') return false; // 団体プランと重なったら即不可
    if (r[3] === roomType) overlapCount++;
  }
  return overlapCount < capacity;
}

// 部屋タイプごとの予約済み日程一覧（LPのカレンダー表示用）
// 団体プランはstandard/dormにも波及し、standard/dormは団体プランにも波及する
function getBookedRanges() {
  var rows = getBookingRows();
  var byRoom = {};
  Object.keys(ROOM_NAMES).forEach(function(k){ byRoom[k] = []; });

  rows.forEach(function(row){
    var roomType = row[3];
    if (row[7] === 'キャンセル') return;
    if (!byRoom[roomType]) return;
    var range = { checkin: formatDate_(row[4]), checkout: formatDate_(row[5]) };

    if (roomType === 'group') {
      GROUND_FLOOR_TYPES.forEach(function(t){ byRoom[t].push(range); });
    } else {
      byRoom[roomType].push(range);
      if (byRoom['group']) byRoom['group'].push(range);
    }
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
    d.guests ? ('人数: ' + d.guests + '名') : ''
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
  body += '人数　　：' + (d.guests || '') + '名\n';
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
  body += '人数　　：' + (d.guests || '') + '名\n';
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

  MailApp.sendEmail({
    to: d.email, replyTo: HOST_EMAIL,
    subject: '【COMPA VILLAGE】ご予約が確定しました',
    body: buildConfirmationText_(d),
    htmlBody: buildConfirmationHtml_(d)
  });

  SpreadsheetApp.getUi().alert('確定メールを送信しました');
}

// 確定メール：プレーンテキスト版（HTML非対応クライアント向けフォールバック）
function buildConfirmationText_(d) {
  var payLink = STRIPE_LINKS[d.roomType];
  var body = d.name + ' 様\n\n';
  body += 'ご予約が確定しましたのでご連絡いたします。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '部屋タイプ：' + (ROOM_NAMES[d.roomType] || d.roomType) + '\n';
  body += 'チェックイン：' + d.checkin + '\n';
  body += 'チェックアウト：' + d.checkout + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += '【チェックインについて】\n';
  body += 'チェックイン：16:00〜22:00 / チェックアウト：10:00まで\n';
  body += '駐車場は無料でご利用いただけます。全室・共用スペースで高速無料WiFiもご利用いただけます。\n\n';
  body += '【アクセス】\n';
  body += '山口県下関市（下関市立美術館・長府庭園エリア）／最寄り：JR新下関駅\n';
  body += ACCESS_MAP_URL + '\n\n';
  if (payLink) {
    body += '【お支払い】\n';
    body += 'まだお支払いがお済みでない場合は、以下より決済をお願いいたします。\n';
    body += payLink + '\n\n';
  }
  body += 'ご不明な点があれば、このメールに返信いただくか、\n';
  body += 'Instagram（@compa0601）のDMでもお気軽にご連絡ください。\n\n';
  body += '当日を楽しみにお待ちしております！\n\n';
  body += 'COMPA VILLAGE\n' + IG_URL;
  return body;
}

// 確定メール：HTML版（ブランドカラーで整形。Gmail等HTML対応クライアントで表示）
function buildConfirmationHtml_(d) {
  var roomName = ROOM_NAMES[d.roomType] || d.roomType;
  var payLink = STRIPE_LINKS[d.roomType];
  var esc = function(s){ return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };

  var payButton = '';
  if (payLink) {
    payButton =
      '<tr><td style="padding:24px 28px 0;">' +
        '<p style="margin:0 0 10px; font-size:14px; color:#1F2B1E;">まだお支払いがお済みでない場合は、以下より決済をお願いいたします。</p>' +
        '<a href="' + esc(payLink) + '" style="display:inline-block; background:#8FBE3F; color:#1D3420; text-decoration:none; font-weight:bold; padding:12px 22px; border-radius:999px; font-size:14px;">お支払いはこちら →</a>' +
      '</td></tr>';
  }

  return '' +
  '<div style="font-family:\'Hiragino Sans\',\'Yu Gothic\',sans-serif; background:#F6F1E4; padding:28px 16px;">' +
    '<table role="presentation" width="100%" style="max-width:520px; margin:0 auto; background:#ffffff; border-radius:16px; overflow:hidden; border:1px solid #DCCFAE;">' +
      '<tr><td style="background:#1D3420; padding:22px 28px;">' +
        '<p style="margin:0; color:#8FBE3F; font-size:12px; letter-spacing:.14em; text-transform:uppercase;">Booking Confirmed</p>' +
        '<p style="margin:6px 0 0; color:#F6F1E4; font-size:20px; font-weight:bold;">COMPA VILLAGE</p>' +
      '</td></tr>' +
      '<tr><td style="padding:26px 28px 0;">' +
        '<p style="margin:0 0 4px; font-size:15px; color:#1F2B1E;">' + esc(d.name) + ' 様</p>' +
        '<p style="margin:0; font-size:14px; color:rgba(31,43,30,.75); line-height:1.7;">ご予約が確定しましたのでご連絡いたします。当日を楽しみにお待ちしております！</p>' +
      '</td></tr>' +
      '<tr><td style="padding:20px 28px 0;">' +
        '<table role="presentation" width="100%" style="background:#EFE7D3; border-radius:12px;">' +
          '<tr><td style="padding:16px 20px; font-size:13px; color:#1F2B1E;">部屋タイプ</td><td style="padding:16px 20px; font-size:13px; color:#1D3420; font-weight:bold; text-align:right;">' + esc(roomName) + '</td></tr>' +
          '<tr><td style="padding:0 20px 16px; font-size:13px; color:#1F2B1E;">チェックイン</td><td style="padding:0 20px 16px; font-size:13px; color:#1D3420; font-weight:bold; text-align:right;">' + esc(d.checkin) + '</td></tr>' +
          '<tr><td style="padding:0 20px 16px; font-size:13px; color:#1F2B1E;">チェックアウト</td><td style="padding:0 20px 16px; font-size:13px; color:#1D3420; font-weight:bold; text-align:right;">' + esc(d.checkout) + '</td></tr>' +
        '</table>' +
      '</td></tr>' +
      '<tr><td style="padding:20px 28px 0;">' +
        '<p style="margin:0 0 6px; font-size:13px; font-weight:bold; color:#1D3420;">チェックインについて</p>' +
        '<p style="margin:0; font-size:13px; color:rgba(31,43,30,.75); line-height:1.8;">チェックイン 16:00〜22:00 ／ チェックアウト 10:00まで<br>駐車場：無料 ／ WiFi：全室・共用スペースで高速無料</p>' +
      '</td></tr>' +
      '<tr><td style="padding:18px 28px 0;">' +
        '<p style="margin:0 0 6px; font-size:13px; font-weight:bold; color:#1D3420;">アクセス</p>' +
        '<p style="margin:0; font-size:13px; color:rgba(31,43,30,.75); line-height:1.8;">山口県下関市（下関市立美術館・長府庭園エリア）／最寄り：JR新下関駅<br><a href="' + esc(ACCESS_MAP_URL) + '" style="color:#2B4A2E;">地図で見る →</a></p>' +
      '</td></tr>' +
      payButton +
      '<tr><td style="padding:24px 28px 26px;">' +
        '<p style="margin:0; font-size:12px; color:rgba(31,43,30,.6); line-height:1.7;">ご不明な点があれば、このメールに返信いただくか、Instagram（<a href="' + esc(IG_URL) + '" style="color:#2B4A2E;">@compa0601</a>）のDMでもお気軽にご連絡ください。<br><br>旅の途中に、最高の出会いを。<br>COMPA VILLAGE</p>' +
      '</td></tr>' +
    '</table>' +
  '</div>';
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
