// =============================================
//  COMPA VILLAGE ゲストハウス — 予約管理 GAS
//  予約リクエスト受付 / リアルタイム空室管理 / メール通知
// =============================================

var HOST_EMAIL = 'jannenenenene@gmail.com';
var CC_EMAIL   = 'yamanishishinsuke19840623@gmail.com';
var SHEET_ID   = '1z8eWW_b7Xa40eRgY7fGZAfVOI58J9Vt1TKFIsHTNQfU';

// 各プランの同時受付可能数（2026-08-12 ゆうさん確認済み：3プラン構成、ファミリールームは廃止し団体プランに統合）
// 2026-08-13 団体様用プランは床面積の都合で15名収容が不可と判明したため「一棟貸しプラン」に改称、10名様まで・¥30,000/泊に変更
// 2026-08-19 一棟貸しプランの料金を「6名様まで¥20,000/泊、7名様以降1名ごとに+¥3,000」に変更（最大10名様は変更なし）
var ROOM_CAPACITY = {
  standard: 3, // スタンダード ツインルーム（個室3室、¥3,250/泊）
  dorm:     4, // ドミトリールーム（4ベッド、1ベッドから予約可、¥2,250/泊）
  group:    1  // 一棟貸しプラン（10名様まで・一棟貸し、6名様まで¥20,000/泊+追加1名ごとに¥3,000）
};
var ROOM_NAMES = {
  standard: 'スタンダード ツインルーム',
  dorm:     'ドミトリールーム',
  group:    '一棟貸しプラン'
};
// 確定メールの決済案内に使うStripe本番決済リンク（2026-08-19更新、index.htmlのPAYMENT_LINKSと同一）
// groupは基本料金¥20,000(数量固定)+追加人数料金¥3,000(数量調整可・最大4)の2商品を含む1本のリンク
var STRIPE_LINKS = {
  standard: 'https://buy.stripe.com/14A8wOewD4yC4mG0qM87K01',
  dorm:     'https://buy.stripe.com/00w5kCfAHc147ySa1m87K03',
  group:    'https://buy.stripe.com/dRmaEW3RZ7KO06qc9u87K05'
};
var ACCESS_MAP_URL = 'https://www.google.com/maps?q=COMPA+VILLAGE+%E4%B8%8B%E9%96%A2%E5%B8%82';
var IG_URL = 'https://www.instagram.com/compa0601/';
// StripeのWebhookエンドポイントURLに ?wh=このトークン を付けて登録する簡易認証（2026-08-15追加）
// GASのdoPost(e)はStripe-Signatureヘッダーを読み取れないため、標準的な署名検証の代わりに使用
var STRIPE_WEBHOOK_SECRET = '2922dae4fb4470280cb48a0ace5778231f81d463c94879e0';

// =============================================
//  エントリーポイント
// =============================================

// LP の予約フォームから送信される
// kind === 'stripe_pending' の場合は、スタンダード/ドミトリーが直接Stripe決済へ進む際の
// 空室確保用の仮予約ログ（名前・メール不要、確定メールは送らない。2026-08-13追加）
function doPost(e) {
  try {
    if (e.parameter && e.parameter.wh === STRIPE_WEBHOOK_SECRET) {
      return handleStripeWebhook_(e);
    }

    var d = JSON.parse(e.postData.contents);
    var roomType = d.roomType;
    var checkin  = d.checkin;   // 'YYYY-MM-DD'
    var checkout = d.checkout;  // 'YYYY-MM-DD'
    var isPending = d.kind === 'stripe_pending';

    if (!roomType || !checkin || !checkout || (!isPending && (!d.name || !d.email))) {
      return res({ok: false, reason: 'invalid_input'});
    }
    if (!ROOM_NAMES[roomType]) {
      return res({ok: false, reason: 'invalid_room'});
    }
    if (!isAvailable(roomType, checkin, checkout)) {
      return res({ok: false, reason: 'unavailable'});
    }

    if (isPending) {
      logPendingPayment(d);
      sendPendingPaymentNotification(d);
    } else {
      logBooking(d);
      sendHostNotification(d);
      sendGuestConfirmation(d);
    }

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
    .requireValueInList(['リクエスト受付', '決済へ進行中', '確定', 'キャンセル', 'オーナーブロック'], true).build();
  sheet.getRange('H2:H1000').setDataValidation(statusRule);

  sheet.autoResizeColumns(1, headers.length);

  Logger.log('セットアップ完了');
  Logger.log('Sheet ID: ' + ss.getId());
  Logger.log('Sheet URL: ' + ss.getUrl());
  Logger.log('');
  Logger.log('↑ この Sheet ID を、コード冒頭の SHEET_ID に貼り付けてください');
}

// このGASはsetupSheet()で新規作成した台帳スプレッドシートに「コンテナバインド」されていない
// スタンドアロン型スクリプトのため、onOpen(e)は自動発火せず「🌵 COMPA予約」メニューが出ない。
// インストール型トリガーで明示的にこの台帳へ紐付ける一度きりの関数（関数選択で実行すればOK）
function installOnOpenTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onOpen') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onOpen').forSpreadsheet(SHEET_ID).onOpen().create();
  Logger.log('台帳スプレッドシートのonOpenトリガーを登録しました。台帳を開き直すとメニューが表示されます');
}

// 既に稼働中の台帳（setupSheetは初回しか使えないため）のステータス列プルダウンに
// 「オーナーブロック」を追加する一度きりの移行用関数。実行メニューから1回だけ動かせばOK
// （動かさなくても、スクリプトからの書き込みはプルダウンの制約を受けないため動作に支障はない）
function addBlockStatusToValidation() {
  if (!SHEET_ID) return;
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('予約台帳');
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['リクエスト受付', '決済へ進行中', '確定', 'キャンセル', 'オーナーブロック'], true).build();
  sheet.getRange('H2:H1000').setDataValidation(statusRule);
  Logger.log('ステータスのプルダウンに「オーナーブロック」を追加しました');
}

// =============================================
//  空室判定・台帳操作
// =============================================

// 一棟貸しプラン（旧・団体様用プラン）は建物全体（standard/dormと同じ部屋）を使うため、
// 「一棟貸しプランが入っている日はstandard/dormも予約不可」「standard/dormが1件でも入っていたら一棟貸しプランは予約不可」
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

// スタンダード/ドミトリー用：決済ページを開くのと同時に記録する仮予約
// お名前・メールはStripe Webhook（handleStripeWebhook_）が決済完了時に自動で埋める。
// refIdは10列目に保存し、WebhookからのStripeイベントとこの行を突き合わせるためのキーになる
function logPendingPayment(d) {
  if (!SHEET_ID) return;
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('予約台帳');
  sheet.appendRow([
    new Date(),
    '（Stripe決済へ・お名前未取得）',
    '',
    d.roomType,
    d.checkin,
    d.checkout,
    d.nights || '',
    '決済へ進行中',
    d.guests ? ('数量: ' + d.guests) : '',
    d.refId || ''
  ]);
}

// 予約台帳の10列目（refId）を全行スキャンして一致する行番号(1-based)を返す。無ければnull
function findRowByRefId_(refId) {
  if (!SHEET_ID || !refId) return null;
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('予約台帳');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][9] || '') === refId) return i + 1;
  }
  return null;
}

// StripeのWebhook（checkout.session.completed）を受け取り、該当する仮予約行に
// 実際のお名前・メールを書き込んで「確定」にする
function handleStripeWebhook_(e) {
  var event = JSON.parse(e.postData.contents);
  if (event.type !== 'checkout.session.completed') {
    return res({ok: true, ignored: true});
  }

  var session = event.data && event.data.object;
  var refId = session && session.client_reference_id;
  if (!refId) return res({ok: false, reason: 'no_ref_id'});

  var row = findRowByRefId_(refId);
  if (!row) return res({ok: false, reason: 'row_not_found'});

  var customerName  = (session.customer_details && session.customer_details.name)  || '';
  var customerEmail = (session.customer_details && session.customer_details.email) || '';

  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('予約台帳');
  if (customerName)  sheet.getRange(row, 2).setValue(customerName);
  if (customerEmail) sheet.getRange(row, 3).setValue(customerEmail);
  sheet.getRange(row, 8).setValue('確定');

  var rowData = sheet.getRange(row, 1, 1, 9).getValues()[0];
  var confirmedGuest = {
    name: customerName || rowData[1],
    email: customerEmail,
    roomType: rowData[3],
    checkin: formatDate_(rowData[4]),
    checkout: formatDate_(rowData[5])
  };
  sendPaymentConfirmedNotification_(confirmedGuest);

  // 即決済型（standard/dorm）はこれまでゲストへのご案内メールが無かったため、
  // 決済完了と同時にチェックイン詳細等の案内メールを送る（2026-08-19追加）
  if (customerEmail) {
    var guideData = {
      name: confirmedGuest.name, roomType: confirmedGuest.roomType,
      checkin: confirmedGuest.checkin, checkout: confirmedGuest.checkout,
      alreadyPaid: true
    };
    MailApp.sendEmail({
      to: customerEmail,
      replyTo: HOST_EMAIL,
      subject: '【COMPA VILLAGE】ご予約が確定しました',
      body: buildConfirmationText_(guideData),
      htmlBody: buildConfirmationHtml_(guideData)
    });
  }

  return res({ok: true});
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

// スタンダード/ドミトリー用：決済ページへの遷移があったことをホストへ軽く知らせる（お客様へのメールなし）
function sendPendingPaymentNotification(d) {
  var body = 'Stripe決済ページへ進んだ方がいます（お名前・メールはまだ取得できません）。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '部屋タイプ：' + (ROOM_NAMES[d.roomType] || d.roomType) + '\n';
  body += 'チェックイン：' + d.checkin + '\n';
  body += 'チェックアウト：' + d.checkout + '\n';
  body += '泊数　　：' + (d.nights || '') + '泊\n';
  body += '数量　　：' + (d.guests || '') + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += '実際にStripeで決済が完了すると、お名前・メールが自動で予約台帳に書き込まれ、\n';
  body += 'ステータスも自動で「確定」になります（このメール以降、通常は操作不要です）。\n\n';
  body += 'もし決済されずこのまま放置された場合は、スプレッドシートの該当行を選んで\n';
  body += '「🌵 COMPA予約」→「選択行の決済保留を解除する」で空室に戻せます。\n\n';
  body += '▶ 予約台帳: ' + sheetUrl();

  MailApp.sendEmail({
    to: HOST_EMAIL,
    cc: CC_EMAIL,
    subject: '【COMPA VILLAGE】決済ページへ進みました — ' + (ROOM_NAMES[d.roomType] || ''),
    body: body
  });
}

// Stripe Webhookが決済完了を検知した時にホストへ送る、実名入りの確定通知
function sendPaymentConfirmedNotification_(d) {
  var body = 'Stripeで決済が完了し、自動的に「確定」にしました。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += 'お名前　：' + (d.name || '（取得できませんでした）') + '\n';
  body += 'メール　：' + (d.email || '（取得できませんでした）') + '\n';
  body += '部屋タイプ：' + (ROOM_NAMES[d.roomType] || d.roomType) + '\n';
  body += 'チェックイン：' + d.checkin + '\n';
  body += 'チェックアウト：' + d.checkout + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += 'この操作は自動なので、特にやることはありません。\n\n';
  body += '▶ 予約台帳: ' + sheetUrl();

  MailApp.sendEmail({
    to: HOST_EMAIL,
    cc: CC_EMAIL,
    subject: '【COMPA VILLAGE】決済が確定しました — ' + (ROOM_NAMES[d.roomType] || ''),
    body: body
  });
}

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
    .addItem('❌ 選択行を「お断り」にしてお詫びメールを送信', 'menuRejectBooking')
    .addSeparator()
    .addItem('💳 選択行を「決済確認済み」にする（メール送信なし）', 'menuMarkPaymentReceived')
    .addItem('🗑 選択行の決済保留を解除する（メール送信なし）', 'menuCancelPendingHold')
    .addSeparator()
    .addItem('📅 空室ブロック管理（オーナー不在・臨時休業など）', 'openBlockSidebar')
    .addToUi();
}

// =============================================
//  オーナーブロック（Booking.com等の「予約不可に設定」相当）
//  ゆうさんが旅行・メンテナンス等で部屋を予約不可にしたい時に使う。
//  実装上は「（オーナーブロック）」という名前の仮予約行として台帳に記録するだけで、
//  既存のisAvailable()/getBookedRanges()がそのまま空室カウントに反映してくれる
//  （新規ロジックを増やさず、既存の予約行カウント方式に相乗りする設計）。
// =============================================

function openBlockSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('BlockSidebar')
    .setTitle('空室ブロック管理')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

// サイドバーの初期表示用：部屋タイプごとの最大数（＝全室ブロック時のデフォルト台数）
function getRoomMeta() {
  return { capacity: ROOM_CAPACITY, names: ROOM_NAMES };
}

// 台数分だけ「オーナーブロック」行を積む。1回の操作で作った行は同じblockIdでまとめて管理し、
// 「解除」ボタン1回でまとめて空室に戻せるようにする
function addOwnerBlock(payload) {
  if (!SHEET_ID) return { ok: false, reason: 'no_sheet' };
  var roomType = payload && payload.roomType;
  if (!ROOM_NAMES[roomType]) return { ok: false, reason: 'invalid_room' };

  var checkin = payload.checkin;
  var checkout = payload.checkout;
  if (!checkin || !checkout || !(new Date(checkin) < new Date(checkout))) {
    return { ok: false, reason: 'invalid_dates' };
  }

  // 全体ブロック（オーナー不在等）はgroupタイプ1行でstandard/dorm/groupを丸ごと塞げる
  // （GROUND_FLOOR_TYPESの相互ブロックに乗るため）。standard/dormの部分ブロックは台数分だけ行を積む
  var maxQty = roomType === 'group' ? 1 : (ROOM_CAPACITY[roomType] || 1);
  var qty = parseInt(payload.quantity, 10);
  if (!qty || qty < 1) qty = maxQty;
  if (qty > maxQty) qty = maxQty;

  var blockId = 'blk-' + new Date().getTime() + '-' + Math.floor(Math.random() * 10000);
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('予約台帳');
  for (var i = 0; i < qty; i++) {
    sheet.appendRow([
      new Date(), '（オーナーブロック）', '', roomType,
      checkin, checkout, '', 'オーナーブロック',
      payload.memo || '', blockId
    ]);
  }
  return { ok: true, blockId: blockId, quantity: qty };
}

// 現在有効なオーナーブロックを一覧取得（サイドバー表示用。同じblockIdの行はまとめて1件にする）
function getActiveBlocks() {
  if (!SHEET_ID) return [];
  var rows = getBookingRows();
  var byBlockId = {};
  rows.forEach(function (row) {
    if (row[7] !== 'オーナーブロック') return;
    var blockId = row[9];
    if (!blockId) return;
    if (!byBlockId[blockId]) {
      byBlockId[blockId] = {
        blockId: blockId,
        roomType: row[3],
        roomName: ROOM_NAMES[row[3]] || row[3],
        checkin: formatDate_(row[4]),
        checkout: formatDate_(row[5]),
        memo: row[8] || '',
        quantity: 0
      };
    }
    byBlockId[blockId].quantity++;
  });
  return Object.keys(byBlockId).map(function (k) { return byBlockId[k]; })
    .sort(function (a, b) { return a.checkin < b.checkin ? -1 : (a.checkin > b.checkin ? 1 : 0); });
}

// blockIdが一致する行を全て「キャンセル」にして空室へ戻す
function removeOwnerBlock(blockId) {
  if (!SHEET_ID || !blockId) return { ok: false };
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('予約台帳');
  var values = sheet.getDataRange().getValues();
  var released = 0;
  for (var i = 1; i < values.length; i++) {
    if (values[i][7] === 'オーナーブロック' && String(values[i][9] || '') === blockId) {
      sheet.getRange(i + 1, 8).setValue('キャンセル');
      released++;
    }
  }
  return { ok: true, released: released };
}

// スタンダード/ドミトリーの仮予約（決済へ進行中）で、Stripe側の入金を確認できた時に使う。
// お名前・メールが無い行なので確認メールは送らない
function menuMarkPaymentReceived() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert('データ行を選択してください'); return; }

  var confirm = SpreadsheetApp.getUi().alert(
    'この行を「決済確認済み（確定）」にしますか？（メールは送信されません）',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  sheet.getRange(row, 8).setValue('確定');
  SpreadsheetApp.getUi().alert('決済確認済みにしました');
}

// 仮予約のまま決済が完了しなかった行を、空室に戻す（メール送信なし）
function menuCancelPendingHold() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert('データ行を選択してください'); return; }

  var confirm = SpreadsheetApp.getUi().alert(
    'この行の決済保留を解除して空室に戻しますか？（メールは送信されません）',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  sheet.getRange(row, 8).setValue('キャンセル');
  SpreadsheetApp.getUi().alert('決済保留を解除しました');
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

function menuRejectBooking() {
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
    d.name + ' 様（' + (ROOM_NAMES[d.roomType] || d.roomType) + '）へお詫びメールを送り、キャンセル扱いにしますか？',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  sheet.getRange(row, 8).setValue('キャンセル');

  var body = d.name + ' 様\n\n';
  body += 'このたびはCOMPA VILLAGEへご予約リクエストをいただき、誠にありがとうございました。\n';
  body += '大変申し訳ございませんが、今回はご希望に添うことができませんでした。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '部屋タイプ：' + (ROOM_NAMES[d.roomType] || d.roomType) + '\n';
  body += 'チェックイン：' + d.checkin + '\n';
  body += 'チェックアウト：' + d.checkout + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += 'またの機会がございましたら、ぜひよろしくお願いいたします。\n\n';
  body += 'COMPA VILLAGE';

  MailApp.sendEmail({
    to: d.email, replyTo: HOST_EMAIL,
    subject: '【COMPA VILLAGE】ご予約リクエストについて',
    body: body
  });

  SpreadsheetApp.getUi().alert('お詫びメールを送信しました');
}

// 確定メール：プレーンテキスト版（HTML非対応クライアント向けフォールバック）
// d.alreadyPaid=true の場合は決済案内を省略する（Stripe決済直後の即決済型で使用）
function buildConfirmationText_(d) {
  var payLink = STRIPE_LINKS[d.roomType];
  var body = d.name + ' 様\n\n';
  body += 'COMPA VILLAGEへようこそ🌵\n\n';
  body += 'このたびはご予約いただきありがとうございます。\n';
  body += '以下、ご滞在に関する大切なご案内ですので、ご確認をお願いいたします。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '部屋タイプ：' + (ROOM_NAMES[d.roomType] || d.roomType) + '\n';
  body += 'チェックイン：' + d.checkin + '\n';
  body += 'チェックアウト：' + d.checkout + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';

  body += '🕒 チェックイン・チェックアウト\n';
  body += '・チェックイン：16:00〜22:00（22時以降は別途¥1,000）\n';
  body += '・チェックアウト：10:00まで\n\n';

  body += '🚗 駐車について\n';
  body += '・敷地内に合計4台（満車の場合は近隣駐車場をご案内する場合がございます）\n';
  body += '・お車はガレージ横にお停めください\n\n';

  body += '⚠️ 立ち入り禁止エリア\n';
  body += '・納屋、蔵（1階・2階とも）※安全面のため立ち入りはご遠慮ください\n\n';

  body += '⚡ 電気について\n';
  body += '・キッチンの電力容量に限りがございます。電子レンジ・ポットなどを同時に2つ以上使用すると、ブレーカーが落ちる場合があります\n';
  body += '・万が一落ちた際は、脱衣所内の分電盤をご確認ください\n\n';

  body += '🔇 夜間のお願い\n';
  body += '・23時以降の室内では会話にご配慮をお願いいたします\n\n';

  body += '🍖 BBQをされる場合\n';
  body += '・BBQセット（コンロ・網・トング）レンタル代¥2,000を頂きます（現金のみ）\n';
  body += '・炭や食材などはご自由にお持ち込みください\n';
  body += '・宿泊者以外の方が参加される場合は、場所代として1人¥2,500を頂いています\n\n';

  body += '📍 道案内について\n';
  body += '道が少し分かりづらく、Googleマップのナビだとお隣のご自宅に案内されてしまう場合がございます。\n';
  body += 'iPhoneをお持ちでしたら、Appleマップで「下関市蒲生野280」と入力していただくと正確に到着できます\n\n';

  body += '🐾 ペットと宿泊される場合\n';
  body += '・追加1匹ごとに¥2,000を頂きます（小型犬のみ、1室2匹まで）\n';
  body += '・ケージは必ずご持参ください\n';
  body += '・共用部や他のお客様がいらっしゃる場所での放し飼いは禁止です（ガレージや屋外での開放は可能です）\n';
  body += '・客室内のベッドの上に乗せることはできません\n';
  body += '・室内外で排泄した際は、他のお客様もいらっしゃいますので速やかに片付けをお願いします\n';
  body += '・シーツや布団、室内に著しい汚損が見られた場合、清掃料金を請求する場合がございます\n\n';

  body += '※貴重品（財布・パスポート等）の管理はご自身にてしっかりとお願いいたします\n\n';

  body += '【アクセス】\n';
  body += ACCESS_MAP_URL + '\n\n';

  if (payLink && !d.alreadyPaid) {
    body += '【お支払い】\n';
    body += 'まだお支払いがお済みでない場合は、以下より決済をお願いいたします。\n';
    body += payLink + '\n\n';
  }
  body += '🐾 ご不明な点があれば、いつでもこのメールへの返信、またはInstagram（@compa0601）のDMでご連絡ください。\n\n';
  body += '当日を楽しみにお待ちしております！\n\n';
  body += 'COMPA VILLAGE\n' + IG_URL;
  return body;
}

// 確定メール：HTML版（ブランドカラーで整形。Gmail等HTML対応クライアントで表示）
// d.alreadyPaid=true の場合は決済案内を省略する（Stripe決済直後の即決済型で使用）
function buildConfirmationHtml_(d) {
  var roomName = ROOM_NAMES[d.roomType] || d.roomType;
  var payLink = STRIPE_LINKS[d.roomType];
  var esc = function(s){ return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };

  var payButton = '';
  if (payLink && !d.alreadyPaid) {
    payButton =
      '<tr><td style="padding:24px 28px 0;">' +
        '<p style="margin:0 0 10px; font-size:14px; color:#1F2B1E;">まだお支払いがお済みでない場合は、以下より決済をお願いいたします。</p>' +
        '<a href="' + esc(payLink) + '" style="display:inline-block; background:#8FBE3F; color:#1D3420; text-decoration:none; font-weight:bold; padding:12px 22px; border-radius:999px; font-size:14px;">お支払いはこちら →</a>' +
      '</td></tr>';
  }

  var guideItems = [
    { icon: '🕒', title: 'チェックイン・チェックアウト', lines: [
      'チェックイン：16:00〜22:00（22時以降は別途¥1,000）',
      'チェックアウト：10:00まで'
    ]},
    { icon: '🚗', title: '駐車について', lines: [
      '敷地内に合計4台（満車の場合は近隣駐車場をご案内する場合がございます）',
      'お車はガレージ横にお停めください'
    ]},
    { icon: '⚠️', title: '立ち入り禁止エリア', lines: [
      '納屋、蔵（1階・2階とも）※安全面のため立ち入りはご遠慮ください'
    ]},
    { icon: '⚡', title: '電気について', lines: [
      'キッチンの電力容量に限りがございます。電子レンジ・ポットなどを同時に2つ以上使用すると、ブレーカーが落ちる場合があります',
      '万が一落ちた際は、脱衣所内の分電盤をご確認ください'
    ]},
    { icon: '🔇', title: '夜間のお願い', lines: [
      '23時以降の室内では会話にご配慮をお願いいたします'
    ]},
    { icon: '🍖', title: 'BBQをされる場合', lines: [
      'BBQセット（コンロ・網・トング）レンタル代¥2,000（現金のみ）',
      '炭や食材などはご自由にお持ち込みください',
      '宿泊者以外の方が参加される場合は、場所代として1人¥2,500を頂いています'
    ]},
    { icon: '📍', title: '道案内について', lines: [
      'Googleマップのナビだとお隣のご自宅に案内される場合がございます',
      'iPhoneをお持ちでしたら、Appleマップで「下関市蒲生野280」と検索いただくと正確に到着できます'
    ]},
    { icon: '🐾', title: 'ペットと宿泊される場合', lines: [
      '追加1匹ごとに¥2,000を頂きます（小型犬のみ、1室2匹まで）',
      'ケージは必ずご持参ください',
      '共用部や他のお客様がいらっしゃる場所での放し飼いは禁止です（ガレージや屋外での開放は可能です）',
      '客室内のベッドの上に乗せることはできません',
      '室内外で排泄した際は、他のお客様もいらっしゃいますので速やかに片付けをお願いします',
      'シーツや布団、室内に著しい汚損が見られた場合、清掃料金を請求する場合がございます'
    ]}
  ];
  var guideHtml = guideItems.map(function (g) {
    var lines = g.lines.map(function (l) { return '<li style="margin:0 0 4px;">' + esc(l) + '</li>'; }).join('');
    return '<tr><td style="padding:18px 28px 0;">' +
      '<p style="margin:0 0 6px; font-size:13px; font-weight:bold; color:#1D3420;">' + g.icon + ' ' + esc(g.title) + '</p>' +
      '<ul style="margin:0; padding-left:18px; font-size:12.5px; color:rgba(31,43,30,.75); line-height:1.7;">' + lines + '</ul>' +
    '</td></tr>';
  }).join('');

  return '' +
  '<div style="font-family:\'Hiragino Sans\',\'Yu Gothic\',sans-serif; background:#F6F1E4; padding:28px 16px;">' +
    '<table role="presentation" width="100%" style="max-width:520px; margin:0 auto; background:#ffffff; border-radius:16px; overflow:hidden; border:1px solid #DCCFAE;">' +
      '<tr><td style="background:#1D3420; padding:22px 28px;">' +
        '<p style="margin:0; color:#8FBE3F; font-size:12px; letter-spacing:.14em; text-transform:uppercase;">Booking Confirmed</p>' +
        '<p style="margin:6px 0 0; color:#F6F1E4; font-size:20px; font-weight:bold;">COMPA VILLAGE</p>' +
      '</td></tr>' +
      '<tr><td style="padding:26px 28px 0;">' +
        '<p style="margin:0 0 4px; font-size:15px; color:#1F2B1E;">' + esc(d.name) + ' 様</p>' +
        '<p style="margin:0; font-size:14px; color:rgba(31,43,30,.75); line-height:1.7;">COMPA VILLAGEへようこそ🌵　ご予約が確定しましたのでご連絡いたします。以下、ご滞在に関する大切なご案内ですので、ご確認をお願いいたします。</p>' +
      '</td></tr>' +
      '<tr><td style="padding:20px 28px 0;">' +
        '<table role="presentation" width="100%" style="background:#EFE7D3; border-radius:12px;">' +
          '<tr><td style="padding:16px 20px; font-size:13px; color:#1F2B1E;">部屋タイプ</td><td style="padding:16px 20px; font-size:13px; color:#1D3420; font-weight:bold; text-align:right;">' + esc(roomName) + '</td></tr>' +
          '<tr><td style="padding:0 20px 16px; font-size:13px; color:#1F2B1E;">チェックイン</td><td style="padding:0 20px 16px; font-size:13px; color:#1D3420; font-weight:bold; text-align:right;">' + esc(d.checkin) + '</td></tr>' +
          '<tr><td style="padding:0 20px 16px; font-size:13px; color:#1F2B1E;">チェックアウト</td><td style="padding:0 20px 16px; font-size:13px; color:#1D3420; font-weight:bold; text-align:right;">' + esc(d.checkout) + '</td></tr>' +
        '</table>' +
      '</td></tr>' +
      guideHtml +
      '<tr><td style="padding:18px 28px 0;">' +
        '<p style="margin:0; font-size:12px; color:rgba(31,43,30,.6); line-height:1.7;">※貴重品（財布・パスポート等）の管理はご自身にてしっかりとお願いいたします</p>' +
      '</td></tr>' +
      '<tr><td style="padding:18px 28px 0;">' +
        '<p style="margin:0 0 6px; font-size:13px; font-weight:bold; color:#1D3420;">アクセス</p>' +
        '<p style="margin:0; font-size:13px; color:rgba(31,43,30,.75); line-height:1.8;">山口県下関市（下関市立美術館・長府庭園エリア）／最寄り：JR新下関駅<br><a href="' + esc(ACCESS_MAP_URL) + '" style="color:#2B4A2E;">地図で見る →</a></p>' +
      '</td></tr>' +
      payButton +
      '<tr><td style="padding:24px 28px 26px;">' +
        '<p style="margin:0; font-size:12px; color:rgba(31,43,30,.6); line-height:1.7;">🐾 ご不明な点があれば、このメールに返信いただくか、Instagram（<a href="' + esc(IG_URL) + '" style="color:#2B4A2E;">@compa0601</a>）のDMでもお気軽にご連絡ください。<br><br>旅の途中に、最高の出会いを。<br>COMPA VILLAGE</p>' +
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
