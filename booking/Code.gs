/**
 * アットインテリア 無料相談 予約システム（Google Apps Script）
 * ------------------------------------------------------------
 * 流れ：
 *   1. 相談者が日時を選んで予約（名前・メール・電話）
 *   2. 予約完了 → そのまま事前アンケートへ（あとで回答も可。確認メールに回答URLあり）
 *   3. 予約時・アンケート回答時に、それぞれあなたへ即メール通知
 *   4. 相談日の前日までにアンケート未回答なら、自動でお願いメールを送信
 *
 * 設定はこの CONFIG だけ変えればOKです。
 */
const CONFIG = {
  BRAND: 'アットインテリア',
  TITLE: '無料相談のご予約',
  SENDER_NAME: 'アットインテリア',

  // 予約通知を受け取るメールアドレス（空欄ならGASを動かしているGoogleアカウント宛て）
  NOTIFY_EMAIL: '',

  // 公開後に表示される予約ページのURL（.../exec）を入れてください（アンケートURLに使います）
  WEB_APP_URL: '',

  // お客様に案内する予約ページのURL（Instagram・LINEからでも開けるページ）
  PAGE_URL: 'https://atinterior.online/yoyaku/',

  // 相談に使う固定のZoom URL（必ず書き換えてください）
  ZOOM_URL: 'https://zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXX',
  ZOOM_NOTE: '', // 例: 'ミーティングID: 123 456 7890'

  // 受付時間（0=日 1=月 ... 6=土）
  WEEKDAYS: [1, 2, 3, 5, 6], // 月・火・水・金・土
  START_HOUR: 16,          // 16:00 から
  END_HOUR: 16.5,          // 16:30 まで（16:00〜16:30 の1枠だけ）
  SLOT_MINUTES: 30,        // 1枠の長さ
  BUFFER_MINUTES: 0,       // 予約の前後に空ける時間
  OPEN_FROM: '2026-10-26', // 受付期間の初日（空欄 '' なら制限なし）
  OPEN_UNTIL: '2026-11-30',// 受付期間の最終日（空欄 '' なら制限なし）
  DAYS_AHEAD: 60,          // 何日先まで予約を受け付けるか
  MIN_NOTICE_HOURS: 3,     // 今から何時間後以降の枠を出すか
  SKIP_JP_HOLIDAYS: true,  // 日本の祝日は受付しない

  // 空き判定に使うカレンダー（'primary' = メインのカレンダー）
  CALENDAR_ID: 'primary',
  IGNORE_ALL_DAY_EVENTS: true, // 終日予定は「空き」とみなす

  // アンケート未回答の方へ、相談日の前日にお願いメールを送る
  SURVEY_REMINDER: true,

  // 事前アンケート
  //   type: 'radio'（1つ選ぶ） / 'checkbox'（複数選べる） / 'textarea'（自由記載）
  //   質問や選択肢はここを書き換えるだけで、画面・メール・スプレッドシートに反映されます
  SURVEY: [
    { key: 'house', label: 'お住まいの種類', type: 'radio', required: true,
      options: ['戸建て（持ち家）', 'マンション（持ち家）', '賃貸', '新築・購入予定', 'その他'] },
    { key: 'family', label: 'ご家族の人数', type: 'radio', required: true,
      options: ['1人', '2人', '3人', '4人', '5人以上'] },
    { key: 'concern', label: '今のお悩み・ご検討内容（複数選択可）', type: 'checkbox', required: true,
      options: ['引越し', 'リノベーション（部屋全体）', 'リノベーション（キッチン）', 'リノベーション（洗面）', 'リノベーション（トイレ）', 'その他'] },
    { key: 'budget', label: 'ご希望のご予算', type: 'radio', required: true,
      options: ['50万円以下', '100万円以下', '100万円以上', '特に決めていない'] },
    { key: 'source', label: 'アットインテリアをどこで知りましたか？（複数選択可）', type: 'checkbox', required: true,
      options: ['Instagram', 'YouTube', '広告', 'ホームページ', '知り合いの紹介', 'その他'] },
    { key: 'style', label: 'アットインテリアのデザインで好きなところ（複数選択可）', type: 'checkbox', required: false,
      options: ['北欧的', 'モダン', '女性らしい', 'おしゃれ', 'その他'] },
    { key: 'note', label: 'お悩みの詳細・ご希望など（自由記載）', type: 'textarea', required: false },
  ],
};

const TZ = 'Asia/Tokyo';
const HOLIDAY_CAL_ID = 'ja.japanese#holiday@group.v.calendar.google.com';
const SHEET_TITLE = 'アットインテリア 無料相談 予約一覧';
const BASE_HEADERS = ['受付日時', '予約日時', '開始(内部用)', 'お名前', 'メール', '電話番号',
  '予約ID', 'カレンダーID', 'アンケート回答日時', 'アンケートお願いメール'];
const COL = { created: 1, when: 2, start: 3, name: 4, email: 5, phone: 6, id: 7, eventId: 8, surveyAt: 9, reminded: 10 };
const SURVEY_COL = BASE_HEADERS.length + 1;
const WEEK = '日月火水木金土';

/* ============ Webページ ============ */

function doGet(e) {
  const p = (e && e.parameter) || {};

  // 予約ページ（PAGE_URL）からの問い合わせ
  if (p.api) {
    return api_(function () {
      if (p.api === 'init') return pageData_(p.id);
      if (p.api === 'slots') return getAvailableSlots();
      throw new Error('不正なリクエストです。');
    });
  }

  const t = HtmlService.createTemplateFromFile('Index');
  const data = pageData_(p.id);
  t.dataJson = JSON.stringify(data).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle(CONFIG.BRAND + ' ' + CONFIG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 自分のサイトに埋め込み可
}

/** 予約ページ（PAGE_URL）から予約・アンケートを受け取る */
function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) {}
  return api_(function () {
    if (req.action === 'book') return bookSlot(req.form);
    if (req.action === 'survey') return submitSurvey(req.id, req.answers);
    throw new Error('不正なリクエストです。');
  });
}

function api_(fn) {
  let out;
  try {
    out = { ok: true, data: fn() };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err).replace(/^Error:\s*/, '') };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function pageData_(id) {
  const data = {
    mode: 'book',
    brand: CONFIG.BRAND,
    title: CONFIG.TITLE,
    slotMinutes: CONFIG.SLOT_MINUTES,
    survey: CONFIG.SURVEY,
  };
  if (id) {
    const rec = findBooking_(String(id));
    if (rec) {
      data.mode = 'survey';
      data.booking = { id: rec.id, name: rec.name, when: rec.when, answered: !!rec.surveyAt, answers: rec.answers };
    } else {
      data.mode = 'notfound';
    }
  }
  return data;
}

/* ============ 空き枠の取得 ============ */

/**
 * 予約できる枠の一覧。表示を速くするため結果を数分間キャッシュします。
 * 予約確定時は fresh=true で、必ずカレンダーを読み直して確認します。
 */
function getAvailableSlots(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) {
    const hit = cache.get('slots');
    if (hit) return JSON.parse(hit);
  }

  const now = new Date();
  const earliest = new Date(now.getTime() + CONFIG.MIN_NOTICE_HOURS * 3600 * 1000);

  // 調べる範囲を受付期間に絞る（カレンダーの読み込みを少なくして速くする）
  let rangeStart = startOfDay_(now);
  let rangeEnd = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), rangeStart.getDate() + CONFIG.DAYS_AHEAD + 1);
  if (CONFIG.OPEN_FROM && parseDate_(CONFIG.OPEN_FROM) > rangeStart) rangeStart = parseDate_(CONFIG.OPEN_FROM);
  if (CONFIG.OPEN_UNTIL) {
    const last = parseDate_(CONFIG.OPEN_UNTIL);
    const after = new Date(last.getFullYear(), last.getMonth(), last.getDate() + 1);
    if (after < rangeEnd) rangeEnd = after;
  }

  // 候補の枠を先に作る
  const candidates = [];
  for (let day = rangeStart; day < rangeEnd; day = new Date(day.getFullYear(), day.getMonth(), day.getDate() + 1)) {
    if (CONFIG.WEEKDAYS.indexOf(day.getDay()) === -1) continue;
    for (let m = CONFIG.START_HOUR * 60; m + CONFIG.SLOT_MINUTES <= CONFIG.END_HOUR * 60; m += CONFIG.SLOT_MINUTES) {
      const start = new Date(day.getFullYear(), day.getMonth(), day.getDate(), 0, m);
      if (start < earliest) continue;
      candidates.push({ day: day, start: start, end: new Date(start.getTime() + CONFIG.SLOT_MINUTES * 60000) });
    }
  }

  const days = [];
  if (candidates.length) {
    const busy = getBusyRanges_(rangeStart, rangeEnd, candidates);
    const holidays = CONFIG.SKIP_JP_HOLIDAYS ? getHolidaySet_(rangeStart, rangeEnd) : {};
    const byDate = {};
    candidates.forEach(function (c) {
      const key = fmt_(c.day, 'yyyy-MM-dd');
      if (holidays[key] || isBusy_(c.start, c.end, busy)) return;
      if (!byDate[key]) {
        byDate[key] = { date: key, label: fmt_(c.day, 'M/d') + '（' + WEEK[c.day.getDay()] + '）', slots: [] };
        days.push(byDate[key]);
      }
      byDate[key].slots.push({ start: c.start.getTime(), label: fmt_(c.start, 'HH:mm') });
    });
  }

  cache.put('slots', JSON.stringify(days), 600); // 10分
  return days;
}

/** 5分ごとに自動で動く：空き状況を先に読み込んでおき、お客様が開いたときすぐ表示できるようにする */
function warmSlots() {
  getAvailableSlots(true);
}

function parseDate_(s) {
  const p = String(s).split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

/* ============ 予約の確定 ============ */

function bookSlot(form) {
  const name = clean_(form && form.name, 100);
  const email = clean_(form && form.email, 200);
  const phone = clean_(form && form.phone, 30);
  const startMs = Number(form && form.start);

  if (!name) throw new Error('お名前を入力してください。');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('メールアドレスの形式が正しくありません。');
  if (phone && !/^[0-9０-９+\-－() ]{6,30}$/.test(phone)) throw new Error('電話番号の形式が正しくありません。');
  if (!startMs) throw new Error('日時を選択してください。');

  const start = new Date(startMs);
  const end = new Date(startMs + CONFIG.SLOT_MINUTES * 60000);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // 選ばれた枠が本当に今も空いているか再確認
    const stillOpen = getAvailableSlots(true).some(function (d) {
      return d.slots.some(function (s) { return s.start === startMs; });
    });
    if (!stillOpen) throw new Error('申し訳ありません。この時間はすでに埋まってしまいました。別の時間をお選びください。');

    const id = Utilities.getUuid();
    const rec = {
      id: id, name: name, email: email, phone: phone,
      when: fmt_(start, 'yyyy年M月d日') + '（' + WEEK[start.getDay()] + '）' + fmt_(start, 'HH:mm') + '〜' + fmt_(end, 'HH:mm'),
    };

    // 1. カレンダーに登録
    const event = getCalendar_().createEvent('【無料相談】' + name + ' 様', start, end, {
      description: eventDescription_(rec, null),
      location: CONFIG.ZOOM_URL,
    });

    // 2. スプレッドシートに記録（電話番号の先頭0が消えないよう ' を付ける）
    const row = [new Date(), rec.when, startMs, name, email, phone ? "'" + phone : '', id, event.getId(), '', ''];
    CONFIG.SURVEY.forEach(function () { row.push(''); });
    getSheet_().appendRow(row);

    // 3. あなたへの通知メール（即時）
    MailApp.sendEmail({
      to: notifyEmail_(),
      subject: '【新規予約】' + rec.when + ' ' + name + ' 様',
      body: [
        '無料相談の予約が入りました。', '',
        '日時: ' + rec.when,
        'お名前: ' + name,
        'メール: ' + email,
        '電話番号: ' + (phone || '（なし）'),
        '', '事前アンケートが回答されたら、あらためてお知らせします。',
        '', '予約一覧: ' + getSheet_().getParent().getUrl(),
      ].join('\n'),
      replyTo: email,
      name: CONFIG.SENDER_NAME,
    });

    // 4. 相談者への確認メール（アンケートURL付き）
    MailApp.sendEmail({
      to: email,
      subject: '【' + CONFIG.BRAND + '】無料相談のご予約ありがとうございます（' + rec.when + '）',
      body: [
        name + ' 様', '',
        CONFIG.BRAND + 'の無料相談にご予約いただき、ありがとうございます。',
        '以下の内容で承りました。', '',
        '日時: ' + rec.when,
        '', '■ 事前アンケートのお願い（1分ほどで終わります）',
        'ご相談をより充実した時間にするため、相談日までに下記からご回答ください。',
        surveyUrl_(id),
        '', '■ 当日のご参加方法',
        zoomText_(),
        '', '日時の変更・キャンセルは、このメールにご返信ください。',
        '', CONFIG.BRAND,
      ].join('\n'),
      replyTo: notifyEmail_(),
      name: CONFIG.SENDER_NAME,
    });

    CacheService.getScriptCache().remove('slots'); // 次に開いた人には最新の空き状況を出す
    return { ok: true, id: id, name: name, when: rec.when };
  } finally {
    lock.releaseLock();
  }
}

/* ============ 事前アンケート ============ */

function submitSurvey(id, answers) {
  answers = answers || {};
  const rec = findBooking_(String(id || ''));
  if (!rec) throw new Error('ご予約が見つかりませんでした。お手数ですが確認メールのURLから開き直してください。');

  const values = CONFIG.SURVEY.map(function (q) {
    const v = answers[q.key];
    let out = '';
    if (q.type === 'checkbox') {
      const arr = (Array.isArray(v) ? v : []).filter(function (o) { return q.options.indexOf(o) !== -1; });
      out = arr.join('、');
    } else if (q.type === 'radio') {
      const s = clean_(v, 100);
      out = q.options.indexOf(s) !== -1 ? s : '';
    } else {
      out = clean_(v, 2000);
    }
    if (q.required && !out) throw new Error('「' + q.label + '」にご回答ください。');
    return out;
  });

  const sheet = getSheet_();
  sheet.getRange(rec.row, SURVEY_COL, 1, values.length).setValues([values]);
  sheet.getRange(rec.row, COL.surveyAt).setValue(new Date());

  const lines = surveyLines_(values);

  // カレンダーの予定にも回答を書き込む
  try {
    const ev = getCalendar_().getEventById(rec.eventId);
    if (ev) ev.setDescription(eventDescription_(rec, lines));
  } catch (err) {}

  // あなたへの通知メール（即時）
  MailApp.sendEmail({
    to: notifyEmail_(),
    subject: '【アンケート回答' + (rec.surveyAt ? '・修正' : '') + '】' + rec.when + ' ' + rec.name + ' 様',
    body: [
      rec.name + ' 様から事前アンケートの回答が届きました。', '',
      '相談日時: ' + rec.when,
      'メール: ' + rec.email,
      '電話番号: ' + (rec.phone || '（なし）'),
      '', lines,
      '', '予約一覧: ' + sheet.getParent().getUrl(),
    ].join('\n'),
    replyTo: rec.email,
    name: CONFIG.SENDER_NAME,
  });

  return { ok: true, when: rec.when };
}

/** 毎日動く：相談日が近いのにアンケート未回答の方へ、お願いメールを1回だけ送る */
function sendSurveyReminders() {
  if (!CONFIG.SURVEY_REMINDER) return;
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const now = Date.now();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const start = Number(r[COL.start - 1]);
    const created = new Date(r[COL.created - 1]).getTime();
    if (!start || r[COL.surveyAt - 1] || r[COL.reminded - 1]) continue;
    if (start - now <= 0 || start - now > 36 * 3600 * 1000) continue; // 36時間以内の相談だけ
    if (now - created < 6 * 3600 * 1000) continue;                     // 予約直後の人には送らない

    MailApp.sendEmail({
      to: String(r[COL.email - 1]),
      subject: '【' + CONFIG.BRAND + '】事前アンケートのお願い（' + r[COL.when - 1] + ' 無料相談）',
      body: [
        r[COL.name - 1] + ' 様', '',
        'まもなく無料相談の日となりました。',
        'ご相談を充実した時間にするため、よろしければ事前アンケートにご回答ください（1分ほどで終わります）。',
        '', surveyUrl_(String(r[COL.id - 1])),
        '', '日時: ' + r[COL.when - 1],
        zoomText_(),
        '', CONFIG.BRAND,
      ].join('\n'),
      replyTo: notifyEmail_(),
      name: CONFIG.SENDER_NAME,
    });
    sheet.getRange(i + 1, COL.reminded).setValue(new Date());
  }
}

/* ============ 初回セットアップ用 ============ */

/** エディタから一度実行して、権限の許可・動作確認・毎日のお願いメールの設定をします */
function setup() {
  getSheet_();
  const days = getAvailableSlots(true);
  Logger.log('通知先: ' + notifyEmail_());
  Logger.log('予約一覧シート: ' + getSheet_().getParent().getUrl());
  Logger.log('空き枠のある日数: ' + days.length + (days[0] ? '（最初: ' + days[0].label + ' ' + days[0].slots[0].label + '）' : ''));

  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'sendSurveyReminders' || fn === 'warmSlots') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('warmSlots').timeBased().everyMinutes(5).create();
  Logger.log('空き状況の先読み: 5分ごとに自動で更新します');
  if (CONFIG.SURVEY_REMINDER) {
    ScriptApp.newTrigger('sendSurveyReminders').timeBased().everyDays(1).atHour(9).create();
    Logger.log('アンケートのお願いメール: 毎日9時台に自動チェックします');
  }

  if (CONFIG.SKIP_JP_HOLIDAYS && !CalendarApp.getCalendarById(HOLIDAY_CAL_ID)) {
    Logger.log('注意: 「日本の祝日」カレンダーが見つかりません。Googleカレンダーで追加すると祝日を自動で除外します。');
  }
  if (!CONFIG.WEB_APP_URL) {
    Logger.log('注意: WEB_APP_URL が空です。公開後に予約ページのURL（.../exec）を入れてください。');
  }
  if (/XXXX/.test(CONFIG.ZOOM_URL)) {
    Logger.log('注意: ZOOM_URL がまだ仮のままです。');
  }
}

/* ============ 内部処理 ============ */

function findBooking_(id) {
  if (!id) return null;
  const rows = getSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[COL.id - 1]) !== id) continue;
    const answers = {};
    CONFIG.SURVEY.forEach(function (q, j) { answers[q.key] = String(r[SURVEY_COL - 1 + j] || ''); });
    return {
      row: i + 1,
      id: id,
      name: String(r[COL.name - 1]),
      email: String(r[COL.email - 1]),
      phone: String(r[COL.phone - 1] || ''),
      when: String(r[COL.when - 1]),
      eventId: String(r[COL.eventId - 1]),
      surveyAt: r[COL.surveyAt - 1] || '',
      answers: answers,
    };
  }
  return null;
}

function eventDescription_(rec, surveyText) {
  return [
    'お名前: ' + rec.name,
    'メール: ' + rec.email,
    '電話番号: ' + (rec.phone || '（なし）'),
    '',
    '--- 事前アンケート ---',
    surveyText || '未回答',
    '',
    zoomText_(),
  ].join('\n');
}

function surveyLines_(values) {
  return CONFIG.SURVEY.map(function (q, i) {
    return '■ ' + q.label + '\n' + (values[i] || '（未回答）');
  }).join('\n');
}

function zoomText_() {
  return 'Zoom: ' + CONFIG.ZOOM_URL + (CONFIG.ZOOM_NOTE ? '\n' + CONFIG.ZOOM_NOTE : '');
}

function surveyUrl_(id) {
  const base = CONFIG.PAGE_URL || CONFIG.WEB_APP_URL || ScriptApp.getService().getUrl();
  return base + '?id=' + encodeURIComponent(id);
}

function getCalendar_() {
  return CONFIG.CALENDAR_ID === 'primary'
    ? CalendarApp.getDefaultCalendar()
    : CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
}

function getBusyRanges_(from, to, candidates) {
  const buf = CONFIG.BUFFER_MINUTES * 60000;
  const busy = [];
  getCalendar_().getEvents(from, to).forEach(function (e) {
    const r = [e.getStartTime().getTime() - buf, e.getEndTime().getTime() + buf];
    // 受付枠に重ならない予定は、詳しく調べずに飛ばす（速くするため）
    if (candidates && !candidates.some(function (c) { return isBusy_(c.start, c.end, [r]); })) return;
    if (CONFIG.IGNORE_ALL_DAY_EVENTS && e.isAllDayEvent()) return;
    try { if (e.getMyStatus() === CalendarApp.GuestStatus.NO) return; } catch (err) {}
    busy.push(r);
  });
  return busy;
}

function isBusy_(start, end, busy) {
  const s = start.getTime(), e = end.getTime();
  return busy.some(function (b) { return s < b[1] && e > b[0]; });
}

function getHolidaySet_(from, to) {
  const set = {};
  const cal = CalendarApp.getCalendarById(HOLIDAY_CAL_ID);
  if (!cal) return set;
  cal.getEvents(from, to).forEach(function (e) { set[fmt_(e.getStartTime(), 'yyyy-MM-dd')] = true; });
  return set;
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  let ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(SHEET_TITLE);
    const sh = ss.getSheets()[0];
    sh.appendRow(BASE_HEADERS.concat(CONFIG.SURVEY.map(function (q) { return q.label; })));
    sh.setFrozenRows(1);
    sh.hideColumns(COL.start);
    props.setProperty('SHEET_ID', ss.getId());
  }
  return ss.getSheets()[0];
}

function notifyEmail_() {
  return CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function fmt_(d, pattern) {
  return Utilities.formatDate(d, TZ, pattern);
}

function clean_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}
