/**
 * 無料相談 予約システム（Google Apps Script）
 * ------------------------------------------------------------
 * 設定はこの CONFIG だけ変えればOKです。
 */
const CONFIG = {
  // ページのタイトル・メール差出人名
  TITLE: '無料相談のご予約',
  SENDER_NAME: '無料相談受付',

  // 予約通知を受け取るメールアドレス（空欄ならGASを動かしているGoogleアカウント宛て）
  NOTIFY_EMAIL: '',

  // 相談に使う固定のZoom URL（必ず書き換えてください）
  ZOOM_URL: 'https://zoom.us/j/XXXXXXXXXX?pwd=XXXXXXXX',
  ZOOM_NOTE: '', // 例: 'ミーティングID: 123 456 7890 / パスコード: abc123'

  // 受付時間（0=日 1=月 ... 6=土）
  WEEKDAYS: [1, 2, 3, 4, 5],
  START_HOUR: 10,          // 10:00 から
  END_HOUR: 18,            // 18:00 まで（最終枠は 17:30〜18:00）
  SLOT_MINUTES: 30,        // 1枠の長さ
  BUFFER_MINUTES: 0,       // 予約の前後に空ける時間
  DAYS_AHEAD: 30,          // 何日先まで予約を受け付けるか
  MIN_NOTICE_HOURS: 3,     // 今から何時間後以降の枠を出すか
  SKIP_JP_HOLIDAYS: true,  // 日本の祝日は受付しない

  // 空き判定に使うカレンダー（'primary' = メインのカレンダー）
  CALENDAR_ID: 'primary',
  IGNORE_ALL_DAY_EVENTS: true, // 終日予定は「空き」とみなす

  // 相談内容の選択肢（最後の「その他」は残しておくのがおすすめ）
  TOPICS: [
    'サービス内容について',
    '料金について',
    'お申し込み・導入について',
    'その他',
  ],
};

const TZ = 'Asia/Tokyo';
const HOLIDAY_CAL_ID = 'ja.japanese#holiday@group.v.calendar.google.com';

/* ============ Webページ ============ */

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.config = { title: CONFIG.TITLE, topics: CONFIG.TOPICS, slotMinutes: CONFIG.SLOT_MINUTES };
  return t.evaluate()
    .setTitle(CONFIG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 自分のサイトに埋め込み可
}

/* ============ 空き枠の取得 ============ */

function getAvailableSlots() {
  const now = new Date();
  const earliest = new Date(now.getTime() + CONFIG.MIN_NOTICE_HOURS * 3600 * 1000);
  const rangeStart = startOfDay_(now);
  const rangeEnd = new Date(rangeStart.getTime() + (CONFIG.DAYS_AHEAD + 1) * 86400 * 1000);

  const busy = getBusyRanges_(rangeStart, rangeEnd);
  const holidays = CONFIG.SKIP_JP_HOLIDAYS ? getHolidaySet_(rangeStart, rangeEnd) : {};

  const days = [];
  for (let i = 0; i <= CONFIG.DAYS_AHEAD; i++) {
    const day = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), rangeStart.getDate() + i);
    const key = fmt_(day, 'yyyy-MM-dd');
    if (CONFIG.WEEKDAYS.indexOf(day.getDay()) === -1 || holidays[key]) continue;

    const slots = [];
    for (let m = CONFIG.START_HOUR * 60; m + CONFIG.SLOT_MINUTES <= CONFIG.END_HOUR * 60; m += CONFIG.SLOT_MINUTES) {
      const start = new Date(day.getFullYear(), day.getMonth(), day.getDate(), 0, m);
      const end = new Date(start.getTime() + CONFIG.SLOT_MINUTES * 60000);
      if (start < earliest) continue;
      if (isBusy_(start, end, busy)) continue;
      slots.push({ start: start.getTime(), label: fmt_(start, 'HH:mm') });
    }
    if (slots.length) {
      days.push({ date: key, label: fmt_(day, 'M/d') + '（' + '日月火水木金土'[day.getDay()] + '）', slots: slots });
    }
  }
  return days;
}

/* ============ 予約の確定 ============ */

function bookSlot(form) {
  const name = clean_(form && form.name, 100);
  const email = clean_(form && form.email, 200);
  const topic = clean_(form && form.topic, 100);
  const message = clean_(form && form.message, 2000);
  const startMs = Number(form && form.start);

  if (!name) throw new Error('お名前を入力してください。');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('メールアドレスの形式が正しくありません。');
  if (CONFIG.TOPICS.indexOf(topic) === -1) throw new Error('ご相談内容を選択してください。');
  if (topic === 'その他' && !message) throw new Error('「その他」を選んだ場合は、ご希望欄にご相談内容をご記入ください。');
  if (!startMs) throw new Error('日時を選択してください。');

  const start = new Date(startMs);
  const end = new Date(startMs + CONFIG.SLOT_MINUTES * 60000);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // 選ばれた枠が本当に今も空いているか再確認
    const stillOpen = getAvailableSlots().some(function (d) {
      return d.slots.some(function (s) { return s.start === startMs; });
    });
    if (!stillOpen) throw new Error('申し訳ありません。この時間はすでに埋まってしまいました。別の時間をお選びください。');

    const whenText = fmt_(start, 'yyyy年M月d日') + '（' + '日月火水木金土'[start.getDay()] + '）' +
      fmt_(start, 'HH:mm') + '〜' + fmt_(end, 'HH:mm');
    const zoomText = 'Zoom: ' + CONFIG.ZOOM_URL + (CONFIG.ZOOM_NOTE ? '\n' + CONFIG.ZOOM_NOTE : '');

    // 1. カレンダーに登録
    const cal = getCalendar_();
    const event = cal.createEvent('【無料相談】' + name + ' 様', start, end, {
      description: [
        'お名前: ' + name,
        'メール: ' + email,
        'ご相談内容: ' + topic,
        'ご希望・詳細:', message || '（なし）',
        '', zoomText,
      ].join('\n'),
      location: CONFIG.ZOOM_URL,
    });

    // 2. スプレッドシートに記録
    getSheet_().appendRow([new Date(), whenText, name, email, topic, message, event.getId()]);

    // 3. あなたへの通知メール（即時）
    MailApp.sendEmail({
      to: notifyEmail_(),
      subject: '【新規予約】' + whenText + ' ' + name + ' 様',
      body: [
        '無料相談の予約が入りました。', '',
        '日時: ' + whenText,
        'お名前: ' + name,
        'メール: ' + email,
        'ご相談内容: ' + topic,
        'ご希望・詳細:', message || '（なし）',
        '', zoomText,
        '', '予約一覧: ' + getSheet_().getParent().getUrl(),
      ].join('\n'),
      replyTo: email,
      name: CONFIG.SENDER_NAME,
    });

    // 4. 相談者への確認メール
    MailApp.sendEmail({
      to: email,
      subject: '【予約確定】無料相談 ' + whenText,
      body: [
        name + ' 様', '',
        '無料相談のご予約ありがとうございます。',
        '以下の内容で承りました。', '',
        '日時: ' + whenText,
        'ご相談内容: ' + topic,
        '', '当日は下記のZoomからご参加ください。',
        zoomText, '',
        '日時の変更・キャンセルは、このメールにご返信ください。',
      ].join('\n'),
      replyTo: notifyEmail_(),
      name: CONFIG.SENDER_NAME,
    });

    return { ok: true, when: whenText };
  } finally {
    lock.releaseLock();
  }
}

/* ============ 初回セットアップ用 ============ */

/** エディタから一度だけ実行して、権限の許可と動作確認をします */
function setup() {
  getSheet_();
  const days = getAvailableSlots();
  Logger.log('通知先: ' + notifyEmail_());
  Logger.log('予約一覧シート: ' + getSheet_().getParent().getUrl());
  Logger.log('空き枠のある日数: ' + days.length + (days[0] ? '（最初: ' + days[0].label + ' ' + days[0].slots[0].label + '）' : ''));
  if (CONFIG.SKIP_JP_HOLIDAYS && !CalendarApp.getCalendarById(HOLIDAY_CAL_ID)) {
    Logger.log('注意: 「日本の祝日」カレンダーが見つかりません。Googleカレンダーで追加すると祝日を自動で除外します。');
  }
}

/* ============ 内部処理 ============ */

function getCalendar_() {
  return CONFIG.CALENDAR_ID === 'primary'
    ? CalendarApp.getDefaultCalendar()
    : CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
}

function getBusyRanges_(from, to) {
  const buf = CONFIG.BUFFER_MINUTES * 60000;
  return getCalendar_().getEvents(from, to)
    .filter(function (e) {
      if (CONFIG.IGNORE_ALL_DAY_EVENTS && e.isAllDayEvent()) return false;
      try { if (e.getMyStatus() === CalendarApp.GuestStatus.NO) return false; } catch (err) {}
      return true;
    })
    .map(function (e) {
      return [e.getStartTime().getTime() - buf, e.getEndTime().getTime() + buf];
    });
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
  let id = props.getProperty('SHEET_ID');
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('無料相談 予約一覧');
    ss.getSheets()[0].appendRow(['受付日時', '予約日時', 'お名前', 'メール', 'ご相談内容', 'ご希望・詳細', 'カレンダーID']);
    ss.getSheets()[0].setFrozenRows(1);
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
