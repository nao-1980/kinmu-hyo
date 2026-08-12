/**
 * ============================================================
 * 勤務表作成ツール ── サーバー側（Google Apps Script）
 * ------------------------------------------------------------
 * 部署ごとのパスワードでログインします。
 * Googleアカウントを持っていない職員でも利用できます。
 *
 * 使うシート
 *   設定        … 部署・必要出勤人数・管理者メール
 *   パスワード  … 管理者と部署ごとのパスワード（非表示）
 *   _data       … 保存の本体（年月×部署ごとのJSON。非表示）
 *   勤務予定     … 人が読む用（1行＝1人1日）
 *   アクセスログ … ログインの記録
 *
 * ※ 初回は setup() を1回だけ実行してください。
 * ============================================================
 */

/* サーバー側のバージョン。index.html の VERSION と一致していれば
   デプロイが正しく反映されています。 */
const SERVER_VERSION = 'v5.3';

const DATA_SHEET = '_data';
const CONF_SHEET = '設定';
const PASS_SHEET = 'パスワード';
const PASS_SHEET_OLD = '合言葉';   /* 旧バージョンで作られたシート名 */
const PLAN_SHEET = '勤務予定';
const LOG_SHEET  = 'アクセスログ';

const DEFAULT_DEPARTMENTS =
  ['管理部','事務','診療アシスタント','看護師','ドライバー','戸塚院','港南院'];

/* ログインの有効時間（秒）。既定は6時間 */
const TOKEN_TTL = 6 * 60 * 60;

/* パスワードを間違えられる回数と、超えた場合の待ち時間（秒） */
const MAX_FAIL = 5;
const LOCK_SEC = 10 * 60;


/* ============================================================
   画面の表示
   ============================================================ */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('勤務表作成ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/* ============================================================
   初回セットアップ
   ============================================================ */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('スプレッドシートに紐づいていません。\n' +
      'スプレッドシートを開き、拡張機能 → Apps Script から作り直してください。');
  }

  /* --- 設定 --- */
  let conf = ss.getSheetByName(CONF_SHEET);
  if (!conf) {
    conf = ss.insertSheet(CONF_SHEET);
    conf.getRange(1, 1, 5, 2).setValues([
      ['項目', '値'],
      ['管理者メール', Session.getActiveUser().getEmail() || ''],
      ['部署', DEFAULT_DEPARTMENTS.join(', ')],
      ['必要出勤人数', 2],
      ['最初に表示する月', '当月'],
    ]);
    conf.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8edf5');
    conf.setColumnWidth(1, 150);
    conf.setColumnWidth(2, 520);
    conf.getRange('A7').setValue(
      '※「管理者メール」に登録した方は、Googleアカウントでログイン済みならパスワードなしで管理者として入れます。');
    conf.getRange('A8').setValue(
      '※「最初に表示する月」は 当月 / 翌月 / 前月 のいずれかを入力してください。');
    conf.getRange('A7:A8').setFontColor('#5b6470').setFontSize(10);
  }

  /* すでに設定シートがある場合、足りない項目を追加します */
  ensureConf_(conf, '最初に表示する月', '当月');

  /* --- パスワード --- */
  let pass = ss.getSheetByName(PASS_SHEET) || ss.getSheetByName(PASS_SHEET_OLD);
  if (!pass) {
    pass = ss.insertSheet(PASS_SHEET);
    const rows = [['対象', 'パスワード', '更新日時']];
    rows.push(['管理者', makePasscode_(), new Date()]);
    getDepartments_().forEach(function (d) { rows.push([d, makePasscode_(), new Date()]); });
    pass.getRange(1, 1, rows.length, 3).setValues(rows);
    pass.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8edf5');
    pass.setColumnWidth(1, 160);
    pass.setColumnWidth(2, 180);
    pass.setColumnWidth(3, 180);
    pass.getRange('B:B').setNumberFormat('@');
    pass.getRange(2, 2, rows.length - 1, 1).setFontFamily('Courier New').setFontSize(13);
    pass.getRange(rows.length + 2, 1).setValue(
      '※ パスワードは自由に書き換えられます。変更すると、次回のログインから新しいパスワードが必要になります。');
    pass.getRange(rows.length + 2, 1).setFontColor('#5b6470').setFontSize(10);
    pass.hideSheet();
  }

  /* --- データ --- */
  let data = ss.getSheetByName(DATA_SHEET);
  if (!data) {
    data = ss.insertSheet(DATA_SHEET);
    data.getRange(1, 1, 1, 4)
        .setValues([['キー', '年月', '部署', '内容（JSON）']])
        .setFontWeight('bold').setBackground('#e8edf5');
    data.setColumnWidth(1, 160);
    data.setColumnWidth(4, 600);
    /* 「2026.10」が数値2026.1に変換されるのを防ぎます */
    data.getRange('A:C').setNumberFormat('@');
    data.hideSheet();
  }

  /* --- 勤務予定 --- */
  let plan = ss.getSheetByName(PLAN_SHEET);
  if (!plan) {
    plan = ss.insertSheet(PLAN_SHEET);
    plan.getRange(1, 1, 1, 8)
        .setValues([['年月', '部署', '社員番号', '氏名', '日付', '曜日', '勤務区分', '区分名称']])
        .setFontWeight('bold').setBackground('#e8edf5');
    /* 年月・部署・社員番号は文字列として扱います */
    plan.getRange('A:D').setNumberFormat('@');
    plan.setFrozenRows(1);
  }

  /* --- アクセスログ --- */
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.getRange(1, 1, 1, 5)
       .setValues([['日時', '結果', '部署', '氏名', '備考']])
       .setFontWeight('bold').setBackground('#e8edf5');
    log.setFrozenRows(1);
  }

  const msg = 'セットアップが完了しました。\n' +
    '「パスワード」シート（非表示）に、管理者と部署ごとのパスワードが発行されています。\n' +
    '表示 → 非表示のシート から開いて確認してください。';
  Logger.log(msg);
  return msg;
}

/** 設定シートに項目がなければ追加します */
function ensureConf_(conf, item, value) {
  const rows = conf.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === item) return;
  }
  conf.appendRow([item, value]);
}

/** 最初に表示する月のずれ（0=当月 / 1=翌月 / -1=前月） */
function startOffset_() {
  const v = String(getConf_('最初に表示する月', '当月')).trim();
  if (v.indexOf('翌') >= 0) return 1;
  if (v.indexOf('前') >= 0) return -1;
  return 0;
}

/** パスワードを発行します（紛らわしい文字を除いた8桁） */
function makePasscode_() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

/**
 * 部署を追加したときに実行します。
 * 「設定」シートの部署のうち、パスワードがまだない部署に発行します。
 * すでにあるパスワードは変更しません。
 */
function syncPasscodes() {
  const sh = passSheet_();
  const rows = sh.getDataRange().getValues();
  const have = {};
  for (let i = 1; i < rows.length; i++) {
    const k = String(rows[i][0]).trim();
    if (k) have[k] = true;
  }

  const added = [];
  if (!have['管理者']) {
    sh.appendRow(['管理者', makePasscode_(), new Date()]);
    added.push('管理者');
  }
  getDepartments_().forEach(function (d) {
    if (!have[d]) {
      sh.appendRow([d, makePasscode_(), new Date()]);
      added.push(d);
    }
  });

  const msg = added.length
    ? 'パスワードを発行しました：' + added.join('、') + '\n「パスワード」シートを確認してください。'
    : '追加が必要な部署はありませんでした。';
  Logger.log(msg);
  return msg;
}

/**
 * すべてのパスワードを作り直します。
 * 漏洩が疑われるときや、年度替わりの一斉更新に使います。
 * 実行すると、全員が次回から新しいパスワードを使うことになります。
 */
function resetPasscodes() {
  syncPasscodes();
  const sh = passSheet_();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (!String(rows[i][0]).trim()) continue;
    sh.getRange(i + 1, 2).setValue(makePasscode_());
    sh.getRange(i + 1, 3).setValue(new Date());
  }
  const msg = 'すべてのパスワードを作り直しました。「パスワード」シートを確認してください。';
  Logger.log(msg);
  return msg;
}


/* ============================================================
   ログイン
   ============================================================ */

/** 画面を開いたときに呼ばれます */
function getAuthInfo() {
  const email = getEmail_();

  /* 管理者メールに登録済みで、Googleアカウントでログイン済みならパスワードは不要 */
  if (email && isAdminMail_(email)) {
    const token = issueToken_({ role: 'admin', dept: '', name: email });
    logAccess_('自動ログイン', '', email, '管理者メール');
    return {
      authed: true, token: token, role: 'admin', dept: '', name: email,
      departments: getDepartments_(), minStaff: getConf_('必要出勤人数', 2),
      serverVersion: SERVER_VERSION, startOffset: startOffset_()
    };
  }
  return { authed: false, departments: getDepartments_(), serverVersion: SERVER_VERSION };
}

/** パスワードでログインします */
function login(dept, passcode) {
  const cache = CacheService.getScriptCache();
  const fkey = 'fail_' + String(dept);
  const fails = Number(cache.get(fkey) || 0);

  if (fails >= MAX_FAIL) {
    logAccess_('ロック中', dept, '', '連続失敗');
    throw new Error('入力の失敗が続いたため、しばらくログインできません。\n10分ほど待ってからお試しください。');
  }

  const code = String(passcode || '').trim();
  if (!code) throw new Error('パスワードを入力してください。');

  const table = getPasscodes_();

  /* 管理者のパスワードなら、部署を問わず管理者として入ります */
  if (table['管理者'] && code === table['管理者']) {
    cache.remove(fkey);
    const token = issueToken_({ role: 'admin', dept: '', name: '管理者' });
    logAccess_('ログイン', '', '管理者', 'パスワード');
    return {
      authed: true, token: token, role: 'admin', dept: '', name: '管理者',
      departments: getDepartments_(), minStaff: getConf_('必要出勤人数', 2),
      serverVersion: SERVER_VERSION, startOffset: startOffset_()
    };
  }

  /* 部署のパスワード */
  const want = table[String(dept)];
  if (want && code === want) {
    cache.remove(fkey);
    const token = issueToken_({ role: 'staff', dept: String(dept), name: '' });
    logAccess_('ログイン', dept, '', 'パスワード');
    return {
      authed: true, token: token, role: 'staff', dept: String(dept), name: '',
      departments: [String(dept)], minStaff: getConf_('必要出勤人数', 2),
      names: rosterNames_(String(dept)), serverVersion: SERVER_VERSION,
      startOffset: startOffset_()
    };
  }

  cache.put(fkey, String(fails + 1), LOCK_SEC);
  logAccess_('失敗', dept, '', 'パスワードが違います');
  throw new Error('パスワードが違います。（あと ' + Math.max(0, MAX_FAIL - fails - 1) + ' 回）');
}

/** ログイン後に氏名を選びます（変更履歴に残すため） */
function setMyName(token, name) {
  const me = verify_(token);
  me.name = String(name || '').trim();
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify(me), TOKEN_TTL);
  logAccess_('氏名選択', me.dept, me.name, '');
  return { ok: true };
}

function issueToken_(me) {
  const token = Utilities.getUuid();
  me.at = new Date().toISOString();
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify(me), TOKEN_TTL);
  return token;
}

/** すべてのデータ操作の入口で必ず確認します */
function verify_(token) {
  const raw = CacheService.getScriptCache().get('tok_' + String(token || ''));
  if (!raw) {
    throw new Error('ログインの有効期限が切れました。画面を再読み込みして、もう一度ログインしてください。');
  }
  return JSON.parse(raw);
}

/** その部署を扱ってよいかを確認します */
function allow_(me, dept) {
  if (me.role === 'admin') return;
  if (String(me.dept) !== String(dept)) {
    throw new Error('この部署を開く権限がありません。');
  }
}


/* ============================================================
   読み込み・保存
   ============================================================ */
function loadMonth(token, ym, dept) {
  const me = verify_(token);
  allow_(me, dept);

  const key = makeKey_(ym, dept);
  const rows = sheet_(DATA_SHEET).getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== key) continue;
    let d;
    try { d = JSON.parse(rows[i][3]); }
    catch (e) { throw new Error('保存データを読み取れませんでした（' + key + '）'); }

    /* 氏名が1件もない記録は「空」とみなし、引き継ぎに回します。
       過去に空のまま保存された月でも、名簿を引き継げるようにするためです。 */
    const named = (d.staff || []).filter(function (st) {
      return String(st.name || '').trim();
    }).length;
    if (named > 0) return d;
    break;
  }

  /* 未作成、または中身が空の月は、直近の月から名簿を引き継ぎます */
  return carryRoster_(rows, ym, dept);
}

/**
 * 過去12か月をさかのぼり、最初に見つかった月の名簿
 * （氏名・社員番号）だけを返します。勤務予定は空にします。
 */
function carryRoster_(rows, ym, dept) {
  const parts = String(ym).split('.');
  let y = Number(parts[0]), m = Number(parts[1]);
  if (!y || !m) return null;

  for (let back = 1; back <= 12; back++) {
    m--;
    if (m < 1) { m = 12; y--; }
    const k = makeKey_(y + '.' + m, dept);

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== k) continue;
      try {
        const d = JSON.parse(rows[i][3]);
        const staff = (d.staff || [])
          .filter(function (st) { return String(st.name || '').trim(); })
          .map(function (st) {
            return { emp: String(st.emp || ''), name: String(st.name || ''), days: {} };
          });
        if (!staff.length) return null;
        return {
          staff: staff, events: {}, archived: [], history: [],
          actuals: {}, actualsMeta: null,
          carriedFrom: y + '.' + m
        };
      } catch (e) { return null; }
    }
  }
  return null;
}

function saveMonth(token, ym, dept, payload) {
  const me = verify_(token);
  allow_(me, dept);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('他の方が保存中です。少し待ってからもう一度お試しください。');
  }

  try {
    const key = makeKey_(ym, dept);
    const json = JSON.stringify(payload);
    if (json.length > 45000) {
      throw new Error('データが大きすぎます。職員数を分けてご利用ください。');
    }

    const sh = sheet_(DATA_SHEET);
    const rows = sh.getDataRange().getValues();
    let found = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === key) { found = i + 1; break; }
    }

    const rec = [key, ym, dept, json];
    const row = (found > 0) ? found : sh.getLastRow() + 1;
    const rng = sh.getRange(row, 1, 1, 4);
    rng.setNumberFormat('@');     /* 数値に変換されないようにします */
    rng.setValues([rec]);

    writePlanSheet_(ym, dept, payload);

    return {
      ok: true,
      at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm'),
      by: me.name || me.role
    };
  } finally {
    lock.releaseLock();
  }
}

function writePlanSheet_(ym, dept, payload) {
  const sh = sheet_(PLAN_SHEET);
  const all = sh.getDataRange().getValues();
  const head = all[0];
  const keep = all.slice(1).filter(function (r) {
    return !(String(r[0]) === ym && String(r[1]) === dept);
  });

  const dow = ['日','月','火','水','木','金','土'];
  const parts = String(ym).split('.');
  const y = Number(parts[0]), m = Number(parts[1]);
  const add = [];

  (payload.staff || []).forEach(function (st) {
    const name = String(st.name || '').trim();
    if (!name) return;
    const days = st.days || {};
    Object.keys(days).forEach(function (d) {
      const code = days[d];
      if (!code) return;
      const date = new Date(y, m - 1, Number(d));
      add.push([ym, dept, String(st.emp || ''), name,
        Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd'),
        dow[date.getDay()], code, (payload.labels || {})[code] || '']);
    });
  });

  add.sort(function (a, b) {
    return (a[3] === b[3]) ? (a[4] < b[4] ? -1 : 1) : (a[3] < b[3] ? -1 : 1);
  });

  const out = [head].concat(keep, add);
  sh.clearContents();
  if (out.length) {
    const rng = sh.getRange(1, 1, out.length, head.length);
    rng.setNumberFormat('@');
    rng.setValues(out);
  }
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8edf5');
}


/**
 * すでに保存済みのデータを、正しい書式で作り直します。
 * 「勤務予定」シートを _data から全部作り直すので、
 * 10月と1月が混ざってしまった状態も直ります。
 * 一度だけ実行してください。
 */
function repairSheets() {
  const data = sheet_(DATA_SHEET);
  data.getRange('A:C').setNumberFormat('@');

  const rows = data.getDataRange().getValues();

  /* _data の年月・部署を、キー（A列）から作り直します */
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]);
    if (key.indexOf('|') < 0) continue;
    const parts = key.split('|');
    data.getRange(i + 1, 2, 1, 2).setNumberFormat('@').setValues([[parts[0], parts[1]]]);
  }

  /* 勤務予定シートを全部作り直します */
  const plan = sheet_(PLAN_SHEET);
  const head = plan.getDataRange().getValues()[0];
  plan.clearContents();
  plan.getRange('A:D').setNumberFormat('@');
  plan.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#e8edf5');

  let n = 0;
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]);
    if (key.indexOf('|') < 0) continue;
    const parts = key.split('|');
    try {
      writePlanSheet_(parts[0], parts[1], JSON.parse(rows[i][3]));
      n++;
    } catch (e) { /* 壊れている行は飛ばします */ }
  }

  const msg = '書式を修正し、勤務予定シートを ' + n + ' 件ぶん作り直しました。';
  Logger.log(msg);
  return msg;
}


/* ============================================================
   補助
   ============================================================ */
/** パスワードのシートを取得します（旧名「合言葉」にも対応） */
function passSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(PASS_SHEET) || ss.getSheetByName(PASS_SHEET_OLD)
      || sheet_(PASS_SHEET);
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。setup() を実行してください。');
  return sh;
}

function makeKey_(ym, dept) { return String(ym) + '|' + String(dept); }

function getEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function getConf_(item, fallback) {
  try {
    const rows = sheet_(CONF_SHEET).getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === item) return rows[i][1];
    }
  } catch (e) { /* 未設定 */ }
  return fallback;
}

function isAdminMail_(email) {
  if (!email) return false;
  const raw = String(getConf_('管理者メール', ''));
  const list = raw.split(/[,、\s]+/).map(function (s) { return s.trim().toLowerCase(); })
                  .filter(String);
  return list.indexOf(email.toLowerCase()) >= 0;
}

function getDepartments_() {
  const raw = String(getConf_('部署', DEFAULT_DEPARTMENTS.join(',')));
  const list = raw.split(/[,、]/).map(function (s) { return s.trim(); }).filter(String);
  return list.length ? list : DEFAULT_DEPARTMENTS;
}

/** パスワードの一覧を { 対象: パスワード } で返します */
function getPasscodes_() {
  const map = {};
  const rows = passSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const k = String(rows[i][0]).trim();
    const v = String(rows[i][1]).trim();
    if (k && v) map[k] = v;
  }
  return map;
}

/** その部署に登録済みの職員名を返します（氏名選択用） */
function rosterNames_(dept) {
  const names = [];
  try {
    const rows = sheet_(DATA_SHEET).getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]) !== String(dept)) continue;
      const d = JSON.parse(rows[i][3]);
      (d.staff || []).forEach(function (s) {
        const n = String(s.name || '').trim();
        if (n && names.indexOf(n) < 0) names.push(n);
      });
    }
  } catch (e) { /* データがなければ空で返します */ }
  return names.sort();
}

function logAccess_(result, dept, name, note) {
  try {
    sheet_(LOG_SHEET).appendRow([new Date(), result, dept || '', name || '', note || '']);
  } catch (e) { /* ログは失敗しても処理を止めません */ }
}
