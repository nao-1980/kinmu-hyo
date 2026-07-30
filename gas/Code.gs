/**
 * ============================================================
 * 勤務表作成ツール ── サーバー側（Google Apps Script）
 * ------------------------------------------------------------
 * スプレッドシートに保存し、Googleアカウントで権限を判定します。
 *
 * 使うシート
 *   設定        … 管理者・部署・必要出勤人数
 *   _data       … 保存の本体（年月×部署ごとのJSON）
 *   勤務予定     … 人が読む用（1行＝1人1日。保存のたびに作り直します）
 *   アクセスログ … 開いた人の記録
 *
 * ※ 初回は setup() を1回だけ実行してください。
 * ============================================================
 */

const DATA_SHEET = '_data';
const CONF_SHEET = '設定';
const PLAN_SHEET = '勤務予定';
const LOG_SHEET  = 'アクセスログ';

/* 既定値。setup() で「設定」シートに書き出されます */
const DEFAULT_DEPARTMENTS =
  ['管理部','事務','診療アシスタント','看護師','ドライバー','戸塚院','港南院'];


/* ============================================================
   画面の表示
   ============================================================ */
function doGet() {
  logAccess_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('勤務表作成ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/* ============================================================
   初回セットアップ
   スクリプトエディタで一度だけ実行してください
   ============================================================ */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /* --- 設定シート --- */
  let conf = ss.getSheetByName(CONF_SHEET);
  if (!conf) {
    conf = ss.insertSheet(CONF_SHEET);
    conf.getRange(1, 1, 4, 2).setValues([
      ['項目', '値'],
      ['管理者', Session.getActiveUser().getEmail()],
      ['部署', DEFAULT_DEPARTMENTS.join(', ')],
      ['必要出勤人数', 2],
    ]);
    conf.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8edf5');
    conf.setColumnWidth(1, 140);
    conf.setColumnWidth(2, 520);
    conf.getRange('A5').setValue(
      '※ 管理者はカンマ区切りで複数指定できます。ここに載っていない人は「職員」になります。');
    conf.getRange('A5').setFontColor('#5b6470').setFontSize(10);
  }

  /* --- データシート --- */
  let data = ss.getSheetByName(DATA_SHEET);
  if (!data) {
    data = ss.insertSheet(DATA_SHEET);
    data.getRange(1, 1, 1, 4)
        .setValues([['キー', '年月', '部署', '内容（JSON）']])
        .setFontWeight('bold').setBackground('#e8edf5');
    data.setColumnWidth(1, 160);
    data.setColumnWidth(4, 600);
    data.hideSheet();
  }

  /* --- 勤務予定シート --- */
  let plan = ss.getSheetByName(PLAN_SHEET);
  if (!plan) {
    plan = ss.insertSheet(PLAN_SHEET);
    plan.getRange(1, 1, 1, 8)
        .setValues([['年月', '部署', '社員番号', '氏名', '日付', '曜日', '勤務区分', '区分名称']])
        .setFontWeight('bold').setBackground('#e8edf5');
    plan.setFrozenRows(1);
  }

  /* --- アクセスログ --- */
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.getRange(1, 1, 1, 3)
       .setValues([['日時', 'メールアドレス', '権限']])
       .setFontWeight('bold').setBackground('#e8edf5');
    log.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert(
    'セットアップが完了しました。\n\n' +
    '「設定」シートで管理者と部署をご確認ください。\n' +
    'その後、デプロイ → 新しいデプロイ → ウェブアプリ で公開します。');
}


/* ============================================================
   画面が最初に呼ぶ処理
   ============================================================ */
function getBootstrap() {
  const email = getEmail_();
  return {
    email: email,
    role: isAdmin_(email) ? 'admin' : 'staff',
    departments: getDepartments_(),
    minStaff: getConf_('必要出勤人数', 2),
  };
}


/* ============================================================
   読み込み
   ============================================================ */
function loadMonth(ym, dept) {
  const key = makeKey_(ym, dept);
  const sh = sheet_(DATA_SHEET);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === key) {
      try {
        return JSON.parse(rows[i][3]);
      } catch (e) {
        throw new Error('保存データを読み取れませんでした（' + key + '）');
      }
    }
  }
  return null;   /* 未作成の月・部署 */
}


/* ============================================================
   保存
   payload = { staff, events, archived, history, actuals, actualsMeta }
   ============================================================ */
function saveMonth(ym, dept, payload) {
  const email = getEmail_();
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
    if (found > 0) sh.getRange(found, 1, 1, 4).setValues([rec]);
    else sh.appendRow(rec);

    writePlanSheet_(ym, dept, payload);

    return { ok: true, at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm'), by: email };

  } finally {
    lock.releaseLock();
  }
}


/* ------------------------------------------------------------
   人が読む用のシートを作り直します（1行＝1人1日）
   ------------------------------------------------------------ */
function writePlanSheet_(ym, dept, payload) {
  const sh = sheet_(PLAN_SHEET);
  const all = sh.getDataRange().getValues();
  const head = all[0];

  /* 対象の年月×部署を取り除きます */
  const keep = all.slice(1).filter(r => !(String(r[0]) === ym && String(r[1]) === dept));

  /* 追加ぶんを作ります */
  const dow = ['日','月','火','水','木','金','土'];
  const parts = String(ym).split('.');
  const y = Number(parts[0]), m = Number(parts[1]);
  const add = [];

  (payload.staff || []).forEach(st => {
    const name = String(st.name || '').trim();
    if (!name) return;
    const days = st.days || {};
    Object.keys(days).forEach(d => {
      const code = days[d];
      if (!code) return;
      const date = new Date(y, m - 1, Number(d));
      add.push([
        ym, dept, String(st.emp || ''), name,
        Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd'),
        dow[date.getDay()], code, (payload.labels || {})[code] || '',
      ]);
    });
  });

  add.sort((a, b) => (a[3] === b[3] ? (a[4] < b[4] ? -1 : 1) : (a[3] < b[3] ? -1 : 1)));

  const out = [head].concat(keep, add);
  sh.clearContents();
  if (out.length) sh.getRange(1, 1, out.length, head.length).setValues(out);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8edf5');
}


/* ============================================================
   補助
   ============================================================ */
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。setup() を実行してください。');
  return sh;
}

function makeKey_(ym, dept) {
  return String(ym) + '|' + String(dept);
}

function getEmail_() {
  return Session.getActiveUser().getEmail() || '';
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

function isAdmin_(email) {
  if (!email) return false;
  const raw = String(getConf_('管理者', ''));
  const list = raw.split(/[,、\s]+/).map(s => s.trim().toLowerCase()).filter(String);
  return list.indexOf(email.toLowerCase()) >= 0;
}

function getDepartments_() {
  const raw = String(getConf_('部署', DEFAULT_DEPARTMENTS.join(',')));
  const list = raw.split(/[,、]/).map(s => s.trim()).filter(String);
  return list.length ? list : DEFAULT_DEPARTMENTS;
}

function logAccess_() {
  try {
    const email = getEmail_();
    sheet_(LOG_SHEET).appendRow([new Date(), email, isAdmin_(email) ? '管理者' : '職員']);
  } catch (e) { /* ログは失敗しても処理を止めません */ }
}
