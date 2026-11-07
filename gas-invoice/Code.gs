/**
 * 請求書PDF 自動作成ツール（Google Apps Script）
 *
 * 「明細」シートに入力した請求データから、請求番号ごとに請求書を作成し、
 * PDFにして Google ドライブに保存します。インボイス制度の記載項目
 * （登録番号・税率ごとの合計と消費税額）に対応しています。
 *
 * 使い方：メニュー「請求書」→「① 初期設定（サンプル作成）」→「② 未発行の請求書を作成」
 * メニューが出ない場合：Apps Script の画面上部で関数「setup」や「createPendingInvoices」を
 * 選んで「実行」しても同じ処理ができます（結果は画面下の実行ログに表示されます）。
 */

const SHEET = {
  ITEMS: '明細',
  SETTINGS: '設定',
  TEMPLATE: '請求書テンプレート',
};

// 明細シートの列（1始まり）
const COL = {
  NO: 1,        // 請求番号
  DATE: 2,      // 請求日
  DUE: 3,       // お支払期限
  CLIENT: 4,    // 取引先名
  PERSON: 5,    // ご担当者
  ITEM: 6,      // 品目
  QTY: 7,       // 数量
  PRICE: 8,     // 単価（税抜）
  TAX: 9,       // 税率
  STATUS: 10,   // ステータス
  PDF: 11,      // PDFリンク
};
const HEADERS = ['請求番号', '請求日', 'お支払期限', '取引先名', 'ご担当者', '品目', '数量', '単価（税抜）', '税率', 'ステータス', 'PDFリンク'];
const STATUS_DONE = '発行済み';
const TEMPLATE_ITEM_START = 18; // テンプレートの明細開始行
const TEMPLATE_ITEM_ROWS = 10;  // 1枚に載せられる明細の行数

/** スプレッドシートを開いたときにメニューを追加 */
function onOpen() {
  const ui = getUi_();
  if (!ui) return;
  ui.createMenu('請求書')
    .addItem('① 初期設定（サンプル作成）', 'setup')
    .addSeparator()
    .addItem('② 未発行の請求書を作成', 'createPendingInvoices')
    .addItem('③ 選択中の行の請求書を作成', 'createSelectedInvoice')
    .addToUi();
}

/* =========================================================
 * 初期設定：シートとサンプルデータを作る
 * （メニューからでも、Apps Script のエディタから「実行」しても動く）
 * ======================================================= */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const ui = getUi_();
  if (ss.getSheetByName(SHEET.ITEMS)) {
    if (ui) {
      const res = ui.alert('すでに「明細」シートがあります。設定とテンプレートだけ作り直しますか？（明細は消えません）', ui.ButtonSet.OK_CANCEL);
      if (res !== ui.Button.OK) return;
    }
  } else {
    createItemsSheet_(ss);
  }
  createSettingsSheet_(ss);
  createTemplateSheet_(ss);
  ss.setActiveSheet(ss.getSheetByName(SHEET.ITEMS));
  alert_('初期設定が完了しました。\n「設定」シートに自社情報を入力してから、②を実行してください。');
}

function createItemsSheet_(ss) {
  const sh = ss.insertSheet(SHEET.ITEMS, 0);
  const sample = [
    ['INV-2026-001', '2026/11/01', '2026/11/30', '株式会社サンプル商事', '山田 太郎', 'Webサイト制作一式', 1, 150000, 0.1, '', ''],
    ['INV-2026-001', '2026/11/01', '2026/11/30', '株式会社サンプル商事', '山田 太郎', 'サーバー設定作業', 2, 10000, 0.1, '', ''],
    ['INV-2026-002', '2026/11/01', '2026/11/30', '合同会社テスト食品', '佐藤 花子', '商品撮影', 1, 30000, 0.1, '', ''],
    ['INV-2026-002', '2026/11/01', '2026/11/30', '合同会社テスト食品', '佐藤 花子', '打合せ用お茶菓子（軽減税率）', 3, 1200, 0.08, '', ''],
  ];
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#4a3326').setFontColor('#ffffff');
  sh.getRange(2, 1, sample.length, HEADERS.length).setValues(sample);
  sh.getRange('B:C').setNumberFormat('yyyy/mm/dd');
  sh.getRange('H:H').setNumberFormat('#,##0');
  sh.getRange('I:I').setNumberFormat('0%');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, HEADERS.length);
}

function createSettingsSheet_(ss) {
  let sh = ss.getSheetByName(SHEET.SETTINGS);
  const current = sh ? readSettings_() : {};
  if (sh) sh.clear(); else sh = ss.insertSheet(SHEET.SETTINGS);
  const rows = [
    ['項目', '内容', '説明'],
    ['会社名', current['会社名'] || '〇〇デザイン事務所', '請求書の発行者名'],
    ['住所', current['住所'] || '〒000-0000 東京都〇〇区〇〇 1-2-3', ''],
    ['電話番号', current['電話番号'] || '00-0000-0000', ''],
    ['メール', current['メール'] || 'info@example.com', ''],
    ['登録番号', current['登録番号'] || 'T0000000000000', '適格請求書発行事業者の登録番号（T＋13桁）'],
    ['振込先', current['振込先'] || '〇〇銀行 〇〇支店 普通 1234567 ﾏﾙﾏﾙﾃﾞｻﾞｲﾝ', ''],
    ['保存先フォルダID', current['保存先フォルダID'] || '', '空欄なら「請求書PDF」フォルダを自動作成'],
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4a3326').setFontColor('#ffffff');
  sh.getRange(2, 3, rows.length - 1, 1).setFontColor('#888888');
  sh.setColumnWidth(1, 140).setColumnWidth(2, 360).setColumnWidth(3, 320);
}

function createTemplateSheet_(ss) {
  let sh = ss.getSheetByName(SHEET.TEMPLATE);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET.TEMPLATE);
  sh.setHiddenGridlines(true);
  [30, 230, 70, 100, 70, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A1:F60').setFontFamily('Noto Sans JP').setFontSize(10).setVerticalAlignment('middle');

  sh.getRange('A2:F2').merge().setValue('請 求 書').setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('E4').setValue('請求番号'); sh.getRange('F4').setHorizontalAlignment('right');
  sh.getRange('E5').setValue('請求日');   sh.getRange('F5').setHorizontalAlignment('right').setNumberFormat('yyyy年m月d日');

  sh.getRange('A7:C7').merge().setFontSize(14).setFontWeight('bold')
    .setBorder(null, null, true, null, null, null);
  sh.getRange('A8:C8').merge();

  // 発行者（右側）
  sh.getRange('D7:F7').merge().setFontWeight('bold').setHorizontalAlignment('right');
  ['D8:F8', 'D9:F9', 'D10:F10', 'D11:F11'].forEach((a) => sh.getRange(a).merge().setHorizontalAlignment('right').setFontSize(9));

  sh.getRange('A13:C13').merge().setValue('下記のとおりご請求申し上げます。');
  sh.getRange('A14:B15').merge().setValue('ご請求金額（税込）').setFontWeight('bold')
    .setBackground('#f3ece2').setHorizontalAlignment('center');
  sh.getRange('C14:D15').merge().setFontSize(18).setFontWeight('bold').setNumberFormat('"¥"#,##0')
    .setHorizontalAlignment('center').setBorder(true, true, true, true, null, null);
  sh.getRange('E14').setValue('お支払期限'); sh.getRange('F14').setNumberFormat('yyyy/m/d').setHorizontalAlignment('right');

  // 明細ヘッダー
  sh.getRange(TEMPLATE_ITEM_START - 1, 1, 1, 6)
    .setValues([['No', '品目', '数量', '単価', '税率', '金額（税抜）']])
    .setBackground('#4a3326').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  const items = sh.getRange(TEMPLATE_ITEM_START, 1, TEMPLATE_ITEM_ROWS, 6);
  items.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(TEMPLATE_ITEM_START, 1, TEMPLATE_ITEM_ROWS, 1).setHorizontalAlignment('center');
  sh.getRange(TEMPLATE_ITEM_START, 4, TEMPLATE_ITEM_ROWS, 1).setNumberFormat('#,##0');
  sh.getRange(TEMPLATE_ITEM_START, 5, TEMPLATE_ITEM_ROWS, 1).setNumberFormat('0%').setHorizontalAlignment('center');
  sh.getRange(TEMPLATE_ITEM_START, 6, TEMPLATE_ITEM_ROWS, 1).setNumberFormat('#,##0');

  // 合計欄（税率ごと）
  const sumTop = TEMPLATE_ITEM_START + TEMPLATE_ITEM_ROWS + 1;
  const labels = [['10%対象（税抜）'], ['消費税（10%）'], ['8%対象（税抜）※軽減税率'], ['消費税（8%）'], ['合計（税込）']];
  sh.getRange(sumTop, 4, labels.length, 2).mergeAcross();
  sh.getRange(sumTop, 4, labels.length, 1).setValues(labels).setHorizontalAlignment('right').setFontSize(9);
  sh.getRange(sumTop, 6, labels.length, 1).setNumberFormat('"¥"#,##0').setHorizontalAlignment('right');
  sh.getRange(sumTop + labels.length - 1, 4, 1, 3).setFontWeight('bold').setBackground('#f3ece2');

  const noteTop = sumTop + labels.length + 1;
  sh.getRange(noteTop, 1, 1, 6).merge().setValue('【お振込先】').setFontWeight('bold');
  sh.getRange(noteTop + 1, 1, 1, 6).merge();
  sh.getRange(noteTop + 2, 1, 1, 6).merge().setValue('※お振込手数料は貴社にてご負担くださいますようお願いいたします。').setFontSize(9).setFontColor('#666666');
}

/* =========================================================
 * 請求書の作成
 * ======================================================= */

/** ステータスが空の請求番号をすべて作成 */
function createPendingInvoices() {
  const groups = groupItems_().filter((g) => g.rows.some((r) => r.values[COL.STATUS - 1] !== STATUS_DONE));
  if (!groups.length) {
    alert_('未発行の請求書はありません。');
    return;
  }
  const done = groups.map(createInvoice_);
  alert_(`${done.length}件の請求書を作成しました。\n\n${done.join('\n')}`);
}

/** 選択中のセルがある行の請求番号を作成（再発行にも使える） */
function createSelectedInvoice() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  if (sh.getName() !== SHEET.ITEMS) {
    ui.alert('「明細」シートで、作成したい請求書の行を選択してから実行してください。');
    return;
  }
  const row = sh.getActiveRange().getRow();
  const no = row > 1 ? sh.getRange(row, COL.NO).getValue() : '';
  const group = no && groupItems_().find((g) => g.no === no);
  if (!group) {
    ui.alert('請求番号が入った行を選択してください。');
    return;
  }
  ui.alert(`請求書を作成しました。\n\n${createInvoice_(group)}`);
}

/** 明細を請求番号ごとにまとめる */
function groupItems_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ITEMS);
  if (!sh) throw new Error('「明細」シートがありません。メニューの「① 初期設定」を実行してください。');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const map = new Map();
  values.forEach((v, i) => {
    const no = String(v[COL.NO - 1]).trim();
    if (!no) return;
    if (!map.has(no)) map.set(no, { no, rows: [] });
    map.get(no).rows.push({ rowIndex: i + 2, values: v });
  });
  return [...map.values()];
}

/** 1件分の請求書を作ってPDF保存し、明細シートに結果を書き戻す */
function createInvoice_(group) {
  const ss = SpreadsheetApp.getActive();
  const settings = readSettings_();
  const tpl = ss.getSheetByName(SHEET.TEMPLATE);
  if (!tpl) throw new Error('「請求書テンプレート」シートがありません。①初期設定を実行してください。');
  if (group.rows.length > TEMPLATE_ITEM_ROWS) {
    throw new Error(`${group.no}：明細が${TEMPLATE_ITEM_ROWS}行を超えています。請求番号を分けてください。`);
  }

  const first = group.rows[0].values;
  const items = group.rows.map(({ values: v }) => {
    const qty = Number(v[COL.QTY - 1]) || 0;
    const price = Number(v[COL.PRICE - 1]) || 0;
    const tax = normalizeTax_(v[COL.TAX - 1]);
    return { name: v[COL.ITEM - 1], qty, price, tax, amount: qty * price };
  });

  // 税率ごとに合計し、消費税は税率ごとに1回だけ端数処理（切り捨て）する
  const base10 = sum_(items.filter((x) => x.tax === 0.1).map((x) => x.amount));
  const base8 = sum_(items.filter((x) => x.tax === 0.08).map((x) => x.amount));
  const tax10 = Math.floor(base10 * 0.1);
  const tax8 = Math.floor(base8 * 0.08);
  const total = base10 + tax10 + base8 + tax8;

  // テンプレートをコピーして値を入れる
  const sh = tpl.copyTo(ss).setName(`tmp_${group.no}_${Date.now()}`);
  try {
    sh.getRange('F4').setValue(group.no);
    sh.getRange('F5').setValue(first[COL.DATE - 1]);
    sh.getRange('A7').setValue(`${first[COL.CLIENT - 1]}　御中`);
    sh.getRange('A8').setValue(first[COL.PERSON - 1] ? `${first[COL.PERSON - 1]} 様` : '');
    sh.getRange('D7').setValue(settings['会社名'] || '');
    sh.getRange('D8').setValue(settings['住所'] || '');
    sh.getRange('D9').setValue(`TEL ${settings['電話番号'] || ''}`);
    sh.getRange('D10').setValue(settings['メール'] || '');
    sh.getRange('D11').setValue(`登録番号：${settings['登録番号'] || ''}`);
    sh.getRange('C14').setValue(total);
    sh.getRange('F14').setValue(first[COL.DUE - 1]);

    const rows = items.map((x, i) => [i + 1, x.name, x.qty, x.price, x.tax, x.amount]);
    sh.getRange(TEMPLATE_ITEM_START, 1, rows.length, 6).setValues(rows);

    const sumTop = TEMPLATE_ITEM_START + TEMPLATE_ITEM_ROWS + 1;
    sh.getRange(sumTop, 6, 5, 1).setValues([[base10], [tax10], [base8], [tax8], [total]]);
    sh.getRange(sumTop + 5 + 2, 1).setValue(settings['振込先'] || '');

    SpreadsheetApp.flush();
    const file = exportSheetAsPdf_(ss, sh, `${group.no}_${first[COL.CLIENT - 1]}.pdf`, settings['保存先フォルダID']);

    // 明細シートに結果を書き戻す
    const items_ = ss.getSheetByName(SHEET.ITEMS);
    group.rows.forEach(({ rowIndex }) => {
      items_.getRange(rowIndex, COL.STATUS).setValue(STATUS_DONE);
      items_.getRange(rowIndex, COL.PDF).setValue(file.getUrl());
    });
    return `${group.no}（${first[COL.CLIENT - 1]}） ¥${total.toLocaleString()}`;
  } finally {
    ss.deleteSheet(sh);
  }
}

/** 指定シートだけをA4縦のPDFにしてドライブに保存 */
function exportSheetAsPdf_(ss, sheet, fileName, folderId) {
  const params = {
    format: 'pdf', size: 'A4', portrait: 'true', fitw: 'true',
    gridlines: 'false', printtitle: 'false', sheetnames: 'false', pagenum: 'UNDEFINED',
    top_margin: '0.6', bottom_margin: '0.6', left_margin: '0.6', right_margin: '0.6',
    gid: sheet.getSheetId(),
  };
  const query = Object.keys(params).map((k) => `${k}=${encodeURIComponent(params[k])}`).join('&');
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?${query}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`PDFの作成に失敗しました（HTTP ${res.getResponseCode()}）`);
  }
  const blob = res.getBlob().setName(fileName);
  return getOutputFolder_(folderId).createFile(blob);
}

/* =========================================================
 * ユーティリティ
 * ======================================================= */

/** メニューから実行したときは画面のUI、エディタから実行したときは null */
function getUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

/** 画面にメッセージを出す（エディタから実行したときは実行ログに出す） */
function alert_(message) {
  const ui = getUi_();
  if (ui) {
    try {
      ui.alert(message);
      return;
    } catch (e) {
      // エディタ実行時は alert が使えないため、ログに出す
    }
  }
  console.log(message);
}
function readSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return {};
  return Object.fromEntries(sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().map(([k, v]) => [k, v]));
}

function getOutputFolder_(folderId) {
  if (folderId) return DriveApp.getFolderById(folderId);
  const name = '請求書PDF';
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/** 「10%」「0.1」「10」などの入力を 0.1 / 0.08 にそろえる */
function normalizeTax_(v) {
  let n = typeof v === 'string' ? parseFloat(v.replace('%', '')) : Number(v);
  if (n >= 1) n = n / 100;
  if (Math.abs(n - 0.08) < 1e-9) return 0.08;
  return 0.1;
}

function sum_(arr) {
  return arr.reduce((a, b) => a + b, 0);
}
