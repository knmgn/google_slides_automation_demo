// =============================================================================
// テスト用ダミーデータ生成スクリプト
// =============================================================================
// このファイルは Code.gs の動作確認専用です。本番環境では不要です。
//
// 【使い方】
//   Step 1. createTestEnvironment() を実行
//           → 旧SS・新SS・テスト用スライドが自動生成される
//           → ログに表示される 3 つの ID を Code.gs の CONFIG に設定する
//   Step 2. replaceChartSources() を実行（Code.gs のメイン処理）
//   Step 3. verifyReplacement() を実行して置換結果を確認する
//   Step 4. 確認後、cleanupTestFiles() でテスト用ファイルをゴミ箱へ移動する
// =============================================================================

/** Script Properties に保存するキー名 */
const _PROP = {
  OLD_SS_ID: 'SETUP_TEST_OLD_SS_ID',
  NEW_SS_ID: 'SETUP_TEST_NEW_SS_ID',
  PRES_ID:   'SETUP_TEST_PRES_ID',
};

// =============================================================================
// 公開関数
// =============================================================================

/**
 * テスト用ファイルを一式生成する。
 *
 * 生成物:
 *   - 旧スプレッドシート: 3シート・各1チャート（棒グラフ / 円グラフ / 折れ線グラフ）
 *   - 新スプレッドシート: 旧SSのコピー + 値・タイトル更新（チャートIDは同一）
 *   - テスト用スライド:   旧SSのチャートをリンク済みで3枚のスライドに埋め込み済み
 *
 * 実行後、ログ（または画面）に表示された ID を Code.gs の CONFIG へ貼り付けてください。
 */
function createTestEnvironment() {
  try {
    Logger.log('テスト環境を作成中...');

    Logger.log('（1/3）旧スプレッドシートを作成中...');
    const oldSS = _createOldSpreadsheet();
    Logger.log('旧スプレッドシート完了: ' + oldSS.getId());

    Logger.log('（2/3）新スプレッドシートを作成中（旧のコピー＋データ更新）...');
    const newSS = _createNewSpreadsheet(oldSS);
    Logger.log('新スプレッドシート完了: ' + newSS.getId());

    Logger.log('（3/3）テスト用プレゼンテーションを作成中...');
    const pres = _createTestPresentation(oldSS);
    Logger.log('プレゼンテーション完了: ' + pres.getId());

    // クリーンアップ用に ID を保存
    PropertiesService.getScriptProperties().setProperties({
      [_PROP.OLD_SS_ID]: oldSS.getId(),
      [_PROP.NEW_SS_ID]: newSS.getId(),
      [_PROP.PRES_ID]:   pres.getId(),
    });

    const msg = [
      '=== テスト環境の作成が完了しました ===',
      '',
      '【次のステップ】Code.gs の CONFIG を以下の値に設定してください:',
      '',
      "  OLD_SPREADSHEET_ID: '" + oldSS.getId() + "'",
      "  NEW_SPREADSHEET_ID: '" + newSS.getId() + "'",
      "  PRESENTATION_ID:    '" + pres.getId() + "'",
      '',
      '設定後に replaceChartSources() を実行し、',
      '完了したら verifyReplacement() で結果を確認してください。',
    ].join('\n');

    Logger.log(msg);
    _setupAlert('テスト環境作成完了', msg);

  } catch (e) {
    const msg = 'テスト環境の作成中にエラーが発生しました:\n' + e.message;
    Logger.log(msg);
    _setupAlert('エラー', msg);
  }
}

/**
 * テスト用プレゼンテーション内のチャートが新スプレッドシートを参照しているか確認する。
 * replaceChartSources() 実行後に呼び出してください。
 */
function verifyReplacement() {
  const props  = PropertiesService.getScriptProperties();
  const presId = props.getProperty(_PROP.PRES_ID);
  const oldId  = props.getProperty(_PROP.OLD_SS_ID);
  const newId  = props.getProperty(_PROP.NEW_SS_ID);

  if (!presId || !oldId || !newId) {
    _setupAlert('エラー', 'createTestEnvironment() を先に実行してください。');
    return;
  }

  const pres = SlidesApp.openById(presId);
  const rows = [];
  let total = 0;

  pres.getSlides().forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHEETS_CHART) return;
      const chart = el.asSheetsChart();
      const refId = chart.getSpreadsheetId();
      total++;

      let status;
      if      (refId === newId) status = '✅ 新SS（置換済み）';
      else if (refId === oldId) status = '❌ 旧SS（未置換）';
      else                      status = '⚠️ 不明 (' + refId.substring(0, 10) + '...)';

      rows.push('  スライド' + (i + 1) + ' / チャートID ' + chart.getChartId() + ': ' + status);
    });
  });

  const msg = total === 0
    ? 'チャートが見つかりませんでした。'
    : 'チャート確認結果（計 ' + total + ' 件）:\n\n' + rows.join('\n');

  Logger.log('[verifyReplacement]\n' + msg);
  _setupAlert('置換確認結果', msg);
}

/**
 * createTestEnvironment() で生成したファイルをゴミ箱へ移動する。
 * Code.gs の CONFIG も手動でリセットしてください。
 */
function cleanupTestFiles() {
  const props = PropertiesService.getScriptProperties();
  const ids   = [
    props.getProperty(_PROP.OLD_SS_ID),
    props.getProperty(_PROP.NEW_SS_ID),
    props.getProperty(_PROP.PRES_ID),
  ].filter(Boolean);

  let deleted = 0;
  ids.forEach(function(id) {
    try {
      DriveApp.getFileById(id).setTrashed(true);
      deleted++;
      Logger.log('ゴミ箱へ移動: ' + id);
    } catch (e) {
      Logger.log('削除失敗 (ID: ' + id + '): ' + e.message);
    }
  });

  Object.values(_PROP).forEach(function(key) {
    props.deleteProperty(key);
  });

  const msg = deleted + ' 件のテストファイルをゴミ箱に移動しました。\nCode.gs の CONFIG も手動でリセットしてください。';
  Logger.log(msg);
  _setupAlert('クリーンアップ完了', msg);
}

/**
 * Script Properties に保存されているチャートIDと参照先を一覧表示する。
 * 置換前後の状態比較やデバッグに使用する。
 */
function listChartIds() {
  const props  = PropertiesService.getScriptProperties();
  const oldId  = props.getProperty(_PROP.OLD_SS_ID);
  const newId  = props.getProperty(_PROP.NEW_SS_ID);
  const presId = props.getProperty(_PROP.PRES_ID);

  if (!oldId || !newId) {
    _setupAlert('エラー', 'createTestEnvironment() を先に実行してください。');
    return;
  }

  const lines = ['=== チャートID 一覧 ===', ''];

  // 旧スプレッドシート
  lines.push('【旧スプレッドシート: ' + oldId + '】');
  SpreadsheetApp.openById(oldId).getSheets().forEach(function(sheet) {
    sheet.getCharts().forEach(function(c) {
      lines.push('  シート「' + sheet.getName() + '」: チャートID = ' + c.getChartId());
    });
  });

  lines.push('');

  // 新スプレッドシート
  lines.push('【新スプレッドシート: ' + newId + '】');
  SpreadsheetApp.openById(newId).getSheets().forEach(function(sheet) {
    sheet.getCharts().forEach(function(c) {
      lines.push('  シート「' + sheet.getName() + '」: チャートID = ' + c.getChartId());
    });
  });

  if (presId) {
    lines.push('');
    lines.push('【プレゼンテーション: ' + presId + '】');
    SlidesApp.openById(presId).getSlides().forEach(function(slide, i) {
      slide.getPageElements().forEach(function(el) {
        if (el.getPageElementType() !== SlidesApp.PageElementType.SHEETS_CHART) return;
        const c = el.asSheetsChart();
        lines.push(
          '  スライド' + (i + 1) + ': チャートID = ' + c.getChartId() +
          '  参照SS: ' + c.getSpreadsheetId().substring(0, 10) + '...'
        );
      });
    });
  }

  const msg = lines.join('\n');
  Logger.log(msg);
  _setupAlert('チャートID 一覧', msg);
}

// =============================================================================
// 内部ヘルパー — スプレッドシート作成
// =============================================================================

/**
 * ダミーデータとチャートを 3 シート分含む旧スプレッドシートを作成して返す。
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function _createOldSpreadsheet() {
  const ss = SpreadsheetApp.create('[テスト] 旧スプレッドシート（参照元）');

  // ---- Sheet 1: 月別売上（棒グラフ） ----
  const s1 = ss.getActiveSheet();
  s1.setName('月別売上');
  s1.getRange('A1:D7').setValues([
    ['月',  '商品A', '商品B', '商品C'],
    ['1月',  120,     95,      60],
    ['2月',  135,    110,      75],
    ['3月',  150,    125,      90],
    ['4月',  165,    140,     105],
    ['5月',  180,    155,     120],
    ['6月',  195,    170,     135],
  ]);
  s1.getRange('A1:D1').setFontWeight('bold').setBackground('#e8f0fe');

  s1.insertChart(
    s1.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(s1.getRange('A1:D7'))
      .setPosition(9, 1, 0, 0)
      .setOption('title', '月別売上【旧データ】')
      .setOption('legend', { position: 'right' })
      .build()
  );

  // ---- Sheet 2: カテゴリ別（円グラフ） ----
  const s2 = ss.insertSheet('カテゴリ別');
  s2.getRange('A1:B5').setValues([
    ['カテゴリ', '売上金額'],
    ['食品',     450000],
    ['衣料品',   320000],
    ['電子機器', 580000],
    ['その他',   150000],
  ]);
  s2.getRange('A1:B1').setFontWeight('bold').setBackground('#e8f0fe');

  s2.insertChart(
    s2.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(s2.getRange('A1:B5'))
      .setPosition(7, 1, 0, 0)
      .setOption('title', 'カテゴリ別売上【旧データ】')
      .build()
  );

  // ---- Sheet 3: 月別達成率（折れ線グラフ） ----
  const s3 = ss.insertSheet('月別達成率');
  s3.getRange('A1:B7').setValues([
    ['月',  '達成率(%)'],
    ['1月',  92],
    ['2月',  88],
    ['3月', 105],
    ['4月',  97],
    ['5月', 110],
    ['6月', 103],
  ]);
  s3.getRange('A1:B1').setFontWeight('bold').setBackground('#e8f0fe');

  s3.insertChart(
    s3.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(s3.getRange('A1:B7'))
      .setPosition(9, 1, 0, 0)
      .setOption('title', '月別目標達成率【旧データ】')
      .setOption('series', { 0: { color: '#e67c73' } })
      .build()
  );

  return ss;
}

/**
 * 旧スプレッドシートをコピーして新スプレッドシートを作成する。
 * makeCopy() によりチャートIDが引き継がれ、その後データとタイトルのみ更新する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} oldSS
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function _createNewSpreadsheet(oldSS) {
  const copy  = DriveApp.getFileById(oldSS.getId()).makeCopy('[テスト] 新スプレッドシート（参照先）');
  const newSS = SpreadsheetApp.openById(copy.getId());

  // Sheet 1: 月別売上 — 値を微増・タイトル更新
  const s1 = newSS.getSheetByName('月別売上');
  if (s1) {
    s1.getRange('B2:D7').setValues([
      [130, 105,  68],
      [148, 122,  83],
      [162, 139,  98],
      [178, 153, 113],
      [195, 167, 128],
      [210, 182, 143],
    ]);
    _updateChartTitle(s1, '月別売上【新データ】');
  }

  // Sheet 2: カテゴリ別 — 値更新・タイトル更新
  const s2 = newSS.getSheetByName('カテゴリ別');
  if (s2) {
    s2.getRange('B2:B5').setValues([[490000], [355000], [630000], [175000]]);
    _updateChartTitle(s2, 'カテゴリ別売上【新データ】');
  }

  // Sheet 3: 月別達成率 — 値更新・タイトル更新
  const s3 = newSS.getSheetByName('月別達成率');
  if (s3) {
    s3.getRange('B2:B7').setValues([[95], [91], [108], [100], [114], [107]]);
    _updateChartTitle(s3, '月別目標達成率【新データ】');
  }

  return newSS;
}

/**
 * シート内の最初のチャートのタイトルを更新する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} newTitle
 */
function _updateChartTitle(sheet, newTitle) {
  const charts = sheet.getCharts();
  if (charts.length > 0) {
    sheet.updateChart(charts[0].modify().setOption('title', newTitle).build());
  }
}

// =============================================================================
// 内部ヘルパー — プレゼンテーション作成
// =============================================================================

/**
 * 旧スプレッドシートのチャートをリンク済みで埋め込んだ
 * テスト用プレゼンテーションを作成して返す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} oldSS
 * @returns {GoogleAppsScript.Slides.Presentation}
 */
function _createTestPresentation(oldSS) {
  const pres = SlidesApp.create('[テスト] スライド（チャート参照置換テスト用）');
  const W = pres.getPageWidth();
  const H = pres.getPageHeight();

  const PADDING   = 40;
  const LABEL_H   = 36;
  const LABEL_TOP = 8;

  // スライド 1: 月別売上（棒グラフ）
  const slide1 = pres.getSlides()[0];
  _clearSlide(slide1);
  _insertChartFromSheet(slide1, oldSS, '月別売上', PADDING, LABEL_TOP + LABEL_H, W - PADDING * 2, H - LABEL_H - PADDING - LABEL_TOP);
  _addSlideLabel(slide1, '月別売上', W, LABEL_TOP, W - PADDING * 2, LABEL_H);

  // スライド 2: カテゴリ別（円グラフ）
  const slide2 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const pieW = W * 0.6;
  _insertChartFromSheet(slide2, oldSS, 'カテゴリ別', (W - pieW) / 2, LABEL_TOP + LABEL_H, pieW, H - LABEL_H - PADDING - LABEL_TOP);
  _addSlideLabel(slide2, 'カテゴリ別売上', W, LABEL_TOP, W - PADDING * 2, LABEL_H);

  // スライド 3: 月別達成率（折れ線グラフ）
  const slide3 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  _insertChartFromSheet(slide3, oldSS, '月別達成率', PADDING, LABEL_TOP + LABEL_H, W - PADDING * 2, H - LABEL_H - PADDING - LABEL_TOP);
  _addSlideLabel(slide3, '月別目標達成率', W, LABEL_TOP, W - PADDING * 2, LABEL_H);

  return pres;
}

/**
 * 指定シートの最初のチャートをスライドにリンク済みで挿入する。
 */
function _insertChartFromSheet(slide, ss, sheetName, left, top, width, height) {
  const sheet  = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const charts = sheet.getCharts();
  if (charts.length === 0) return;
  slide.insertSheetsChart(charts[0], left, top, width, height);
}

/**
 * スライドの既存要素をすべて削除する（デフォルトのプレースホルダを含む）。
 */
function _clearSlide(slide) {
  slide.getPageElements().forEach(function(el) {
    try { el.remove(); } catch (_) {}
  });
}

/**
 * スライド上部にラベルのテキストボックスを追加する。
 */
function _addSlideLabel(slide, text, slideWidth, top, width, height) {
  var left = (slideWidth - width) / 2;
  var box  = slide.insertTextBox(text, left, top, width, height);
  var style = box.getText().getTextStyle();
  style.setFontSize(13);
  style.setForegroundColor('#444444');
  box.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// =============================================================================
// ユーティリティ
// =============================================================================

/**
 * アラートを表示する（スクリプトエディタ直接実行時はログのみ）。
 */
function _setupAlert(title, message) {
  Logger.log('[' + title + ']\n' + message);
  try {
    var ui = SlidesApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (_) {}
}
