// =============================================================================
// Googleスライド チャート参照先一括置換スクリプト
// =============================================================================
// 使い方:
//   1. CONFIG の OLD_SPREADSHEET_ID・NEW_SPREADSHEET_ID を設定する
//   2. スクリプトエディタの「実行」メニューから replaceChartSources() を実行する
//      または、スライドのメニュー「チャート置換」から操作する
// =============================================================================

/** ユーザー設定 */
const CONFIG = {
  /** 置換前スプレッドシートのID（現在チャートが参照しているスプレッドシート） */
  OLD_SPREADSHEET_ID: 'YOUR_OLD_SPREADSHEET_ID',

  /** 置換後スプレッドシートのID（新しく参照させたいスプレッドシート） */
  NEW_SPREADSHEET_ID: 'YOUR_NEW_SPREADSHEET_ID',

  /**
   * 対象プレゼンテーションのID。
   * 空文字列の場合、スクリプトが紐付いているプレゼンテーションを自動で使用する。
   */
  PRESENTATION_ID: '',
};

// =============================================================================
// UI メニュー
// =============================================================================

/**
 * スライドを開いたときにカスタムメニューを追加する。
 * このトリガーはスライドが開かれたときに自動で呼ばれる。
 * スクリプトエディタから手動実行した場合は SlidesApp.getUi() が使えないため
 * エラーを無視して終了する（メニューが追加されないだけで問題ない）。
 * ※ replaceChartSources() を直接実行したい場合は、ドロップダウンで
 *    「replaceChartSources」を選択してから「実行」ボタンを押してください。
 */
function onOpen() {
  try {
    SlidesApp.getUi()
      .createMenu('チャート置換')
      .addItem('ドライラン（置換対象の確認）', 'dryRun')
      .addSeparator()
      .addItem('チャート参照先を置換する', 'replaceChartSources')
      .addToUi();
  } catch (e) {
    // スクリプトエディタや非Slidesコンテキストから実行された場合は無視
    Logger.log('onOpen: メニューの追加をスキップしました（' + e.message + '）');
  }
}

// =============================================================================
// メイン処理
// =============================================================================

/**
 * スライド内のすべてのリンク済みチャートについて、
 * OLD_SPREADSHEET_ID を参照しているものを NEW_SPREADSHEET_ID へ一括置換する。
 * チャートの位置・サイズは元のまま維持される。
 */
function replaceChartSources() {
  const { OLD_SPREADSHEET_ID: oldId, NEW_SPREADSHEET_ID: newId, PRESENTATION_ID: presId } = CONFIG;

  if (!_validateConfig(oldId, newId)) return;

  const presentation = _openPresentation(presId);
  if (!presentation) return;

  const newSpreadsheet = _openSpreadsheet(newId);
  if (!newSpreadsheet) return;

  const results = { replaced: 0, skipped: 0, notFound: [], errors: [] };

  presentation.getSlides().forEach((slide, i) => {
    _processSlide(slide, i + 1, oldId, newSpreadsheet, results);
  });

  _showReport(results);
}

/**
 * 実際の置換を行わず、置換対象となるチャートの一覧をログとアラートで表示する。
 * 本番実行前の確認に使用する。
 */
function dryRun() {
  const { OLD_SPREADSHEET_ID: oldId, PRESENTATION_ID: presId } = CONFIG;

  if (oldId === 'YOUR_OLD_SPREADSHEET_ID') {
    _alert('設定エラー', 'CONFIG の OLD_SPREADSHEET_ID を設定してください。');
    return;
  }

  const presentation = _openPresentation(presId);
  if (!presentation) return;

  const found = [];

  presentation.getSlides().forEach((slide, i) => {
    _collectSheetsCharts(slide).forEach(({ element }) => {
      const chart = element.asSheetsChart();
      if (chart.getSpreadsheetId() === oldId) {
        found.push(`スライド ${i + 1}：チャートID = ${chart.getChartId()}`);
      }
    });
  });

  const msg =
    found.length > 0
      ? `置換対象チャート ${found.length} 件\n\n${found.join('\n')}`
      : '置換対象のチャートは見つかりませんでした。\nOLD_SPREADSHEET_ID の設定を確認してください。';

  Logger.log('[ドライラン]\n' + msg);
  _alert('ドライラン結果', msg);
}

// =============================================================================
// 内部ヘルパー
// =============================================================================

/**
 * 1枚のスライドに含まれるすべてのチャート要素を対象に置換処理を行う。
 * グループ直下のチャートも処理するが、グループのネストが深い場合は対応外。
 *
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {number} slideNumber 1始まりのスライド番号（ログ用）
 * @param {string} oldId 置換前スプレッドシートID
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} newSpreadsheet 置換後スプレッドシート
 * @param {Object} results 集計オブジェクト
 */
function _processSlide(slide, slideNumber, oldId, newSpreadsheet, results) {
  _collectSheetsCharts(slide).forEach(({ element, inGroup }) => {
    const chart = element.asSheetsChart();

    if (chart.getSpreadsheetId() !== oldId) {
      results.skipped++;
      return;
    }

    _replaceChart(element, chart, slideNumber, newSpreadsheet, inGroup, results);
  });
}

/**
 * スライド上の SheetsChart 要素をすべて収集して返す。
 * トップレベル要素とグループ直下の要素を対象にする。
 *
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @returns {{ element: GoogleAppsScript.Slides.PageElement, inGroup: boolean }[]}
 */
function _collectSheetsCharts(slide) {
  const results = [];
  const CHART = SlidesApp.PageElementType.SHEETS_CHART;
  const GROUP = SlidesApp.PageElementType.GROUP;

  slide.getPageElements().forEach(el => {
    const type = el.getPageElementType();
    if (type === CHART) {
      results.push({ element: el, inGroup: false });
    } else if (type === GROUP) {
      el.asGroup().getChildren().forEach(child => {
        if (child.getPageElementType() === CHART) {
          results.push({ element: child, inGroup: true });
        }
      });
    }
  });

  return results;
}

/**
 * 指定されたチャート要素を新スプレッドシートのチャートで置換する。
 *
 * @param {GoogleAppsScript.Slides.PageElement} element
 * @param {GoogleAppsScript.Slides.SheetsChart} chart
 * @param {number} slideNumber
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} newSpreadsheet
 * @param {boolean} inGroup グループ内要素かどうか（ログ用）
 * @param {Object} results 集計オブジェクト
 */
function _replaceChart(element, chart, slideNumber, newSpreadsheet, inGroup, results) {
  const chartId = chart.getChartId();
  const label = `スライド ${slideNumber}、チャートID ${chartId}${inGroup ? '（グループ内）' : ''}`;

  try {
    const sourceChart = _findChartById(newSpreadsheet, chartId);
    if (!sourceChart) {
      const msg = `${label}: 新スプレッドシートにチャートID ${chartId} が見つかりません`;
      Logger.log(msg);
      results.notFound.push(msg);
      return;
    }

    // 位置・サイズを記録してから削除
    const left   = element.getLeft();
    const top    = element.getTop();
    const width  = element.getWidth();
    const height = element.getHeight();
    const slide  = element.getParentPage();

    element.remove();

    // 同一位置・サイズで新チャートを挿入（リンク済みとして挿入される）
    slide.insertSheetsChart(sourceChart, left, top, width, height);

    results.replaced++;
    Logger.log(`置換完了: ${label}`);
  } catch (e) {
    const msg = `${label}: ${e.message}`;
    Logger.log('エラー: ' + msg);
    results.errors.push(msg);
  }
}

/**
 * スプレッドシート内の全シートを横断してチャートIDで検索する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {number} chartId
 * @returns {GoogleAppsScript.Spreadsheet.EmbeddedChart|null}
 */
function _findChartById(spreadsheet, chartId) {
  for (const sheet of spreadsheet.getSheets()) {
    for (const c of sheet.getCharts()) {
      if (c.getChartId() === chartId) return c;
    }
  }
  return null;
}

// =============================================================================
// ユーティリティ
// =============================================================================

/**
 * CONFIG の基本バリデーション。問題があれば false を返す。
 */
function _validateConfig(oldId, newId) {
  if (oldId === 'YOUR_OLD_SPREADSHEET_ID' || newId === 'YOUR_NEW_SPREADSHEET_ID') {
    _alert('設定エラー', 'CONFIG の OLD_SPREADSHEET_ID と NEW_SPREADSHEET_ID を正しく設定してください。');
    return false;
  }
  if (oldId === newId) {
    _alert('設定エラー', 'OLD_SPREADSHEET_ID と NEW_SPREADSHEET_ID が同じ値です。');
    return false;
  }
  return true;
}

/**
 * プレゼンテーションを開く。失敗した場合は null を返す。
 *
 * @param {string} presId プレゼンテーションID（空の場合はアクティブなものを使用）
 * @returns {GoogleAppsScript.Slides.Presentation|null}
 */
function _openPresentation(presId) {
  try {
    return presId ? SlidesApp.openById(presId) : SlidesApp.getActivePresentation();
  } catch (e) {
    _alert('エラー', `プレゼンテーションを開けませんでした:\n${e.message}`);
    return null;
  }
}

/**
 * スプレッドシートを開く。失敗した場合は null を返す。
 *
 * @param {string} id スプレッドシートID
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null}
 */
function _openSpreadsheet(id) {
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    _alert('エラー', `スプレッドシートを開けませんでした（ID: ${id}）:\n${e.message}`);
    return null;
  }
}

/**
 * 置換結果のサマリーをログとアラートで表示する。
 *
 * @param {{ replaced: number, skipped: number, notFound: string[], errors: string[] }} results
 */
function _showReport(results) {
  const lines = [
    `✅ 置換完了: ${results.replaced} 件`,
    `⏭️ スキップ（参照先が異なる）: ${results.skipped} 件`,
  ];

  if (results.notFound.length > 0) {
    lines.push('\n⚠️ 新スプレッドシートでチャートが見つからなかったもの:');
    results.notFound.forEach(m => lines.push('  • ' + m));
  }

  if (results.errors.length > 0) {
    lines.push('\n❌ エラーが発生したもの:');
    results.errors.forEach(m => lines.push('  • ' + m));
  }

  const msg = lines.join('\n');
  Logger.log('[置換レポート]\n' + msg);
  _alert('置換レポート', msg);
}

/**
 * アラートダイアログを表示する。
 * スクリプトエディタから直接実行した場合はUIが使えないため、ログのみ出力する。
 *
 * @param {string} title
 * @param {string} message
 */
function _alert(title, message) {
  Logger.log(`[${title}] ${message}`);
  try {
    const ui = SlidesApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (_) {
    // スクリプトエディタ直接実行時はUIが利用できないため無視
  }
}
