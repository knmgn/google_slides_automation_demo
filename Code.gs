// =============================================================================
// Bulk Replace Chart Sources in Google Slides
// =============================================================================
// Usage:
//   1. Set OLD_SPREADSHEET_ID and NEW_SPREADSHEET_ID in CONFIG below.
//   2. Run replaceChartSources() from the script editor's Run menu,
//      or use the "Replace Charts" menu in the Slides UI.
// =============================================================================

/** User configuration */
const CONFIG = {
  /** ID of the spreadsheet currently referenced by the charts (the one to replace). */
  OLD_SPREADSHEET_ID: 'YOUR_OLD_SPREADSHEET_ID',

  /** ID of the new spreadsheet to redirect all charts to. */
  NEW_SPREADSHEET_ID: 'YOUR_NEW_SPREADSHEET_ID',

  /**
   * ID of the target presentation.
   * If left empty, the script uses the presentation it is bound to.
   */
  PRESENTATION_ID: '',
};

// =============================================================================
// UI Menu
// =============================================================================

/** Adds a custom menu to the Slides UI when the presentation is opened. */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('Replace Charts')
    .addItem('Dry run (preview targets)', 'dryRun')
    .addSeparator()
    .addItem('Replace chart sources', 'replaceChartSources')
    .addToUi();
}

// =============================================================================
// Main Functions
// =============================================================================

/**
 * Scans all linked Sheets charts in the presentation and re-links those
 * referencing OLD_SPREADSHEET_ID to NEW_SPREADSHEET_ID.
 * Chart position and size are preserved.
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
 * Lists all charts that would be replaced without making any changes.
 * Use this to verify targets before running the actual replacement.
 */
function dryRun() {
  const { OLD_SPREADSHEET_ID: oldId, PRESENTATION_ID: presId } = CONFIG;

  if (oldId === 'YOUR_OLD_SPREADSHEET_ID') {
    _alert('Config Error', 'Please set OLD_SPREADSHEET_ID in CONFIG.');
    return;
  }

  const presentation = _openPresentation(presId);
  if (!presentation) return;

  const found = [];

  presentation.getSlides().forEach((slide, i) => {
    _collectSheetsCharts(slide).forEach(({ element }) => {
      const chart = element.asSheetsChart();
      if (chart.getSpreadsheetId() === oldId) {
        found.push(`Slide ${i + 1}: Chart ID = ${chart.getChartId()}`);
      }
    });
  });

  const msg =
    found.length > 0
      ? `Found ${found.length} chart(s) to replace:\n\n${found.join('\n')}`
      : 'No charts found matching OLD_SPREADSHEET_ID.\nPlease verify the ID in CONFIG.';

  Logger.log('[Dry Run]\n' + msg);
  _alert('Dry Run Results', msg);
}

// =============================================================================
// Internal Helpers
// =============================================================================

/**
 * Processes all chart elements on a single slide.
 * Charts directly inside a group are included, but deeper nesting is not supported.
 *
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {number} slideNumber 1-based slide number (for logging)
 * @param {string} oldId Source spreadsheet ID to replace
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} newSpreadsheet Target spreadsheet
 * @param {Object} results Aggregation object
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
 * Collects all SheetsChart elements on the slide.
 * Includes top-level elements and those directly inside a group.
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
 * Replaces a single chart element with its counterpart from the new spreadsheet.
 *
 * @param {GoogleAppsScript.Slides.PageElement} element
 * @param {GoogleAppsScript.Slides.SheetsChart} chart
 * @param {number} slideNumber
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} newSpreadsheet
 * @param {boolean} inGroup Whether the element is inside a group (for logging)
 * @param {Object} results Aggregation object
 */
function _replaceChart(element, chart, slideNumber, newSpreadsheet, inGroup, results) {
  const chartId = chart.getChartId();
  const label = `Slide ${slideNumber}, Chart ID ${chartId}${inGroup ? ' (in group)' : ''}`;

  try {
    const sourceChart = _findChartById(newSpreadsheet, chartId);
    if (!sourceChart) {
      const msg = `${label}: chart ID ${chartId} not found in the new spreadsheet`;
      Logger.log(msg);
      results.notFound.push(msg);
      return;
    }

    // Record position and size before removing
    const left   = element.getLeft();
    const top    = element.getTop();
    const width  = element.getWidth();
    const height = element.getHeight();
    const slide  = element.getParentPage();

    element.remove();

    // Re-insert at the same position and size, linked to the new spreadsheet
    slide.insertSheetsChart(sourceChart, left, top, width, height);

    results.replaced++;
    Logger.log(`Replaced: ${label}`);
  } catch (e) {
    const msg = `${label}: ${e.message}`;
    Logger.log('Error: ' + msg);
    results.errors.push(msg);
  }
}

/**
 * Searches all sheets in the spreadsheet for a chart matching the given chart ID.
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
// Utilities
// =============================================================================

/**
 * Validates CONFIG. Returns false and shows an alert if something is wrong.
 */
function _validateConfig(oldId, newId) {
  if (oldId === 'YOUR_OLD_SPREADSHEET_ID' || newId === 'YOUR_NEW_SPREADSHEET_ID') {
    _alert('Config Error', 'Please set both OLD_SPREADSHEET_ID and NEW_SPREADSHEET_ID in CONFIG.');
    return false;
  }
  if (oldId === newId) {
    _alert('Config Error', 'OLD_SPREADSHEET_ID and NEW_SPREADSHEET_ID must be different.');
    return false;
  }
  return true;
}

/**
 * Opens the presentation. Returns null on failure.
 *
 * @param {string} presId Presentation ID. If empty, the active presentation is used.
 * @returns {GoogleAppsScript.Slides.Presentation|null}
 */
function _openPresentation(presId) {
  try {
    return presId ? SlidesApp.openById(presId) : SlidesApp.getActivePresentation();
  } catch (e) {
    _alert('Error', `Could not open the presentation:\n${e.message}`);
    return null;
  }
}

/**
 * Opens a spreadsheet by ID. Returns null on failure.
 *
 * @param {string} id Spreadsheet ID
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null}
 */
function _openSpreadsheet(id) {
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    _alert('Error', `Could not open the spreadsheet (ID: ${id}):\n${e.message}`);
    return null;
  }
}

/**
 * Logs and displays a summary of the replacement results.
 *
 * @param {{ replaced: number, skipped: number, notFound: string[], errors: string[] }} results
 */
function _showReport(results) {
  const lines = [
    `✅ Replaced: ${results.replaced}`,
    `⏭️ Skipped (different source): ${results.skipped}`,
  ];

  if (results.notFound.length > 0) {
    lines.push('\n⚠️ Charts not found in the new spreadsheet:');
    results.notFound.forEach(m => lines.push('  • ' + m));
  }

  if (results.errors.length > 0) {
    lines.push('\n❌ Errors:');
    results.errors.forEach(m => lines.push('  • ' + m));
  }

  const msg = lines.join('\n');
  Logger.log('[Replacement Report]\n' + msg);
  _alert('Replacement Report', msg);
}

/**
 * Shows an alert dialog. Falls back to log-only output when running
 * directly from the script editor, where no UI context is available.
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
    // No UI available when run directly from the script editor
  }
}
