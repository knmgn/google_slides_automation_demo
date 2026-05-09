// =============================================================================
// Test Data Generation Script
// =============================================================================
// This file is only for verifying Code.gs behavior. Not needed in production.
//
// Usage:
//   Step 1. Run createTestEnvironment()
//           → Creates the old SS, new SS, and a test presentation automatically.
//           → Copy the 3 IDs from the log into CONFIG in Code.gs.
//   Step 2. Run replaceChartSources() (the main function in Code.gs).
//   Step 3. Run verifyReplacement() to confirm the results.
//   Step 4. Run cleanupTestFiles() to move the test files to trash.
// =============================================================================

/** Script Property keys used for cleanup */
const _PROP = {
  OLD_SS_ID: 'SETUP_TEST_OLD_SS_ID',
  NEW_SS_ID: 'SETUP_TEST_NEW_SS_ID',
  PRES_ID:   'SETUP_TEST_PRES_ID',
};

// =============================================================================
// Public Functions
// =============================================================================

/**
 * Generates all test files in one shot.
 *
 * Creates:
 *   - Old spreadsheet: 3 sheets, one chart each (column / pie / line)
 *   - New spreadsheet: copy of the old one with updated values and titles
 *                      (chart IDs are identical due to makeCopy())
 *   - Test presentation: all 3 charts embedded as linked Sheets charts
 *                        from the old spreadsheet
 *
 * After running, paste the IDs shown in the log into CONFIG in Code.gs.
 */
function createTestEnvironment() {
  try {
    Logger.log('Creating test environment...');

    Logger.log('(1/3) Creating old spreadsheet...');
    const oldSS = _createOldSpreadsheet();
    Logger.log('Old spreadsheet created: ' + oldSS.getId());

    Logger.log('(2/3) Creating new spreadsheet (copy of old + data update)...');
    const newSS = _createNewSpreadsheet(oldSS);
    Logger.log('New spreadsheet created: ' + newSS.getId());

    Logger.log('(3/3) Creating test presentation...');
    const pres = _createTestPresentation(oldSS);
    Logger.log('Presentation created: ' + pres.getId());

    // Store IDs for later cleanup
    PropertiesService.getScriptProperties().setProperties({
      [_PROP.OLD_SS_ID]: oldSS.getId(),
      [_PROP.NEW_SS_ID]: newSS.getId(),
      [_PROP.PRES_ID]:   pres.getId(),
    });

    const msg = [
      '=== Test environment created ===',
      '',
      'Next step: update CONFIG in Code.gs with the following IDs:',
      '',
      "  OLD_SPREADSHEET_ID: '" + oldSS.getId() + "'",
      "  NEW_SPREADSHEET_ID: '" + newSS.getId() + "'",
      "  PRESENTATION_ID:    '" + pres.getId() + "'",
      '',
      'Then run replaceChartSources(), and use',
      'verifyReplacement() to confirm the results.',
    ].join('\n');

    Logger.log(msg);
    _setupAlert('Test Environment Ready', msg);

  } catch (e) {
    const msg = 'An error occurred while creating the test environment:\n' + e.message;
    Logger.log(msg);
    _setupAlert('Error', msg);
  }
}

/**
 * Checks whether all charts in the test presentation now reference the new
 * spreadsheet. Call this after running replaceChartSources().
 */
function verifyReplacement() {
  const props  = PropertiesService.getScriptProperties();
  const presId = props.getProperty(_PROP.PRES_ID);
  const oldId  = props.getProperty(_PROP.OLD_SS_ID);
  const newId  = props.getProperty(_PROP.NEW_SS_ID);

  if (!presId || !oldId || !newId) {
    _setupAlert('Error', 'Please run createTestEnvironment() first.');
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
      if      (refId === newId) status = '✅ New SS (replaced)';
      else if (refId === oldId) status = '❌ Old SS (not replaced)';
      else                      status = '⚠️ Unknown (' + refId.substring(0, 10) + '...)';

      rows.push('  Slide ' + (i + 1) + ' / Chart ID ' + chart.getChartId() + ': ' + status);
    });
  });

  const msg = total === 0
    ? 'No charts found.'
    : 'Verification results (' + total + ' chart(s)):\n\n' + rows.join('\n');

  Logger.log('[verifyReplacement]\n' + msg);
  _setupAlert('Verification Results', msg);
}

/**
 * Moves all files created by createTestEnvironment() to trash.
 * Remember to reset CONFIG in Code.gs manually afterward.
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
      Logger.log('Moved to trash: ' + id);
    } catch (e) {
      Logger.log('Failed to delete (ID: ' + id + '): ' + e.message);
    }
  });

  Object.values(_PROP).forEach(function(key) {
    props.deleteProperty(key);
  });

  const msg = deleted + ' test file(s) moved to trash.\nRemember to reset CONFIG in Code.gs manually.';
  Logger.log(msg);
  _setupAlert('Cleanup Complete', msg);
}

/**
 * Lists all chart IDs and their source spreadsheets for the old SS, new SS,
 * and test presentation. Useful for comparing IDs before and after replacement.
 */
function listChartIds() {
  const props  = PropertiesService.getScriptProperties();
  const oldId  = props.getProperty(_PROP.OLD_SS_ID);
  const newId  = props.getProperty(_PROP.NEW_SS_ID);
  const presId = props.getProperty(_PROP.PRES_ID);

  if (!oldId || !newId) {
    _setupAlert('Error', 'Please run createTestEnvironment() first.');
    return;
  }

  const lines = ['=== Chart ID List ===', ''];

  // Old spreadsheet
  lines.push('[Old Spreadsheet: ' + oldId + ']');
  SpreadsheetApp.openById(oldId).getSheets().forEach(function(sheet) {
    sheet.getCharts().forEach(function(c) {
      lines.push('  Sheet "' + sheet.getName() + '": Chart ID = ' + c.getChartId());
    });
  });

  lines.push('');

  // New spreadsheet
  lines.push('[New Spreadsheet: ' + newId + ']');
  SpreadsheetApp.openById(newId).getSheets().forEach(function(sheet) {
    sheet.getCharts().forEach(function(c) {
      lines.push('  Sheet "' + sheet.getName() + '": Chart ID = ' + c.getChartId());
    });
  });

  if (presId) {
    lines.push('');
    lines.push('[Presentation: ' + presId + ']');
    SlidesApp.openById(presId).getSlides().forEach(function(slide, i) {
      slide.getPageElements().forEach(function(el) {
        if (el.getPageElementType() !== SlidesApp.PageElementType.SHEETS_CHART) return;
        const c = el.asSheetsChart();
        lines.push(
          '  Slide ' + (i + 1) + ': Chart ID = ' + c.getChartId() +
          '  Source SS: ' + c.getSpreadsheetId().substring(0, 10) + '...'
        );
      });
    });
  }

  const msg = lines.join('\n');
  Logger.log(msg);
  _setupAlert('Chart ID List', msg);
}

// =============================================================================
// Internal Helpers — Spreadsheet Creation
// =============================================================================

/**
 * Creates the old spreadsheet with dummy data and charts across 3 sheets.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function _createOldSpreadsheet() {
  const ss = SpreadsheetApp.create('[Test] Old Spreadsheet (source)');

  // ---- Sheet 1: Monthly Sales (column chart) ----
  const s1 = ss.getActiveSheet();
  s1.setName('Monthly Sales');
  s1.getRange('A1:D7').setValues([
    ['Month', 'Product A', 'Product B', 'Product C'],
    ['Jan',    120,          95,           60],
    ['Feb',    135,         110,           75],
    ['Mar',    150,         125,           90],
    ['Apr',    165,         140,          105],
    ['May',    180,         155,          120],
    ['Jun',    195,         170,          135],
  ]);
  s1.getRange('A1:D1').setFontWeight('bold').setBackground('#e8f0fe');

  s1.insertChart(
    s1.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(s1.getRange('A1:D7'))
      .setPosition(9, 1, 0, 0)
      .setOption('title', 'Monthly Sales [Old Data]')
      .setOption('legend', { position: 'right' })
      .build()
  );

  // ---- Sheet 2: By Category (pie chart) ----
  const s2 = ss.insertSheet('By Category');
  s2.getRange('A1:B5').setValues([
    ['Category',    'Revenue'],
    ['Food',         450000],
    ['Apparel',      320000],
    ['Electronics',  580000],
    ['Other',        150000],
  ]);
  s2.getRange('A1:B1').setFontWeight('bold').setBackground('#e8f0fe');

  s2.insertChart(
    s2.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(s2.getRange('A1:B5'))
      .setPosition(7, 1, 0, 0)
      .setOption('title', 'Sales by Category [Old Data]')
      .build()
  );

  // ---- Sheet 3: Monthly Achievement Rate (line chart) ----
  const s3 = ss.insertSheet('Achievement Rate');
  s3.getRange('A1:B7').setValues([
    ['Month', 'Achievement Rate (%)'],
    ['Jan',    92],
    ['Feb',    88],
    ['Mar',   105],
    ['Apr',    97],
    ['May',   110],
    ['Jun',   103],
  ]);
  s3.getRange('A1:B1').setFontWeight('bold').setBackground('#e8f0fe');

  s3.insertChart(
    s3.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(s3.getRange('A1:B7'))
      .setPosition(9, 1, 0, 0)
      .setOption('title', 'Monthly Achievement Rate [Old Data]')
      .setOption('series', { 0: { color: '#e67c73' } })
      .build()
  );

  return ss;
}

/**
 * Copies the old spreadsheet to create the new one, then updates the data and
 * chart titles. Chart IDs are preserved by makeCopy().
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} oldSS
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function _createNewSpreadsheet(oldSS) {
  const copy  = DriveApp.getFileById(oldSS.getId()).makeCopy('[Test] New Spreadsheet (target)');
  const newSS = SpreadsheetApp.openById(copy.getId());

  // Sheet 1: Monthly Sales — update values and title
  const s1 = newSS.getSheetByName('Monthly Sales');
  if (s1) {
    s1.getRange('B2:D7').setValues([
      [130, 105,  68],
      [148, 122,  83],
      [162, 139,  98],
      [178, 153, 113],
      [195, 167, 128],
      [210, 182, 143],
    ]);
    _updateChartTitle(s1, 'Monthly Sales [New Data]');
  }

  // Sheet 2: By Category — update values and title
  const s2 = newSS.getSheetByName('By Category');
  if (s2) {
    s2.getRange('B2:B5').setValues([[490000], [355000], [630000], [175000]]);
    _updateChartTitle(s2, 'Sales by Category [New Data]');
  }

  // Sheet 3: Achievement Rate — update values and title
  const s3 = newSS.getSheetByName('Achievement Rate');
  if (s3) {
    s3.getRange('B2:B7').setValues([[95], [91], [108], [100], [114], [107]]);
    _updateChartTitle(s3, 'Monthly Achievement Rate [New Data]');
  }

  return newSS;
}

/**
 * Updates the title of the first chart on the given sheet.
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
// Internal Helpers — Presentation Creation
// =============================================================================

/**
 * Creates a test presentation with all charts from the old spreadsheet
 * embedded as linked Sheets charts.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} oldSS
 * @returns {GoogleAppsScript.Slides.Presentation}
 */
function _createTestPresentation(oldSS) {
  const pres = SlidesApp.create('[Test] Presentation (chart source replacement test)');
  const W = pres.getPageWidth();
  const H = pres.getPageHeight();

  const PADDING   = 40;
  const LABEL_H   = 36;
  const LABEL_TOP = 8;

  // Slide 1: Monthly Sales (column chart)
  const slide1 = pres.getSlides()[0];
  _clearSlide(slide1);
  _insertChartFromSheet(slide1, oldSS, 'Monthly Sales', PADDING, LABEL_TOP + LABEL_H, W - PADDING * 2, H - LABEL_H - PADDING - LABEL_TOP);
  _addSlideLabel(slide1, 'Monthly Sales', W, LABEL_TOP, W - PADDING * 2, LABEL_H);

  // Slide 2: By Category (pie chart)
  const slide2 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const pieW = W * 0.6;
  _insertChartFromSheet(slide2, oldSS, 'By Category', (W - pieW) / 2, LABEL_TOP + LABEL_H, pieW, H - LABEL_H - PADDING - LABEL_TOP);
  _addSlideLabel(slide2, 'Sales by Category', W, LABEL_TOP, W - PADDING * 2, LABEL_H);

  // Slide 3: Monthly Achievement Rate (line chart)
  const slide3 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  _insertChartFromSheet(slide3, oldSS, 'Achievement Rate', PADDING, LABEL_TOP + LABEL_H, W - PADDING * 2, H - LABEL_H - PADDING - LABEL_TOP);
  _addSlideLabel(slide3, 'Monthly Achievement Rate', W, LABEL_TOP, W - PADDING * 2, LABEL_H);

  return pres;
}

/**
 * Inserts the first chart from the named sheet into the slide as a linked Sheets chart.
 */
function _insertChartFromSheet(slide, ss, sheetName, left, top, width, height) {
  const sheet  = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const charts = sheet.getCharts();
  if (charts.length === 0) return;
  slide.insertSheetsChart(charts[0], left, top, width, height);
}

/**
 * Removes all existing elements from the slide, including default placeholders.
 */
function _clearSlide(slide) {
  slide.getPageElements().forEach(function(el) {
    try { el.remove(); } catch (_) {}
  });
}

/**
 * Adds a centered label text box at the top of the slide.
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
// Utilities
// =============================================================================

/**
 * Shows an alert dialog. Falls back to log-only output when running
 * directly from the script editor, where no UI context is available.
 */
function _setupAlert(title, message) {
  Logger.log('[' + title + ']\n' + message);
  try {
    var ui = SlidesApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (_) {}
}
