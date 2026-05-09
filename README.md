# Google Slides — Bulk Chart Source Replacer

A Google Apps Script that batch-replaces the source spreadsheet of all linked
Sheets charts in a Google Slides presentation, redirecting them from an old
spreadsheet ID to a new one.

---

## Demo

https://github.com/knmgn/google_slides_automation_demo/raw/refs/heads/main/demo.mp4

---

## Use Cases

- You copied a spreadsheet to create a new environment (production / staging /
  new fiscal year) and want your existing presentation to reference the new copy
  instead of the original.
- The source spreadsheet ID changed due to an org migration or permission
  restructuring.

---

## File Structure

```
.
├── Code.gs   # Main script (production use)
├── Setup.gs  # Test data generation script
└── README.md # This file
```

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| Account permissions | The account running the script needs **edit access** to the presentation and both spreadsheets |
| Chart type | Only **linked** Sheets charts are processed. Charts that have been converted to static images are ignored |
| Chart IDs | The old and new spreadsheets must contain charts with **matching chart IDs** (this is automatically satisfied when the new spreadsheet is created via copy) |

---

## Setup

### 1. Open the Script Editor

1. Open the target Google Slides presentation.
2. Click **Extensions → Apps Script**.
3. The script editor opens.

### 2. Paste the Script

1. Select `Code.gs` (or any `.gs` file) in the editor's file list.
2. Delete the existing code and paste the contents of `Code.gs`.
3. Save with **Ctrl+S**.

### 3. Configure CONFIG

Edit the `CONFIG` object at the top of `Code.gs`:

```javascript
const CONFIG = {
  OLD_SPREADSHEET_ID: 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx', // currently referenced spreadsheet
  NEW_SPREADSHEET_ID: 'yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy', // new spreadsheet to point to
  PRESENTATION_ID: '', // leave empty to use the presentation the script is bound to
};
```

#### Finding a Spreadsheet ID

The ID appears in the spreadsheet URL:

```
https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit
```

Copy the string between `/d/` and `/edit`.

#### PRESENTATION_ID

- **Leave empty (default)**: the script targets the presentation it is bound to.
- **Set an ID**: use this when targeting a presentation other than the one the
  script is bound to. Find the ID in the presentation URL:
  `https://docs.google.com/presentation/d/<ID>/edit`

---

## Running the Script

### Option A: Custom Menu (recommended)

After reloading the presentation, a **Replace Charts** menu appears in the
menu bar.

| Menu item | What it does |
|-----------|-------------|
| Dry run (preview targets) | Lists charts that would be replaced, without making any changes |
| Replace chart sources | Runs the actual replacement |

> **Tip**: Always run a dry run first to verify the target charts before
> committing to the replacement.

### Option B: Script Editor

1. Select the function you want to run from the **dropdown at the top** of the
   editor, then click the **Run** button (▶).

   | What you want to do | Function to select |
   |--------------------|--------------------|
   | Preview targets | `dryRun` |
   | Run replacement | `replaceChartSources` |

   > **Note**: `onOpen` is a trigger function that runs automatically when the
   > presentation is opened — do not select it manually. Running it from the
   > editor produces a `Cannot call SlidesApp.getUi() from this context` error,
   > which is expected behavior.

2. On the first run, an authorization dialog will appear — click **Allow**.

---

## Reading the Results

After execution, full details are written to the **execution log**
(View → Logs in the script editor).

```
[Replacement Report]
✅ Replaced: 5
⏭️ Skipped (different source): 2
⚠️ Charts not found in the new spreadsheet:
  • Slide 3, Chart ID 987654321: chart ID 987654321 not found in the new spreadsheet
```

---

## Testing with Setup.gs

`Setup.gs` lets you generate a complete set of dummy files and safely verify
the script's behavior before using it on real data.

### Adding Setup.gs to the Project

1. In the script editor, click **Add a file → Script** and name it `Setup`.
2. Paste the contents of `Setup.gs` and save.

### Step-by-Step Test Workflow

#### Step 1: Generate the test environment

Run `createTestEnvironment()`. Three files are created in Google Drive:

| File name | Contents |
|-----------|---------|
| `[Test] Old Spreadsheet (source)` | 3 sheets (Monthly Sales, By Category, Achievement Rate), one chart each |
| `[Test] New Spreadsheet (target)` | Copy of the old one (chart IDs identical) with updated values and titles |
| `[Test] Presentation (chart source replacement test)` | 3 slides, each with a chart linked to the old spreadsheet |

The execution log prints the IDs — paste them into `CONFIG` in `Code.gs`:

```
OLD_SPREADSHEET_ID: 'xxxxxxxxxxxxxxxxxxxx...'
NEW_SPREADSHEET_ID: 'yyyyyyyyyyyyyyyyyyyy...'
PRESENTATION_ID:    'zzzzzzzzzzzzzzzzzzzz...'
```

#### Step 2: Inspect chart IDs (optional)

Run `listChartIds()` to display all chart IDs in both spreadsheets and the
presentation. If the old and new spreadsheets show the same IDs, you are ready
to proceed.

#### Step 3: Dry run

Run `dryRun()` to confirm which charts will be replaced.

#### Step 4: Replace

Run `replaceChartSources()`.

#### Step 5: Verify

Run `verifyReplacement()` to check that every chart in the presentation now
references the new spreadsheet.

```
Verification results (3 chart(s)):

  Slide 1 / Chart ID 123456789: ✅ New SS (replaced)
  Slide 2 / Chart ID 234567890: ✅ New SS (replaced)
  Slide 3 / Chart ID 345678901: ✅ New SS (replaced)
```

#### Step 6: Clean up

Run `cleanupTestFiles()` to move the 3 test files to trash, then manually
reset `CONFIG` in `Code.gs`.

### Setup.gs Function Reference

| Function | Description |
|----------|-------------|
| `createTestEnvironment()` | Generates old SS, new SS, and test presentation; logs all IDs |
| `verifyReplacement()` | Inspects the presentation and reports each chart's current source |
| `cleanupTestFiles()` | Moves generated test files to trash |
| `listChartIds()` | Lists all chart IDs across both spreadsheets and the presentation (debug) |

---

## Important Notes

### Chart ID Matching

The replacement relies on the old and new spreadsheets having **identical chart
IDs**.

- Spreadsheet **created via copy**: chart IDs are preserved ✅
- Charts **manually recreated** in the new spreadsheet: IDs will differ and
  those charts will not be found ⚠️

### Charts Inside Groups

Charts that are direct children of a group are included in the replacement.
Charts nested more than one level deep inside groups are not processed — ungroup
them first if needed.

### After Replacement

- Replaced charts are inserted as **linked** Sheets charts pointing to the new
  spreadsheet.
- Position and size are preserved exactly.
- Chart appearance (title, legend, etc.) reflects the chart settings in the
  new spreadsheet.

### Undo

Changes made by GAS cannot be undone with **Ctrl+Z** in Slides. Make a backup
copy of the presentation before running the replacement.

---

## Troubleshooting

| Symptom | What to check |
|---------|---------------|
| "Could not open the presentation" | Verify `PRESENTATION_ID` is correct, or leave it empty |
| "Could not open the spreadsheet" | Verify the ID is correct and the account has access |
| Replaced count is 0 | Confirm `OLD_SPREADSHEET_ID` matches the actual source (use dry run to inspect) |
| Chart not found in new spreadsheet | Confirm the new spreadsheet has a chart with the same chart ID |
| Permission error | Ensure the account running the script has edit access to both spreadsheets and the presentation |

---

## Function Reference

### Code.gs

| Function | Type | Description |
|----------|------|-------------|
| `onOpen()` | trigger | Adds the custom menu when the presentation is opened |
| `replaceChartSources()` | public | Main function — bulk-replaces chart sources |
| `dryRun()` | public | Lists replacement targets without making any changes |
| `_processSlide()` | internal | Processes all charts on a single slide |
| `_collectSheetsCharts()` | internal | Collects SheetsChart elements from a slide |
| `_replaceChart()` | internal | Handles replacement for a single chart element |
| `_findChartById()` | internal | Searches a spreadsheet for a chart by ID |
| `_validateConfig()` | internal | Validates CONFIG values |
| `_openPresentation()` | internal | Opens the presentation |
| `_openSpreadsheet()` | internal | Opens a spreadsheet |
| `_showReport()` | internal | Displays the replacement summary |
| `_alert()` | internal | Shows an alert dialog |

### Setup.gs

| Function | Type | Description |
|----------|------|-------------|
| `createTestEnvironment()` | public | Generates old SS, new SS, and test presentation |
| `verifyReplacement()` | public | Checks each chart's source after replacement |
| `cleanupTestFiles()` | public | Moves test files to trash |
| `listChartIds()` | public | Lists chart IDs across both spreadsheets and the presentation |
| `_createOldSpreadsheet()` | internal | Creates the old spreadsheet with dummy data |
| `_createNewSpreadsheet()` | internal | Copies the old SS and updates its data |
| `_createTestPresentation()` | internal | Creates the test presentation |

---

## License

MIT License
