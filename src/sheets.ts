// for license and source, visit https://github.com/3096/primorina

var dashboardEditRange = [
  "I5", // Status cell
  "AB28", // Document Version
  "AT40", // Current Document Version
  "T29", // Document Status
  "AV1", // Name of drop down 1 (import)
  "AV2", // Name of drop down 2 (auto import)
  "AG14", // Selection
  "AG18" // URL
];

// Cells that needs refreshing
var dashboardRefreshRange = [
];

function importButtonScript() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
  if (dashboardSheet && settingsSheet) {
    var userImportSelection = dashboardSheet.getRange(dashboardEditRange[4]).getValue();
    var importSelectionText = dashboardSheet.getRange(dashboardEditRange[6]).getValue();
    var urlInput = dashboardSheet.getRange(dashboardEditRange[7]).getValue();
    dashboardSheet.getRange(dashboardEditRange[7]).setValue(""); //Clear input
    if (userImportSelection == importSelectionText) {
      settingsSheet.getRange("D6").setValue(urlInput);
      importDataManagement();
    } else {
      settingsSheet.getRange("D17").setValue(urlInput);
      importFromAPI();
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Unable to find 'Dashboard' or 'Settings'", "Missing Sheets");
  }
}

function displayReadme() {
  var sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
  if (sheetSource) {
    // Avoid Exception: You can't remove all the sheets in a document.Details
    var placeHolderSheet = null;
    if (SpreadsheetApp.getActive().getSheets().length == 1) {
      placeHolderSheet = SpreadsheetApp.getActive().insertSheet();
    }
    var sheetToRemove = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_README);
    if (sheetToRemove) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
    }
    var sheetREADMESource;

    // Add Language
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      sheetREADMESource = sheetSource.getSheetByName(SHEET_NAME_README + "-" + languageFound);
    }
    if (sheetREADMESource) {
      // Found language
    } else {
      // Default
      sheetREADMESource = sheetSource.getSheetByName(SHEET_NAME_README);
    }

    sheetREADMESource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(SHEET_NAME_README);

    // Remove placeholder if available
    if (placeHolderSheet) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(placeHolderSheet);
    }
    var sheetREADME = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_README);
    // Refresh Contents Links
    var contentsAvailable = sheetREADME.getRange(13, 1).getValue();
    var contentsStartIndex = 15;

    for (var i = 0; i < contentsAvailable; i++) {
      var valueRange = sheetREADME.getRange(contentsStartIndex + i, 1).getValue();
      var formulaRange = sheetREADME.getRange(contentsStartIndex + i, 1).getFormula();
      // Display for user, doesn't do anything
      sheetREADME.getRange(contentsStartIndex + i, 1).setFormula(formulaRange);

      // Grab URL RichTextValue from Source
      const range = sheetREADMESource.getRange(contentsStartIndex + i, 1);
      const RichTextValue = range.getRichTextValue().getRuns();
      const res = RichTextValue.reduce((ar, e) => {
        const url = e.getLinkUrl();
        if (url) ar.push(url);
        return ar;
      }, []);
      //  Convert to string
      var resString = res + "";
      var arrayString = resString.split("=");
      if (arrayString.length > 1) {
        var text = arrayString[2];
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(valueRange)
          .setLinkUrl("#gid=" + sheetREADME.getSheetId() + 'range=' + text)
          .build();
        sheetREADME.getRange(contentsStartIndex + i, 1).setRichTextValue(richText);
      }
    }
    reorderSheets();
    SpreadsheetApp.getActive().setActiveSheet(sheetREADME);
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function displayAbout() {
  var titleString = "About";
  var htmlString = "Script version: " + SCRIPT_VERSION;
  var widthSize = 500;
  var heightSize = 400;

  var htmlOutput = HtmlService
    .createHtmlOutput(htmlString)
    .setWidth(widthSize) //optional
    .setHeight(heightSize); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, titleString);
}

function moveToSheetByName(nameOfSheet: string) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(nameOfSheet);
  if (sheet) {
    sheet.activate();
  } else {
    const title = "Error";
    const message = "Unable to find sheet named '" + nameOfSheet + "'.";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

const moveToDashboardSheet = () => moveToSheetByName(SHEET_NAME_DASHBOARD);
const moveToSettingsSheet = () => moveToSheetByName(SHEET_NAME_SETTINGS);
const moveToChangelogSheet = () => moveToSheetByName(SHEET_NAME_CHANGELOG);
const moveToReadmeSheet = () => moveToSheetByName(SHEET_NAME_README);
const moveToPrimogemLogSheet = () => moveToSheetByName(SHEET_NAME_PRIMOGEM_LOG);
const moveToPrimogemYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_PRIMOGEM_YEARLY_REPORT);
const moveToPrimogemMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_PRIMOGEM_MONTHLY_REPORT);
const moveToCrystalLogSheet = () => moveToSheetByName(SHEET_NAME_CRYSTAL_LOG);
const moveToCrystalYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_CRYSTAL_YEARLY_REPORT);
const moveToCrystalMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_CRYSTAL_MONTHLY_REPORT);
const moveToResinLogSheet = () => moveToSheetByName(SHEET_NAME_RESIN_LOG);
const moveToResinYearlyReportSheet = () => moveToSheetByName(SHEET_NAME_RESIN_YEARLY_REPORT);
const moveToResinMonthlyReportSheet = () => moveToSheetByName(SHEET_NAME_RESIN_MONTHLY_REPORT);