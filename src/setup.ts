// for license and source, visit https://github.com/3096/primorina
function onInstall(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    generateInitialiseToolbar();
  } else {
    onOpen(e);
  }
}

function onOpen(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    generateInitialiseToolbar();
  } else {
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
    if (!settingsSheet) {
      generateInitialiseToolbar();
    } else {
      getDefaultMenu();
    }
    checkLocaleIsSetCorrectly();
  }
}

function generateInitialiseToolbar() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TCBS')
  .addItem('Initialise', 'updateItemsList')
  .addToUi();
}

function displayUserPrompt(titlePrompt: string, messagePrompt: string) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    titlePrompt,
    messagePrompt,
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  return result;
}

function displayUserAlert(titleAlert: string, messageAlert: string) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    titleAlert,
    messageAlert,
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  return result;
}

/* Ensure Sheets is set to the supported locale due to source document formula */
function checkLocaleIsSetCorrectly() {
  var currentLocale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  if (currentLocale != SHEET_SOURCE_SUPPORTED_LOCALE) {
    SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetLocale(SHEET_SOURCE_SUPPORTED_LOCALE);
    var message = 'To ensure compatibility with formula from source document, your locale "'+currentLocale+'" has been changed to "'+SHEET_SOURCE_SUPPORTED_LOCALE+'"';
    var title = 'Sheets Locale Changed';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function getDefaultMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TCBS')
  .addSeparator()
  .addSubMenu(ui.createMenu('Data Management')
          .addItem('Import', 'importDataManagement')
          .addSeparator()
          .addItem('Auto Import from miHoYo', 'importFromAPI')
          .addItem('Auto Import from HoYoLAB', 'importFromHoYoLAB')
          )
  .addSeparator()
  .addItem('Quick Update', 'quickUpdate')
  .addItem('Update Items', 'updateItemsList')
  .addItem('Get Latest README', 'displayReadme')
  .addToUi();
}

function getSettingsSheet() {
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SETTINGS);
    var sheetSource;
    if (!settingsSheet) {
      sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
      var sheetSettingSource = sheetSource.getSheetByName(SHEET_NAME_SETTINGS);
      if (sheetSettingSource) {
        settingsSheet = sheetSettingSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
        settingsSheet.setName(SHEET_NAME_SETTINGS);
        getDefaultMenu();
      }
    } else {
      settingsSheet.getRange("H1").setValue(SCRIPT_VERSION);
    }
    var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
    if (!dashboardSheet) {
      if (!sheetSource) {
        sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
      }
      var sheetDashboardSource = sheetSource.getSheetByName(SHEET_NAME_DASHBOARD);
      if (sheetDashboardSource) {
        dashboardSheet = sheetDashboardSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
        dashboardSheet.setName(SHEET_NAME_DASHBOARD);
      }
    } else {
      if (SHEET_SCRIPT_IS_ADD_ON) {
        dashboardSheet.getRange(dashboardEditRange[10]).setFontColor("green").setFontWeight("bold").setHorizontalAlignment("left").setValue("Add-On Enabled");
      } else {
        dashboardSheet.getRange(dashboardEditRange[10]).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("left").setValue("Embedded Script");
      }
    }
    return settingsSheet;
}