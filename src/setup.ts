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
      /* Check migration for settings required */
      var bannerSettings = LOG_RANGES[SHEET_NAME_ARTIFACT_LOG];
      var isToggled = settingsSheet.getRange(bannerSettings['range_toggle']).getValue();
      if(isToggled == ""){
        // Migration step required, missing Artifact toggle
        settingsSheet.getRange(13,4,1,2).setBorder(false, false, false, false, false, false).breakApart();
        settingsSheet.getRange("D13").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_MORA_LOG);
        settingsSheet.getRange("E13").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setValue(settingsSheet.getRange("E12").getValue());
        settingsSheet.getRange("D12").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_ARTIFACT_LOG);
        settingsSheet.getRange("E12").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setValue('NOT DONE');
        settingsSheet.getRange(14,4,2,2).breakApart();

        settingsSheet.getRange(14,4,1,2).setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).mergeAcross();
        settingsSheet.getRange("D14").setFontColor("black").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setValue('Auto Import');
        settingsSheet.getRange(15,4,1,2).mergeAcross();
        settingsSheet.getRange("D15").setFontColor("red").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue('Using Genshin Impact Feedback URL to call API with AUTH_KEY');

        settingsSheet.getRange("D22").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_ARTIFACT_LOG);
        settingsSheet.getRange("E22").setFontSize(10).setBackground("white").insertCheckboxes().setValue(true);
        settingsSheet.getRange("D29").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(SHEET_NAME_ARTIFACT_LOG);

        // set data validation
        var listOfSheets = [
          SHEET_NAME_DASHBOARD,
          SHEET_NAME_README,
          SHEET_NAME_CHANGELOG,
          SHEET_NAME_SETTINGS,
          SHEET_NAME_PRIMOGEM_MONTHLY_REPORT,
          SHEET_NAME_PRIMOGEM_YEARLY_REPORT,
          SHEET_NAME_PRIMOGEM_LOG,
          SHEET_NAME_CRYSTAL_MONTHLY_REPORT,
          SHEET_NAME_CRYSTAL_YEARLY_REPORT,
          SHEET_NAME_CRYSTAL_LOG,
          SHEET_NAME_RESIN_MONTHLY_REPORT,
          SHEET_NAME_RESIN_YEARLY_REPORT,
          SHEET_NAME_RESIN_LOG,
          SHEET_NAME_MORA_MONTHLY_REPORT,
          SHEET_NAME_MORA_YEARLY_REPORT,
          SHEET_NAME_MORA_LOG,
          SHEET_NAME_ARTIFACT_MONTHLY_REPORT,
          SHEET_NAME_ARTIFACT_YEARLY_REPORT,
          SHEET_NAME_ARTIFACT_LOG
        ];
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(listOfSheets, true).build();
        settingsSheet.getRange(11,2,20,1).setBackground("white").setFontSize(10).setFontWeight(null).setHorizontalAlignment("center").setDataValidation(rule);;
        settingsSheet.getRange("A28").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(18);
        settingsSheet.getRange("A29").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(19);
        settingsSheet.getRange("A30").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setValue(20);
        
        settingsSheet.getRange("B27").setValue(SHEET_NAME_ARTIFACT_MONTHLY_REPORT);
        settingsSheet.getRange("B28").setValue(SHEET_NAME_ARTIFACT_YEARLY_REPORT);
        settingsSheet.getRange("B29").setValue(SHEET_NAME_ARTIFACT_LOG);
        // Load Artifact Log Sheet if missing
        var artifactLogSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ARTIFACT_LOG);
        if (!artifactLogSheet) {
          if (!sheetSource) {
            sheetSource = SpreadsheetApp.openById(SHEET_SOURCE_ID);
          }
          var sheetArtifactLogSource = sheetSource.getSheetByName(SHEET_NAME_ARTIFACT_LOG);
          artifactLogSheet = sheetArtifactLogSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
          artifactLogSheet.setName(SHEET_NAME_ARTIFACT_LOG);
        }
        // Remove old Dashboard if exist
        var removeDashboardSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_DASHBOARD);
        if (removeDashboardSheet) {
          SpreadsheetApp.getActiveSpreadsheet().deleteSheet(removeDashboardSheet);
        }
      }
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
        updateDashboard(dashboardSheet);
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

function updateDashboard(dashboardSheet) {
  // Go through the available logs sheet list
  const availableSheets = NAME_OF_LOG_HISTORIES.concat(NAME_OF_LOG_HISTORIES_HOYOLAB);
  for (var i = 0; i < availableSheets.length; i++) {
    var logSheet = SpreadsheetApp.getActive().getSheetByName(availableSheets[i]);

    if (logSheet) {
      var iLastRow = logSheet.getRange(2, 2, logSheet.getLastRow(), 1).getValues().filter(String).length;
      dashboardSheet.getRange(LOG_RANGES[availableSheets[i]]['range_dashboard_length']).setValue(iLastRow);
    }
  }
}